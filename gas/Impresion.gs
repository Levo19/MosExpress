// ============================================================
// MosExpress — Impresion.gs
// Proxy de impresión PrintNode y generación ESC/POS interna.
// Script Property requerida: PRINTNODE_API_KEY
// ============================================================

// [v2.5.33] Estado bulk de impresoras — usado por el wizard moderno para mostrar
// el icono 🟢/🔴/🟡 en cada card de estación EN VIVO. Una sola llamada GET
// /printers/{ids} consulta varias IDs separadas por coma. Si PrintNode falla,
// devolvemos cada impresora en estado "checking" para que la UI sea graceful.
function estadoImpresoras(idsCsv) {
  var printNodeKey = PropertiesService.getScriptProperties().getProperty('PRINTNODE_API_KEY');
  if (!printNodeKey) return generarRespuestaError('PRINTNODE_API_KEY no configurada');
  var ids = String(idsCsv || '').split(',').map(function(x){ return String(x).trim(); }).filter(Boolean);
  if (!ids.length) return ContentService.createTextOutput(JSON.stringify({ status: 'success', impresoras: [] })).setMimeType(ContentService.MimeType.JSON);
  try {
    var resp = UrlFetchApp.fetch('https://api.printnode.com/printers/' + ids.join(','), {
      method:  'get',
      headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(printNodeKey + ':') },
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    var lista = [];
    if (code === 200) {
      var arr = JSON.parse(resp.getContentText());
      arr.forEach(function(p){
        var pState    = String(p.state || 'unknown').toLowerCase();
        var compState = String((p.computer && p.computer.state) || 'unknown').toLowerCase();
        var online    = (pState === 'online') && (compState === 'connected');
        lista.push({
          id:       p.id,
          name:     String(p.name || ('#' + p.id)),
          online:   online,
          state:    pState,
          computerState: compState,
          computer: String((p.computer && p.computer.name) || ''),
          reason:   online ? 'Lista para imprimir'
                  : (pState !== 'online' ? 'Impresora ' + pState
                  : 'Agente PrintNode ' + compState)
        });
      });
    }
    // Cualquier ID que no haya devuelto PrintNode → marcar offline
    var devueltos = {};
    lista.forEach(function(p){ devueltos[String(p.id)] = true; });
    ids.forEach(function(id){
      if (!devueltos[id]) lista.push({ id: id, name: '#' + id, online: false, state: 'not_found', reason: 'No registrada en PrintNode' });
    });
    return ContentService.createTextOutput(JSON.stringify({ status: 'success', impresoras: lista })).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return generarRespuestaError('Error consultando PrintNode (bulk): ' + err.toString());
  }
}

// [v2.5.16] Verifica estado de impresora en PrintNode antes de enviar print job.
// Si la impresora local está offline (agente PrintNode caído, USB desconectada,
// computador apagado), PrintNode acepta el job pero NUNCA imprime → el cajero
// no se entera. Esta función consulta /printers/{id} para validar state ANTES.
//
// Devuelve { status, state, computer, nombre, online: bool }
function verificarEstadoImpresora(data) {
  var printNodeKey = PropertiesService.getScriptProperties().getProperty('PRINTNODE_API_KEY');
  if (!printNodeKey) return generarRespuestaError("PRINTNODE_API_KEY no configurada");
  if (!data.printerId) return generarRespuestaError("printerId requerido");
  var printerId = parseInt(data.printerId, 10);
  if (isNaN(printerId) || printerId <= 0) return generarRespuestaError("printerId inválido: " + data.printerId);
  try {
    var resp = UrlFetchApp.fetch('https://api.printnode.com/printers/' + printerId, {
      method:  'get',
      headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(printNodeKey + ':') },
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    if (code !== 200) {
      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        online: false,
        state:  'http_error_' + code,
        nombre: '',
        computer: '',
        mensaje: 'PrintNode HTTP ' + code + ': ' + resp.getContentText().substring(0, 120)
      })).setMimeType(ContentService.MimeType.JSON);
    }
    var arr = JSON.parse(resp.getContentText());
    var p = (arr && arr[0]) || null;
    if (!p) {
      return ContentService.createTextOutput(JSON.stringify({
        status: 'success', online: false, state: 'not_found', nombre: '', computer: '',
        mensaje: 'Impresora ' + printerId + ' no existe en PrintNode'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    // State puede ser 'online' | 'offline' | 'unknown'.
    // computer.state similar — si el agente PrintNode está caído, computer.state='disconnected'
    var pState   = String(p.state || 'unknown').toLowerCase();
    var compState = String((p.computer && p.computer.state) || 'unknown').toLowerCase();
    var online = (pState === 'online') && (compState === 'connected');
    return ContentService.createTextOutput(JSON.stringify({
      status:   'success',
      online:   online,
      state:    pState,
      computerState: compState,
      nombre:   String(p.name || ''),
      computer: String((p.computer && p.computer.name) || ''),
      mensaje:  online ? 'OK' : 'Impresora ' + (pState !== 'online' ? pState : 'agente ' + compState)
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return generarRespuestaError("Error consultando PrintNode: " + err.toString());
  }
}

// Proxy para impresión manual desde el browser (fallback / etiquetas especiales)
function procesarImpresion(data) {
  var printNodeKey = PropertiesService.getScriptProperties().getProperty('PRINTNODE_API_KEY');
  if (!printNodeKey) return generarRespuestaError("PRINTNODE_API_KEY no configurada en Propiedades del script.");
  if (!data.printerId || !data.content) return generarRespuestaError("Faltan datos de impresión (printerId o content).");

  var printerId = parseInt(data.printerId, 10);
  if (isNaN(printerId) || printerId <= 0) {
    return generarRespuestaError("printerId inválido: '" + data.printerId + "'. Verifica el campo PrintNode_ID en la hoja ZONAS_CONFIG.");
  }

  try {
    var resp = UrlFetchApp.fetch('https://api.printnode.com/printjobs', {
      method:       'post',
      headers:      { 'Authorization': 'Basic ' + Utilities.base64Encode(printNodeKey + ':') },
      contentType:  'application/json',
      payload:      JSON.stringify({
        printerId:   printerId,
        title:       data.title || 'MOSexpress',
        contentType: 'raw_base64',
        content:     data.content,
        source:      'MOSexpress'
      }),
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    if (code !== 201) {
      return generarRespuestaError("PrintNode respondió " + code +
        " (printerId=" + printerId + "): " + resp.getContentText());
    }
    return ContentService.createTextOutput(JSON.stringify({
      status: "success", printJobId: resp.getContentText()
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return generarRespuestaError("Error llamando a PrintNode: " + err.toString());
  }
}

// Construye el ticket ESC/POS en GAS y lo envía a PrintNode directamente.
// Elimina el segundo round-trip browser→GAS→PrintNode.
// nfResult: objeto de emitirNubeFact (puede ser null para NOTA_DE_VENTA).
function imprimirTicketInternamente(data, correlativo, printerId, nfResult) {
  var printNodeKey = PropertiesService.getScriptProperties().getProperty('PRINTNODE_API_KEY');
  if (!printNodeKey || !printerId) return false;

  var auth   = data.auth   || {};
  var header = data.header || {};
  var items  = data.items  || [];

  var W    = 48;
  var SEP  = new Array(W + 1).join('=') + '\n';
  var SEPd = new Array(W + 1).join('-') + '\n';

  var tipoLabel = header.tipoDoc === 'NOTA_DE_VENTA' ? 'NOTA DE VENTA' :
                  header.tipoDoc === 'BOLETA'         ? 'BOLETA'         :
                  header.tipoDoc === 'FACTURA'        ? 'FACTURA'        :
                  normalizarTextoGAS(header.tipoDoc || '');

  var txt = '\x1b\x40';
  txt += '\x1b\x61\x01';
  // [v40.3] Sello "PAGADO · COBRO DIFERIDO" arriba si es reimpresión post-cobro
  if (data.esPagoDiferido) {
    txt += '\x1b\x21\x10';
    txt += '*** PAGADO ***\n';
    txt += 'COBRO DIFERIDO\n';
    txt += '\x1b\x21\x00';
    txt += SEPd;
  }
  txt += '\x1b\x21\x30MOSexpress\x1b\x21\x00\n';
  txt += '\x1b\x21\x10' + tipoLabel + '\x1b\x21\x00\n';
  txt += 'Tk: ' + correlativo + '\n';
  txt += SEP;
  txt += '\x1b\x61\x00';
  txt += 'FECHA   : ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss') + '\n';

  var clienteNombre = normalizarTextoGAS((header.cliente && header.cliente.nombre) || '');
  var clienteDoc    = (header.cliente && header.cliente.doc) ? String(header.cliente.doc) : '';
  var clienteTipo   = (header.cliente && header.cliente.tipo) || 0; // 0=sin,1=DNI,4=CE,6=RUC,7=PAS
  if (clienteNombre) txt += 'CLIENTE : ' + clienteNombre.substring(0, 38) + '\n';
  if (clienteDoc) {
    var labelDoc = clienteTipo === 1 ? 'DNI     '
                 : clienteTipo === 4 ? 'C.E.    '
                 : clienteTipo === 6 ? 'RUC     '
                 : clienteTipo === 7 ? 'PASAP.  '
                 : 'DOC     ';
    txt += labelDoc + ': ' + clienteDoc + '\n';
  }
  txt += (auth.esCajero ? 'CAJERO  ' : 'VENDEDOR') + ': ' + normalizarTextoGAS(auth.vendedor || '') + '\n';
  txt += SEP;
  txt += 'CANT  DESCRIPCION                      SUBTOTAL \n';
  txt += SEPd;

  items.forEach(function(item) {
    var nombre   = normalizarTextoGAS(item.nombre || '');
    var m        = nombre.match(/^(.+?)\s+\((.+)\)$/);
    var baseName = m ? m[1] : nombre;
    var empaque  = m ? m[2] : null;
    var desc = baseName.substring(0, 31);
    while (desc.length < 31) desc += ' ';
    var cant = String(item.cantidad || '').substring(0, 4);
    while (cant.length < 5) cant += ' ';
    var sub = parseFloat(item.subtotal || 0).toFixed(2);
    while (sub.length < 10) sub = ' ' + sub;
    txt += cant + ' ' + desc + ' ' + sub + '\n';
    if (empaque) txt += '        ' + empaque.substring(0, 38) + '\n';
  });

  txt += SEPd;
  txt += '\x1b\x61\x02';
  txt += '\x1b\x21\x10TOTAL: S/ ' + parseFloat(header.total || 0).toFixed(2) + '\x1b\x21\x00\n';
  txt += 'METODO: ' + normalizarTextoGAS(header.metodo || 'EFECTIVO') + '\n';
  // [v40.3] Si es reimpresión por pago diferido, agregar detalle del cobro
  if (data.esPagoDiferido && data.pagoDiferido) {
    var pd = data.pagoDiferido;
    txt += '\x1b\x61\x01';
    txt += SEPd;
    txt += '\x1b\x21\x10COBRO RECIBIDO\x1b\x21\x00\n';
    txt += normalizarTextoGAS('Caja: ' + (pd.cajaCobro || '')).substring(0, W) + '\n';
    txt += normalizarTextoGAS('Cajero: ' + (pd.cajeroCobro || '')).substring(0, W) + '\n';
    if (pd.adminAsig) {
      txt += normalizarTextoGAS('Asignado por: ' + pd.adminAsig).substring(0, W) + '\n';
    }
    txt += normalizarTextoGAS('Fecha cobro: ' + (pd.fechaCobro || '')).substring(0, W) + '\n';
    txt += SEPd;
  }
  txt += '\n\x1b\x61\x01*** GRACIAS POR SU COMPRA ***\n';
  var qrData = (nfResult && nfResult.qrString) ? nfResult.qrString : correlativo;
  txt += qrESCPOSGas(qrData);
  if (nfResult && nfResult.hash) {
    txt += '\x1b\x61\x01';
    txt += normalizarTextoGAS('Hash: ' + nfResult.hash).substring(0, W) + '\n';
  }
  if (nfResult && !nfResult.ok && nfResult.error) {
    txt += '\x1b\x61\x01[CPE pendiente de emision]\n';
  }
  txt += '\n\n\n\n\n\x1d\x56\x00\x1b\x6d\x1b\x69\x1b\x42\x05\x02';

  var bytes = [];
  for (var ci = 0; ci < txt.length; ci++) bytes.push(txt.charCodeAt(ci) & 0xFF);
  var content = Utilities.base64Encode(bytes);

  try {
    var resp = UrlFetchApp.fetch('https://api.printnode.com/printjobs', {
      method:       'post',
      headers:      { 'Authorization': 'Basic ' + Utilities.base64Encode(printNodeKey + ':') },
      contentType:  'application/json',
      payload:      JSON.stringify({
        printerId:   parseInt(printerId, 10),
        title:       tipoLabel + ' ' + correlativo,
        contentType: 'raw_base64',
        content:     content,
        source:      'MOSexpress-GAS'
      }),
      muteHttpExceptions: true
    });
    return resp.getResponseCode() === 201;
  } catch (e) {
    Logger.log('imprimirTicketInternamente error: ' + e.toString());
    return false;
  }
}

// ── Helpers ESC/POS ───────────────────────────────────────────
function normalizarTextoGAS(str) {
  return String(str || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^\x20-\x7E]/g, '?');
}

function qrESCPOSGas(text) {
  var len = text.length + 3;
  var pL  = len & 0xFF;
  var pH  = (len >> 8) & 0xFF;
  return '\x1d\x28\x6b\x04\x00\x31\x41\x32\x00' +
         '\x1d\x28\x6b\x03\x00\x31\x43\x05' +
         '\x1d\x28\x6b\x03\x00\x31\x45\x31' +
         '\x1d\x28\x6b' + String.fromCharCode(pL) + String.fromCharCode(pH) +
         '\x31\x50\x30' + text +
         '\x1d\x28\x6b\x03\x00\x31\x51\x30' +
         '\n';
}
