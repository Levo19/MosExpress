// ============================================================
// MosExpress — Creditos.gs
// Sistema de cobro asignado de créditos:
//   - asignarCobroACajero: admin/master asigna un crédito a una caja
//   - confirmarCobroAsignado: cajero confirma el cobro recibido
//   - rechazarCobroAsignado: cajero rechaza (cliente no llegó, etc)
//   - getCobrosAsignadosCajero: lista cobros pendientes de una caja
//   - getCreditosPendientes: lista todos los créditos sin cobrar
//   - escalarCobrosVencidos: trigger 5min detecta > 1h sin resolver
//
// Hoja CREDITOS_COBRO_ASIGNADO (se autocreaa):
//   ID_Cobro | ID_Venta | Caja_Destino | Vendedor_Dest | Metodo_Sug |
//   Estado | Admin_Asignador | Fecha_Asig | Fecha_Res | Razon |
//   ID_Caja_Origen | Monto | Cliente_Nombre | Correlativo
//
// Estados: ASIGNADO | COBRADO | RECHAZADO | TIMEOUT
// ============================================================

// [v2.5.27] Columnas nuevas: Fecha_Vencimiento + Horas_TTL.
// Permite vencimiento configurable (1h default · 2h · 4h · 6h) y la
// columna explícita evita recalcular Fecha_Asig + N horas en cada poll.
var _CREDITO_COBRO_HEADERS = [
  'ID_Cobro','ID_Venta','Caja_Destino','Vendedor_Dest','Metodo_Sug',
  'Estado','Admin_Asignador','Fecha_Asig','Fecha_Res','Razon',
  'ID_Caja_Origen','Monto','Cliente_Nombre','Correlativo',
  'Fecha_Vencimiento','Horas_TTL'
];

function _getHojaCobrosAsignados() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('CREDITOS_COBRO_ASIGNADO');
  if (!sheet) {
    sheet = ss.insertSheet('CREDITOS_COBRO_ASIGNADO');
    sheet.appendRow(_CREDITO_COBRO_HEADERS);
    sheet.setFrozenRows(1);
    return sheet;
  }
  // [v2.5.27] Migrar headers — agregar columnas nuevas si faltan
  var firstRow = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), _CREDITO_COBRO_HEADERS.length)).getValues()[0];
  var current = firstRow.map(function(h){ return String(h || '').trim(); });
  var faltan = _CREDITO_COBRO_HEADERS.filter(function(h) { return current.indexOf(h) === -1; });
  if (faltan.length > 0) {
    var startCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, startCol, 1, faltan.length).setValues([faltan]);
    sheet.getRange(1, startCol, 1, faltan.length)
         .setFontWeight('bold').setBackground('#1f2937').setFontColor('#fff');
  }
  return sheet;
}

// ============================================================
// asignarCobroACajero — admin/master desde MOS asigna un crédito
// ============================================================
// Recibe via POST desde MOS (con clave admin validada).
// Crea row en CREDITOS_COBRO_ASIGNADO + push al cajero destino.
//
// payload:
// { tipoEvento: 'ASIGNAR_COBRO_CAJERO',
//   idVenta, cajaDestino, metodoSugerido ('EFECTIVO'|'VIRTUAL'|'MIXTO'),
//   adminAuth: { nombre, rol, via:'PIN_8DIG' } }
function asignarCobroACajero(data) {
  if (!data.idVenta)         return generarRespuestaError('idVenta requerido');
  if (!data.cajaDestino)     return generarRespuestaError('cajaDestino requerida');
  if (!data.metodoSugerido)  return generarRespuestaError('metodoSugerido requerido');
  if (!data.adminAuth || !data.adminAuth.nombre) {
    return generarRespuestaError('adminAuth requerido (esta acción requiere admin/master)');
  }

  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var ventas = ss.getSheetByName('VENTAS_CABECERA');
  var cajas  = ss.getSheetByName('CAJAS');
  if (!ventas || !cajas) return generarRespuestaError('Hojas requeridas no encontradas');

  // 1. Validar que la venta exista, sea CREDITO o POR_COBRAR, no anulada
  var fv = ventas.getDataRange().getValues();
  var ventaData = null;
  for (var i = fv.length - 1; i > 0; i--) {
    if (String(fv[i][0]) === String(data.idVenta)) {
      ventaData = {
        formaPago:   String(fv[i][8] || ''),
        correlativo: String(fv[i][9] || ''),
        cliente:     String(fv[i][5] || ''),
        total:       parseFloat(fv[i][6]) || 0,
        cajaOriginal:String(fv[i][10] || '')
      };
      break;
    }
  }
  if (!ventaData) return generarRespuestaError('Venta ' + data.idVenta + ' no encontrada');
  var fpUpper = ventaData.formaPago.toUpperCase();
  if (fpUpper !== 'CREDITO' && fpUpper !== 'POR_COBRAR') {
    return generarRespuestaError('Venta no está pendiente (estado: ' + ventaData.formaPago + ')');
  }

  // 2. Validar caja destino ABIERTA y obtener vendedor (cajero)
  var fc = cajas.getDataRange().getValues();
  var cajaInfo = null;
  for (var j = fc.length - 1; j > 0; j--) {
    if (String(fc[j][0]) === String(data.cajaDestino)) {
      cajaInfo = {
        vendedor: String(fc[j][1] || ''),
        estado:   String(fc[j][5] || '')
      };
      break;
    }
  }
  if (!cajaInfo || cajaInfo.estado !== 'ABIERTA') {
    return generarRespuestaError('Caja destino no está abierta');
  }

  // 3. Verificar que no haya OTRA asignación ASIGNADO para esta venta (idempotencia)
  var hoja = _getHojaCobrosAsignados();
  var fa = hoja.getDataRange().getValues();
  for (var k = 1; k < fa.length; k++) {
    if (String(fa[k][1]) === String(data.idVenta) && String(fa[k][5]) === 'ASIGNADO') {
      return generarRespuestaError('Ya hay un cobro ASIGNADO para esta venta (' + String(fa[k][0]) + ')');
    }
  }

  // 4. [v2.5.27] Calcular Fecha_Vencimiento según horasTTL (default 1h)
  var horasTTL = parseInt(data.horasTTL, 10);
  if (![1, 2, 4, 6].includes(horasTTL)) horasTTL = 1; // sanitizar
  var ahora = new Date();
  var fechaVencimiento = new Date(ahora.getTime() + horasTTL * 60 * 60 * 1000);

  // 5. Crear row de asignación
  var idCobro = 'CB-' + ahora.getTime();
  hoja.appendRow([
    idCobro, data.idVenta, data.cajaDestino, cajaInfo.vendedor,
    String(data.metodoSugerido).toUpperCase(),
    'ASIGNADO',
    String(data.adminAuth.nombre || '').replace(/^admin:/i, ''),
    ahora, '', '',
    ventaData.cajaOriginal, ventaData.total, ventaData.cliente, ventaData.correlativo,
    fechaVencimiento, horasTTL
  ]);
  SpreadsheetApp.flush();

  // 5. Push al cajero destino via MOS
  try {
    var url = PropertiesService.getScriptProperties().getProperty('MOS_WEB_APP_URL');
    if (url) {
      var titulo = '💳 Cobro pendiente · ' + (ventaData.cliente || 'cliente');
      var cuerpo = (data.adminAuth.nombre || 'Admin').replace(/^admin:/i, '') +
                   ' te asignó un crédito de S/ ' + ventaData.total.toFixed(2) +
                   ' (' + String(data.metodoSugerido).toUpperCase() + '). Tocá para cobrar.';
      UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          action: 'enviarPushUsuario',
          usuario: cajaInfo.vendedor,
          titulo: titulo,
          cuerpo: cuerpo,
          idNotif: 'CREDITO_COBRO_PENDIENTE',
          // Datos para que el cliente PWA pueda abrir el modal directo
          extra: { idCobro: idCobro, idVenta: data.idVenta, total: ventaData.total }
        }),
        muteHttpExceptions: true
      });
    }
  } catch(ePush) { Logger.log('Push asignar cobro: ' + ePush.message); }

  // 6. Log auditoría
  try {
    auditarLog('CREDITOS_COBRO_ASIGNADO', idCobro, {
      usuario: 'MOS-Admin', rol: 'ADMIN',
      source: 'MOS_ASIGNAR_COBRO',
      accion: 'crear',
      autorizadoPor: { nombre: data.adminAuth.nombre, rol: data.adminAuth.rol || 'ADMIN', via: 'PIN_8DIG' },
      ref: { idVenta: data.idVenta, cajaDestino: data.cajaDestino, vendedor: cajaInfo.vendedor, monto: ventaData.total },
      motivo: ''
    });
  } catch(eLog) { Logger.log('Log asignar: ' + eLog.message); }

  // [v2.5.27] Auto-instalar trigger horario (idempotente)
  try { _ensureTriggerEscalarCobros(); } catch(_){}

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    idCobro: idCobro,
    cajeroDestino: cajaInfo.vendedor,
    horasTTL: horasTTL,
    fechaVencimiento: fechaVencimiento.toISOString(),
    mensaje: 'Cobro asignado a ' + cajaInfo.vendedor + ' · vence en ' + horasTTL + 'h'
  })).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// confirmarCobroAsignado — el cajero recibe el dinero y confirma
// ============================================================
// Internamente llama a cobrarCreditoConExtra (que ya maneja todo:
// MOVIMIENTOS_EXTRA + cambio formaPago + audit). Después marca el
// row de CREDITOS_COBRO_ASIGNADO como COBRADO.
//
// payload:
// { tipoEvento: 'CONFIRMAR_COBRO_ASIGNADO',
//   idCobro, metodoFinal ('EFECTIVO'|'VIRTUAL'|'MIXTO (...)'),
//   montoEfectivo, montoVirtual (si MIXTO),
//   auth: { vendedor, esCajero, deviceId } }
function confirmarCobroAsignado(data) {
  if (!data.idCobro)     return generarRespuestaError('idCobro requerido');
  if (!data.metodoFinal) return generarRespuestaError('metodoFinal requerido');

  var hoja = _getHojaCobrosAsignados();
  var fa = hoja.getDataRange().getValues();
  var cRow = -1, cobroData = null;
  for (var i = 1; i < fa.length; i++) {
    if (String(fa[i][0]) === String(data.idCobro)) {
      cRow = i + 1;
      cobroData = {
        idVenta:    String(fa[i][1]),
        cajaDest:   String(fa[i][2]),
        estado:     String(fa[i][5]),
        adminAsig:  String(fa[i][6]),
        monto:      parseFloat(fa[i][11]) || 0,
        cliente:    String(fa[i][12])
      };
      break;
    }
  }
  if (cRow < 2) return generarRespuestaError('Cobro ' + data.idCobro + ' no encontrado');
  if (cobroData.estado !== 'ASIGNADO') {
    return generarRespuestaError('El cobro no está en estado ASIGNADO (actual: ' + cobroData.estado + ')');
  }

  // Llamar internamente al endpoint existente cobrarCreditoConExtra
  var cobroResp = cobrarCreditoConExtra({
    idVenta:       cobroData.idVenta,
    cajaReceptora: cobroData.cajaDest,
    metodo:        data.metodoFinal,
    montoEfectivo: data.montoEfectivo,
    montoVirtual:  data.montoVirtual,
    obs:           'Cobro asignado ' + data.idCobro,
    auth:          data.auth || {},
    adminAuth:     null   // ya autorizado al asignar, el cajero solo confirma
  });

  // Detectar error del cobrarCreditoConExtra (ContentService no se puede leer directo)
  var cobroResult;
  try {
    cobroResult = JSON.parse(cobroResp.getContent());
  } catch(eP) {
    return generarRespuestaError('Error parsing cobro response');
  }
  if (cobroResult.status !== 'success') {
    return generarRespuestaError('Error procesando cobro: ' + (cobroResult.mensaje || ''));
  }

  // Marcar row CREDITOS_COBRO_ASIGNADO como COBRADO
  hoja.getRange(cRow, 6).setValue('COBRADO');                            // Estado
  hoja.getRange(cRow, 9).setValue(new Date());                            // Fecha_Res
  hoja.getRange(cRow, 5).setValue(String(data.metodoFinal).toUpperCase());// Metodo final
  SpreadsheetApp.flush();

  // Push de vuelta a admin (cierre del ciclo)
  try {
    var url = PropertiesService.getScriptProperties().getProperty('MOS_WEB_APP_URL');
    if (url && cobroData.adminAsig) {
      UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          action: 'enviarPushUsuario',
          usuario: cobroData.adminAsig,
          titulo: '✅ Cobro confirmado · ' + cobroData.cliente,
          cuerpo: 'S/ ' + cobroData.monto.toFixed(2) + ' cobrado en ' + cobroData.cajaDest +
                  ' (' + String(data.metodoFinal).toUpperCase() + ')',
          idNotif: 'CREDITO_COBRO_CONFIRMADO'
        }),
        muteHttpExceptions: true
      });
    }
  } catch(_){}

  // Reimprimir ticket con sello "PAGADO" si la caja tiene printer
  try {
    var pos = _resolverPrinterCaja(cobroData.cajaDest);
    if (pos && pos.printerId) {
      _reimprimirTicketConSello(cobroData.idVenta, pos.printerId, {
        metodoFinal: data.metodoFinal,
        cajaDest:    cobroData.cajaDest,
        cajero:      (data.auth && data.auth.vendedor) || '',
        adminAsig:   cobroData.adminAsig
      });
    }
  } catch(eImp) { Logger.log('Reimpresion sello: ' + eImp.message); }

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    mensaje: 'Cobro confirmado',
    idCobro: data.idCobro,
    idVenta: cobroData.idVenta,
    monto:   cobroData.monto
  })).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// rechazarCobroAsignado — el cajero rechaza con razón
// ============================================================
// payload: { tipoEvento:'RECHAZAR_COBRO_ASIGNADO', idCobro, razon, auth }
function rechazarCobroAsignado(data) {
  if (!data.idCobro) return generarRespuestaError('idCobro requerido');
  if (!data.razon)   return generarRespuestaError('razon es obligatoria');

  var hoja = _getHojaCobrosAsignados();
  var fa = hoja.getDataRange().getValues();
  var cRow = -1, cobroData = null;
  for (var i = 1; i < fa.length; i++) {
    if (String(fa[i][0]) === String(data.idCobro)) {
      cRow = i + 1;
      cobroData = {
        estado:    String(fa[i][5]),
        adminAsig: String(fa[i][6]),
        cliente:   String(fa[i][12]),
        monto:     parseFloat(fa[i][11]) || 0,
        cajaDest:  String(fa[i][2])
      };
      break;
    }
  }
  if (cRow < 2) return generarRespuestaError('Cobro no encontrado');
  if (cobroData.estado !== 'ASIGNADO') {
    return generarRespuestaError('Solo se puede rechazar un cobro ASIGNADO');
  }

  hoja.getRange(cRow, 6).setValue('RECHAZADO');
  hoja.getRange(cRow, 9).setValue(new Date());
  hoja.getRange(cRow, 10).setValue(String(data.razon).substring(0, 250));
  SpreadsheetApp.flush();

  // Push al admin que asignó (para reaccionar)
  try {
    var url = PropertiesService.getScriptProperties().getProperty('MOS_WEB_APP_URL');
    if (url && cobroData.adminAsig) {
      UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          action: 'enviarPushUsuario',
          usuario: cobroData.adminAsig,
          titulo: '⚠ Cobro rechazado · ' + cobroData.cliente,
          cuerpo: 'S/ ' + cobroData.monto.toFixed(2) + ' rechazado: ' + String(data.razon).substring(0, 100),
          idNotif: 'CREDITO_COBRO_RECHAZADO'
        }),
        muteHttpExceptions: true
      });
    }
  } catch(_){}

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success', mensaje: 'Cobro rechazado'
  })).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// getCobrosAsignadosCajero — el cajero pregunta qué le asignaron
// ============================================================
// GET ?accion=cobros_asignados_cajero&cajaId=CAJA-X
// Devuelve los cobros con estado ASIGNADO para esa caja.
function getCobrosAsignadosCajero(cajaId) {
  if (!cajaId) return generarRespuestaError('cajaId requerido');
  var hoja = _getHojaCobrosAsignados();
  var fa = hoja.getDataRange().getValues();
  var result = [];
  // [v2.5.27] Indexar columnas por header para soportar columnas nuevas
  var hdrs = fa[0].map(function(h){ return String(h || '').trim(); });
  var iIdCobro   = hdrs.indexOf('ID_Cobro');
  var iIdVenta   = hdrs.indexOf('ID_Venta');
  var iCajaDest  = hdrs.indexOf('Caja_Destino');
  var iVendDest  = hdrs.indexOf('Vendedor_Dest');
  var iMetodo    = hdrs.indexOf('Metodo_Sug');
  var iEstado    = hdrs.indexOf('Estado');
  var iAdminAsig = hdrs.indexOf('Admin_Asignador');
  var iFAsig     = hdrs.indexOf('Fecha_Asig');
  var iMonto     = hdrs.indexOf('Monto');
  var iCliente   = hdrs.indexOf('Cliente_Nombre');
  var iCorrel    = hdrs.indexOf('Correlativo');
  var iVenc      = hdrs.indexOf('Fecha_Vencimiento');
  var iHorasTTL  = hdrs.indexOf('Horas_TTL');
  for (var i = 1; i < fa.length; i++) {
    if (String(fa[i][iCajaDest]) !== String(cajaId)) continue;
    if (String(fa[i][iEstado]) !== 'ASIGNADO')      continue;
    // Calcular Fecha_Vencimiento — si no está set, fallback a Fecha_Asig + horasTTL (o 1h)
    var fVencISO = '';
    if (iVenc >= 0 && fa[i][iVenc]) {
      fVencISO = fa[i][iVenc] instanceof Date ? fa[i][iVenc].toISOString() : new Date(fa[i][iVenc]).toISOString();
    } else if (iFAsig >= 0 && fa[i][iFAsig]) {
      var fAsigMs = fa[i][iFAsig] instanceof Date ? fa[i][iFAsig].getTime() : new Date(fa[i][iFAsig]).getTime();
      var ttl = (iHorasTTL >= 0 ? (parseInt(fa[i][iHorasTTL], 10) || 1) : 1);
      fVencISO = new Date(fAsigMs + ttl * 60 * 60 * 1000).toISOString();
    }
    result.push({
      idCobro:           String(fa[i][iIdCobro]),
      idVenta:           String(fa[i][iIdVenta]),
      cajaDestino:       String(fa[i][iCajaDest]),
      vendedorDest:      String(fa[i][iVendDest]),
      metodoSug:         String(fa[i][iMetodo]),
      adminAsig:         String(fa[i][iAdminAsig]),
      fechaAsig:         fa[i][iFAsig] instanceof Date ? fa[i][iFAsig].toISOString() : String(fa[i][iFAsig] || ''),
      fechaVencimiento:  fVencISO,
      horasTTL:          iHorasTTL >= 0 ? (parseInt(fa[i][iHorasTTL], 10) || 1) : 1,
      monto:             parseFloat(fa[i][iMonto]) || 0,
      cliente:           String(fa[i][iCliente]),
      correlativo:       String(fa[i][iCorrel])
    });
  }
  return ContentService.createTextOutput(JSON.stringify({
    status: 'success', cobros: result
  })).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// getCreditosPendientes — para mostrar en MOS las cards
// ============================================================
// GET ?accion=creditos_pendientes&diasAtras=30
// Lista ventas con FormaPago=CREDITO o POR_COBRAR de los últimos N días,
// agrupadas por día. Excluye las que ya tienen una asignación COBRADA.
function getCreditosPendientes(diasAtras) {
  var dias = parseInt(diasAtras, 10) || 30;
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var ventas = ss.getSheetByName('VENTAS_CABECERA');
  if (!ventas) return generarRespuestaError('VENTAS_CABECERA no encontrada');

  // Set de idVenta YA cobradas (estado COBRADO en CREDITOS_COBRO_ASIGNADO)
  var cobradas = {};
  var asigSet  = {};   // {idVenta: idCobro} si está ASIGNADO
  try {
    var hoja = _getHojaCobrosAsignados();
    var fa = hoja.getDataRange().getValues();
    for (var k = 1; k < fa.length; k++) {
      var estK = String(fa[k][5]);
      if (estK === 'COBRADO')  cobradas[String(fa[k][1])] = true;
      if (estK === 'ASIGNADO') asigSet[String(fa[k][1])]  = {
        idCobro: String(fa[k][0]),
        cajaDestino: String(fa[k][2]),
        vendedorDest: String(fa[k][3]),
        fechaAsig: fa[k][7]
      };
    }
  } catch(_){}

  var tz = Session.getScriptTimeZone();
  var hoy = new Date();
  var limite = new Date();
  limite.setDate(hoy.getDate() - dias);

  // [v41.2] Cargar VENTAS_DETALLE una sola vez para mapear items por idVenta.
  // Resolvemos los índices de columna leyendo la cabecera real (no asumir orden).
  var itemsPorVenta = {};
  try {
    var detSh = ss.getSheetByName('VENTAS_DETALLE');
    if (detSh) {
      var fd = detSh.getDataRange().getValues();
      if (fd.length >= 2) {
        var hdrs = fd[0].map(function(h){ return String(h).trim(); });
        var iId   = hdrs.indexOf('ID_Venta');
        var iNom  = hdrs.indexOf('Nombre');
        var iCant = hdrs.indexOf('Cantidad');
        var iPrec = hdrs.indexOf('Precio');
        var iSub  = hdrs.indexOf('Subtotal');
        // Fallbacks defensivos si el header está en otro idioma/case
        if (iId   < 0) iId   = 0;
        if (iNom  < 0) iNom  = 2;
        if (iCant < 0) iCant = 3;
        if (iPrec < 0) iPrec = 4;
        if (iSub  < 0) iSub  = 5;
        for (var d = 1; d < fd.length; d++) {
          var idVD = String(fd[d][iId]);
          if (!idVD) continue;
          if (!itemsPorVenta[idVD]) itemsPorVenta[idVD] = [];
          if (itemsPorVenta[idVD].length >= 12) continue; // tope generoso por ticket
          var cant = parseFloat(fd[d][iCant]) || 0;
          var sub  = parseFloat(fd[d][iSub])  || 0;
          // Si Subtotal está vacío pero hay cantidad × precio, calcular
          if (!sub) sub = cant * (parseFloat(fd[d][iPrec]) || 0);
          itemsPorVenta[idVD].push({
            nombre:   String(fd[d][iNom] || ''),
            cantidad: cant,
            subtotal: sub
          });
        }
      }
    }
  } catch(_){}

  var fv = ventas.getDataRange().getValues();
  // Agrupar por día
  var porDia = {};
  for (var i = 1; i < fv.length; i++) {
    var fp = String(fv[i][8] || '').toUpperCase();
    // [v40.5] Solo CREDITO entra a la baraja. POR_COBRAR es del flow del
    // turno del vendedor (no es un crédito formal otorgado al cliente).
    if (fp !== 'CREDITO') continue;
    var idV = String(fv[i][0]);
    if (cobradas[idV]) continue;

    var fecha = fv[i][1] instanceof Date ? fv[i][1] : new Date(fv[i][1]);
    if (isNaN(fecha.getTime())) continue;
    if (fecha < limite) continue;

    var diaKey = Utilities.formatDate(fecha, tz, 'yyyy-MM-dd');
    if (!porDia[diaKey]) porDia[diaKey] = [];
    var itemsTicket = itemsPorVenta[idV] || [];
    porDia[diaKey].push({
      idVenta:        idV,
      correlativo:    String(fv[i][9] || ''),
      cliente:        String(fv[i][5] || ''),
      clienteDoc:     String(fv[i][4] || ''),
      vendedor:       String(fv[i][2] || ''),
      total:          parseFloat(fv[i][6]) || 0,
      formaPago:      String(fv[i][8] || ''),
      obs:            String(fv[i][14] || ''),
      idCaja:         String(fv[i][10] || ''),
      fechaISO:       Utilities.formatDate(fecha, tz, 'yyyy-MM-dd HH:mm:ss'),
      asignado:       asigSet[idV] || null,  // si ya está siendo cobrado
      items:          itemsTicket,           // [v41] top 5 items para mostrar resumen
      itemsCount:     itemsTicket.length
    });
  }

  // Construir array ordenado de días (más reciente primero)
  var dias_ = Object.keys(porDia).sort(function(a,b){ return b.localeCompare(a); });
  var grupos = dias_.map(function(d) {
    return {
      fecha:   d,
      tickets: porDia[d],
      total:   porDia[d].reduce(function(s,t){ return s + t.total; }, 0),
      cuenta:  porDia[d].length
    };
  });
  var totalAcum = grupos.reduce(function(s,g){ return s + g.total; }, 0);

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    grupos: grupos,
    totalAcumulado: totalAcum,
    totalTickets: grupos.reduce(function(s,g){ return s + g.cuenta; }, 0)
  })).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// escalarCobrosVencidos — trigger 5min
// ============================================================
// Marca como TIMEOUT los cobros ASIGNADO > 1 hora y dispara push
// de escalación a admin+master en MOS.
// [v2.5.27] Procesador de vencimientos — debe correr via trigger horario.
// Compara Fecha_Vencimiento contra ahora. Si venció:
//   1. Estado: EXPIRADO (semánticamente "vencido sin cobrarse")
//   2. VENTAS_CABECERA.formaPago = 'CREDITO' (vuelve al pool original)
//   3. VENTAS_CABECERA.cobrado = false
//   4. Push push al admin asignador: "expiró, re-asignar?"
// Ya NO usa el TTL hardcodeado 1h — lee Fecha_Vencimiento de cada row.
function escalarCobrosVencidos() {
  var hoja = _getHojaCobrosAsignados();
  var fa = hoja.getDataRange().getValues();
  if (fa.length < 2) return { ok: true, vencidos: 0 };
  // Indexar columnas por header (resiliente a reordenamiento)
  var hdrs = fa[0].map(function(h){ return String(h || '').trim(); });
  var iEstado    = hdrs.indexOf('Estado');
  var iAsig      = hdrs.indexOf('Fecha_Asig');
  var iRes       = hdrs.indexOf('Fecha_Res');
  var iRazon     = hdrs.indexOf('Razon');
  var iIdVenta   = hdrs.indexOf('ID_Venta');
  var iMonto     = hdrs.indexOf('Monto');
  var iCliente   = hdrs.indexOf('Cliente_Nombre');
  var iVendedor  = hdrs.indexOf('Vendedor_Dest');
  var iAdminAsig = hdrs.indexOf('Admin_Asignador');
  var iVenc      = hdrs.indexOf('Fecha_Vencimiento');
  var iHorasTTL  = hdrs.indexOf('Horas_TTL');
  var ahora = new Date().getTime();
  var UNA_HORA = 60 * 60 * 1000;
  var url = PropertiesService.getScriptProperties().getProperty('MOS_WEB_APP_URL');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ventasSh = ss.getSheetByName('VENTAS_CABECERA');
  var n = 0;

  for (var i = 1; i < fa.length; i++) {
    if (String(fa[i][iEstado]) !== 'ASIGNADO') continue;
    // Determinar la fecha de vencimiento:
    // 1. Si Fecha_Vencimiento está set → usarla
    // 2. Sino fallback Fecha_Asig + (Horas_TTL || 1h)
    var fVencMs;
    var fVenc = iVenc >= 0 ? fa[i][iVenc] : null;
    if (fVenc) {
      fVencMs = fVenc instanceof Date ? fVenc.getTime() : new Date(fVenc).getTime();
    } else {
      var fAsigMs = fa[i][iAsig] instanceof Date ? fa[i][iAsig].getTime() : new Date(fa[i][iAsig]).getTime();
      if (isNaN(fAsigMs)) continue;
      var ttlMs = (iHorasTTL >= 0 ? (parseInt(fa[i][iHorasTTL], 10) || 1) : 1) * UNA_HORA;
      fVencMs = fAsigMs + ttlMs;
    }
    if (isNaN(fVencMs) || ahora < fVencMs) continue;

    // VENCIDO — marcar y restaurar
    hoja.getRange(i + 1, iEstado + 1).setValue('EXPIRADO');
    hoja.getRange(i + 1, iRes + 1).setValue(new Date());
    hoja.getRange(i + 1, iRazon + 1).setValue('Vencido sin cobrarse · cliente no llegó');

    // [v2.5.27] Restaurar VENTAS_CABECERA → vuelve a estado CREDITO
    // Cols esperadas en VENTAS_CABECERA: 0=ID, 7=cobrado?, 8=formaPago
    // (puede variar — buscar por header)
    try {
      if (ventasSh) {
        var vdAll = ventasSh.getDataRange().getValues();
        var vHdrs = vdAll[0].map(function(h){ return String(h || '').trim(); });
        var iVId   = vHdrs.indexOf('ID') >= 0 ? vHdrs.indexOf('ID') : 0;
        var iVCob  = vHdrs.indexOf('Cobrado') >= 0 ? vHdrs.indexOf('Cobrado') : vHdrs.indexOf('cobrado');
        var iVFp   = vHdrs.indexOf('FormaPago') >= 0 ? vHdrs.indexOf('FormaPago') : vHdrs.indexOf('Forma_Pago');
        var idV = String(fa[i][iIdVenta]);
        for (var k = vdAll.length - 1; k > 0; k--) {
          if (String(vdAll[k][iVId]) === idV) {
            if (iVFp >= 0)  ventasSh.getRange(k + 1, iVFp + 1).setValue('CREDITO');
            if (iVCob >= 0) ventasSh.getRange(k + 1, iVCob + 1).setValue(false);
            break;
          }
        }
      }
    } catch(eR) { Logger.log('Restore venta a CREDITO: ' + eR.message); }

    n++;

    // Push al admin asignador (no a todos, solo al que asignó)
    if (url) {
      try {
        UrlFetchApp.fetch(url, {
          method: 'post',
          contentType: 'application/json',
          payload: JSON.stringify({
            action: 'enviarPushUsuario',
            usuario: String(fa[i][iAdminAsig] || ''),
            titulo: '⏰ Cobro expirado · ' + String(fa[i][iCliente] || 'cliente'),
            cuerpo: 'S/ ' + parseFloat(fa[i][iMonto] || 0).toFixed(2) +
                    ' asignado a ' + String(fa[i][iVendedor]) + ' venció sin cobrarse. ' +
                    'Volvió a CRÉDITO. Re-asignar?',
            idNotif: 'CREDITO_COBRO_EXPIRADO',
            extra: { idVenta: String(fa[i][iIdVenta]) }
          }),
          muteHttpExceptions: true
        });
      } catch(_){}
    }
  }
  Logger.log('escalarCobrosVencidos · vencidos: ' + n);
  return { ok: true, vencidos: n };
}

// ============================================================
// configurarTriggerEscalacion — ejecutar UNA vez desde editor GAS
// ============================================================
function configurarTriggerEscalacionCobros() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'escalarCobrosVencidos') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('escalarCobrosVencidos').timeBased().everyMinutes(5).create();
  Logger.log('Trigger escalarCobrosVencidos configurado · cada 5 min');
  return { ok: true };
}

// ============================================================
// Helpers internos
// ============================================================

// Devuelve { printerId } de la caja consultando ZONAS_CONFIG (estación de la caja).
function _resolverPrinterCaja(cajaId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var cajas = ss.getSheetByName('CAJAS');
    if (!cajas) return null;
    var fc = cajas.getDataRange().getValues();
    var estacionNombre = '';
    for (var i = fc.length - 1; i > 0; i--) {
      if (String(fc[i][0]) === String(cajaId)) { estacionNombre = String(fc[i][2] || ''); break; }
    }
    if (!estacionNombre) return null;

    var zonas = ss.getSheetByName('ZONAS_CONFIG');
    if (!zonas) return null;
    var fz = zonas.getDataRange().getValues();
    var hdrs = fz[0].map(function(h){ return String(h).trim(); });
    var iEst = hdrs.indexOf('Estacion_Nombre');
    var iPid = hdrs.indexOf('PrintNode_ID');
    if (iEst < 0 || iPid < 0) return null;
    for (var j = 1; j < fz.length; j++) {
      if (String(fz[j][iEst]) === estacionNombre) {
        return { printerId: fz[j][iPid] };
      }
    }
  } catch(e) { Logger.log('_resolverPrinterCaja: ' + e.message); }
  return null;
}

// Reimprime el ticket original con un sello "PAGADO · COBRO DIFERIDO" arriba.
// Reutiliza imprimirTicketInternamente con un payload reconstruido + flag esPagoDiferido.
function _reimprimirTicketConSello(idVenta, printerId, ctx) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ventas = ss.getSheetByName('VENTAS_CABECERA');
  var detSheet = ss.getSheetByName('VENTAS_DETALLE');
  if (!ventas || !detSheet) return false;

  var fv = ventas.getDataRange().getValues();
  var venta = null;
  for (var i = fv.length - 1; i > 0; i--) {
    if (String(fv[i][0]) === String(idVenta)) {
      venta = {
        correlativo: String(fv[i][9]),
        tipoDoc:     String(fv[i][7]),
        cliente: {
          doc:    String(fv[i][4] || ''),
          nombre: String(fv[i][5] || ''),
          tipo:   parseInt(fv[i][15] || 0, 10)
        },
        metodo:      ctx.metodoFinal,
        total:       parseFloat(fv[i][6]) || 0,
        vendedor:    String(fv[i][2] || '')
      };
      break;
    }
  }
  if (!venta) return false;

  // Reconstruir items
  var fd = detSheet.getDataRange().getValues();
  var items = [];
  for (var k = 1; k < fd.length; k++) {
    if (String(fd[k][0]) !== String(idVenta)) continue;
    items.push({
      sku:      String(fd[k][1] || ''),
      nombre:   String(fd[k][2] || ''),
      cantidad: parseFloat(fd[k][3]) || 0,
      precio:   parseFloat(fd[k][4]) || 0,
      subtotal: parseFloat(fd[k][5]) || 0
    });
  }

  var data = {
    auth:   { vendedor: ctx.cajero || venta.vendedor },
    header: {
      tipoDoc: venta.tipoDoc,
      total:   venta.total,
      metodo:  venta.metodo,
      cliente: venta.cliente
    },
    items: items,
    // [v40.3] Flag para que imprimirTicketInternamente agregue el sello
    esPagoDiferido: true,
    pagoDiferido: {
      cajaCobro:  ctx.cajaDest,
      cajeroCobro: ctx.cajero,
      adminAsig:  ctx.adminAsig,
      fechaCobro: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
    }
  };

  return imprimirTicketInternamente(data, venta.correlativo, printerId, null);
}

// ============================================================
// [v2.5.27] Auto-instalación del trigger horario que procesa
// vencimientos. Se llama desde asignarCobroACajero — si nadie
// asigna no se instala, pero al primer asignamiento queda activo.
// ============================================================
function _ensureTriggerEscalarCobros() {
  try {
    var existe = ScriptApp.getProjectTriggers().some(function(t) {
      return t.getHandlerFunction() === 'escalarCobrosVencidos';
    });
    if (!existe) {
      ScriptApp.newTrigger('escalarCobrosVencidos').timeBased().everyMinutes(15).create();
      Logger.log('[Trigger] escalarCobrosVencidos auto-instalado · cada 15min');
    }
  } catch(e) { Logger.log('[Trigger escalarCobrosVencidos] auto fallo: ' + e.message); }
}

// Setup público — llamable desde Apps Script editor manualmente
function setupTriggerEscalarCobros() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'escalarCobrosVencidos') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('escalarCobrosVencidos').timeBased().everyMinutes(15).create();
  return { ok: true, mensaje: 'Trigger escalarCobrosVencidos creado · cada 15min' };
}
