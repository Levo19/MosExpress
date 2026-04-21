// ============================================================
// MosExpress — Ventas.gs
// Registro de ventas, correlativo O(1), consultas de ventas.
// ============================================================

function procesarVenta(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetCabecera = ss.getSheetByName("VENTAS_CABECERA");
  var sheetDetalle  = ss.getSheetByName("VENTAS_DETALLE");

  if (!sheetCabecera) throw new Error("Pestaña VENTAS_CABECERA no encontrada.");
  if (!sheetDetalle)  throw new Error("Pestaña VENTAS_DETALLE no encontrada.");

  var auth   = data.auth        || {};
  var pos    = data.pos_config  || {};
  var header = data.header      || {};
  var items  = data.items       || [];

  // Seguridad de rol: dispositivo cajero sin caja abierta → POR_COBRAR
  if (auth.esCajero && !pos.cajaId) {
    header.metodo = 'POR_COBRAR';
  }

  var fechaActual = new Date();
  var idVenta = "V-" + fechaActual.getTime();

  // ── Idempotencia: evita duplicados cuando el browser reintenta ──────────
  var refLocal = (data.data_sync && data.data_sync.last_sync) ? String(data.data_sync.last_sync) : '';
  if (refLocal) {
    var totalFilas  = sheetCabecera.getLastRow();
    var buscarDesde = Math.max(2, totalFilas - 199);
    var filasBuscar = sheetCabecera.getRange(buscarDesde, 1, totalFilas - buscarDesde + 1, 16).getValues();
    for (var fi = filasBuscar.length - 1; fi >= 0; fi--) {
      if (String(filasBuscar[fi][13]) === refLocal) {
        return { idVenta: String(filasBuscar[fi][0]), correlativo: String(filasBuscar[fi][9]), printDispatched: false };
      }
    }
  }

  // ── Correlativo O(1) con LockService ────────────────────────────────────
  var correlativoNumero = obtenerSiguienteCorrelativoRapido(ss, pos.serieActual);
  var correlativoFinal  = pos.serieActual + "-" + ("000000" + correlativoNumero).slice(-6);

  // ── VENTAS_CABECERA (19 columnas) ────────────────────────────────────────
  // ID_Venta | Fecha | Vendedor | Estacion | Cliente_Doc | Cliente_Nombre | Total
  // | Tipo_Doc | FormaPago | Correlativo | ID_Caja | ID_Dispositivo | Estado_Envio
  // | Ref_Local | Obs | Tipo_Doc_Cliente | NF_Estado | NF_Hash | NF_Enlace
  var tipoDocCliente = parseInt((header.cliente && header.cliente.tipo) || 0, 10);
  sheetCabecera.appendRow([
    idVenta, fechaActual, auth.vendedor, auth.estacion,
    (header.cliente && header.cliente.doc)    || '',
    (header.cliente && header.cliente.nombre) || '',
    header.total,
    header.tipoDoc,
    header.metodo || 'EFECTIVO',
    correlativoFinal, pos.cajaId, auth.deviceId, "COMPLETADO",
    refLocal,
    String(header.obs || ''),
    tipoDocCliente,
    '', '', ''   // NF_Estado, NF_Hash, NF_Enlace — se llenan después si aplica
  ]);

  // ── VENTAS_DETALLE (10 columnas) — escritura por lote ────────────────────
  // ID_Venta | SKU | Nombre | Cantidad | Precio | Subtotal | Cod_Barras
  // | Valor_Unitario | Tipo_IGV | Unidad_Medida
  if (items.length > 0) {
    var detalleRows = items.map(function(item) {
      var valorUnitario = parseFloat(item.valor_unitario) ||
                          Math.round(parseFloat(item.precio || 0) / 1.18 * 100) / 100;
      return [
        idVenta,
        item.sku,
        item.nombre,
        item.cantidad,
        item.precio,
        item.subtotal,
        String(item.codBarras || ''),
        Math.round(valorUnitario * 100) / 100,
        parseInt(item.tipo_igv || 1, 10),
        String(item.unidad_de_medida || 'NIU')
      ];
    });
    var lastRow = sheetDetalle.getLastRow();
    var rangeDetalle = sheetDetalle.getRange(lastRow + 1, 1, detalleRows.length, detalleRows[0].length);
    // Forzar texto en col 7 (Cod_Barras) y col 2 (SKU) antes de escribir
    // para que Sheets no elimine ceros a la izquierda de códigos numéricos
    sheetDetalle.getRange(lastRow + 1, 7, detalleRows.length, 1).setNumberFormat('@STRING@');
    sheetDetalle.getRange(lastRow + 1, 2, detalleRows.length, 1).setNumberFormat('@STRING@');
    rangeDetalle.setValues(detalleRows);
  }

  // ── Registrar cliente frecuente ──────────────────────────────────────────
  if (header.tipoDoc !== 'NOTA_DE_VENTA') {
    verificarYAgregaCliente(
      (header.cliente && header.cliente.doc)       || '',
      (header.cliente && header.cliente.nombre)    || '',
      header.tipoDoc,
      (header.cliente && header.cliente.direccion) || ''
    );
  }

  // ── Emitir CPE en NubeFact (solo BOLETA y FACTURA) ───────────────────────
  var nfEstado = 'NA';
  var nfHash   = '';
  var nfEnlace = '';
  var nfResult = null;

  if (header.tipoDoc === 'BOLETA' || header.tipoDoc === 'FACTURA') {
    nfResult = emitirNubeFact(data, correlativoFinal);
    nfEstado = nfResult.ok ? 'EMITIDO' : 'ERROR';
    nfHash   = nfResult.hash   || '';
    nfEnlace = nfResult.enlace || '';
    var nfRow = sheetCabecera.getLastRow();
    sheetCabecera.getRange(nfRow, 17, 1, 3).setValues([[nfEstado, nfHash, nfEnlace]]);
    if (!nfResult.ok) Logger.log('NubeFact error venta ' + idVenta + ': ' + (nfResult.error || ''));
  }

  // ── Imprimir si el browser lo pidió explícitamente ───────────────────────
  // pos.print_request=true: solo index.html v39+. Sin este flag → browser imprime por su cuenta.
  var printDispatched = false;
  if (pos.print_request === true && pos.printerId) {
    printDispatched = imprimirTicketInternamente(data, correlativoFinal, pos.printerId, nfResult);
  }

  // ── Auto-registro de jornada en MOS (idempotente por nombre + fecha) ───────
  try { _registrarJornadaEnMOS(String(auth.vendedor || '')); } catch(eJ) {
    Logger.log('Auto-jornada MOS: ' + eJ.message);
  }

  return { idVenta: idVenta, correlativo: correlativoFinal, printDispatched: printDispatched,
           nfEstado: nfEstado, nfHash: nfHash, nfEnlace: nfEnlace };
}

// Registra la jornada del vendedor en ProyectoMOS al procesar su primera venta del día.
// Idempotente: si ya existe una jornada con el mismo nombre y fecha no inserta duplicados.
function _registrarJornadaEnMOS(nombreVendedor) {
  if (!nombreVendedor) return;
  var mosSsId = PropertiesService.getScriptProperties().getProperty('MOS_SS_ID');
  if (!mosSsId) return;

  var tz    = Session.getScriptTimeZone();
  var fecha = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var ss    = SpreadsheetApp.openById(mosSsId);
  var sheet = ss.getSheetByName('JORNADAS');
  if (!sheet) return;

  // Idempotencia: verificar si ya existe la jornada hoy
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][3]).toLowerCase() === nombreVendedor.toLowerCase() &&
        String(data[i][1]).substring(0, 10) === fecha) return;
  }

  // Buscar montoBase en PERSONAL_MASTER de MOS
  var monto = 0;
  try {
    var pm    = ss.getSheetByName('PERSONAL_MASTER');
    if (pm) {
      var pmData = pm.getDataRange().getValues();
      var pmHdrs = pmData[0].map(function(h){ return String(h).trim(); });
      var idxNom = pmHdrs.indexOf('nombre');
      var idxMon = pmHdrs.indexOf('montoBase');
      for (var j = 1; j < pmData.length; j++) {
        if (String(pmData[j][idxNom]).toLowerCase() === nombreVendedor.toLowerCase()) {
          monto = parseFloat(pmData[j][idxMon]) || 0;
          break;
        }
      }
    }
  } catch(e2) {}

  sheet.appendRow([
    'JOR' + new Date().getTime(), fecha, '', nombreVendedor,
    'VENDEDOR', 'mosExpress', '', monto, '', 'AUTO', 'AUTO_VENTA'
  ]);
}

// Devuelve todas las ventas de hoy de la zona del cajero (filtradas por prefijos de serie)
function ventasHoyZona(prefijosStr, desdeStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("VENTAS_CABECERA");
  if (!sheet) return generarRespuestaError("VENTAS_CABECERA no encontrada");

  var prefijos = prefijosStr ? prefijosStr.split(',').map(function(p) { return p.trim(); }) : [];
  var data = sheet.getDataRange().getValues();
  var hoy   = new Date().toDateString();
  // Si se envía "desde" (ISO datetime de apertura de caja), filtrar por turno.
  // Si no, usar el filtro legacy de "hoy".
  var desde = (desdeStr && desdeStr.trim()) ? new Date(desdeStr.trim()) : null;
  var result = [];

  for (var i = 1; i < data.length; i++) {
    var fechaTk = data[i][1] instanceof Date ? data[i][1] : new Date(data[i][1]);
    if (desde) {
      // Solo tickets emitidos en o después de la apertura del turno
      if (fechaTk < desde) continue;
    } else {
      if (fechaTk.toDateString() !== hoy) continue;
    }
    var correlativo = String(data[i][9]);
    if (prefijos.length > 0) {
      var enZona = prefijos.some(function(p) { return correlativo.indexOf(p) === 0; });
      if (!enZona) continue;
    }
    result.push({
      id_venta:       data[i][0],
      fecha:          data[i][1],
      vendedor:       String(data[i][2] || ''),
      estacion:       String(data[i][3] || ''),
      cliente_doc:    String(data[i][4] || ''),
      cliente_nombre: String(data[i][5] || ''),
      total:          parseFloat(data[i][6]) || 0,
      tipo_doc:       String(data[i][7] || ''),
      forma_pago:     String(data[i][8] || ''),
      correlativo:    correlativo,
      id_caja:        String(data[i][10] || ''),
      id_dispositivo: String(data[i][11] || ''),
      status:         String(data[i][12] || ''),
      ref_local:      String(data[i][13] || ''),
      obs:            String(data[i][14] || '')
    });
  }

  return ContentService.createTextOutput(JSON.stringify({
    status: "success", ventas: result
  })).setMimeType(ContentService.MimeType.JSON);
}

function detalleVenta(idVenta) {
  if (!idVenta) return generarRespuestaError("id_venta requerido");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("VENTAS_DETALLE");
  if (!sheet) return generarRespuestaError("VENTAS_DETALLE no encontrada");

  var data = sheet.getDataRange().getValues();
  var items = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(idVenta)) {
      items.push({
        sku:      String(data[i][1] || ''),
        nombre:   String(data[i][2] || ''),
        cantidad: parseFloat(data[i][3]) || 0,
        precio:   parseFloat(data[i][4]) || 0,
        subtotal: parseFloat(data[i][5]) || 0
      });
    }
  }
  return ContentService.createTextOutput(JSON.stringify({
    status: "success", items: items
  })).setMimeType(ContentService.MimeType.JSON);
}

function verificarYAgregaCliente(doc, nombre, tipoDoc, direccion) {
  if (!doc) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("CLIENTES_FRECUENTES");
  if (!sheet) return;
  var data = obtenerDatosHojaComoJSON(sheet);
  for (var i = 0; i < data.length; i++) {
    if (String(data[i].Documento) === String(doc)) return; // ya existe
  }
  // Esquema: Documento | Nombre | Tipo | Fecha | Direccion
  sheet.appendRow([doc, nombre, tipoDoc, new Date(), String(direccion || '')]);
}

// ── Correlativo O(1) con LockService ─────────────────────────────────────────
// Hoja CORRELATIVOS: encabezados Serie | Siguiente
// Crea la hoja automáticamente si no existe.
function obtenerSiguienteCorrelativoRapido(ss, serie) {
  var sheet = ss.getSheetByName('CORRELATIVOS');
  if (!sheet) {
    var sheetCab = ss.getSheetByName('VENTAS_CABECERA');
    var initial  = sheetCab ? obtenerSiguienteCorrelativo(sheetCab, serie) : 1;
    sheet = ss.insertSheet('CORRELATIVOS');
    sheet.appendRow(['Serie', 'Siguiente']);
    sheet.appendRow([serie, initial + 1]);
    return initial;
  }

  var lock = LockService.getScriptLock();
  try { lock.waitLock(6000); } catch (e) {
    // Contención: tomar el mayor entre CORRELATIVOS y O(n) para no retroceder
    var sheetCabFb = ss.getSheetByName('VENTAS_CABECERA');
    var numON = sheetCabFb ? obtenerSiguienteCorrelativo(sheetCabFb, serie) : 1;
    try {
      var dataFb = sheet.getDataRange().getValues();
      for (var fi = 1; fi < dataFb.length; fi++) {
        if (String(dataFb[fi][0]) === serie) {
          numON = Math.max(numON, parseInt(dataFb[fi][1], 10) || 1);
          break;
        }
      }
    } catch(e2) {}
    return numON;
  }

  try {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === serie) {
        var siguiente = parseInt(data[i][1], 10) || 1;
        // Anti-duplicado: si CORRELATIVOS quedó detrás de la realidad, avanzar
        while (correlativoYaExiste(ss, serie, siguiente)) { siguiente++; }
        sheet.getRange(i + 1, 2).setValue(siguiente + 1);
        return siguiente;
      }
    }
    // Serie nueva
    var sheetCab2 = ss.getSheetByName('VENTAS_CABECERA');
    var initial2  = sheetCab2 ? obtenerSiguienteCorrelativo(sheetCab2, serie) : 1;
    sheet.appendRow([serie, initial2 + 1]);
    return initial2;
  } finally {
    lock.releaseLock();
  }
}

// Fallback O(n): scan de VENTAS_CABECERA. Solo se usa cuando CORRELATIVOS no existe todavía.
function obtenerSiguienteCorrelativo(sheet, serie) {
  var data = sheet.getDataRange().getValues();
  var prefijo = serie + "-";
  var maxCorrelativo = 0;
  for (var i = 1; i < data.length; i++) {
    var valorSerie = String(data[i][9]);
    if (valorSerie.indexOf(prefijo) === 0) {
      var num = parseInt(valorSerie.substring(prefijo.length), 10);
      if (!isNaN(num) && num > maxCorrelativo) maxCorrelativo = num;
    }
  }
  return maxCorrelativo + 1;
}

// Verifica si un correlativo ya existe en las últimas 100 filas (bounded O(100))
function correlativoYaExiste(ss, serie, numero) {
  var sheetCab = ss.getSheetByName('VENTAS_CABECERA');
  if (!sheetCab) return false;
  var objetivo  = serie + '-' + ('000000' + numero).slice(-6);
  var totalRows = sheetCab.getLastRow();
  if (totalRows < 2) return false;
  var desde = Math.max(2, totalRows - 99);
  var filas = sheetCab.getRange(desde, 10, totalRows - desde + 1, 1).getValues();
  for (var i = 0; i < filas.length; i++) {
    if (String(filas[i][0]) === objetivo) return true;
  }
  return false;
}
