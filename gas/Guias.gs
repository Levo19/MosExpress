// ============================================================
// MosExpress — Guias.gs
// Guías de stock por zona, auditorías físicas y traslados.
// ============================================================

// Auto-genera una guía SALIDA_VENTAS al cerrar caja y descuenta STOCK_ZONAS
function generarGuiaSalidaVentas(ss, cajaId, vendedor, zona) {
  var sheetVC    = ss.getSheetByName("VENTAS_CABECERA");
  var sheetVD    = ss.getSheetByName("VENTAS_DETALLE");
  var sheetGC    = ss.getSheetByName("GUIAS_CABECERA");
  var sheetGD    = ss.getSheetByName("GUIAS_DETALLE");
  var sheetStock = ss.getSheetByName("STOCK_ZONAS");
  if (!sheetVC || !sheetVD || !sheetGC || !sheetGD || !sheetStock) return;

  // 1. IDs de ventas no anuladas de esta caja
  var ventas = sheetVC.getDataRange().getValues();
  var idsVenta = [];
  for (var i = 1; i < ventas.length; i++) {
    if (String(ventas[i][10]) === String(cajaId) && String(ventas[i][8]) !== 'ANULADO') {
      idsVenta.push(String(ventas[i][0]));
    }
  }
  if (!idsVenta.length) return;

  // 2. Sumar cantidades por Cod_Barras
  // Col 7 (índice 6) = Cod_Barras; col 1 (índice 1) = SKU (fallback ventas antiguas)
  var detalle = sheetVD.getDataRange().getValues();
  var totales = {};
  for (var j = 1; j < detalle.length; j++) {
    if (idsVenta.indexOf(String(detalle[j][0])) === -1) continue;
    var cod = String(detalle[j][6] || detalle[j][1]).trim();
    if (!cod) continue;
    totales[cod] = (totales[cod] || 0) + (parseFloat(detalle[j][3]) || 0);
  }

  var cods = Object.keys(totales);
  if (!cods.length) return;

  // 3. Registrar guía y descontar stock
  var idGuia = "G-VENTAS-" + new Date().getTime();
  // GUIAS_CABECERA: ID_Guia | Fecha | Vendedor | Zona_ID | Tipo | Observacion | Zona_Destino | Estado
  sheetGC.appendRow([idGuia, new Date(), vendedor, zona, 'SALIDA_VENTAS',
    'Auto cierre de caja · ' + cajaId, '', 'CONFIRMADO']);
  cods.forEach(function(cod) {
    sheetGD.appendRow([idGuia, String(cod), totales[cod]]);
    actualizarStockFila(sheetStock, cod, zona, -totales[cod]);
  });
}

function listarGuias(zona) {
  if (!zona) return generarRespuestaError("zona requerida");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("GUIAS_CABECERA");
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'success', guias: [] }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][3]) === zona || String(data[i][6] || '') === zona) {
      result.push({
        id_guia:      String(data[i][0]),
        fecha:        data[i][1],
        vendedor:     String(data[i][2]),
        zona:         String(data[i][3]),
        tipo:         String(data[i][4]),
        observacion:  String(data[i][5] || ''),
        zona_destino: String(data[i][6] || ''),
        estado:       String(data[i][7] || '')
      });
    }
  }
  result.sort(function(a, b) { return new Date(b.fecha) - new Date(a.fecha); });
  return ContentService.createTextOutput(JSON.stringify({ status: 'success', guias: result }))
    .setMimeType(ContentService.MimeType.JSON);
}

function detalleGuia(idGuia) {
  if (!idGuia) return generarRespuestaError("id_guia requerido");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("GUIAS_DETALLE");
  if (!sheet) return generarRespuestaError("GUIAS_DETALLE no encontrada");
  var data = sheet.getDataRange().getValues();
  var items = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(idGuia)) {
      items.push({ cod_barras: String(data[i][1]), cantidad: parseFloat(data[i][2]) || 0 });
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ status: 'success', items: items }))
    .setMimeType(ContentService.MimeType.JSON);
}

function trasladosEntrantes(zona, desde) {
  if (!zona) return generarRespuestaError("zona requerida");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("GUIAS_CABECERA");
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'success', traslados: [] }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var desdeDate = desde ? new Date(parseInt(desde)) : new Date(Date.now() - 86400000);
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][4]) !== 'ENTRADA_TRASLADO') continue;
    if (String(data[i][3]) !== zona) continue;
    if (new Date(data[i][1]) > desdeDate) {
      result.push({
        id_guia:     String(data[i][0]),
        fecha:       data[i][1],
        origen:      String(data[i][6] || ''),
        observacion: String(data[i][5] || '')
      });
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ status: 'success', traslados: result }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getStockZonas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("STOCK_ZONAS");
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'success', stock: [] }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var data = obtenerDatosHojaComoJSON(sheet);
  return ContentService.createTextOutput(JSON.stringify({ status: 'success', stock: data }))
    .setMimeType(ContentService.MimeType.JSON);
}

function registrarGuia(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetCab   = ss.getSheetByName("GUIAS_CABECERA");
  var sheetDet   = ss.getSheetByName("GUIAS_DETALLE");
  var sheetStock = ss.getSheetByName("STOCK_ZONAS");
  if (!sheetCab)   return generarRespuestaError("Pestaña GUIAS_CABECERA no encontrada.");
  if (!sheetDet)   return generarRespuestaError("Pestaña GUIAS_DETALLE no encontrada.");
  if (!sheetStock) return generarRespuestaError("Pestaña STOCK_ZONAS no encontrada.");

  // Tipos: SALIDA_JEFA | SALIDA_MOVIMIENTO | SALIDA_VENTAS
  //        ENTRADA_ALMACEN | ENTRADA_TRASLADO | ENTRADA_LIBRE
  var tipo        = data.tipo;
  var esSalida    = (tipo.indexOf('SALIDA') === 0);
  var signo       = esSalida ? -1 : 1;
  var zonaDestino = String(data.zona_destino || '');

  // Validar stock para salidas
  if (esSalida) {
    var stockData = sheetStock.getDataRange().getValues();
    for (var si = 0; si < (data.items || []).length; si++) {
      var siItem = data.items[si];
      var siCb   = String(siItem.cod_barras);
      var stockActual = 0;
      for (var sr = 1; sr < stockData.length; sr++) {
        if (String(stockData[sr][0]) === siCb && String(stockData[sr][1]) === String(data.zona)) {
          stockActual = parseFloat(stockData[sr][2]) || 0;
          break;
        }
      }
      if (stockActual < siItem.cantidad) {
        return generarRespuestaError('Stock insuficiente para ' + siCb +
          ': disponible=' + stockActual + ', solicitado=' + siItem.cantidad);
      }
    }
  }

  var idGuia = "G-" + new Date().getTime();
  sheetCab.appendRow([idGuia, new Date(), data.vendedor, data.zona, tipo,
    data.observacion || '', zonaDestino, 'CONFIRMADO']);

  var stockResult = [];
  (data.items || []).forEach(function(item) {
    var cb = String(item.cod_barras);
    sheetDet.appendRow([idGuia, cb, item.cantidad]);
    var nuevaCant = actualizarStockFila(sheetStock, cb, data.zona, signo * item.cantidad);
    stockResult.push({ cod_barras: cb, cantidad: nuevaCant });
  });

  // SALIDA_MOVIMIENTO → genera ENTRADA_TRASLADO automática en zona destino
  var idGuiaEntrada = null;
  if (tipo === 'SALIDA_MOVIMIENTO' && zonaDestino) {
    idGuiaEntrada = "G-TRA-" + (new Date().getTime() + 1);
    sheetCab.appendRow([idGuiaEntrada, new Date(), data.vendedor, zonaDestino, 'ENTRADA_TRASLADO',
      'Traslado desde ' + data.zona + ' — Guía origen: ' + idGuia, data.zona, 'CONFIRMADO']);
    (data.items || []).forEach(function(item) {
      var cb = String(item.cod_barras);
      sheetDet.appendRow([idGuiaEntrada, cb, item.cantidad]);
      actualizarStockFila(sheetStock, cb, zonaDestino, item.cantidad);
    });
  }

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success', idGuia: idGuia, idGuiaEntrada: idGuiaEntrada, stock: stockResult
  })).setMimeType(ContentService.MimeType.JSON);
}

// Returns up to 30 stock items for audit: prioritises products not audited in 7+ days,
// fills remainder from PRESENTACIONES catalog (items not yet in zone stock).
function getListaAuditoria(zona, usuario) {
  if (!zona) return ContentService.createTextOutput(JSON.stringify({ status: 'error', mensaje: 'zona requerida' }))
    .setMimeType(ContentService.MimeType.JSON);

  var ss         = SpreadsheetApp.getActiveSpreadsheet();
  var sheetStock = ss.getSheetByName('STOCK_ZONAS');
  var sheetPres  = ss.getSheetByName('PRESENTACIONES');
  var items      = [];
  var codsEnZona = {};

  if (sheetStock && sheetStock.getLastRow() > 1) {
    var stockData = sheetStock.getDataRange().getValues();
    var hdrs  = stockData[0].map(function(h) { return String(h).trim(); });
    var colCB   = hdrs.indexOf('Cod_Barras');           if (colCB   < 0) colCB   = 0;
    var colZona = hdrs.indexOf('Zona_ID');              if (colZona < 0) colZona = 1;
    var colCant = hdrs.indexOf('Cantidad');             if (colCant < 0) colCant = 2;
    var colFech = hdrs.indexOf('Fecha_Ultimo_Registro');

    for (var i = 1; i < stockData.length; i++) {
      if (String(stockData[i][colZona]) !== String(zona)) continue;
      var cb = String(stockData[i][colCB]);
      if (!cb) continue;
      codsEnZona[cb] = true;

      var fechaReg = colFech >= 0 ? stockData[i][colFech] : null;
      var diasSin  = fechaReg ? (Date.now() - new Date(fechaReg).getTime()) / 86400000 : 9999;

      if (diasSin >= 7) {
        items.push({ cod_barras: cb, cantSistema: parseFloat(stockData[i][colCant]) || 0, diasSin: diasSin, esCatalogo: false });
      }
    }
  }

  // Oldest audit first
  items.sort(function(a, b) { return b.diasSin - a.diasSin; });
  var seleccionados = items.slice(0, 30);

  // Fill remainder from catalog (products not yet in zone stock)
  if (seleccionados.length < 30 && sheetPres && sheetPres.getLastRow() > 1) {
    var presData = sheetPres.getDataRange().getValues();
    var presHdrs = presData[0].map(function(h) { return String(h).trim(); });
    var presColCB = presHdrs.indexOf('Cod_Barras'); if (presColCB < 0) presColCB = 0;
    // Shuffle catalog rows for variety
    var presRows = presData.slice(1).sort(function() { return Math.random() - 0.5; });
    for (var p = 0; p < presRows.length && seleccionados.length < 30; p++) {
      var pCb = String(presRows[p][presColCB]);
      if (!pCb || codsEnZona[pCb]) continue;
      seleccionados.push({ cod_barras: pCb, cantSistema: 0, esCatalogo: true });
    }
  }

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    items: seleccionados.map(function(x) {
      return { cod_barras: x.cod_barras, cantSistema: x.cantSistema, esCatalogo: x.esCatalogo || false };
    })
  })).setMimeType(ContentService.MimeType.JSON);
}

function registrarAuditoria(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetAudit = ss.getSheetByName("AUDITORIAS");
  var sheetStock = ss.getSheetByName("STOCK_ZONAS");
  if (!sheetAudit) return generarRespuestaError("Pestaña AUDITORIAS no encontrada.");
  if (!sheetStock) return generarRespuestaError("Pestaña STOCK_ZONAS no encontrada.");

  // Auto-add audit tracking columns if missing
  _ensureStockZonasAuditCols(sheetStock);

  var idAudit  = "A-" + new Date().getTime();
  var usuario  = String(data.vendedor || '');
  var ahora    = new Date();

  // Columnas AUDITORIAS: ID_Auditoria | Fecha | Vendedor | Zona_ID | Cod_Barras | Cant_Sistema | Cant_Real | Diferencia
  (data.items || []).forEach(function(item) {
    var cb       = String(item.cod_barras);
    var cantReal = parseFloat(item.cantReal) || 0;
    var diff     = cantReal - (parseFloat(item.cantSistema) || 0);
    sheetAudit.appendRow([idAudit, ahora, usuario, data.zona, cb, item.cantSistema, cantReal, diff]);
    _actualizarStockAuditoria(sheetStock, cb, data.zona, diff, usuario, ahora);
  });

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success', idAuditoria: idAudit
  })).setMimeType(ContentService.MimeType.JSON);
}

// Adds Usuario and Fecha_Ultimo_Registro columns to STOCK_ZONAS if not present
function _ensureStockZonasAuditCols(sheet) {
  var hdrs = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return String(h).trim(); });
  if (hdrs.indexOf('Usuario') < 0) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue('Usuario');
    hdrs = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return String(h).trim(); });
  }
  if (hdrs.indexOf('Fecha_Ultimo_Registro') < 0) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue('Fecha_Ultimo_Registro');
  }
}

// Updates stock for an audit: applies diff and records who audited + when.
// Barcode stored as string to preserve leading zeros.
function _actualizarStockAuditoria(sheet, codBarras, zonaId, diff, usuario, fecha) {
  var data = sheet.getDataRange().getValues();
  var hdrs = data[0].map(function(h) { return String(h).trim(); });
  var colCB   = hdrs.indexOf('Cod_Barras');           if (colCB   < 0) colCB   = 0;
  var colZona = hdrs.indexOf('Zona_ID');              if (colZona < 0) colZona = 1;
  var colCant = hdrs.indexOf('Cantidad');             if (colCant < 0) colCant = 2;
  var colUser = hdrs.indexOf('Usuario');
  var colFech = hdrs.indexOf('Fecha_Ultimo_Registro');

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][colCB]) === String(codBarras) && String(data[i][colZona]) === String(zonaId)) {
      var nuevaCant = (parseFloat(data[i][colCant]) || 0) + diff;
      sheet.getRange(i + 1, colCant + 1).setValue(nuevaCant);
      if (colUser >= 0) sheet.getRange(i + 1, colUser + 1).setValue(usuario);
      if (colFech >= 0) sheet.getRange(i + 1, colFech + 1).setValue(fecha);
      sheet.getRange(i + 1, colCB + 1).setNumberFormat('@STRING@');
      return nuevaCant;
    }
  }
  // New row — build it respecting column positions
  var totalCols = Math.max(colCant, colUser >= 0 ? colUser : 0, colFech >= 0 ? colFech : 0) + 1;
  var newRow = new Array(totalCols).fill('');
  newRow[colCB]   = String(codBarras);
  newRow[colZona] = String(zonaId);
  newRow[colCant] = Math.max(0, diff);
  if (colUser >= 0) newRow[colUser] = usuario;
  if (colFech >= 0) newRow[colFech] = fecha;
  sheet.appendRow(newRow);
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, colCB + 1).setNumberFormat('@STRING@');
  return Math.max(0, diff);
}

// Actualiza (o crea) la fila de stock para un código+zona. Devuelve la cantidad resultante.
function actualizarStockFila(sheet, codBarras, zonaId, delta) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(codBarras) && String(data[i][1]) === String(zonaId)) {
      var nuevaCant = (parseFloat(data[i][2]) || 0) + delta;
      sheet.getRange(i + 1, 3).setValue(nuevaCant);
      return nuevaCant;
    }
  }
  var cantInicial = Math.max(0, delta);
  sheet.appendRow([String(codBarras), String(zonaId), cantInicial]);
  return cantInicial;
}
