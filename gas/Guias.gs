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

function registrarAuditoria(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetAudit = ss.getSheetByName("AUDITORIAS");
  var sheetStock = ss.getSheetByName("STOCK_ZONAS");
  if (!sheetAudit) return generarRespuestaError("Pestaña AUDITORIAS no encontrada.");
  if (!sheetStock) return generarRespuestaError("Pestaña STOCK_ZONAS no encontrada.");

  var idAudit = "A-" + new Date().getTime();
  // Columnas AUDITORIAS: ID_Auditoria | Fecha | Vendedor | Zona_ID | Cod_Barras | Cant_Sistema | Cant_Real | Diferencia
  (data.items || []).forEach(function(item) {
    var cb       = String(item.cod_barras);
    var cantReal = parseFloat(item.cantReal) || 0;
    var diff     = cantReal - (parseFloat(item.cantSistema) || 0);
    sheetAudit.appendRow([idAudit, new Date(), data.vendedor, data.zona, cb, item.cantSistema, cantReal, diff]);
    actualizarStockFila(sheetStock, cb, data.zona, diff);
  });

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success', idAuditoria: idAudit
  })).setMimeType(ContentService.MimeType.JSON);
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
