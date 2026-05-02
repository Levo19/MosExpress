// ============================================================
// MosExpress — Guias.gs
// Guías de stock por zona, auditorías físicas y traslados.
// ============================================================

// Auto-genera una guía SALIDA_VENTAS al cerrar caja y descuenta STOCK_ZONAS
// Optimizada: lee STOCK_ZONAS una sola vez, hace updates en memoria,
// y escribe el GUIAS_DETALLE + STOCK_ZONAS modificado en batch.
// DEFENSA EN PROFUNDIDAD: chequea si ya existe una guía SALIDA_VENTAS para
// esta caja antes de generar — evita duplicación incluso si la idempotencia
// de procesarCierreCaja falla por algún motivo.
function generarGuiaSalidaVentas(ss, cajaId, vendedor, zona) {
  var sheetVC    = ss.getSheetByName("VENTAS_CABECERA");
  var sheetVD    = ss.getSheetByName("VENTAS_DETALLE");
  var sheetGC    = ss.getSheetByName("GUIAS_CABECERA");
  var sheetGD    = ss.getSheetByName("GUIAS_DETALLE");
  var sheetStock = ss.getSheetByName("STOCK_ZONAS");
  if (!sheetVC || !sheetVD || !sheetGC || !sheetGD || !sheetStock) return;

  // 0. DEFENSA: ¿ya existe guía SALIDA_VENTAS para esta caja? Si sí, abortar.
  // Identificamos la guía por: Tipo='SALIDA_VENTAS' + Observacion contiene cajaId
  var gcData = sheetGC.getDataRange().getValues();
  for (var g = 1; g < gcData.length; g++) {
    if (String(gcData[g][4]) === 'SALIDA_VENTAS' &&
        String(gcData[g][5] || '').indexOf(String(cajaId)) >= 0) {
      Logger.log('generarGuiaSalidaVentas: ya existe guía para caja ' + cajaId + ' (id=' + gcData[g][0] + ') — saltando.');
      return;
    }
  }

  // 1. IDs de ventas no anuladas de esta caja
  var ventas = sheetVC.getDataRange().getValues();
  var idsVentaSet = {};
  for (var i = 1; i < ventas.length; i++) {
    if (String(ventas[i][10]) === String(cajaId) && String(ventas[i][8]) !== 'ANULADO') {
      idsVentaSet[String(ventas[i][0])] = true;
    }
  }
  var idsVenta = Object.keys(idsVentaSet);
  if (!idsVenta.length) return;

  // 2. Sumar cantidades por Cod_Barras
  var detalle = sheetVD.getDataRange().getValues();
  var totales = {};
  for (var j = 1; j < detalle.length; j++) {
    if (!idsVentaSet[String(detalle[j][0])]) continue;
    var cod = String(detalle[j][6] || detalle[j][1]).trim();
    if (!cod) continue;
    totales[cod] = (totales[cod] || 0) + (parseFloat(detalle[j][3]) || 0);
  }

  var cods = Object.keys(totales);
  if (!cods.length) return;

  // 3. Registrar cabecera de guía
  var idGuia = "G-VENTAS-" + new Date().getTime();
  sheetGC.appendRow([idGuia, new Date(), vendedor, zona, 'SALIDA_VENTAS',
    'Auto cierre de caja · ' + cajaId, '', 'CONFIRMADO']);

  // 4. Detalle de guía — batch append en una sola escritura
  var detalleRows = cods.map(function(cod) { return [idGuia, String(cod), totales[cod]]; });
  var startRow = sheetGD.getLastRow() + 1;
  sheetGD.getRange(startRow, 2, detalleRows.length, 1).setNumberFormat('@STRING@');
  sheetGD.getRange(startRow, 1, detalleRows.length, 3).setValues(detalleRows);

  // 5. Stock — leer una sola vez, modificar en memoria, escribir cambios en batch
  var stockData = sheetStock.getDataRange().getValues();
  var stockHdr  = stockData[0];
  var stockMap  = {}; // "cod|zona" → indice de fila (0-based desde header)
  for (var s = 1; s < stockData.length; s++) {
    var key = String(stockData[s][0]) + '|' + String(stockData[s][1]);
    stockMap[key] = s;
  }

  var nuevasFilas = [];
  cods.forEach(function(cod) {
    var key = String(cod) + '|' + String(zona);
    var idx = stockMap[key];
    if (idx !== undefined) {
      stockData[idx][2] = (parseFloat(stockData[idx][2]) || 0) - totales[cod];
    } else {
      nuevasFilas.push([String(cod), String(zona), -totales[cod]]);
    }
  });

  // Re-escribir solo las filas modificadas (saltando header)
  if (stockData.length > 1) {
    sheetStock.getRange(2, 1, stockData.length - 1, stockHdr.length).setValues(stockData.slice(1));
  }
  // Append filas nuevas si hay
  if (nuevasFilas.length > 0) {
    var newStart = sheetStock.getLastRow() + 1;
    sheetStock.getRange(newStart, 1, nuevasFilas.length, 1).setNumberFormat('@STRING@');
    sheetStock.getRange(newStart, 1, nuevasFilas.length, 3).setValues(nuevasFilas);
  }
}

// ════════════════════════════════════════════════════════════════════════
// HERRAMIENTA DE LIMPIEZA: borra guías SALIDA_VENTAS duplicadas para una caja
// y revierte el stock descontado de más.
//
// USO MANUAL desde el editor de Apps Script:
//   1. Abrir el archivo Guias.gs
//   2. Seleccionar función "limpiarGuiasDuplicadasCaja"
//   3. Ejecutar (▶) — debes editar el cajaId hardcoded primero
// O invocar como Web App:
//   POST { tipoEvento: 'LIMPIAR_DUPLICADOS', cajaId: 'CAJA-XXX' }
// ════════════════════════════════════════════════════════════════════════
function limpiarGuiasDuplicadasCaja(cajaIdParam) {
  var cajaId = cajaIdParam || 'CAJA-EDITAR-AQUI'; // editar antes de correr manual
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetGC    = ss.getSheetByName('GUIAS_CABECERA');
  var sheetGD    = ss.getSheetByName('GUIAS_DETALLE');
  var sheetStock = ss.getSheetByName('STOCK_ZONAS');
  if (!sheetGC || !sheetGD || !sheetStock) {
    return { ok: false, error: 'Hojas no encontradas' };
  }

  Logger.log('limpiarGuiasDuplicadasCaja: buscando cajaId="' + cajaId + '"');

  // 1. Buscar todas las guías SALIDA_VENTAS para esta caja
  // Match flexible: por si el cajaId no aparece en la observación, también
  // matcheamos contra columna H (Caja_ID) si existe (índice 7).
  var gcData = sheetGC.getDataRange().getValues();
  var guiasCaja = []; // {idGuia, rowSheet, fecha, zona, obs}
  var totalSalidaVentas = 0;
  var primeras5Obs = []; // para debug si no encuentra match
  var cajaIdNorm = String(cajaId).trim();
  for (var i = 1; i < gcData.length; i++) {
    if (String(gcData[i][4]) !== 'SALIDA_VENTAS') continue;
    totalSalidaVentas++;
    var obs = String(gcData[i][5] || '').trim();
    if (primeras5Obs.length < 5) primeras5Obs.push({ row: i + 1, obs: obs });
    // Match: cajaId aparece en la observación
    if (obs.indexOf(cajaIdNorm) >= 0) {
      guiasCaja.push({
        idGuia: String(gcData[i][0]),
        rowSheet: i + 1,
        fecha: gcData[i][1],
        zona: String(gcData[i][3]),
        obs: obs
      });
    }
  }

  Logger.log('Total SALIDA_VENTAS en sheet: ' + totalSalidaVentas);
  Logger.log('Match para cajaId "' + cajaIdNorm + '": ' + guiasCaja.length);
  if (guiasCaja.length === 0 && totalSalidaVentas > 0) {
    Logger.log('No matcheó. Primeras 5 observaciones encontradas:');
    primeras5Obs.forEach(function(p) {
      Logger.log('  Fila ' + p.row + ': "' + p.obs + '"');
    });
  }

  if (guiasCaja.length <= 1) {
    return {
      ok: true,
      mensaje: 'Solo hay ' + guiasCaja.length + ' guía para "' + cajaIdNorm + '". Nada que limpiar.',
      cajaIdBuscado: cajaIdNorm,
      totalSalidaVentas: totalSalidaVentas,
      ejemplosObservaciones: primeras5Obs,
      hint: guiasCaja.length === 0 ? 'cajaId no encontrado — revisar formato exacto en columna F (Observacion) de GUIAS_CABECERA. Usar diagnosticarSalidaVentas() para ver cajaIds disponibles.' : ''
    };
  }

  // 2. Conservar la PRIMERA (más antigua), eliminar las demás y revertir stock
  guiasCaja.sort(function(a, b){ return new Date(a.fecha) - new Date(b.fecha); });
  var guiaConservada = guiasCaja[0];
  var guiasAEliminar = guiasCaja.slice(1);
  var idsAEliminar = guiasAEliminar.map(function(g){ return g.idGuia; });
  var zona = guiaConservada.zona;

  // 3. Leer GUIAS_DETALLE de las guías a eliminar y sumar al stock de vuelta
  var gdData = sheetGD.getDataRange().getValues();
  var revertStock = {}; // codBarras → cantidad a sumar de vuelta
  var detalleRowsAEliminar = []; // filas a borrar de GUIAS_DETALLE
  for (var j = gdData.length - 1; j >= 1; j--) {
    if (idsAEliminar.indexOf(String(gdData[j][0])) >= 0) {
      var cod = String(gdData[j][1]);
      var cant = parseFloat(gdData[j][2]) || 0;
      revertStock[cod] = (revertStock[cod] || 0) + cant;
      detalleRowsAEliminar.push(j + 1);
    }
  }

  // 4. Sumar de vuelta al stock (todo en memoria + un setValues batch)
  var stockData = sheetStock.getDataRange().getValues();
  var stockHdr  = stockData[0];
  var stockMap  = {};
  for (var s = 1; s < stockData.length; s++) {
    stockMap[String(stockData[s][0]) + '|' + String(stockData[s][1])] = s;
  }
  Object.keys(revertStock).forEach(function(cod) {
    var key = String(cod) + '|' + String(zona);
    var idx = stockMap[key];
    if (idx !== undefined) {
      stockData[idx][2] = (parseFloat(stockData[idx][2]) || 0) + revertStock[cod];
    }
  });
  if (stockData.length > 1) {
    sheetStock.getRange(2, 1, stockData.length - 1, stockHdr.length).setValues(stockData.slice(1));
  }

  // 5. Eliminar filas de GUIAS_DETALLE (de mayor a menor para no descuadrar índices)
  detalleRowsAEliminar.sort(function(a, b){ return b - a; });
  detalleRowsAEliminar.forEach(function(r){ sheetGD.deleteRow(r); });

  // 6. Eliminar filas de GUIAS_CABECERA (de mayor a menor)
  var cabRows = guiasAEliminar.map(function(g){ return g.rowSheet; }).sort(function(a, b){ return b - a; });
  cabRows.forEach(function(r){ sheetGC.deleteRow(r); });

  return {
    ok: true,
    mensaje: 'Limpieza exitosa',
    conservada: guiaConservada.idGuia,
    eliminadas: idsAEliminar,
    cantidadGuiasEliminadas: idsAEliminar.length,
    productosRevertidos: Object.keys(revertStock).length,
    detalleRevertido: revertStock
  };
}

// ════════════════════════════════════════════════════════════════════════
// DIAGNÓSTICO: lista todas las cajas que tienen guías SALIDA_VENTAS
// y cuántas duplicadas hay por cada una. Útil cuando limpiarGuias...
// no encuentra match y necesitas ver el cajaId exacto.
//
// USO MANUAL: ejecutar la función desde Apps Script editor → ver Logs
// O Web App: GET ?accion=diagnosticar_salida_ventas
// ════════════════════════════════════════════════════════════════════════
function diagnosticarSalidaVentas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetGC = ss.getSheetByName('GUIAS_CABECERA');
  if (!sheetGC) return { ok: false, error: 'GUIAS_CABECERA no encontrada' };

  var gcData = sheetGC.getDataRange().getValues();
  var porObservacion = {};   // observacion completa → count
  var todasSalidaVentas = []; // lista completa con detalles
  for (var i = 1; i < gcData.length; i++) {
    if (String(gcData[i][4]) !== 'SALIDA_VENTAS') continue;
    var obs = String(gcData[i][5] || '').trim();
    porObservacion[obs] = (porObservacion[obs] || 0) + 1;
    todasSalidaVentas.push({
      row: i + 1,
      idGuia: String(gcData[i][0]),
      fecha: gcData[i][1] instanceof Date
        ? Utilities.formatDate(gcData[i][1], Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
        : String(gcData[i][1] || ''),
      vendedor: String(gcData[i][2]),
      zona: String(gcData[i][3]),
      obs: obs,
      estado: String(gcData[i][7] || '')
    });
  }

  // Detectar duplicados (misma observación = misma caja)
  var duplicados = [];
  Object.keys(porObservacion).forEach(function(obs) {
    if (porObservacion[obs] > 1) {
      // Extraer cajaId de la observación si tiene patrón "Auto cierre de caja · CAJA-..."
      var m = obs.match(/CAJA-[\d-]+/);
      duplicados.push({
        observacion: obs,
        cantidad: porObservacion[obs],
        cajaId: m ? m[0] : null
      });
    }
  });

  Logger.log('=== DIAGNÓSTICO SALIDA_VENTAS ===');
  Logger.log('Total guías SALIDA_VENTAS: ' + todasSalidaVentas.length);
  Logger.log('Cajas con duplicados: ' + duplicados.length);
  duplicados.forEach(function(d) {
    Logger.log('  - "' + d.observacion + '" → ' + d.cantidad + ' guías' + (d.cajaId ? ' [cajaId: ' + d.cajaId + ']' : ''));
  });
  Logger.log('--- Últimas 10 guías SALIDA_VENTAS ---');
  todasSalidaVentas.slice(-10).forEach(function(g) {
    Logger.log('  Fila ' + g.row + ' | ' + g.fecha + ' | ' + g.idGuia + ' | obs: "' + g.obs + '"');
  });

  return {
    ok: true,
    total: todasSalidaVentas.length,
    duplicados: duplicados,
    ultimas10: todasSalidaVentas.slice(-10)
  };
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

  var idGuia = "G-" + new Date().getTime();
  sheetCab.appendRow([idGuia, new Date(), data.vendedor, data.zona, tipo,
    data.observacion || '', zonaDestino, 'CONFIRMADO']);

  var stockResult = [];
  (data.items || []).forEach(function(item) {
    var cb = String(item.cod_barras);
    var nextDet = sheetDet.getLastRow() + 1;
    sheetDet.getRange(nextDet, 2).setNumberFormat('@STRING@');
    sheetDet.getRange(nextDet, 1, 1, 3).setValues([[idGuia, cb, item.cantidad]]);
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
      var nextDetE = sheetDet.getLastRow() + 1;
      sheetDet.getRange(nextDetE, 2).setNumberFormat('@STRING@');
      sheetDet.getRange(nextDetE, 1, 1, 3).setValues([[idGuiaEntrada, cb, item.cantidad]]);
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

  // Fill remainder from MOS catalog (products not yet in zone stock)
  if (seleccionados.length < 30) {
    try {
      var mosSsId2 = PropertiesService.getScriptProperties().getProperty('MOS_SS_ID') || '';
      if (mosSsId2) {
        var mosSS2   = SpreadsheetApp.openById(mosSsId2);
        var prodRows = _obtenerHojaMOS(mosSS2, 'PRODUCTOS_MASTER');
        prodRows = prodRows.filter(function(p) { return String(p.estado) !== '0'; })
                           .sort(function() { return Math.random() - 0.5; });
        for (var p = 0; p < prodRows.length && seleccionados.length < 30; p++) {
          var pCb = String(prodRows[p].codigoBarra || prodRows[p].idProducto || '').trim();
          if (!pCb || codsEnZona[pCb]) continue;
          seleccionados.push({ cod_barras: pCb, cantSistema: 0, esCatalogo: true });
        }
      }
    } catch(eCat) {
      Logger.log('getListaAuditoria catalog fill ERROR: ' + eCat.message);
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

  _ensureStockZonasAuditCols(sheetStock);

  var tz       = Session.getScriptTimeZone();
  var ahora    = new Date();
  var ahoraStr = Utilities.formatDate(ahora, tz, 'yyyy-MM-dd HH:mm:ss');
  var hoy      = Utilities.formatDate(ahora, tz, 'yyyy-MM-dd');
  var usuario  = String(data.vendedor || '');
  var idAudit  = "A-" + ahora.getTime();

  // ── Leer tabla AUDITORIAS UNA vez y construir índice para dedup ──
  // key = vendedor|zona|cod_barras → fila (1-based) si ya existe HOY
  var auditData  = sheetAudit.getDataRange().getValues();
  var auditIndex = {};
  for (var r = 1; r < auditData.length; r++) {
    var rawFecha = auditData[r][1];
    var rowDate  = rawFecha instanceof Date
      ? Utilities.formatDate(rawFecha, tz, 'yyyy-MM-dd')
      : String(rawFecha).substring(0, 10);
    if (rowDate !== hoy) continue;
    var k = String(auditData[r][2]) + '|' + String(auditData[r][3]) + '|' + String(auditData[r][4]);
    auditIndex[k] = r + 1; // fila real en Sheets (1-based)
  }

  // Columnas AUDITORIAS: ID_Auditoria(1) | Fecha(2) | Vendedor(3) | Zona_ID(4) | Cod_Barras(5) | Cant_Sistema(6) | Cant_Real(7) | Diferencia(8)
  (data.items || []).forEach(function(item) {
    var cb      = String(item.cod_barras);
    var cantSis = parseFloat(item.cantSistema) || 0;
    var cantReal = parseFloat(item.cantReal) || 0;
    var diff    = cantReal - cantSis;
    var key     = usuario + '|' + data.zona + '|' + cb;

    if (auditIndex[key]) {
      // ── Ya existe fila hoy → ACTUALIZAR (no duplicar) ──
      var existingRow = auditIndex[key];
      sheetAudit.getRange(existingRow, 2).setValue(ahoraStr); // Fecha con hora
      sheetAudit.getRange(existingRow, 6).setValue(cantSis);
      sheetAudit.getRange(existingRow, 7).setValue(cantReal);
      sheetAudit.getRange(existingRow, 8).setValue(diff);
    } else {
      // ── Fila nueva: formatear Cod_Barras como texto ANTES de escribir ──
      var nextAuditRow = sheetAudit.getLastRow() + 1;
      sheetAudit.getRange(nextAuditRow, 5).setNumberFormat('@STRING@');
      sheetAudit.getRange(nextAuditRow, 1, 1, 8).setValues(
        [[idAudit, ahoraStr, usuario, data.zona, cb, cantSis, cantReal, diff]]
      );
      auditIndex[key] = nextAuditRow; // evitar duplicado si el mismo item llega dos veces en el batch
    }

    // Stock: establecer cantidad DIRECTAMENTE al valor real auditado
    _actualizarStockAuditoria(sheetStock, cb, data.zona, cantReal, usuario, ahoraStr);
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

// Establece el stock de un producto a la cantidad real auditada (SET, no delta).
// Barcode stored as string to preserve leading zeros.
function _actualizarStockAuditoria(sheet, codBarras, zonaId, cantReal, usuario, fecha) {
  var data = sheet.getDataRange().getValues();
  var hdrs = data[0].map(function(h) { return String(h).trim(); });
  var colCB   = hdrs.indexOf('Cod_Barras');           if (colCB   < 0) colCB   = 0;
  var colZona = hdrs.indexOf('Zona_ID');              if (colZona < 0) colZona = 1;
  var colCant = hdrs.indexOf('Cantidad');             if (colCant < 0) colCant = 2;
  var colUser = hdrs.indexOf('Usuario');
  var colFech = hdrs.indexOf('Fecha_Ultimo_Registro');

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][colCB]) === String(codBarras) && String(data[i][colZona]) === String(zonaId)) {
      // SET directamente la cantidad auditada (no aplicar delta)
      sheet.getRange(i + 1, colCant + 1).setValue(cantReal);
      if (colUser >= 0) sheet.getRange(i + 1, colUser + 1).setValue(usuario);
      if (colFech >= 0) sheet.getRange(i + 1, colFech + 1).setValue(fecha);
      // Re-escribir barcode como string (corrige filas antiguas guardadas como número)
      sheet.getRange(i + 1, colCB + 1).setNumberFormat('@STRING@');
      sheet.getRange(i + 1, colCB + 1).setValue(String(codBarras));
      return cantReal;
    }
  }
  // Fila nueva: formatear Cod_Barras como texto ANTES de escribir el valor
  var totalCols = Math.max(colCant, colUser >= 0 ? colUser : 0, colFech >= 0 ? colFech : 0) + 1;
  var newRow = new Array(totalCols).fill('');
  newRow[colCB]   = String(codBarras);
  newRow[colZona] = String(zonaId);
  newRow[colCant] = Math.max(0, cantReal);
  if (colUser >= 0) newRow[colUser] = usuario;
  if (colFech >= 0) newRow[colFech] = fecha;
  var nextStockRow = sheet.getLastRow() + 1;
  sheet.getRange(nextStockRow, colCB + 1).setNumberFormat('@STRING@');
  sheet.getRange(nextStockRow, 1, 1, totalCols).setValues([newRow]);
  return Math.max(0, cantReal);
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
  // Formatear Cod_Barras como texto ANTES de escribir para preservar ceros a la izquierda
  var nextRow = sheet.getLastRow() + 1;
  sheet.getRange(nextRow, 1).setNumberFormat('@STRING@');
  sheet.getRange(nextRow, 1, 1, 3).setValues([[String(codBarras), String(zonaId), cantInicial]]);
  return cantInicial;
}
