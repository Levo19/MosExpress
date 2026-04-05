// ESTRUCTURA DE BASE DE DATOS ESPERADA EN GOOGLE SHEETS
// Pestañas requeridas: PRODUCTO_BASE, PRESENTACIONES, EQUIVALENCIAS, PROMOCIONES,
//                      ZONAS_CONFIG, CLIENTES_FRECUENTES, VENTAS_CABECERA, VENTAS_DETALLE,
//                      CAJAS, DISPOSITIVOS
// Opcional: LOG_IMPRESIONES

// ============================================================
// CLAVE DE PRINTNODE: guardarla en Proyecto > Propiedades del script
// Nombre de la propiedad: PRINTNODE_API_KEY
// Valor: tu_clave_de_printnode
// ============================================================

function doGet(e) {
  var accion = e.parameter.accion;

  if (accion === 'descargar') {
    return descargarCatalogo();
  }

  if (accion === 'verificar_dispositivo') {
    return verificarDispositivo(e.parameter.id);
  }

  if (accion === 'ventas_hoy_zona') {
    return ventasHoyZona(e.parameter.prefijos);
  }

  if (accion === 'detalle_venta') {
    return detalleVenta(e.parameter.id_venta);
  }

  if (accion === 'stock_zonas') {
    return getStockZonas();
  }

  if (accion === 'cajero_activo') {
    return cajeroActivo(e.parameter.zona);
  }

  if (accion === 'listar_guias') {
    return listarGuias(e.parameter.zona);
  }

  if (accion === 'detalle_guia') {
    return detalleGuia(e.parameter.id_guia);
  }

  if (accion === 'traslados_entrantes') {
    return trasladosEntrantes(e.parameter.zona, e.parameter.desde);
  }

  if (accion === 'consultar_cliente') {
    return consultarCliente(e.parameter.doc);
  }

  return ContentService.createTextOutput(JSON.stringify({ error: "Acción no válida" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var rawData = e.postData.contents;
    var data = JSON.parse(rawData);

    if (data.tipoEvento === 'APERTURA_CAJA') {
      return procesarAperturaCaja(data);
    }

    if (data.tipoEvento === 'CIERRE_CAJA') {
      return procesarCierreCaja(data);
    }

    if (data.tipoEvento === 'COBRAR_VENTA') {
      return cobrarVentaExistente(data);
    }

    if (data.tipoEvento === 'ANULACION_MASIVA') {
      return anulacionMasiva(data);
    }

    if (data.tipoEvento === 'ANULACION') {
      return anularVentaIndividual(data);
    }

    // NUEVO: proxy de impresión — la clave de PrintNode vive aquí, no en el browser
    if (data.accion === 'imprimir') {
      return procesarImpresion(data);
    }

    if (data.tipoEvento === 'REGISTRAR_GUIA') {
      return registrarGuia(data);
    }

    if (data.tipoEvento === 'REGISTRAR_AUDITORIA') {
      return registrarAuditoria(data);
    }

    // Default: registrar venta
    var response = procesarVenta(data);
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      idVenta: response.idVenta,
      correlativo: response.correlativo,
      mensaje: "Venta procesada con éxito"
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return generarRespuestaError(error.toString());
  }
}

// ------ FUNCIONES DE NEGOCIO ------ //

function procesarAperturaCaja(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetCajas = ss.getSheetByName("CAJAS");
  if (!sheetCajas) return generarRespuestaError("Pestaña CAJAS no encontrada.");

  // Verificar cajero único por zona — solo un cajero activo a la vez por zona
  if (data.zona) {
    var filasCheck = sheetCajas.getDataRange().getValues();
    for (var c = 1; c < filasCheck.length; c++) {
      if (String(filasCheck[c][5]) === 'ABIERTA' && String(filasCheck[c][8] || '') === String(data.zona)) {
        return generarRespuestaError(
          "Ya hay un turno activo en " + data.zona + " (cajero: " + filasCheck[c][1] + "). Cierra ese turno primero."
        );
      }
    }
  }

  var fechaActual = new Date();
  var idCaja = "CAJA-" + fechaActual.getTime();

  // Columnas: ID_Caja | Vendedor | Estacion | Fecha_Apertura | Monto_Inicial | Estado | Monto_Final | Fecha_Cierre | Zona_ID
  sheetCajas.appendRow([idCaja, data.vendedor, data.estacion, fechaActual, data.montoInicial, "ABIERTA", "", "", data.zona || '']);

  return ContentService.createTextOutput(JSON.stringify({
    status: "success", idCaja: idCaja, mensaje: "Caja aperturada exitosamente"
  })).setMimeType(ContentService.MimeType.JSON);
}

function procesarCierreCaja(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetCajas = ss.getSheetByName("CAJAS");
  if (!sheetCajas) return generarRespuestaError("Pestaña CAJAS no encontrada.");

  var filas = sheetCajas.getDataRange().getValues();
  var cajaEncontrada = false;
  var cajaVendedor = '';
  var cajaZona = '';

  for (var i = 1; i < filas.length; i++) {
    if (String(filas[i][0]) === String(data.cajaId)) {
      sheetCajas.getRange(i + 1, 6).setValue("CERRADA");
      sheetCajas.getRange(i + 1, 7).setValue(data.montoFinal);
      sheetCajas.getRange(i + 1, 8).setValue(new Date()); // Fecha_Cierre col 8
      cajaVendedor = String(filas[i][1]);
      cajaZona = String(filas[i][8] || ''); // col 9 = Zona_ID
      cajaEncontrada = true;
      break;
    }
  }

  if (!cajaEncontrada) return generarRespuestaError("Caja con ID " + data.cajaId + " no encontrada.");

  // Auto-generar SALIDA_VENTAS en segundo plano (no bloquea la respuesta si falla)
  if (cajaZona) {
    try { generarGuiaSalidaVentas(ss, data.cajaId, cajaVendedor, cajaZona); }
    catch(e) { Logger.log("Error guia ventas: " + e.toString()); }
  }

  return ContentService.createTextOutput(JSON.stringify({
    status: "success", mensaje: "Caja cerrada correctamente"
  })).setMimeType(ContentService.MimeType.JSON);
}

// Agrega registros a GUIAS y descuenta STOCK_ZONAS con todas las ventas del turno
function generarGuiaSalidaVentas(ss, cajaId, vendedor, zona) {
  var sheetVC    = ss.getSheetByName("VENTAS_CABECERA");
  var sheetVD    = ss.getSheetByName("VENTAS_DETALLE");
  var sheetGC    = ss.getSheetByName("GUIAS_CABECERA");
  var sheetGD    = ss.getSheetByName("GUIAS_DETALLE");
  var sheetStock = ss.getSheetByName("STOCK_ZONAS");
  if (!sheetVC || !sheetVD || !sheetGC || !sheetGD || !sheetStock) return;

  // 1. Recopilar IDs de ventas no anuladas de esta caja
  var ventas = sheetVC.getDataRange().getValues();
  var idsVenta = [];
  for (var i = 1; i < ventas.length; i++) {
    if (String(ventas[i][10]) === String(cajaId) && String(ventas[i][8]) !== 'ANULADO') {
      idsVenta.push(String(ventas[i][0]));
    }
  }
  if (!idsVenta.length) return;

  // 2. Sumar cantidades por Cod_Barras
  // Col 6 (índice 6) = Cod_Barras real; col 1 (índice 1) = SKU (fallback para ventas antiguas sin col 7)
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
  sheetGC.appendRow([idGuia, new Date(), vendedor, zona, 'SALIDA_VENTAS', 'Auto cierre de caja · ' + cajaId, '', 'CONFIRMADO']);
  cods.forEach(function(cod) {
    sheetGD.appendRow([idGuia, String(cod), totales[cod]]);
    actualizarStockFila(sheetStock, cod, zona, -totales[cod]);
  });
}

function procesarVenta(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetCabecera = ss.getSheetByName("VENTAS_CABECERA");
  var sheetDetalle = ss.getSheetByName("VENTAS_DETALLE");

  if (!sheetCabecera) throw new Error("Pestaña VENTAS_CABECERA no encontrada.");
  if (!sheetDetalle) throw new Error("Pestaña VENTAS_DETALLE no encontrada.");

  var auth = data.auth;
  var pos = data.pos_config;
  var header = data.header;
  var items = data.items;

  var fechaActual = new Date();
  var idVenta = "V-" + fechaActual.getTime();

  var correlativoNumero = obtenerSiguienteCorrelativo(sheetCabecera, pos.serieActual);
  var correlativoFinal = pos.serieActual + "-" + ("000000" + correlativoNumero).slice(-6);

  // Esquema VENTAS_CABECERA (14 columnas):
  // ID_Venta | Fecha | Vendedor | Estacion | Cliente_Doc | Cliente_Nombre | Total
  // | Tipo_Doc | FormaPago | Correlativo | ID_Caja | ID_Dispositivo | Estado_Envio | Ref_Local
  var refLocal = (data.data_sync && data.data_sync.last_sync) ? String(data.data_sync.last_sync) : '';
  sheetCabecera.appendRow([
    idVenta, fechaActual, auth.vendedor, auth.estacion,
    header.cliente.doc, header.cliente.nombre, header.total,
    header.tipoDoc,                    // col 8: Tipo_Doc  (limpio)
    header.metodo || 'EFECTIVO',       // col 9: FormaPago (EFECTIVO/VIRTUAL/MIXTO/POR_COBRAR)
    correlativoFinal, pos.cajaId, auth.deviceId, "COMPLETADO",
    refLocal                           // col 14: Ref_Local (ID del dispositivo para cross-ref QR)
  ]);

  // Esquema VENTAS_DETALLE (7 columnas):
  // ID_Venta | SKU | Nombre | Cantidad | Precio | Subtotal | Cod_Barras
  items.forEach(function(item) {
    sheetDetalle.appendRow([
      idVenta, item.sku, item.nombre, item.cantidad, item.precio, item.subtotal,
      String(item.codBarras || '')   // col 7: código de barras real — forzar texto
    ]);
  });

  if (header.tipoDoc !== 'NOTA_DE_VENTA') {
    verificarYAgregaCliente(header.cliente.doc, header.cliente.nombre, header.tipoDoc);
  }

  return { idVenta: idVenta, correlativo: correlativoFinal };
}

// NUEVO: proxy de impresión para PrintNode — la clave nunca sale del servidor
function procesarImpresion(data) {
  var printNodeKey = PropertiesService.getScriptProperties().getProperty('PRINTNODE_API_KEY');
  if (!printNodeKey) return generarRespuestaError("PRINTNODE_API_KEY no configurada en Propiedades del script.");
  if (!data.printerId || !data.content) return generarRespuestaError("Faltan datos de impresión (printerId o content).");

  var printerId = parseInt(data.printerId, 10);
  if (isNaN(printerId) || printerId <= 0) {
    return generarRespuestaError("printerId inválido: '" + data.printerId + "'. Verifica el campo PrintNode_ID en la hoja ZONAS_CONFIG.");
  }

  var options = {
    method: 'post',
    headers: {
      'Authorization': 'Basic ' + Utilities.base64Encode(printNodeKey + ':')
    },
    contentType: 'application/json',
    payload: JSON.stringify({
      printerId: printerId,
      title: data.title || 'MOSexpress',
      contentType: 'raw_base64',
      content: data.content,
      source: 'MOSexpress'
    }),
    muteHttpExceptions: true
  };

  try {
    var resp = UrlFetchApp.fetch('https://api.printnode.com/printjobs', options);
    var code = resp.getResponseCode();
    if (code !== 201) {
      return generarRespuestaError("PrintNode respondió " + code + " (printerId=" + printerId + "): " + resp.getContentText());
    }
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      printJobId: resp.getContentText()
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return generarRespuestaError("Error llamando a PrintNode: " + err.toString());
  }
}

// ------ FUNCIONES AUXILIARES ------ //

// CORRECCIÓN: eliminada VENTAS_CABECERA del catálogo — no es necesaria en el browser
// y exponía todo el historial de ventas a cualquier dispositivo autorizado
function descargarCatalogo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetsNames = ["PRODUCTO_BASE", "PRESENTACIONES", "EQUIVALENCIAS", "PROMOCIONES", "ZONAS_CONFIG", "CLIENTES_FRECUENTES", "STOCK_ZONAS"];
  var catalogo = {};

  sheetsNames.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    catalogo[name] = sheet ? obtenerDatosHojaComoJSON(sheet) : [];
  });

  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    data: catalogo
  })).setMimeType(ContentService.MimeType.JSON);
}

// Columnas que deben tratarse siempre como texto (nunca como número)
var COLUMNAS_TEXTO = ['Cod_Barras', 'Cod_Barras_Real', 'SKU_Base', 'SKU', 'ID_Dispositivo', 'ID_Venta', 'ID_Caja', 'ID_Guia'];

function obtenerDatosHojaComoJSON(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  var headers = data[0];
  var result = [];

  for (var i = 1; i < data.length; i++) {
    var rowData = {};
    for (var j = 0; j < headers.length; j++) {
      var header = String(headers[j]).trim();
      var val = data[i][j];
      // Columnas de código siempre como texto para no perder ceros a la izquierda
      if (COLUMNAS_TEXTO.indexOf(header) !== -1) {
        val = val === '' || val === null || val === undefined ? '' : String(val).trim();
      }
      rowData[header] = val;
    }
    result.push(rowData);
  }
  return result;
}

// CORRECCIÓN: parsing más robusto — extrae solo el número final después del último guion
function obtenerSiguienteCorrelativo(sheet, serie) {
  var data = sheet.getDataRange().getValues();
  var prefijo = serie + "-";
  var maxCorrelativo = 0;

  for (var i = 1; i < data.length; i++) {
    var valorSerie = String(data[i][9]); // Columna J: correlativo (col 10, índice 9)
    if (valorSerie.indexOf(prefijo) === 0) {
      var parte = valorSerie.substring(prefijo.length); // todo lo que viene después del prefijo
      var num = parseInt(parte, 10);
      if (!isNaN(num) && num > maxCorrelativo) maxCorrelativo = num;
    }
  }
  return maxCorrelativo + 1;
}

function verificarYAgregaCliente(doc, nombre, tipoDoc) {
  if (!doc) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("CLIENTES_FRECUENTES");
  if (!sheet) return;
  var data = obtenerDatosHojaComoJSON(sheet);
  for (var i = 0; i < data.length; i++) {
    if (String(data[i].Documento) === String(doc)) return; // ya existe
  }
  sheet.appendRow([doc, nombre, tipoDoc, new Date()]);
}

// Anula en masa todos los tickets POR_COBRAR que no fueron cobrados al cierre del turno
function anulacionMasiva(data) {
  if (!data.ids || !data.ids.length) return generarRespuestaError("No se enviaron IDs a anular.");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("VENTAS_CABECERA");
  if (!sheet) return generarRespuestaError("VENTAS_CABECERA no encontrada.");

  var filas = sheet.getDataRange().getValues();
  var anulados = 0;

  for (var i = 1; i < filas.length; i++) {
    if (data.ids.indexOf(String(filas[i][0])) !== -1) {
      sheet.getRange(i + 1, 9).setValue('ANULADO'); // col 9: FormaPago = ANULADO
      anulados++;
    }
  }

  return ContentService.createTextOutput(JSON.stringify({
    status: "success", anulados: anulados
  })).setMimeType(ContentService.MimeType.JSON);
}

// Devuelve todas las ventas de hoy que pertenecen a la zona del cajero
// prefijosStr: "NV-001,B-001,F-001" (series de la zona separadas por coma)
function ventasHoyZona(prefijosStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("VENTAS_CABECERA");
  if (!sheet) return generarRespuestaError("VENTAS_CABECERA no encontrada");

  var prefijos = prefijosStr ? prefijosStr.split(',').map(function(p) { return p.trim(); }) : [];
  var data = sheet.getDataRange().getValues();
  var hoy = new Date().toDateString();
  var result = [];

  for (var i = 1; i < data.length; i++) {
    var fechaFila = new Date(data[i][1]).toDateString();
    if (fechaFila !== hoy) continue;

    var correlativo = String(data[i][9]); // col 10 (índice 9)
    if (prefijos.length > 0) {
      var enZona = prefijos.some(function(p) { return correlativo.indexOf(p) === 0; });
      if (!enZona) continue;
    }

    result.push({
      id_venta:       data[i][0],
      fecha:          data[i][1],
      vendedor:       data[i][2],
      estacion:       data[i][3],
      cliente_doc:    String(data[i][4] || ''),
      cliente_nombre: String(data[i][5] || ''),
      total:          parseFloat(data[i][6]) || 0,
      tipo_doc:       String(data[i][7] || ''),   // col 8: Tipo_Doc limpio
      forma_pago:     String(data[i][8] || ''),   // col 9: FormaPago
      correlativo:    correlativo,
      id_caja:        String(data[i][10] || ''),
      id_dispositivo: String(data[i][11] || ''),
      status:         String(data[i][12] || ''),  // col 13: Estado_Envio
      ref_local:      String(data[i][13] || '')   // col 14: Ref_Local (cross-ref QR)
    });
  }

  return ContentService.createTextOutput(JSON.stringify({
    status: "success", ventas: result
  })).setMimeType(ContentService.MimeType.JSON);
}

// Devuelve los items de una venta específica desde VENTAS_DETALLE
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

// Actualiza el método de pago y estado de una venta existente (flujo COBRAR cajero)
function cobrarVentaExistente(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("VENTAS_CABECERA");
  if (!sheet) return generarRespuestaError("VENTAS_CABECERA no encontrada");

  var filas = sheet.getDataRange().getValues();
  for (var i = 1; i < filas.length; i++) {
    if (String(filas[i][0]) === String(data.idVenta)) {
      sheet.getRange(i + 1, 9).setValue(data.metodo); // col 9: FormaPago = método real
      return ContentService.createTextOutput(JSON.stringify({
        status: "success", mensaje: "Venta cobrada correctamente"
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  return generarRespuestaError("Venta con ID " + data.idVenta + " no encontrada.");
}

// Anula un ticket individual enviado desde el frontend (tipoEvento='ANULACION')
function anularVentaIndividual(data) {
  if (!data.ventaId) return generarRespuestaError("No se proporcionó ventaId.");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("VENTAS_CABECERA");
  if (!sheet) return generarRespuestaError("VENTAS_CABECERA no encontrada.");

  var filas = sheet.getDataRange().getValues();
  for (var i = 1; i < filas.length; i++) {
    if (String(filas[i][0]) === String(data.ventaId)) {
      sheet.getRange(i + 1, 9).setValue('ANULADO'); // col 9: FormaPago = ANULADO
      return ContentService.createTextOutput(JSON.stringify({
        status: "success", mensaje: "Venta anulada correctamente"
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  return generarRespuestaError("Venta con ID " + data.ventaId + " no encontrada.");
}

function generarRespuestaError(msg) {
  return ContentService.createTextOutput(JSON.stringify({
    status: "error", mensaje: msg
  })).setMimeType(ContentService.MimeType.JSON);
}

// ------ STOCK: GUÍAS Y AUDITORÍA ------ //

function cajeroActivo(zona) {
  if (!zona) return generarRespuestaError("zona requerida");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("CAJAS");
  if (!sheet) return generarRespuestaError("CAJAS no encontrada");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    // Col 6 (índice 5) = Estado, Col 9 (índice 8) = Zona_ID
    if (String(data[i][5]) === 'ABIERTA' && String(data[i][8] || '') === zona) {
      return ContentService.createTextOutput(JSON.stringify({
        status: 'success', activo: true,
        vendedor: String(data[i][1]), idCaja: String(data[i][0]), desde: data[i][3]
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ status: 'success', activo: false }))
    .setMimeType(ContentService.MimeType.JSON);
}

function listarGuias(zona) {
  if (!zona) return generarRespuestaError("zona requerida");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("GUIAS_CABECERA");
  if (!sheet) return ContentService.createTextOutput(JSON.stringify({ status: 'success', guias: [] })).setMimeType(ContentService.MimeType.JSON);
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    // Incluir guías donde Zona_ID === zona O Zona_Destino === zona (traslados entrantes)
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
  if (!sheet) return ContentService.createTextOutput(JSON.stringify({ status: 'success', traslados: [] })).setMimeType(ContentService.MimeType.JSON);
  var data = sheet.getDataRange().getValues();
  // Default: últimas 24 horas si no se pasa timestamp
  var desdeDate = desde ? new Date(parseInt(desde)) : new Date(Date.now() - 86400000);
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][4]) !== 'ENTRADA_TRASLADO') continue;
    if (String(data[i][3]) !== zona) continue;
    var fechaGuia = new Date(data[i][1]);
    if (fechaGuia > desdeDate) {
      result.push({
        id_guia:   String(data[i][0]),
        fecha:     data[i][1],
        origen:    String(data[i][6] || ''),
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
  if (!sheet) return ContentService.createTextOutput(JSON.stringify({ status: 'success', stock: [] })).setMimeType(ContentService.MimeType.JSON);
  var data = obtenerDatosHojaComoJSON(sheet);
  return ContentService.createTextOutput(JSON.stringify({ status: 'success', stock: data })).setMimeType(ContentService.MimeType.JSON);
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
  // No existe — crear fila nueva
  var cantInicial = Math.max(0, delta);
  sheet.appendRow([String(codBarras), String(zonaId), cantInicial]);
  return cantInicial;
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
  // GUIAS_CABECERA: ID_Guia | Fecha | Vendedor | Zona_ID | Tipo | Observacion | Zona_Destino | Estado
  sheetCab.appendRow([idGuia, new Date(), data.vendedor, data.zona, tipo, data.observacion || '', zonaDestino, 'CONFIRMADO']);

  var stockResult = [];
  (data.items || []).forEach(function(item) {
    var cb = String(item.cod_barras);
    sheetDet.appendRow([idGuia, cb, item.cantidad]);
    var nuevaCant = actualizarStockFila(sheetStock, cb, data.zona, signo * item.cantidad);
    stockResult.push({ cod_barras: cb, cantidad: nuevaCant });
  });

  // SALIDA_MOVIMIENTO → ENTRADA_TRASLADO automática en zona destino
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
    // Actualizar stock a la cantidad real contada (diff puede ser + o -)
    actualizarStockFila(sheetStock, cb, data.zona, diff);
  });

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success', idAuditoria: idAudit
  })).setMimeType(ContentService.MimeType.JSON);
}

function verificarDispositivo(deviceId) {
  if (!deviceId) return generarRespuestaError("ID de dispositivo no proporcionado");

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("DISPOSITIVOS");

  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "success", autorizado: false, mensaje: "Tabla DISPOSITIVOS no encontrada"
    })).setMimeType(ContentService.MimeType.JSON);
  }

  var data = obtenerDatosHojaComoJSON(sheet);
  var autorizado = false;

  for (var i = 0; i < data.length; i++) {
    if (data[i].ID_Dispositivo === deviceId && data[i].Estado === 'ACTIVO') {
      autorizado = true;
      break;
    }
  }

  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    autorizado: autorizado
  })).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// CONSULTA DNI/RUC — vía APISPeru
// Requiere propiedad de script: APISPERU_TOKEN
// Registro en: https://apisperu.com/servicios/dniruc/
// ============================================================
function consultarCliente(doc) {
  if (!doc) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'Documento requerido' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  doc = String(doc).trim();

  // 1. Buscar primero en la tabla local CLIENTES_FRECUENTES
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('CLIENTES_FRECUENTES');
  if (sheet) {
    var rows = sheet.getDataRange().getValues();
    var headers = rows[0].map(function(h) { return String(h).trim(); });
    var docIdx = headers.indexOf('Documento');
    var nomIdx = headers.indexOf('Nombre');
    if (docIdx >= 0 && nomIdx >= 0) {
      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][docIdx]).trim() === doc) {
          return ContentService.createTextOutput(JSON.stringify({
            status: 'success',
            nombre: String(rows[i][nomIdx]),
            documento: doc,
            tipo: doc.length === 11 ? 'RUC' : 'DNI',
            fuente: 'local'
          })).setMimeType(ContentService.MimeType.JSON);
        }
      }
    }
  }

  // 2. Consultar APISPeru
  var token = PropertiesService.getScriptProperties().getProperty('APISPERU_TOKEN');
  if (!token) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: 'Token no configurado. Agregar APISPERU_TOKEN en Propiedades del script.'
    })).setMimeType(ContentService.MimeType.JSON);
  }

  try {
    var tipo = doc.length === 11 ? 'ruc' : 'dni';
    var url = 'https://dniruc.apisperu.com/api/v1/' + tipo + '/' + doc + '?token=' + token;
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var json = JSON.parse(response.getContentText());

    var nombre = '';
    if (tipo === 'dni') {
      nombre = [json.nombres, json.apellidoPaterno, json.apellidoMaterno].filter(Boolean).join(' ').trim();
    } else {
      nombre = (json.razonSocial || '').trim();
    }

    if (!nombre) {
      return ContentService.createTextOutput(JSON.stringify({
        status: 'not_found',
        message: 'No se encontró información para ' + doc
      })).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      nombre: nombre,
      documento: doc,
      tipo: tipo === 'ruc' ? 'RUC' : 'DNI',
      fuente: 'api'
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: 'Error consultando API: ' + e.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
