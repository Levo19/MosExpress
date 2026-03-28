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

    // NUEVO: proxy de impresión — la clave de PrintNode vive aquí, no en el browser
    if (data.accion === 'imprimir') {
      return procesarImpresion(data);
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

  var fechaActual = new Date();
  var idCaja = "CAJA-" + fechaActual.getTime();

  // Columnas: ID_Caja | Vendedor | Estacion | Fecha_Apertura | Monto_Inicial | Estado | Monto_Final
  sheetCajas.appendRow([
    idCaja,
    data.vendedor,
    data.estacion,
    fechaActual,
    data.montoInicial,
    "ABIERTA",
    ""
  ]);

  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    idCaja: idCaja,
    mensaje: "Caja aperturada exitosamente"
  })).setMimeType(ContentService.MimeType.JSON);
}

// CORRECCIÓN: esta función estaba llamada pero nunca definida — causaba error silencioso en cada cierre de turno
function procesarCierreCaja(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetCajas = ss.getSheetByName("CAJAS");
  if (!sheetCajas) return generarRespuestaError("Pestaña CAJAS no encontrada.");

  var filas = sheetCajas.getDataRange().getValues();
  var cajaEncontrada = false;

  for (var i = 1; i < filas.length; i++) {
    if (filas[i][0] === data.cajaId) {
      // Columna F (índice 5) = Estado, Columna G (índice 6) = Monto_Final
      sheetCajas.getRange(i + 1, 6).setValue("CERRADA");
      sheetCajas.getRange(i + 1, 7).setValue(data.montoFinal);
      sheetCajas.getRange(i + 1, 8).setValue(new Date()); // Fecha_Cierre
      cajaEncontrada = true;
      break;
    }
  }

  if (!cajaEncontrada) {
    return generarRespuestaError("Caja con ID " + data.cajaId + " no encontrada.");
  }

  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    mensaje: "Caja cerrada correctamente"
  })).setMimeType(ContentService.MimeType.JSON);
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

  sheetCabecera.appendRow([
    idVenta, fechaActual, auth.vendedor, auth.estacion,
    header.cliente.doc, header.cliente.nombre, header.total,
    header.tipoDoc + " (" + (header.metodo || 'EFECTIVO') + ")",
    correlativoFinal, pos.cajaId, auth.deviceId, "COMPLETADO"
  ]);

  items.forEach(function(item) {
    sheetDetalle.appendRow([
      idVenta, item.sku, item.nombre, item.cantidad, item.precio, item.subtotal
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

  var options = {
    method: 'post',
    headers: {
      'Authorization': 'Basic ' + Utilities.base64Encode(printNodeKey + ':')
    },
    contentType: 'application/json',
    payload: JSON.stringify({
      printerId: parseInt(data.printerId, 10),
      title: data.title || 'MOSexpress',
      contentType: 'raw_base64',
      content: data.content
    }),
    muteHttpExceptions: true
  };

  try {
    var resp = UrlFetchApp.fetch('https://api.printnode.com/printjobs', options);
    var code = resp.getResponseCode();
    if (code !== 201) {
      return generarRespuestaError("PrintNode respondió " + code + ": " + resp.getContentText());
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
  var sheetsNames = ["PRODUCTO_BASE", "PRESENTACIONES", "EQUIVALENCIAS", "PROMOCIONES", "ZONAS_CONFIG", "CLIENTES_FRECUENTES"];
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

function obtenerDatosHojaComoJSON(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  var headers = data[0];
  var result = [];

  for (var i = 1; i < data.length; i++) {
    var rowData = {};
    for (var j = 0; j < headers.length; j++) {
      rowData[headers[j]] = data[i][j];
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
    var valorSerie = String(data[i][8]); // Columna I: correlativo
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

function generarRespuestaError(msg) {
  return ContentService.createTextOutput(JSON.stringify({
    status: "error", mensaje: msg
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
