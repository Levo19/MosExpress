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
