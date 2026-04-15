// ============================================================
// MosExpress — Caja.gs
// Apertura/cierre de turno, cobros, anulaciones, créditos,
// movimientos extra de caja y query de cajero activo.
// ============================================================

// ── Helper: cierra automáticamente cajas ABIERTA de días anteriores ──────────
// Retorna la cantidad de cajas auto-cerradas.
function _autoCerrarCajasViejas(sheetCajas) {
  var tz  = Session.getScriptTimeZone();
  var hoy = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var filas = sheetCajas.getDataRange().getValues();
  var cerradas = 0;
  for (var c = 1; c < filas.length; c++) {
    if (String(filas[c][5]) !== 'ABIERTA') continue;
    var fApert = filas[c][3];
    if (!fApert) continue;
    var diaApert = Utilities.formatDate(
      fApert instanceof Date ? fApert : new Date(fApert), tz, 'yyyy-MM-dd'
    );
    if (diaApert < hoy) {
      sheetCajas.getRange(c + 1, 6).setValue('CERRADA_AUTO');
      sheetCajas.getRange(c + 1, 8).setValue(new Date());
      cerradas++;
    }
  }
  if (cerradas > 0) SpreadsheetApp.flush(); // garantiza que los setValue se escriban antes del próximo read
  return cerradas;
}

function procesarAperturaCaja(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetCajas = ss.getSheetByName("CAJAS");
  if (!sheetCajas) return generarRespuestaError("Pestaña CAJAS no encontrada.");

  // Auto-cerrar cajas de días anteriores y forzar escritura antes de re-leer
  var _cajasAutoCerradas = _autoCerrarCajasViejas(sheetCajas);

  // Un solo cajero activo por zona a la vez
  if (data.zona) {
    var filasActualizadas = sheetCajas.getDataRange().getValues();
    for (var i = 1; i < filasActualizadas.length; i++) {
      if (String(filasActualizadas[i][5]) === 'ABIERTA' &&
          String(filasActualizadas[i][8] || '') === String(data.zona)) {
        return generarRespuestaError(
          "Ya hay un turno activo en " + data.zona + " (cajero: " + filasActualizadas[i][1] + "). Cierra ese turno primero."
        );
      }
    }
  }

  var idCaja = "CAJA-" + new Date().getTime();
  // Columnas: ID_Caja | Vendedor | Estacion | Fecha_Apertura | Monto_Inicial | Estado | Monto_Final | Fecha_Cierre | Zona_ID
  sheetCajas.appendRow([idCaja, data.vendedor, data.estacion, new Date(), data.montoInicial || 0, "ABIERTA", "", "", data.zona || '']);

  return ContentService.createTextOutput(JSON.stringify({
    status: "success", idCaja: idCaja,
    mensaje: "Caja aperturada exitosamente",
    cajasAutoCerradas: _cajasAutoCerradas
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
      sheetCajas.getRange(i + 1, 8).setValue(new Date());
      cajaVendedor = String(filas[i][1]);
      cajaZona = String(filas[i][8] || '');
      cajaEncontrada = true;
      break;
    }
  }

  if (!cajaEncontrada) return generarRespuestaError("Caja con ID " + data.cajaId + " no encontrada.");

  // Auto-generar guía SALIDA_VENTAS (no bloquea la respuesta si falla)
  if (cajaZona) {
    try { generarGuiaSalidaVentas(ss, data.cajaId, cajaVendedor, cajaZona); }
    catch(e) { Logger.log("Error guia ventas: " + e.toString()); }
  }

  return ContentService.createTextOutput(JSON.stringify({
    status: "success", mensaje: "Caja cerrada correctamente"
  })).setMimeType(ContentService.MimeType.JSON);
}

function cajeroActivo(zona) {
  if (!zona) return generarRespuestaError("zona requerida");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("CAJAS");
  if (!sheet) return generarRespuestaError("CAJAS no encontrada");
  // Auto-cerrar cajas viejas antes de consultar (evita falso positivo de "hay cajero activo")
  var _cerradas = _autoCerrarCajasViejas(sheet);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][5]) === 'ABIERTA' && String(data[i][8] || '') === zona) {
      return ContentService.createTextOutput(JSON.stringify({
        status: 'success', activo: true,
        vendedor: String(data[i][1]), idCaja: String(data[i][0]), desde: data[i][3],
        cajasAutoCerradas: _cerradas
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  return ContentService.createTextOutput(JSON.stringify({
    status: 'success', activo: false, cajasAutoCerradas: _cerradas
  })).setMimeType(ContentService.MimeType.JSON);
}

function cobrarVentaExistente(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("VENTAS_CABECERA");
  if (!sheet) return generarRespuestaError("VENTAS_CABECERA no encontrada");

  var filas = sheet.getDataRange().getValues();
  for (var i = 1; i < filas.length; i++) {
    if (String(filas[i][0]) === String(data.idVenta)) {
      sheet.getRange(i + 1, 9).setValue(data.metodo);
      if (data.cajaId) sheet.getRange(i + 1, 11).setValue(String(data.cajaId));
      return ContentService.createTextOutput(JSON.stringify({
        status: "success", mensaje: "Venta cobrada correctamente"
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  return generarRespuestaError("Venta con ID " + data.idVenta + " no encontrada.");
}

function creditarVenta(data) {
  if (!data.idVenta) return generarRespuestaError("idVenta requerido");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("VENTAS_CABECERA");
  if (!sheet) return generarRespuestaError("VENTAS_CABECERA no encontrada");
  var filas = sheet.getDataRange().getValues();
  for (var i = 1; i < filas.length; i++) {
    if (String(filas[i][0]) === String(data.idVenta)) {
      sheet.getRange(i + 1, 9).setValue('CREDITO');
      sheet.getRange(i + 1, 15).setValue(String(data.obs || ''));
      return ContentService.createTextOutput(JSON.stringify({
        status: "success", mensaje: "Crédito registrado"
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  return generarRespuestaError("Venta " + data.idVenta + " no encontrada.");
}

function anularVentaIndividual(data) {
  if (!data.ventaId) return generarRespuestaError("No se proporcionó ventaId.");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("VENTAS_CABECERA");
  if (!sheet) return generarRespuestaError("VENTAS_CABECERA no encontrada.");
  var filas = sheet.getDataRange().getValues();
  for (var i = 1; i < filas.length; i++) {
    if (String(filas[i][0]) === String(data.ventaId)) {
      sheet.getRange(i + 1, 9).setValue('ANULADO');
      return ContentService.createTextOutput(JSON.stringify({
        status: "success", mensaje: "Venta anulada correctamente"
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  return generarRespuestaError("Venta con ID " + data.ventaId + " no encontrada.");
}

// Anula en masa todos los tickets POR_COBRAR no cobrados al cierre del turno
function anulacionMasiva(data) {
  if (!data.ids || !data.ids.length) return generarRespuestaError("No se enviaron IDs a anular.");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("VENTAS_CABECERA");
  if (!sheet) return generarRespuestaError("VENTAS_CABECERA no encontrada.");

  var filas = sheet.getDataRange().getValues();
  var anulados = 0;
  for (var i = 1; i < filas.length; i++) {
    if (data.ids.indexOf(String(filas[i][0])) !== -1) {
      sheet.getRange(i + 1, 9).setValue('ANULADO');
      anulados++;
    }
  }
  return ContentService.createTextOutput(JSON.stringify({
    status: "success", anulados: anulados
  })).setMimeType(ContentService.MimeType.JSON);
}

function registrarExtraCaja(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("MOVIMIENTOS_EXTRA");
  if (!sheet) {
    sheet = ss.insertSheet("MOVIMIENTOS_EXTRA");
    sheet.appendRow(["ID_Extra","ID_Caja","Timestamp","Tipo","Monto","Concepto","Obs","Registrado_Por"]);
  }
  var id = "EX-" + new Date().getTime();
  sheet.appendRow([
    id,
    String(data.cajaId      || ''),
    new Date(),
    String(data.tipo        || 'EGRESO'),
    parseFloat(data.monto)  || 0,
    String(data.concepto    || ''),
    String(data.obs         || ''),
    String(data.registradoPor || '')
  ]);
  return ContentService.createTextOutput(JSON.stringify({
    status: "success", idExtra: id
  })).setMimeType(ContentService.MimeType.JSON);
}

function getExtrasCaja(cajaId) {
  if (!cajaId) return generarRespuestaError("cajaId requerido");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("MOVIMIENTOS_EXTRA");
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({ status: "success", extras: [] }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(cajaId)) {
      result.push({
        id:            String(data[i][0] || ''),
        tipo:          String(data[i][3] || 'EGRESO'),
        monto:         parseFloat(data[i][4]) || 0,
        concepto:      String(data[i][5] || ''),
        obs:           String(data[i][6] || ''),
        registradoPor: String(data[i][7] || '')
      });
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ status: "success", extras: result }))
    .setMimeType(ContentService.MimeType.JSON);
}
