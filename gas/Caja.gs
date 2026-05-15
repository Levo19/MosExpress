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
      sheetCajas.getRange(c + 1, 8).setValue(Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss'));
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

  // Asegurar que la columna 'PrintNode_ID' existe (col 10) — auto-creación idempotente
  // para que warehouseMos pueda leer a qué impresora mandar avisos de preingreso.
  try {
    var lastCol = sheetCajas.getLastColumn();
    var headers = sheetCajas.getRange(1, 1, 1, Math.max(lastCol, 1)).getValues()[0];
    var hasPrintNode = headers.some(function(h) { return String(h).trim() === 'PrintNode_ID'; });
    if (!hasPrintNode) sheetCajas.getRange(1, lastCol + 1).setValue('PrintNode_ID');
  } catch(e) { /* no-fatal */ }

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
  var _tz    = Session.getScriptTimeZone();
  var _ahora = Utilities.formatDate(new Date(), _tz, 'yyyy-MM-dd HH:mm:ss');
  // Columnas: ID_Caja | Vendedor | Estacion | Fecha_Apertura | Monto_Inicial | Estado | Monto_Final | Fecha_Cierre | Zona_ID | PrintNode_ID
  sheetCajas.appendRow([
    idCaja, data.vendedor, data.estacion, _ahora,
    data.montoInicial || 0, "ABIERTA", "", "", data.zona || '',
    data.printNodeId || ''   // NUEVA: ID PrintNode de la impresora asignada a esta caja
  ]);
  SpreadsheetApp.flush(); // garantiza que appendRow llegue a Sheets antes de retornar el ID al frontend

  // Notificar a admins/master en MOS — solo a ellos, no al cajero mismo
  try {
    var horaStr = Utilities.formatDate(new Date(), _tz, 'HH:mm');
    _notificarMOS(
      '🛒 ' + (data.vendedor || 'Cajero') + ' aperturó caja',
      (data.estacion || data.zona || '') + ' · ' + horaStr,
      data.vendedor || null,
      'ME_CAJA_APERTURA'
    );
  } catch(eP) { Logger.log('Push apertura caja: ' + eP.message); }

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
  var yaCerrada = false;

  for (var i = 1; i < filas.length; i++) {
    if (String(filas[i][0]) === String(data.cajaId)) {
      // Idempotencia: si ya está CERRADA, devolver éxito inmediato sin reprocesar.
      // Permite que el frontend reintente tras timeout sin duplicar guía/stock.
      if (String(filas[i][5]) === 'CERRADA') {
        yaCerrada = true;
        cajaEncontrada = true;
        break;
      }
      sheetCajas.getRange(i + 1, 6).setValue("CERRADA");
      sheetCajas.getRange(i + 1, 7).setValue(data.montoFinal);
      sheetCajas.getRange(i + 1, 8).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'));
      // Flush garantiza que el "CERRADA" se commitee antes de generar la guía.
      // Si la guía/push tardan o fallan, la caja igual queda marcada como cerrada.
      SpreadsheetApp.flush();
      cajaVendedor = String(filas[i][1]);
      cajaZona = String(filas[i][8] || '');
      cajaEncontrada = true;
      break;
    }
  }

  if (!cajaEncontrada) return generarRespuestaError("Caja con ID " + data.cajaId + " no encontrada.");

  // Idempotente: respuesta inmediata si ya estaba cerrada (no reprocesar guía/stock)
  if (yaCerrada) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "success", mensaje: "Caja ya estaba cerrada", yaCerrada: true
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // Auto-generar guía SALIDA_VENTAS (no bloquea la respuesta si falla)
  if (cajaZona) {
    try { generarGuiaSalidaVentas(ss, data.cajaId, cajaVendedor, cajaZona); }
    catch(e) { Logger.log("Error guia ventas: " + e.toString()); }
  }

  // Push a MOS — cierre de caja
  try {
    var hora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm');
    _notificarMOS('🔐 Caja cerrada · ' + hora, cajaVendedor + (cajaZona ? ' · ' + cajaZona : '') + ' · S/ ' + (parseFloat(data.montoFinal) || 0).toFixed(2), cajaVendedor, 'ME_CAJA_CIERRE');
  } catch(eP) { Logger.log('Push cierre: ' + eP.message); }

  return ContentService.createTextOutput(JSON.stringify({
    status: "success", mensaje: "Caja cerrada correctamente"
  })).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// CIERRE_CAJA_FORZADO — admin/master cierra caja remotamente desde MOS
// ============================================================
// Bug original que motiva este endpoint: el flow normal de cierre hace
// 2 POST independientes (ANULACION_MASIVA + CIERRE_CAJA) sin LockService,
// dejando estados inconsistentes si el segundo falla (POR_COBRAR anulados
// pero caja sigue ABIERTA, como pasó con CAJA-1778848407996 el 2026-05-15).
//
// Este endpoint hace TODO en una sola transacción protegida por LockService:
//   1. Anular POR_COBRAR de esa caja (FormaPago → ANULADO)
//   2. Calcular montoFinal automáticamente (apertura + efectivo + ingresos - egresos)
//   3. Marcar caja CERRADA + montoFinal + fechaCierre
//   4. Auditoría en VENTAS_CABECERA.historialCambios + CAJA log
//   5. Push al cajero original: "tu caja fue cerrada por admin X"
//   6. Push a MOS confirmando
function cerrarCajaForzado(data) {
  if (!data || !data.idCaja) return generarRespuestaError('idCaja requerido');
  var adminAuth = data.adminAuth || {};
  var adminNombre = String(adminAuth.nombre || adminAuth.validadoPor || 'admin');

  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch(e) { return generarRespuestaError('LOCK_TIMEOUT: otra operación en curso, reintentá en unos segundos'); }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetCajas = ss.getSheetByName('CAJAS');
    var sheetVentas = ss.getSheetByName('VENTAS_CABECERA');
    var sheetExtra = ss.getSheetByName('MOVIMIENTOS_EXTRA');
    if (!sheetCajas)  return generarRespuestaError('CAJAS no encontrada');
    if (!sheetVentas) return generarRespuestaError('VENTAS_CABECERA no encontrada');

    // ── 1. Localizar la caja ──────────────────────────────────────
    var filasCajas = sheetCajas.getDataRange().getValues();
    var filaCaja = -1;
    var cajaRow = null;
    for (var i = 1; i < filasCajas.length; i++) {
      if (String(filasCajas[i][0]) === String(data.idCaja)) {
        filaCaja = i;
        cajaRow = filasCajas[i];
        break;
      }
    }
    if (filaCaja < 0) return generarRespuestaError('Caja ' + data.idCaja + ' no encontrada');
    var estadoActual = String(cajaRow[5] || '');
    if (estadoActual !== 'ABIERTA') {
      return ContentService.createTextOutput(JSON.stringify({
        status: 'success', yaCerrada: true,
        mensaje: 'La caja ya estaba ' + estadoActual
      })).setMimeType(ContentService.MimeType.JSON);
    }
    var cajaVendedor = String(cajaRow[1] || '');
    var cajaEstacion = String(cajaRow[2] || '');
    var cajaZona     = String(cajaRow[8] || '');
    var montoInicial = parseFloat(cajaRow[4]) || 0;

    // ── 2. Anular POR_COBRAR de esta caja ────────────────────────
    var filasV = sheetVentas.getDataRange().getValues();
    var idsAnulados = [];
    var efectivoVentas = 0;
    for (var v = 1; v < filasV.length; v++) {
      var idCajaV = String(filasV[v][10] || '');
      if (idCajaV !== String(data.idCaja)) continue;
      var formaPago = String(filasV[v][8] || '').toUpperCase();
      var total = parseFloat(filasV[v][6]) || 0;
      if (formaPago === 'POR_COBRAR') {
        sheetVentas.getRange(v + 1, 9).setValue('ANULADO');
        idsAnulados.push(String(filasV[v][0]));
      } else if (formaPago === 'EFECTIVO' || formaPago.indexOf('MIXTO') === 0) {
        // Sumar efectivo neto de la venta. Para MIXTO el formato es
        // "MIXTO (VIR:X.XX/EFE:Y.YY)" — extraemos la parte EFE.
        if (formaPago === 'EFECTIVO') {
          efectivoVentas += total;
        } else {
          var m = formaPago.match(/EFE:([\d.]+)/);
          if (m) efectivoVentas += parseFloat(m[1]) || 0;
        }
      }
    }

    // ── 3. Sumar ingresos/egresos extra de la caja ───────────────
    var ingresosEfe = 0, egresosEfe = 0;
    if (sheetExtra) {
      var filasE = sheetExtra.getDataRange().getValues();
      for (var x = 1; x < filasE.length; x++) {
        if (String(filasE[x][1] || '') !== String(data.idCaja)) continue;
        var tipoE = String(filasE[x][3] || '');
        var mtoE  = parseFloat(filasE[x][4]) || 0;
        if      (tipoE === 'INGRESO') ingresosEfe += mtoE;
        else if (tipoE === 'EGRESO')  egresosEfe  += mtoE;
      }
    }

    // ── 4. Calcular montoFinal automático ────────────────────────
    var montoFinal = Math.round((montoInicial + efectivoVentas + ingresosEfe - egresosEfe) * 100) / 100;

    // ── 5. Cerrar la caja (CERRADA + montoFinal + fechaCierre) ──
    var tz = Session.getScriptTimeZone();
    var fechaCierre = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
    sheetCajas.getRange(filaCaja + 1, 6).setValue('CERRADA');
    sheetCajas.getRange(filaCaja + 1, 7).setValue(montoFinal);
    sheetCajas.getRange(filaCaja + 1, 8).setValue(fechaCierre);
    SpreadsheetApp.flush();

    // ── 6. Auditoría ─────────────────────────────────────────────
    try {
      auditarLog && auditarLog('CAJAS', data.idCaja, {
        usuario: adminNombre, rol: String(adminAuth.rol || 'ADMIN'),
        source: 'MOS_CIERRE_FORZADO', accion: 'cerrar_caja_forzado',
        autorizadoPor: adminAuth,
        cambios: [
          { campo: 'Estado',      antes: 'ABIERTA', despues: 'CERRADA' },
          { campo: 'Monto_Final', antes: '',         despues: montoFinal }
        ],
        ref: {
          idCaja: data.idCaja, vendedor: cajaVendedor,
          zona: cajaZona, idsAnulados: idsAnulados.length,
          montoFinal: montoFinal
        },
        motivo: String(data.motivo || 'Cierre forzado por admin desde MOS'),
        ts: new Date().toISOString()
      });
    } catch(eA) { Logger.log('Audit cierre forzado: ' + eA.message); }

    // ── 7. Notificar al cajero original (push directo) ───────────
    try {
      if (cajaVendedor) {
        enviarPushUsuario && enviarPushUsuario(
          cajaVendedor,
          '🔐 Tu caja fue cerrada por admin',
          adminNombre + ' cerró tu turno · Monto final S/ ' + montoFinal.toFixed(2),
          { idNotif: 'ME_CAJA_CERRADA_POR_ADMIN', idCaja: data.idCaja }
        );
      }
    } catch(eU) { Logger.log('Push cajero: ' + eU.message); }

    // ── 8. Push a MOS confirmando ────────────────────────────────
    try {
      _notificarMOS(
        '🔐 Cierre forzado · ' + cajaVendedor,
        adminNombre + ' · S/ ' + montoFinal.toFixed(2) + (idsAnulados.length ? ' · ' + idsAnulados.length + ' anulados' : ''),
        cajaVendedor, 'ME_CAJA_CIERRE_FORZADO'
      );
    } catch(eM) { Logger.log('Push MOS cierre forzado: ' + eM.message); }

    return ContentService.createTextOutput(JSON.stringify({
      status:      'success',
      idCaja:      data.idCaja,
      vendedor:    cajaVendedor,
      zona:        cajaZona,
      montoInicial: montoInicial,
      efectivoVentas: Math.round(efectivoVentas * 100) / 100,
      ingresos:    Math.round(ingresosEfe * 100) / 100,
      egresos:     Math.round(egresosEfe * 100) / 100,
      montoFinal:  montoFinal,
      anulados:    idsAnulados.length,
      idsAnulados: idsAnulados,
      fechaCierre: fechaCierre,
      cerradoPor:  adminNombre,
      mensaje:     'Caja cerrada forzadamente por admin'
    })).setMimeType(ContentService.MimeType.JSON);

  } finally {
    try { lock.releaseLock(); } catch(_){}
  }
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
  for (var i = filas.length - 1; i > 0; i--) {  // buscar desde el final (más probable)
    if (String(filas[i][0]) === String(data.idVenta)) {
      var formaAnt = String(filas[i][8] || '');
      var cajaAnt  = String(filas[i][10] || '');
      sheet.getRange(i + 1, 9).setValue(data.metodo);
      if (data.cajaId) sheet.getRange(i + 1, 11).setValue(String(data.cajaId));

      // Log de auditoría
      try {
        var actor = _audExtraerActor(data);
        var cambios = [{ campo:'FormaPago', antes: formaAnt, despues: String(data.metodo) }];
        if (data.cajaId && data.cajaId !== cajaAnt) {
          cambios.push({ campo:'ID_Caja', antes: cajaAnt, despues: String(data.cajaId) });
        }
        auditarLog('VENTAS_CABECERA', data.idVenta, {
          usuario: actor.usuario, rol: actor.rol,
          source: 'ME_COBRAR_VENTA',
          accion: 'cobrar_venta',
          cambios: cambios,
          autorizadoPor: actor.autorizadoPor || null,
          motivo: data.motivo || ''
        });
      } catch(_){}

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
  for (var i = filas.length - 1; i > 0; i--) {
    if (String(filas[i][0]) === String(data.idVenta)) {
      var formaAnt = String(filas[i][8] || '');
      var obsAnt   = String(filas[i][14] || '');
      sheet.getRange(i + 1, 9).setValue('CREDITO');
      sheet.getRange(i + 1, 15).setValue(String(data.obs || ''));

      try {
        var actor = _audExtraerActor(data);
        auditarLog('VENTAS_CABECERA', data.idVenta, {
          usuario: actor.usuario, rol: actor.rol,
          source: 'ME_CREDITAR_VENTA',
          accion: 'convertir_a_credito',
          cambios: [
            { campo:'FormaPago', antes: formaAnt, despues:'CREDITO' },
            { campo:'Obs',       antes: obsAnt,   despues: String(data.obs || '') }
          ],
          autorizadoPor: actor.autorizadoPor || null,
          motivo: data.motivo || data.obs || ''
        });
      } catch(_){}

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
  for (var i = filas.length - 1; i > 0; i--) {
    if (String(filas[i][0]) === String(data.ventaId)) {
      var formaAnt = String(filas[i][8] || '');
      sheet.getRange(i + 1, 9).setValue('ANULADO');

      try {
        var actor = _audExtraerActor(data);
        auditarLog('VENTAS_CABECERA', data.ventaId, {
          usuario: actor.usuario, rol: actor.rol,
          source: 'ME_ANULAR_VENTA',
          accion: 'anular_venta_interna',
          cambios: [{ campo:'FormaPago', antes: formaAnt, despues:'ANULADO' }],
          autorizadoPor: actor.autorizadoPor || null,
          motivo: data.motivo || ''
        });
      } catch(_){}

      // Avisar a WH que descuente del pickup origen (si existe y no está cerrado).
      // No bloquea ni rompe la anulación si falla.
      try { notificarAnulacionPickupAWH(data.ventaId); } catch(_){}

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
  var idsAnulados = [];
  for (var i = 1; i < filas.length; i++) {
    if (data.ids.indexOf(String(filas[i][0])) !== -1) {
      sheet.getRange(i + 1, 9).setValue('ANULADO');
      anulados++;
      idsAnulados.push(String(filas[i][0]));
    }
  }
  // Notificar WH para descontar de pickups origen (no bloquea)
  try { idsAnulados.forEach(function(id){ notificarAnulacionPickupAWH(id); }); } catch(_){}
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

  // Alerta de recojo de efectivo: tras INGRESO sube monto → posible cruce de
  // threshold (alerta). Tras EGRESO baja monto → bandera se reajusta sola
  // sin enviar push. Virtuales no cuentan (no tocan caja física).
  try {
    if (data.cajaId) _chequearAlertaEfectivo(data.cajaId);
  } catch(eA) { Logger.log('Alerta efectivo (extra): ' + eA.message); }

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
