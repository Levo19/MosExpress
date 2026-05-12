// ============================================================
// MosExpress — AlertaEfectivo.gs
// Sistema preventivo de robos: notifica a admins+master cuando
// el efectivo en una caja supera S/500 (primera bandera) y
// después cada S/250 adicional (750, 1000, 1250, ...).
//
// Trigger: se llama desde procesarVenta (al cobrar EFECTIVO o
// MIXTO) y desde registrarExtraCaja (cualquier tipo). Si el
// efectivo baja (egreso = recojo del admin), la bandera baja
// sola — permite que cuando vuelva a subir, la alerta vuelva
// a dispararse al cruzar el siguiente threshold.
//
// Idempotencia: hoja CAJA_ALERTAS_EFECTIVO guarda la última
// bandera por idCaja. Si el cálculo da la misma bandera o
// menor, no notifica.
// ============================================================

var _CAJA_ALERTA_INICIAL    = 500;
var _CAJA_ALERTA_INCREMENTO = 250;

// Llamado público desde Ventas.gs y Caja.gs. NO bloquea el flujo
// principal si falla — la alerta es nice-to-have, la venta debe
// completar igual.
function _chequearAlertaEfectivo(idCaja) {
  try {
    if (!idCaja) return;
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // ── 1. Leer la caja (montoInicial + cajero + zona + estación) ──
    var caja = _alertaLeerCaja(ss, idCaja);
    if (!caja) return;
    // Solo alertar para cajas ABIERTAS — si ya cerró, no hay nada que recoger
    if (caja.estado !== 'ABIERTA') return;

    // ── 2. Calcular efectivo actual de la caja ──
    var efectivo = _alertaCalcularEfectivo(ss, idCaja, caja.montoInicial);

    // ── 3. Calcular bandera teórica del monto actual ──
    var banderaTeorica = _alertaBanderaPorMonto(efectivo);

    // ── 4. Comparar con bandera anterior en hoja CAJA_ALERTAS_EFECTIVO ──
    var sheetAlertas = _alertaGetSheet(ss);
    var registro = _alertaLeerRegistro(sheetAlertas, idCaja);
    var banderaAnterior = registro ? registro.bandera : 0;

    if (banderaTeorica > banderaAnterior) {
      // ── 5a. Cruzó hacia arriba — notificar ──
      var thresholdAlertado = _alertaThresholdDeBandera(banderaTeorica);
      _alertaEnviarPush(caja, efectivo, thresholdAlertado);
      _alertaGuardarRegistro(sheetAlertas, registro, idCaja, banderaTeorica, efectivo);
    } else if (banderaTeorica < banderaAnterior) {
      // ── 5b. Bajó (egreso/recojo) — solo actualiza la bandera sin notificar.
      // Próxima vez que suba al siguiente threshold sí alerta de nuevo.
      _alertaGuardarRegistro(sheetAlertas, registro, idCaja, banderaTeorica, efectivo);
    }
    // Si banderaTeorica === banderaAnterior, no hacer nada
  } catch (e) {
    Logger.log('[AlertaEfectivo] error: ' + e.message);
  }
}

// ── Bandera por monto ──────────────────────────────────────
// 0:    < 500
// 1:   >= 500
// 2:   >= 750
// 3:   >= 1000
// N:   >= 500 + (N-1) * 250
function _alertaBanderaPorMonto(monto) {
  if (monto < _CAJA_ALERTA_INICIAL) return 0;
  return 1 + Math.floor((monto - _CAJA_ALERTA_INICIAL) / _CAJA_ALERTA_INCREMENTO);
}

function _alertaThresholdDeBandera(bandera) {
  if (bandera <= 0) return 0;
  return _CAJA_ALERTA_INICIAL + (bandera - 1) * _CAJA_ALERTA_INCREMENTO;
}

// ── Lee fila de la caja desde CAJAS ──────────────────────
function _alertaLeerCaja(ss, idCaja) {
  var sheet = ss.getSheetByName('CAJAS');
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(idCaja)) {
      return {
        idCaja:       String(idCaja),
        cajero:       String(data[i][1] || ''),
        estacion:     String(data[i][2] || ''),
        montoInicial: parseFloat(data[i][4]) || 0,
        estado:       String(data[i][5] || ''),
        zona:         String(data[i][8] || '')
      };
    }
  }
  return null;
}

// ── Calcula efectivo actual: montoInicial + EFE_ventas + INGRESOS - EGRESOS ──
// Solo cuenta lo que toca caja física (no virtuales).
function _alertaCalcularEfectivo(ss, idCaja, montoInicial) {
  var total = parseFloat(montoInicial) || 0;

  // Ventas EFECTIVO o parte EFE de MIXTO
  var vSheet = ss.getSheetByName('VENTAS_CABECERA');
  if (vSheet) {
    var vd = vSheet.getDataRange().getValues();
    for (var v = 1; v < vd.length; v++) {
      if (String(vd[v][10]) !== String(idCaja)) continue;
      var estado = String(vd[v][12] || '');
      if (estado === 'ANULADO') continue;
      var metodo = String(vd[v][8] || '').toUpperCase();
      var monto  = parseFloat(vd[v][6]) || 0;
      if (metodo === 'EFECTIVO') {
        total += monto;
      } else if (metodo.indexOf('MIXTO') === 0) {
        // MIXTO|EFE:X|VIR:Y → contar solo la parte EFE
        var efeM = metodo.match(/EFE:([\d.]+)/i);
        if (efeM) total += parseFloat(efeM[1]) || 0;
      }
      // POR_COBRAR, CREDITO, VIRTUAL → no tocan caja física
    }
  }

  // Extras: INGRESO suma, EGRESO resta (solo efectivo, sin _VIRTUAL)
  var eSheet = ss.getSheetByName('MOVIMIENTOS_EXTRA');
  if (eSheet) {
    var ed = eSheet.getDataRange().getValues();
    for (var e = 1; e < ed.length; e++) {
      if (String(ed[e][1]) !== String(idCaja)) continue;
      var tipo  = String(ed[e][3] || '');
      var monto = parseFloat(ed[e][4]) || 0;
      if (tipo === 'INGRESO')      total += monto;
      else if (tipo === 'EGRESO')  total -= monto;
      // INGRESO_VIRTUAL / EGRESO_VIRTUAL no tocan caja física
    }
  }

  return Math.round(total * 100) / 100;
}

// ── Hoja CAJA_ALERTAS_EFECTIVO ──────────────────────────────
// Columnas: idCaja | bandera | montoUltimo | fechaActualizada
function _alertaGetSheet(ss) {
  var sheet = ss.getSheetByName('CAJA_ALERTAS_EFECTIVO');
  if (!sheet) {
    sheet = ss.insertSheet('CAJA_ALERTAS_EFECTIVO');
    sheet.appendRow(['idCaja', 'bandera', 'montoUltimo', 'fechaActualizada']);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  }
  return sheet;
}

function _alertaLeerRegistro(sheet, idCaja) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(idCaja)) {
      return { fila: i + 1, bandera: parseInt(data[i][1], 10) || 0,
               montoUltimo: parseFloat(data[i][2]) || 0 };
    }
  }
  return null;
}

function _alertaGuardarRegistro(sheet, registroPrev, idCaja, banderaNueva, montoNuevo) {
  var now = new Date();
  if (registroPrev) {
    sheet.getRange(registroPrev.fila, 2, 1, 3).setValues([[banderaNueva, montoNuevo, now]]);
  } else {
    sheet.appendRow([String(idCaja), banderaNueva, montoNuevo, now]);
  }
}

// ── Push a admins+master vía _notificarMOS (ya existente en Code.gs) ──
function _alertaEnviarPush(caja, efectivoActual, thresholdAlcanzado) {
  var titulo = '💰 Recoger efectivo · S/ ' + thresholdAlcanzado.toFixed(2);
  var cuerpo = caja.cajero
             + (caja.zona     ? ' · ' + caja.zona     : '')
             + (caja.estacion ? ' · ' + caja.estacion : '')
             + ' · actual: S/ ' + efectivoActual.toFixed(2);
  try { _notificarMOS(titulo, cuerpo); }
  catch(e) { Logger.log('[AlertaEfectivo] _notificarMOS falló: ' + e.message); }
}

// ── Útil para debug manual desde el editor de GAS ─────────────
// Permite forzar el check de una caja específica sin esperar un evento.
function debugChequearAlertaEfectivo(idCajaParam) {
  var idCaja = idCajaParam || 'PEGAR_ID_CAJA_AQUI';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var caja = _alertaLeerCaja(ss, idCaja);
  if (!caja) { Logger.log('Caja no encontrada: ' + idCaja); return; }
  var efe = _alertaCalcularEfectivo(ss, idCaja, caja.montoInicial);
  var bAct = _alertaBanderaPorMonto(efe);
  Logger.log('Caja=' + idCaja + ' cajero=' + caja.cajero + ' estado=' + caja.estado);
  Logger.log('Efectivo actual = S/ ' + efe.toFixed(2) + ' → bandera ' + bAct
             + ' (threshold S/ ' + _alertaThresholdDeBandera(bAct) + ')');
  _chequearAlertaEfectivo(idCaja);
}
