// _CutoverDiag.gs — Wrappers READ-ONLY para decidir el corte de ventas (Hoja→Supabase).
// Solo hacen VISIBLE (Logger.log) el resultado de verificaciones que ya existen; NO activan
// ni escriben NADA. Correr desde el editor de Apps Script y copiar el "Registro de ejecución".
//
// Orden sugerido (gate money-safe):
//   1) cut1_paridad()      → el número que importa: data.solo_en_sheets_count debe ser 0.
//   2) cut2_reconciliar()  → sin drift en ventas / ventas_detalle.
//   3) cut3_estado()       → estado actual del modo directo (CORRELATIVO_SOURCE / sync-off).
// Las ACTIVACIONES del corte (activarCorrelativoSupabase / instalarTriggerReconciliacionDirectas /
// activarMEVentasDirecto) NO se envuelven acá a propósito: se corren aparte, con paridad en verde.

function cut1_paridad() {
  var r = verificarParidadLectura(3);
  Logger.log('PARIDAD => ' + JSON.stringify(r));
  return r;
}

function cut2_reconciliar() {
  var r = reconciliarME();
  Logger.log('RECONCILIAR => ' + JSON.stringify(r));
  return r;
}

function cut3_estado() {
  var r = estadoMESupabase();
  Logger.log('ESTADO => ' + JSON.stringify(r));
  return r;
}

// READ-ONLY. Integridad de la Hoja de ventas: confirma la hipótesis de que
// VENTAS_CABECERA fue recortada (detalle con id_venta sin cabecera) y si la
// colisión de correlativo NV existe también en la fuente. NO escribe nada.
function cut4_hoja_integridad() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cab = ss.getSheetByName('VENTAS_CABECERA').getDataRange().getValues();
  var det = ss.getSheetByName('VENTAS_DETALLE').getDataRange().getValues();
  var H = {}; cab[0].forEach(function(h, i){ H[String(h).trim()] = i; });
  var iId   = (H['ID_Venta'] != null) ? H['ID_Venta'] : 0;
  var iCorr = (H['Correlativo'] != null) ? H['Correlativo'] : -1;

  var idCab = {}, corrCount = {};
  for (var i = 1; i < cab.length; i++) {
    var id = String(cab[i][iId] || '').trim();
    if (id) idCab[id] = 1;
    if (iCorr >= 0) {
      var corr = String(cab[i][iCorr] || '').trim();
      if (corr && corr.indexOf('undefined') < 0) corrCount[corr] = (corrCount[corr] || 0) + 1;
    }
  }
  var idDet = {};
  for (var j = 1; j < det.length; j++) { var d = String(det[j][0] || '').trim(); if (d) idDet[d] = 1; }

  var detSinCab = Object.keys(idDet).filter(function(id){ return !idCab[id]; });
  var dups = Object.keys(corrCount).filter(function(k){ return corrCount[k] > 1; });

  var r = {
    cabecera_filas: cab.length - 1,
    detalle_filas: det.length - 1,
    cabecera_ids: Object.keys(idCab).length,
    detalle_ids_distinct: Object.keys(idDet).length,
    detalle_ids_SIN_cabecera: detSinCab.length,      // >0 = cabecera recortada (confirma hipótesis)
    correlativos_DUP_en_hoja: dups.length,           // >0 = colisión también en la fuente (histórica)
    muestra_dup_hoja: dups.slice(0, 10),
    muestra_det_sin_cab: detSinCab.slice(0, 5)
  };
  Logger.log('HOJA_INTEGRIDAD => ' + JSON.stringify(r));
  return r;
}

// ───────────── ACTIVACIONES DEL CORTE (correr UNA a la vez, en orden) ─────────────
// Cada wrapper llama a la función real (en Ventas.gs / Fase2Auth.gs / MigracionME.gs) y
// muestra su resultado en el log. Correr 5a, verificar conmigo, y recién 5b, 5c.

// 5a — Flip del correlativo a Supabase (se auto-valida vs la Hoja; ABORTA si atrás; reversible).
function cut5a_flip_correlativo() {
  var r = activarCorrelativoSupabase();
  Logger.log('FLIP_CORRELATIVO => ' + JSON.stringify(r));
  return r;
}

// 5b — Backstop: instala el trigger (cada 10min) que espeja Supabase→Hoja cualquier venta directa.
// Aditivo, idempotente. Correr DESPUÉS del 5a (verificado ✅).
function cut5b_backstop() {
  var r = instalarTriggerReconciliacionDirectas();
  Logger.log('BACKSTOP => ' + JSON.stringify(r));
  return r;
}

// 5c — Sync-off ventas: mete 'ventas' a ME_SYNC_OFF_TABLAS para que el batch Hoja→Supabase
// NO pise las ventas directas. Reversible (revertirMEVentasDirecto). Correr DESPUÉS del 5b.
function cut5c_syncoff_ventas() {
  var r = activarMEVentasDirecto();
  Logger.log('SYNCOFF_VENTAS => ' + JSON.stringify(r));
  return r;
}

// 6 — READ-ONLY. Estado de las fuentes de datos (para decidir si es seguro apagar el backstop/Hoja).
// Si FUENTE_DATOS=supabase (o los Flips ya no se llaman), congelar la Hoja es seguro.
function cut6_estado_fuentes() {
  var p = PropertiesService.getScriptProperties();
  var r = {
    CORRELATIVO_SOURCE: p.getProperty('CORRELATIVO_SOURCE') || '(vacío→sheets)',
    FUENTE_DATOS:       p.getProperty('FUENTE_DATOS')       || '(vacío→sheets)',
    FUENTE_DATOS_OFF:   p.getProperty('FUENTE_DATOS_OFF')   || '(vacío)',
    ME_SYNC_OFF_TABLAS: p.getProperty('ME_SYNC_OFF_TABLAS') || '(vacío)'
  };
  Logger.log('ESTADO_FUENTES => ' + JSON.stringify(r));
  return r;
}

// 8 — Flipea a Supabase los 3 últimos endpoints forzados a Sheets (FUENTE_DATOS_OFF). Post-corte la sombra
// es la verdad (cobros/créditos ya son Supabase-directo). Reversible: desactivarUnoME('<endpoint>').
// Correr, VERIFICAR la app (cajas/cobros/créditos se ven bien), y RECIÉN cut7.
function cut8_flip_restantes() {
  reactivarUnoME('estado_cajas');
  reactivarUnoME('cobros_en_vuelo');
  reactivarUnoME('creditos_pendientes');
  var r = estadoFuenteDatosME();
  Logger.log('FLIP_RESTANTES => ' + JSON.stringify(r));
  return r;
}

// 7 — Apagar el backstop (trigger reconciliarDirectasSheets 10min) = la Hoja de VENTAS deja de actualizarse
// (queda archivo histórico). Correr SOLO tras confirmar (cut6) que nada lee la Hoja en vivo. Reversible:
// re-instalar con instalarTriggerReconciliacionDirectas().
function cut7_apagar_backstop() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'reconciliarDirectasSheets') { ScriptApp.deleteTrigger(t); n++; }
  });
  Logger.log('BACKSTOP_OFF => trigger reconciliarDirectasSheets eliminado: ' + n);
  return { ok: true, eliminados: n };
}
