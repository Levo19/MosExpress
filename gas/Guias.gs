// ============================================================
// MosExpress — Guias.gs
// Guías de stock por zona, auditorías físicas y traslados.
// ============================================================

// ════════════════════════════════════════════════════════════════════════
// CUTOVER NÚCLEO DE STOCK → SUPABASE (escritura directa) — gate + helpers.
// ────────────────────────────────────────────────────────────────────────
// Flag Script Property ME_ESCRITURA_STOCK_DIRECTA: '1'/'true'/'on' → ON.
//   OFF (default) → comportamiento IDÉNTICO a hoy (read-modify-write en la Hoja).
//   ON            → la mutación de SALDO va por RPC ATÓMICA a me.stock_zonas + kardex.
//                   La Hoja se sigue escribiendo (fuente de respaldo / lo que lee el sync HASTA
//                   que se apague ME_SYNC_OFF_TABLAS); si la RPC FALLA, la Hoja es el fallback.
// Idempotencia: descuento de venta por id_caja (clave de kardex); ajuste por localId; guía por idGuia+cod.
// Revertir: poner el flag OFF (vuelve a la Hoja) — sin redeploy si solo cambia la Property.
// ════════════════════════════════════════════════════════════════════════
function _meStockDirecto() {
  try {
    var v = String(PropertiesService.getScriptProperties().getProperty('ME_ESCRITURA_STOCK_DIRECTA') || '').toLowerCase();
    return v === '1' || v === 'true' || v === 'on' || v === 'si';
  } catch (e) { return false; }
}

// ════════════════════════════════════════════════════════════════════════
// CUTOVER GUÍAS ME → MODELO WH (nace ABIERTA, aplica stock AL CERRAR) — gate + helpers.
// ────────────────────────────────────────────────────────────────────────
// Flag mos.config ME_GUIAS_CICLO_ABIERTA (espejo en Script Property opcional para no pegar a la BD en cada guía):
//   ON  → las guías MANUALES (ENTRADA_ALMACEN/ENTRADA_LIBRE/SALIDA_MOVIMIENTO/SALIDA_JEFA/SALIDA_DEVOLUCION_WH)
//         nacen 'ABIERTA' (solo metadata, cantidad_aplicada=0). NO se aplica stock al crear (NO se llama
//         me.zona_registrar_guia). El stock lo aplica me.cerrar_guia_zona_idempotente AL CERRAR (idempotente,
//         delta=cantidad−aplicada, con espejo de traslado OUT origen + IN destino).
//   OFF (DEFAULT) → comportamiento legacy: nacen 'CONFIRMADO' y aplican stock al crear (intacto).
// SALIDA_VENTAS NUNCA cambia (es automática al cerrar caja, descuenta por ticket vía zona_descontar_venta).
// REVERSIBLE: poner el flag OFF vuelve al modelo viejo sin redeploy (las guías ABIERTA en vuelo se cierran solas
//   por el cron / botón; las nuevas nacerán CONFIRMADO de nuevo). Requiere ME_ESCRITURA_STOCK_DIRECTA=ON
//   (sin escritura directa el modelo ABIERTA no aplica stock en ningún lado → se ignora el flag por seguridad).
// La Property local ME_GUIAS_CICLO_ABIERTA (1/0) tiene prioridad; si no está, se consulta mos.config (cache 60s).
// ════════════════════════════════════════════════════════════════════════
function _meGuiasCicloAbierta() {
  // Money-safety: el modelo ABIERTA SOLO tiene sentido con escritura directa ON (el cierre aplica a Supabase).
  if (!_meStockDirecto()) return false;
  try {
    var prop = String(PropertiesService.getScriptProperties().getProperty('ME_GUIAS_CICLO_ABIERTA') || '').toLowerCase();
    if (prop === '1' || prop === 'true' || prop === 'on' || prop === 'si') return true;
    if (prop === '0' || prop === 'false' || prop === 'off' || prop === 'no') return false;
  } catch (e) {}
  // fallback: leer mos.config (cacheado 60s para no pegar a la BD por cada guía)
  try {
    var cache = CacheService.getScriptCache();
    var hit = cache.get('ME_GUIAS_CICLO_ABIERTA');
    if (hit != null) return hit === '1';
    var r = _sbSelect('mos.config', { select: 'valor', filters: { clave: 'eq.ME_GUIAS_CICLO_ABIERTA' }, limit: 1 });
    var on = '0';
    if (r && r.ok && r.data && r.data.length) {
      var v = String(r.data[0].valor || '').toLowerCase();
      on = (v === '1' || v === 'true' || v === 'on' || v === 'si') ? '1' : '0';
    }
    try { cache.put('ME_GUIAS_CICLO_ABIERTA', on, 60); } catch (eC) {}
    return on === '1';
  } catch (e) { return false; }  // ante cualquier duda → modelo viejo (seguro)
}

// Tipos de guía MANUAL que entran al ciclo ABIERTA. SALIDA_VENTAS y ENTRADA_TRASLADO (espejo) quedan FUERA.
function _esGuiaManualCicloAbierta(tipo) {
  var t = String(tipo || '').toUpperCase();
  return t === 'ENTRADA_ALMACEN' || t === 'ENTRADA_LIBRE' ||
         t === 'SALIDA_MOVIMIENTO' || t === 'SALIDA_JEFA' || t === 'SALIDA_DEVOLUCION_WH';
}

// Cierra una guía de zona vía RPC idempotente (aplica stock UNA vez: OUT origen + IN espejo si traslado).
//   La RPC me.cerrar_guia_zona (wrapper, param `p`) → me.cerrar_guia_zona_idempotente. Idempotente: recerrar = delta 0.
//   Devuelve {ok, ...} de la RPC, o {ok:false, error} si falla. NUNCA lanza.
function _meCerrarGuiaDirecto(idGuia) {
  var id = String(idGuia || '').trim();
  if (!id) return { ok: false, error: 'sin idGuia' };
  try {
    // cerrar_guia_zona NO empieza con 'zona_' → _sbRpc envía el body tal cual; mandamos {p:{idGuia}} (param `p`).
    var r = _sbRpc('me', 'cerrar_guia_zona', { p: { idGuia: id } });
    if (!r.ok || !(r.data && r.data.ok)) {
      Logger.log('[cerrar-guia] FALLÓ ' + id + ' HTTP ' + r.code + ' ' + (r.error || JSON.stringify(r.data || {})));
      return { ok: false, error: r.error || (r.data && r.data.error) || 'rpc' };
    }
    return r.data;
  } catch (e) {
    Logger.log('[cerrar-guia] EXCEPCIÓN ' + id + ': ' + e.message);
    return { ok: false, error: String(e.message) };
  }
}

// ════════════════════════════════════════════════════════════════════════
// LECTURA del cierre desde Supabase (delete-safe del Sheet).
// ────────────────────────────────────────────────────────────────────────
// Gate Script Property ME_LECTURA_CIERRE_DIRECTA:
//   '1'/'true'/'on' (DEFAULT ON) → el cierre y la guía SALIDA_VENTAS leen los
//        datos de la caja (ventas/efectivo/POR_COBRAR/totales por cod/guía existente)
//        desde me.cierre_datos_caja (RPC LECTURA). El Sheet ya NO se lee.
//   '0'/'off'/'no'  → comportamiento legacy: lee del Sheet (rollback instantáneo).
// Ante CUALQUIER fallo de la RPC → fallback automático al Sheet (si existe). Si el
// Sheet ya no existe y la RPC falla, se loggea y se sigue (no rompe el cierre/venta).
// MONEY-SAFETY: es 100% LECTURA — el descuento sigue por zona_descontar_venta
// (idempotente por id_caja) y el cierre por _dualWriteCajaME/Sheet. NO doble-cuenta.
// ════════════════════════════════════════════════════════════════════════
function _meLecturaCierreDirecta() {
  try {
    var v = String(PropertiesService.getScriptProperties().getProperty('ME_LECTURA_CIERRE_DIRECTA') || '1').toLowerCase();
    return v === '1' || v === 'true' || v === 'on' || v === 'si';
  } catch (e) { return true; }  // default ON: el resto de ME ya lee directo (ME_LECTURA_DIRECTA=1)
}

// Lee de me.cierre_datos_caja todo lo que el cierre/guía necesitan de la caja.
// Devuelve el objeto data de la RPC ({ok:true, ...}) o null si falla / gate OFF.
// SHAPE: { ok, id_caja, vendedor, estacion, zona, estado, monto_inicial, monto_final,
//   printnode_id, fecha_apertura, fecha_cierre, efectivo_ventas, ingresos_efe,
//   egresos_efe, ids_por_cobrar:[], totales_por_cod:{cod:cant}, guia_salida_existe }
function _meCierreDatosCaja(idCaja) {
  if (!_meLecturaCierreDirecta()) return null;
  var idc = String(idCaja || '').trim();
  if (!idc) return null;
  try {
    var r = _sbRpc('me', 'cierre_datos_caja', { p_id_caja: idc });
    if (r && r.ok && r.data && r.data.ok) return r.data;
    Logger.log('[cierre-lectura] RPC falló caja=' + idc + ' HTTP ' + (r && r.code) + ' ' + ((r && (r.error || (r.data && r.data.error))) || ''));
    return null;
  } catch (e) {
    Logger.log('[cierre-lectura] EXCEPCIÓN caja=' + idc + ': ' + (e && e.message));
    return null;
  }
}

// [delete-safe 166] Lee de me.venta_estado_lectura el estado puntual de UNA venta
// (forma_pago / id_caja / obs) para cobrar/creditar/anular sin tocar la hoja
// VENTAS_CABECERA. Gate = mismo _meLecturaCierreDirecta. Devuelve el data {ok,...}
// de la RPC, o null si gate OFF / RPC falla → el caller cae al Sheet (fallback).
// SHAPE: { ok, id_venta, forma_pago, id_caja, obs }
function _meVentaEstado(idVenta) {
  if (!_meLecturaCierreDirecta()) return null;
  var idv = String(idVenta || '').trim();
  if (!idv) return null;
  try {
    var r = _sbRpc('me', 'venta_estado_lectura', { p_id_venta: idv });
    if (r && r.ok && r.data && r.data.ok) return r.data;
    Logger.log('[venta-estado-lectura] RPC falló venta=' + idv + ' HTTP ' + (r && r.code) + ' ' + ((r && (r.error || (r.data && r.data.error))) || ''));
    return null;
  } catch (e) {
    Logger.log('[venta-estado-lectura] EXCEPCIÓN venta=' + idv + ': ' + (e && e.message));
    return null;
  }
}

// [delete-safe 166] Busca la caja ABIERTA de un dispositivo (printnode_id=deviceId)
// desde me.cajas, sin recorrer la hoja CAJAS. Gate = _meLecturaCierreDirecta.
// Devuelve el data {ok, encontrada, ...} de la RPC, o null si gate OFF / RPC falla.
// SHAPE: { ok, encontrada, id_caja, vendedor, estacion, zona, monto_inicial,
//   estado, printnode_id, fecha_apertura, zombis:[ids] }
function _meCajaAbiertaPorDevice(deviceId) {
  if (!_meLecturaCierreDirecta()) return null;
  var dev = String(deviceId || '').trim();
  if (!dev) return null;
  try {
    var r = _sbRpc('me', 'caja_abierta_por_device', { p_device: dev });
    if (r && r.ok && r.data && r.data.ok) return r.data;
    Logger.log('[caja-device-lectura] RPC falló device=' + dev + ' HTTP ' + (r && r.code) + ' ' + ((r && (r.error || (r.data && r.data.error))) || ''));
    return null;
  } catch (e) {
    Logger.log('[caja-device-lectura] EXCEPCIÓN device=' + dev + ': ' + (e && e.message));
    return null;
  }
}

// Descuento de stock por cierre de caja (venta) — RPC atómica idempotente por id_caja.
// totalesPorCod = { codBarras: cantidad, ... }. Devuelve {ok, ...} de la RPC o {ok:false} si falló.
function _meDescontarVentaDirecto(idCaja, zona, vendedor, totalesPorCod) {
  var items = Object.keys(totalesPorCod || {}).map(function (cb) {
    return { codBarra: String(cb), cantidad: parseFloat(totalesPorCod[cb]) || 0 };
  }).filter(function (it) { return it.codBarra && it.cantidad > 0; });
  if (!items.length) return { ok: true, vacio: true };
  try {
    var r = _sbRpc('me', 'zona_descontar_venta', {
      idCaja: String(idCaja), zona: String(zona), usuario: String(vendedor || ''),
      origen: 'GAS', items: items
    });
    if (!r.ok || !(r.data && r.data.ok)) {
      Logger.log('[stock-directo venta] FALLÓ caja=' + idCaja + ' HTTP ' + r.code + ' ' + (r.error || JSON.stringify(r.data || {})));
      return { ok: false, error: r.error || (r.data && r.data.error) || 'rpc' };
    }
    return r.data;
  } catch (e) {
    Logger.log('[stock-directo venta] EXCEPCIÓN caja=' + idCaja + ': ' + e.message);
    return { ok: false, error: String(e.message) };
  }
}

// Ajuste set-absoluto — RPC me.zona_ajustar_stock (set + log + kardex). Idempotente por localId.
//   origen='AUDITORIA' → la RPC etiqueta el movimiento del kardex como AUDITORIA y RE-ANCLA el saldo
//   (nuevoAbsoluto) al valor contado (fix 🔴#1). Cualquier otro origen → AJUSTE. El código va TAL CUAL
//   (sin upper) para no crear filas fantasma con códigos alfanuméricos (fix 🔴#2).
function _meAjustarStockDirecto(zona, codBarras, nuevo, usuario, localId, origen) {
  try {
    var r = _sbRpc('me', 'zona_ajustar_stock', {
      zona: String(zona), codBarra: String(codBarras), nuevo: nuevo,
      usuario: String(usuario || ''), origen: String(origen || 'GAS'), localId: localId ? String(localId) : null
    });
    if (!r.ok || !(r.data && r.data.ok)) {
      Logger.log('[stock-directo ajuste] FALLÓ ' + codBarras + '@' + zona + ' HTTP ' + r.code + ' ' + (r.error || JSON.stringify(r.data || {})));
      return { ok: false, error: r.error || (r.data && r.data.error) || 'rpc' };
    }
    return r.data;
  } catch (e) {
    Logger.log('[stock-directo ajuste] EXCEPCIÓN ' + codBarras + ': ' + e.message);
    return { ok: false, error: String(e.message) };
  }
}

// Guía manual (SALIDA_JEFA/MOVIMIENTO/ENTRADA_*) — RPC me.zona_registrar_guia (delta firmado + kardex). Idemp por idGuia+cod.
function _meRegistrarGuiaDirecto(idGuia, zona, tipo, items, usuario, idGuiaEntrada, zonaDestino) {
  var its = (items || []).map(function (it) {
    return { codBarra: String(it.cod_barras), cantidad: parseFloat(it.cantidad) || 0 };
  }).filter(function (it) { return it.codBarra && it.cantidad > 0; });
  if (!its.length) return { ok: true, vacio: true };
  try {
    var r = _sbRpc('me', 'zona_registrar_guia', {
      idGuia: String(idGuia), zona: String(zona), tipo: String(tipo), items: its,
      usuario: String(usuario || ''), origen: 'GAS',
      idGuiaEntrada: idGuiaEntrada ? String(idGuiaEntrada) : null,
      zonaDestino: zonaDestino ? String(zonaDestino) : null
    });
    if (!r.ok || !(r.data && r.data.ok)) {
      Logger.log('[stock-directo guia] FALLÓ ' + idGuia + ' HTTP ' + r.code + ' ' + (r.error || JSON.stringify(r.data || {})));
      return { ok: false, error: r.error || (r.data && r.data.error) || 'rpc' };
    }
    return r.data;
  } catch (e) {
    Logger.log('[stock-directo guia] EXCEPCIÓN ' + idGuia + ': ' + e.message);
    return { ok: false, error: String(e.message) };
  }
}

// Metadata de guía (cabecera+detalle) en Supabase — RPC me.zona_guia_registrar_meta.
// IMPORTANTE money-safety: esta RPC NO toca me.stock_zonas ni kardex (solo escribe cabecera/detalle).
// El SALDO ya lo aplican las RPCs de stock (zona_descontar_venta / zona_registrar_guia) en el mismo flujo,
// así que grabar la metadata NO re-aplica stock → SIN doble conteo. Idempotente por idGuia (reaplicar
// el mismo idGuia NO duplica cabecera ni detalle). NUNCA lanza excepción (best-effort).
//   items = [{cod_barras|codBarra, cantidad}, ...]
function _meRegistrarGuiaMetaDirecto(meta) {
  try {
    var its = (meta.items || []).map(function (it) {
      return { codBarra: String(it.codBarra || it.cod_barras || ''), cantidad: parseFloat(it.cantidad) || 0 };
    }).filter(function (it) { return it.codBarra && it.cantidad > 0; });
    var r = _sbRpc('me', 'zona_guia_registrar_meta', {
      idGuia: String(meta.idGuia), zona: String(meta.zona), tipo: String(meta.tipo),
      fecha: meta.fecha != null ? String(meta.fecha) : null,
      vendedor: meta.vendedor ? String(meta.vendedor) : null,
      observacion: meta.observacion != null ? String(meta.observacion) : null,
      zonaDestino: meta.zonaDestino ? String(meta.zonaDestino) : null,
      estado: meta.estado ? String(meta.estado) : 'CONFIRMADO',
      items: its
    });
    if (!r.ok || !(r.data && r.data.ok)) {
      Logger.log('[guia-meta] FALLÓ ' + meta.idGuia + ' HTTP ' + r.code + ' ' + (r.error || JSON.stringify(r.data || {})));
      return { ok: false, error: r.error || (r.data && r.data.error) || 'rpc' };
    }
    return r.data;
  } catch (e) {
    Logger.log('[guia-meta] EXCEPCIÓN ' + (meta && meta.idGuia) + ': ' + e.message);
    return { ok: false, error: String(e.message) };
  }
}

// ════════════════════════════════════════════════════════════════════════
// REPOSICIÓN DE STOCK AL ANULAR UNA VENTA CUYA CAJA YA CERRÓ.
// ────────────────────────────────────────────────────────────────────────
// PROBLEMA: el descuento de stock por venta ocurre UNA sola vez, al CIERRE de
// caja (generarGuiaSalidaVentas), sumando las cantidades de las ventas NO
// anuladas por Cod_Barras. Si una venta se anula DESPUÉS del cierre, su stock
// ya fue descontado y NUNCA se repone → faltante fantasma permanente.
//   · Anular ANTES del cierre → el cierre la filtra (FormaPago='ANULADO'),
//     nunca se descontó → NO hay que reponer.
//   · Anular DESPUÉS del cierre (ya existe guía SALIDA_VENTAS de su caja) →
//     hay que SUMAR de vuelta el stock de esa venta.
//
// DETECCIÓN de "caja ya cerrada": existe una guía Tipo='SALIDA_VENTAS' cuya
// Observacion contiene el ID_Caja de la venta (mismo criterio con que
// generarGuiaSalidaVentas evita duplicar). De esa guía tomamos la ZONA exacta
// con la que se descontó (col Zona_ID), no la inferimos.
//
// VÍA: me.zona_registrar_guia con tipo='ENTRADA' (signo +1 = SUMA) e
// idGuia='ANUL:<idVenta>' → kardex refId 'GUIA:ANUL:<idVenta>:<cod>', ÚNICO e
// IDEMPOTENTE (uq_me_kardex_ref). Anular 2 veces la misma venta NO repone 2x.
// Cantidades = suma por Cod_Barras de VENTAS_DETALLE de esa venta (mismas
// unidades de zona que usó el descuento del cierre).
//
// MONEY-SAFETY: gated por ME_ESCRITURA_STOCK_DIRECTA (OFF → no-op, igual que
// hoy: con la Hoja como verdad y el sync vivo, el cierre re-descuenta sin la
// venta anulada en el próximo recálculo... pero ese recálculo NO existe tras
// el cierre — por eso la reposición SOLO aplica cuando la escritura directa
// está ON, que es el modo donde el saldo NO se re-deriva). Best-effort: si la
// RPC falla, se encola en ME_STOCK_PENDIENTE (tipo 'guia') → reintento
// idempotente por el trigger existente. NUNCA lanza (no rompe la anulación).
// ════════════════════════════════════════════════════════════════════════

// [delete-safe] Lee de me.venta_reposicion_datos TODO lo que la reposición de UNA venta
// anulada necesita (id_caja, caja_cerrada, zona del descuento, totales por cod) SIN tocar
// el Sheet. Devuelve el objeto data de la RPC ({ok:true, ...}) o null si gate OFF / falla.
// SHAPE: { ok, id_venta, id_caja, caja_cerrada:bool, zona, totales_por_cod:{cod:cant}, forma_pago }
function _meVentaReposicionDatos(idVenta) {
  if (!_meLecturaCierreDirecta()) return null;
  var idv = String(idVenta || '').trim();
  if (!idv) return null;
  try {
    var r = _sbRpc('me', 'venta_reposicion_datos', { p_id_venta: idv });
    if (r && r.ok && r.data && r.data.ok) return r.data;
    Logger.log('[reposicion-lectura] RPC falló venta=' + idv + ' HTTP ' + (r && r.code) + ' ' + ((r && (r.error || (r.data && r.data.error))) || ''));
    return null;
  } catch (e) {
    Logger.log('[reposicion-lectura] EXCEPCIÓN venta=' + idv + ': ' + (e && e.message));
    return null;
  }
}

// ¿La caja de esta venta YA está cerrada? (existe guía SALIDA_VENTAS suya).
// Devuelve { cerrada:bool, zona:string } — zona = Zona_ID de la guía SALIDA_VENTAS.
// [delete-safe] Fallback Sheet. La fuente primaria (Supabase) la usa _reponerStockVentaAnulada.
function _meCajaVentaYaCerrada(ss, idCaja) {
  var out = { cerrada: false, zona: '' };
  var idc = String(idCaja || '').trim();
  if (!idc) return out;
  var sheetGC = ss.getSheetByName('GUIAS_CABECERA');
  if (!sheetGC) return out;
  // GUIAS_CABECERA: 0 ID_Guia | 1 Fecha | 2 Vendedor | 3 Zona_ID | 4 Tipo | 5 Observacion | 6 Zona_Destino | 7 Estado
  var gc = sheetGC.getDataRange().getValues();
  for (var g = 1; g < gc.length; g++) {
    if (String(gc[g][4]) === 'SALIDA_VENTAS' &&
        String(gc[g][5] || '').indexOf(idc) >= 0) {
      out.cerrada = true;
      out.zona = String(gc[g][3] || '').trim();   // zona exacta del descuento
      return out;
    }
  }
  return out;
}

// Repone (SUMA) el stock de zona de UNA venta anulada, SOLO si su caja ya cerró.
// Idempotente (idGuia 'ANUL:<idVenta>' → kardex único) + best-effort (encola si falla).
// Llamar tras marcar la venta ANULADO. NUNCA lanza. No-op si stock-directo OFF o caja no cerrada.
//
// [delete-safe] FUENTE PRIMARIA: me.venta_reposicion_datos (id_caja + caja_cerrada + zona del
// descuento + totales por cod, todo de me.ventas/ventas_detalle/guias_cabecera). Así una anulación
// TARDÍA (post-cierre) repone stock aunque el Sheet ya no exista. Si la RPC no está disponible
// (gate OFF / falla) → FALLBACK al Sheet (comportamiento legacy). El cálculo es idéntico en ambos
// caminos (cod_barras col6 || sku col1; caja cerrada = guía SALIDA_VENTAS con obs⊇idCaja; zona = su Zona_ID).
function _reponerStockVentaAnulada(ss, idVenta, usuario) {
  try {
    if (!_meStockDirecto()) return { ok: true, skip: 'stock-directo OFF' };
    var idv = String(idVenta || '').trim();
    if (!idv) return { ok: false, error: 'sin idVenta' };

    var idCaja = '', zona = '', totales = {};

    // ── FUENTE PRIMARIA: Supabase (delete-safe) ──
    var sb = _meVentaReposicionDatos(idv);
    if (sb) {
      idCaja = String(sb.id_caja || '').trim();
      if (!idCaja) return { ok: false, error: 'venta sin ID_Caja: ' + idv };
      // ¿La caja ya cerró? Si NO → el cierre todavía filtrará esta venta (nunca se descontó) → NO reponer.
      if (sb.caja_cerrada !== true) return { ok: true, skip: 'caja aún abierta — el cierre filtrará la venta', idCaja: idCaja };
      zona = String(sb.zona || '').trim();
      if (!zona) return { ok: false, error: 'guía SALIDA_VENTAS sin zona para caja ' + idCaja };
      var tpc = sb.totales_por_cod || {};
      Object.keys(tpc).forEach(function (cod) {
        var q = parseFloat(tpc[cod]) || 0;
        if (cod && q > 0) totales[String(cod)] = q;
      });
    } else {
      // ── FALLBACK: Sheet (gate OFF o RPC falló) ──
      var sheetVC = ss.getSheetByName('VENTAS_CABECERA');
      var sheetVD = ss.getSheetByName('VENTAS_DETALLE');
      if (!sheetVC || !sheetVD) return { ok: false, error: 'faltan hojas VC/VD (RPC no disponible)' };

      // 1) ID_Caja de la venta. VENTAS_CABECERA: ... 10 ID_Caja (igual que generarGuiaSalidaVentas).
      var vc = sheetVC.getDataRange().getValues();
      for (var i = vc.length - 1; i > 0; i--) {
        if (String(vc[i][0]) === idv) { idCaja = String(vc[i][10] || '').trim(); break; }
      }
      if (!idCaja) return { ok: false, error: 'venta sin ID_Caja: ' + idv };

      // 2) ¿La caja ya cerró? Si NO → el cierre todavía filtrará esta venta (nunca se descontó) → NO reponer.
      var est = _meCajaVentaYaCerrada(ss, idCaja);
      if (!est.cerrada) return { ok: true, skip: 'caja aún abierta — el cierre filtrará la venta', idCaja: idCaja };
      zona = est.zona;
      if (!zona) return { ok: false, error: 'guía SALIDA_VENTAS sin zona para caja ' + idCaja };

      // 3) Sumar cantidades por Cod_Barras (idéntico al descuento del cierre: Cod_Barras = col 6, fallback SKU col 1).
      var vd = sheetVD.getDataRange().getValues();
      for (var j = 1; j < vd.length; j++) {
        if (String(vd[j][0]) !== idv) continue;
        var cod = String(vd[j][6] || vd[j][1]).trim();
        if (!cod) continue;
        totales[cod] = (totales[cod] || 0) + (parseFloat(vd[j][3]) || 0);
      }
    }

    var cods = Object.keys(totales).filter(function (c) { return totales[c] > 0; });
    if (!cods.length) return { ok: true, vacio: true, idCaja: idCaja };

    // 4) ENTRADA idempotente (suma). idGuia='ANUL:<idVenta>' → refId 'GUIA:ANUL:<idVenta>:<cod>' único.
    var idGuiaAnul = 'ANUL:' + idv;
    var items = cods.map(function (cod) { return { cod_barras: String(cod), cantidad: totales[cod] }; });
    var r = _meRegistrarGuiaDirecto(idGuiaAnul, zona, 'ENTRADA', items, String(usuario || ''), null, null);
    if (!r || !r.ok) {
      // Money-safety: encolar para reintento IDEMPOTENTE (mismo idGuia → kardex no doblará). Reusa
      // el trigger reintentarStockPendiente (tipo 'guia' → _meRegistrarGuiaDirecto). NUNCA rompe la anulación.
      Logger.log('[reposicion-anulada] RPC falló — encolando ' + idGuiaAnul + ': ' + ((r && r.error) || 'rpc'));
      try {
        _persistirStockPendiente('guia', idGuiaAnul, {
          idGuia: idGuiaAnul, zona: zona, tipo: 'ENTRADA', items: items, usuario: String(usuario || '')
        }, (r && r.error) || 'rpc');
      } catch (eP) { Logger.log('[reposicion-anulada] encolar falló: ' + eP.message); }
      return { ok: false, error: (r && r.error) || 'rpc', encolado: true, idGuia: idGuiaAnul };
    }
    Logger.log('[reposicion-anulada] OK ' + idGuiaAnul + ' zona=' + zona + ' cods=' + cods.length);
    return { ok: true, idGuia: idGuiaAnul, zona: zona, cods: cods.length };
  } catch (e) {
    Logger.log('[reposicion-anulada] EXCEPCIÓN venta=' + idVenta + ': ' + (e && e.message));
    return { ok: false, error: String(e && e.message) };
  }
}

// ── Cola de DESCUENTOS de venta que la RPC directa NO pudo aplicar ─────────
// Money-safety: cuando ME_ESCRITURA_STOCK_DIRECTA está ON pero la RPC
// me.zona_descontar_venta FALLA, el saldo de Supabase queda SIN descontar.
// Mientras el sync Hoja→Supabase siga vivo, el fallback a la Hoja se reconcilia
// solo (≤15min). PERO una vez que ME_SYNC_OFF_TABLAS incluye stock_zonas, la Hoja
// ya no es la fuente de verdad de stock → un fallo de RPC se perdería en silencio
// (eso causó 59 desync en WH). Por eso persistimos el payload en una hoja de cola
// para REINTENTO IDEMPOTENTE. La RPC dedupea por refId 'VENTA-CAJA:<idCaja>:<codBarra>'
// en el kardex → reintentar NUNCA duplica el descuento (aunque el fallback a la Hoja
// ya lo haya escrito mientras el sync estaba vivo). NUNCA lanza excepción (best-effort).
// Cols: idCaja | zona | usuario | payload(JSON {totales}) | intentos | ultimoIntento | ultimoError | estado
function _getColaDescuentoPendiente() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('ME_DESCUENTO_PENDIENTE');
  if (!sheet) {
    sheet = ss.insertSheet('ME_DESCUENTO_PENDIENTE');
    sheet.appendRow(['idCaja','zona','usuario','payload','intentos','ultimoIntento','ultimoError','estado']);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
    sheet.getRange(1, 1, sheet.getMaxRows(), 1).setNumberFormat('@STRING@');  // idCaja como texto
  }
  return sheet;
}

// Persiste (o actualiza) el descuento de UNA caja que la RPC no pudo aplicar.
// Idempotente por idCaja: si ya está en cola PENDIENTE, solo incrementa intentos
// y refresca el error (el payload es el mismo conjunto de totales de la caja).
function _persistirDescuentoPendiente(idCaja, zona, usuario, totales, error) {
  try {
    var sheet = _getColaDescuentoPendiente();
    var data  = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(idCaja) && String(data[i][7]) === 'PENDIENTE') {
        sheet.getRange(i + 1, 5).setValue((parseInt(data[i][4]) || 0) + 1);
        sheet.getRange(i + 1, 6).setValue(new Date());
        sheet.getRange(i + 1, 7).setValue(String(error || '').slice(0, 500));
        return;
      }
    }
    sheet.appendRow([
      String(idCaja), String(zona || ''), String(usuario || ''),
      JSON.stringify({ totales: totales || {} }),
      1, new Date(), String(error || '').slice(0, 500), 'PENDIENTE'
    ]);
    Logger.log('[stock-directo] descuento PERSISTIDO en cola · caja=' + idCaja);
  } catch (e) {
    Logger.log('[stock-directo] falló persistencia cola descuento caja=' + idCaja + ': ' + e.message);
  }
}

// Trigger / manual: reintenta los descuentos PENDIENTES vía la MISMA RPC idempotente.
// Se rinde tras 10 intentos (marca ABANDONADO) para no spammear. 100% lectura/escritura
// de la cola; no toca la Hoja STOCK_ZONAS (la RPC gobierna el saldo en Supabase).
function reintentarDescuentosPendientes() {
  var sheet = _getColaDescuentoPendiente();
  if (sheet.getLastRow() < 2) return { ok: true, mensaje: 'Cola vacía' };
  var data = sheet.getDataRange().getValues();
  var reaplicados = 0, intentados = 0;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][7] || '') !== 'PENDIENTE') continue;
    var intentos = parseInt(data[i][4]) || 0;
    if (intentos >= 10) { sheet.getRange(i + 1, 8).setValue('ABANDONADO'); continue; }
    var idCaja = String(data[i][0]), zona = String(data[i][1]), usuario = String(data[i][2]);
    var totales; try { totales = (JSON.parse(data[i][3]) || {}).totales || {}; } catch (_){ continue; }
    intentados++;
    var r = _meDescontarVentaDirecto(idCaja, zona, usuario, totales);
    sheet.getRange(i + 1, 5).setValue(intentos + 1);
    sheet.getRange(i + 1, 6).setValue(new Date());
    if (r && r.ok) { sheet.getRange(i + 1, 8).setValue('APLICADO'); reaplicados++; }
    else { sheet.getRange(i + 1, 7).setValue(String((r && r.error) || 'rpc').slice(0, 500)); }
  }
  Logger.log('[stock-directo] reintento descuentos · intentados=' + intentados + ' reaplicados=' + reaplicados);
  return { ok: true, intentados: intentados, reaplicados: reaplicados };
}

// Ejecutar 1 vez desde el editor para reintento automático cada 10 min.
function setupDescuentoRetryTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'reintentarDescuentosPendientes') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('reintentarDescuentosPendientes').timeBased().everyMinutes(10).create();
  return { ok: true, mensaje: 'Trigger 10min reintentarDescuentosPendientes creado' };
}

// ── Cola de MUTACIONES de stock-zona (guías manuales + ajustes de auditoría) ──
// que la RPC directa NO pudo aplicar — análoga a ME_DESCUENTO_PENDIENTE.
// Money-safety: con ME_SYNC_OFF_TABLAS apagando stock_zonas/guias_*, el fallback a la
// Hoja YA NO se propaga a Supabase → un fallo de RPC en registrarGuia/registrarAuditoria
// se perdería en silencio (drift). Persistimos el payload EXACTO de la RPC para REINTENTO
// IDEMPOTENTE por la MISMA clave determinista que usa la RPC en su kardex:
//   - guia   → me.zona_registrar_guia, dedup por idGuia+codBarra → reintentar NO duplica.
//   - ajuste → me.zona_ajustar_stock, idempotente por localId → reintentar re-ancla al MISMO valor (SET absoluto, no delta).
// NUNCA lanza excepción (best-effort). NUNCA toca la Hoja STOCK_ZONAS (la RPC gobierna el saldo).
// Cols: tipo('guia'|'ajuste') | clave(dedup) | payload(JSON args RPC) | intentos | ultimoIntento | ultimoError | estado
function _getColaStockPendiente() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('ME_STOCK_PENDIENTE');
  if (!sheet) {
    sheet = ss.insertSheet('ME_STOCK_PENDIENTE');
    sheet.appendRow(['tipo','clave','payload','intentos','ultimoIntento','ultimoError','estado']);
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
    sheet.getRange(1, 2, sheet.getMaxRows(), 1).setNumberFormat('@STRING@');  // clave como texto
  }
  return sheet;
}

// Persiste (o actualiza) UNA mutación que la RPC no pudo aplicar.
// Idempotente por (tipo, clave): si ya está PENDIENTE, solo refresca intentos/error/payload
// (el payload se reemplaza por el más reciente — mismo conjunto de items de esa operación).
function _persistirStockPendiente(tipo, clave, payload, error) {
  try {
    var sheet = _getColaStockPendiente();
    var data  = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(tipo) && String(data[i][1]) === String(clave) && String(data[i][6]) === 'PENDIENTE') {
        sheet.getRange(i + 1, 3).setValue(JSON.stringify(payload || {}));
        sheet.getRange(i + 1, 4).setValue((parseInt(data[i][3]) || 0) + 1);
        sheet.getRange(i + 1, 5).setValue(new Date());
        sheet.getRange(i + 1, 6).setValue(String(error || '').slice(0, 500));
        return;
      }
    }
    var nextRow = sheet.getLastRow() + 1;
    sheet.getRange(nextRow, 2).setNumberFormat('@STRING@');
    sheet.appendRow([
      String(tipo), String(clave), JSON.stringify(payload || {}),
      1, new Date(), String(error || '').slice(0, 500), 'PENDIENTE'
    ]);
    Logger.log('[stock-directo] mutación ' + tipo + ' PERSISTIDA en cola · clave=' + clave);
  } catch (e) {
    Logger.log('[stock-directo] falló persistencia cola stock ' + tipo + ' clave=' + clave + ': ' + e.message);
  }
}

// Trigger / manual: reintenta las mutaciones PENDIENTES vía la MISMA RPC idempotente.
// Reusa los helpers directos (_meRegistrarGuiaDirecto / _meAjustarStockDirecto) que ya
// arman/loggean el payload — NO duplica lógica. Se rinde tras 10 intentos (ABANDONADO).
function reintentarStockPendiente() {
  var sheet = _getColaStockPendiente();
  if (sheet.getLastRow() < 2) return { ok: true, mensaje: 'Cola vacía' };
  var data = sheet.getDataRange().getValues();
  var reaplicados = 0, intentados = 0;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][6] || '') !== 'PENDIENTE') continue;
    var intentos = parseInt(data[i][3]) || 0;
    if (intentos >= 10) { sheet.getRange(i + 1, 7).setValue('ABANDONADO'); continue; }
    var tipo = String(data[i][0]);
    var p; try { p = JSON.parse(data[i][2]) || {}; } catch (_){ continue; }
    intentados++;
    var r = null;
    if (tipo === 'guia') {
      r = _meRegistrarGuiaDirecto(p.idGuia, p.zona, p.tipo, p.items, p.usuario, p.idGuiaEntrada, p.zonaDestino);
    } else if (tipo === 'ajuste') {
      r = _meAjustarStockDirecto(p.zona, p.codBarra, p.nuevo, p.usuario, p.localId, p.origen);
    } else if (tipo === 'guia_meta') {
      // METADATA ONLY (cabecera/detalle Supabase). Idempotente por idGuia → reintentar NO duplica ni toca stock.
      r = _meRegistrarGuiaMetaDirecto({
        idGuia: p.idGuia, zona: p.zona, tipo: p.tipo, vendedor: p.vendedor || p.usuario,
        observacion: p.observacion, zonaDestino: p.zonaDestino, items: p.items
      });
    } else { continue; }
    sheet.getRange(i + 1, 4).setValue(intentos + 1);
    sheet.getRange(i + 1, 5).setValue(new Date());
    if (r && r.ok) { sheet.getRange(i + 1, 7).setValue('APLICADO'); reaplicados++; }
    else { sheet.getRange(i + 1, 6).setValue(String((r && r.error) || 'rpc').slice(0, 500)); }
  }
  Logger.log('[stock-directo] reintento mutaciones · intentados=' + intentados + ' reaplicados=' + reaplicados);
  return { ok: true, intentados: intentados, reaplicados: reaplicados };
}

// Ejecutar 1 vez desde el editor para reintento automático cada 10 min.
function setupStockRetryTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'reintentarStockPendiente') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('reintentarStockPendiente').timeBased().everyMinutes(10).create();
  return { ok: true, mensaje: 'Trigger 10min reintentarStockPendiente creado' };
}

// Auto-genera una guía SALIDA_VENTAS al cerrar caja y descuenta STOCK_ZONAS
// Optimizada: lee STOCK_ZONAS una sola vez, hace updates en memoria,
// y escribe el GUIAS_DETALLE + STOCK_ZONAS modificado en batch.
// DEFENSA EN PROFUNDIDAD: chequea si ya existe una guía SALIDA_VENTAS para
// esta caja antes de generar — evita duplicación incluso si la idempotencia
// de procesarCierreCaja falla por algún motivo.
function generarGuiaSalidaVentas(ss, cajaId, vendedor, zona) {
  // [delete-safe] Las hojas son OPCIONALES: si la lectura directa está ON, los
  // datos del cierre (anti-dup + totales por cod) vienen de me.cierre_datos_caja
  // y las hojas solo se usan como espejo de respaldo (best-effort). Si el Sheet
  // ya no existe, el flujo NO se rompe.
  var sheetVC    = ss.getSheetByName("VENTAS_CABECERA");
  var sheetVD    = ss.getSheetByName("VENTAS_DETALLE");
  var sheetGC    = ss.getSheetByName("GUIAS_CABECERA");
  var sheetGD    = ss.getSheetByName("GUIAS_DETALLE");
  var sheetStock = ss.getSheetByName("STOCK_ZONAS");

  var totales = {};   // { cod_barras: cantidad }  (ventas VIVAS de la caja)
  var fuente  = '';

  // ── FUENTE PRIMARIA: Supabase (me.cierre_datos_caja) ──
  // Trae guia_salida_existe (anti-dup), totales_por_cod (descuento) sin tocar el Sheet.
  var sb = _meCierreDatosCaja(cajaId);
  if (sb) {
    fuente = 'supabase';
    // 0. DEFENSA anti-duplicado (misma semántica que el Sheet: guía SALIDA_VENTAS
    //    cuya observación contiene el id_caja).
    if (sb.guia_salida_existe === true) {
      Logger.log('generarGuiaSalidaVentas: ya existe guía SALIDA_VENTAS para caja ' + cajaId + ' (Supabase) — saltando.');
      return;
    }
    // 1+2. Totales por cod_barras de ventas VIVAS (excluye ANULADO%). Ya agregado en la RPC.
    var tpc = sb.totales_por_cod || {};
    Object.keys(tpc).forEach(function (cod) {
      var q = parseFloat(tpc[cod]) || 0;
      if (cod && q > 0) totales[String(cod)] = q;
    });
  } else if (sheetVC && sheetVD && sheetGC) {
    // ── FALLBACK: Sheet (gate OFF o RPC falló) ──
    fuente = 'sheet';
    // 0. DEFENSA: ¿ya existe guía SALIDA_VENTAS para esta caja? Si sí, abortar.
    var gcData = sheetGC.getDataRange().getValues();
    for (var g = 1; g < gcData.length; g++) {
      if (String(gcData[g][4]) === 'SALIDA_VENTAS' &&
          String(gcData[g][5] || '').indexOf(String(cajaId)) >= 0) {
        Logger.log('generarGuiaSalidaVentas: ya existe guía para caja ' + cajaId + ' (id=' + gcData[g][0] + ') — saltando.');
        return;
      }
    }
    // 1. IDs de ventas VIVAS de esta caja (NO anuladas NI convertidas).
    //    money-safety: excluimos TODA FormaPago que empiece con 'ANULADO' — cubre
    //    'ANULADO' (anulación normal) y 'ANULADO_CONVERSION' (NV→CPE). Una NV convertida
    //    NO debe descontarse aquí: su físico lo descuenta el CPE nuevo en SU cierre.
    var ventas = sheetVC.getDataRange().getValues();
    var idsVentaSet = {};
    for (var i = 1; i < ventas.length; i++) {
      if (String(ventas[i][10]) === String(cajaId) && !/^ANULADO/.test(String(ventas[i][8]))) {
        idsVentaSet[String(ventas[i][0])] = true;
      }
    }
    if (!Object.keys(idsVentaSet).length) return;
    // 2. Sumar cantidades por Cod_Barras
    var detalle = sheetVD.getDataRange().getValues();
    for (var j = 1; j < detalle.length; j++) {
      if (!idsVentaSet[String(detalle[j][0])]) continue;
      var cod = String(detalle[j][6] || detalle[j][1]).trim();
      if (!cod) continue;
      totales[cod] = (totales[cod] || 0) + (parseFloat(detalle[j][3]) || 0);
    }
  } else {
    // Ni Supabase ni Sheet disponibles → no hay forma de computar; loggear y salir sin romper.
    Logger.log('generarGuiaSalidaVentas: sin fuente de datos (RPC falló y Sheet ausente) para caja ' + cajaId + ' — abortando sin error.');
    return;
  }

  var cods = Object.keys(totales);
  if (!cods.length) return;
  Logger.log('generarGuiaSalidaVentas: caja ' + cajaId + ' fuente=' + fuente + ' cods=' + cods.length);

  // 3+4. Cabecera + detalle de guía → SHEET (best-effort, espejo de respaldo).
  //    Si el Sheet ya no existe, igual generamos idGuia (para el descuento + la meta a Supabase).
  var idGuia = "G-VENTAS-" + new Date().getTime();
  if (sheetGC && sheetGD) {
    try {
      sheetGC.appendRow([idGuia, new Date(), vendedor, zona, 'SALIDA_VENTAS',
        'Auto cierre de caja · ' + cajaId, '', 'CONFIRMADO']);
      var detalleRows = cods.map(function(cod) { return [idGuia, String(cod), totales[cod]]; });
      var startRow = sheetGD.getLastRow() + 1;
      sheetGD.getRange(startRow, 2, detalleRows.length, 1).setNumberFormat('@STRING@');
      sheetGD.getRange(startRow, 1, detalleRows.length, 3).setValues(detalleRows);
    } catch (eSheet) { Logger.log('generarGuiaSalidaVentas: escritura espejo a Sheet falló (no bloquea): ' + eSheet.message); }
  }

  // 5. Stock — DESCUENTO.
  //    [cutover] Si la escritura directa está ON, el descuento del SALDO va por RPC ATÓMICA a
  //    me.stock_zonas (idempotente por id_caja → re-cerrar la misma caja NO re-descuenta) + kardex.
  //    En ese caso NO tocamos el cantidad de la Hoja (evita el DOBLE CONTEO confirmado: el sync, mientras
  //    siga vivo, re-upsertearía la Hoja → y la Hoja ya estaría descontada por RMW = resta dos veces).
  //    Si la RPC FALLA → fallback al read-modify-write de la Hoja (comportamiento de hoy).
  var stockDirectoOK = false;
  if (_meStockDirecto()) {
    var rDir = _meDescontarVentaDirecto(cajaId, zona, vendedor, totales);
    stockDirectoOK = !!(rDir && rDir.ok);
    if (!stockDirectoOK) {
      // Money-safety: la RPC falló. NO lanzamos excepción (el cierre/venta ya está
      // persistido y NO debe romperse). NO swallow silencioso (eso causó 59 desync en WH).
      // Persistimos el descuento en cola para REINTENTO IDEMPOTENTE (refId por caja+cod en
      // el kardex). Abajo el fallback a la Hoja sigue (cubre el caso sync-aún-vivo); la cola
      // es la red de seguridad para cuando ME_SYNC_OFF_TABLAS ya apagó el sync de stock_zonas.
      Logger.log('generarGuiaSalidaVentas: RPC directo falló — encolando descuento + fallback a la Hoja para caja ' + cajaId);
      try { _persistirDescuentoPendiente(cajaId, zona, vendedor, totales, (rDir && rDir.error) || 'rpc'); } catch(ePD) { Logger.log('Encolar descuento pendiente: ' + ePD.message); }
    }
  }

  if (!stockDirectoOK && sheetStock) {
    // ── Fallback / modo legacy: read-modify-write en la Hoja (fuente de verdad cuando directo OFF/falla) ──
    // Guardado por sheetStock: si el Sheet ya no existe (delete-safe) y la RPC ya falló,
    // el descuento quedó en cola (ME_DESCUENTO_PENDIENTE) para reintento idempotente.
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

  // 5b. METADATA de la guía → Supabase (cabecera+detalle). SOLO si la escritura directa está ON.
  //    El stock ya lo aplicó zona_descontar_venta arriba → esto es METADATA ONLY (no re-aplica saldo,
  //    SIN doble conteo). La Hoja queda como espejo de seguridad (appendRow de arriba intacto).
  //    Idempotente por idGuia. Si falla → la cola ME_STOCK_PENDIENTE ('guia_meta') reintenta; NUNCA rompe el cierre.
  if (_meStockDirecto()) {
    var metaItemsV = cods.map(function(cod){ return { codBarra: String(cod), cantidad: totales[cod] }; });
    var rMetaV = _meRegistrarGuiaMetaDirecto({
      idGuia: idGuia, zona: zona, tipo: 'SALIDA_VENTAS', vendedor: vendedor,
      observacion: 'Auto cierre de caja · ' + cajaId, items: metaItemsV
    });
    if (!rMetaV || !rMetaV.ok) {
      Logger.log('generarGuiaSalidaVentas: meta Supabase falló — encolando para guía ' + idGuia);
      try { _persistirStockPendiente('guia_meta', idGuia, {
        idGuia: idGuia, zona: zona, tipo: 'SALIDA_VENTAS', vendedor: vendedor,
        observacion: 'Auto cierre de caja · ' + cajaId, items: metaItemsV
      }, (rMetaV && rMetaV.error) || 'rpc'); } catch(eMV) { Logger.log('Encolar guia_meta venta: ' + eMV.message); }
    }
  }

  // 6. Enviar pickup a WH (no bloquea — si falla solo loggea)
  try { enviarPickupAWH(ss, idGuia, cajaId, vendedor, zona, totales); }
  catch(e) { Logger.log('Pickup → WH falló: ' + e.message); }
}

// ════════════════════════════════════════════════════════════════════════
// PICKUP A WH — al cerrar caja, generar lista de reposición agrupada
// por skuBase y enviarla a WH para que el operador despache.
// ════════════════════════════════════════════════════════════════════════
function enviarPickupAWH(ss, idGuia, cajaId, vendedor, zona, totalesPorCodBarras) {
  // [100% Supabase 2026-06-24] El pickup NACE en Supabase: wh.crear_pickup_cierre_caja
  // resuelve el canónico (mos.productos) e inserta en wh.pickups → el trigger v2 lo
  // consolida en la lista acumulada de la zona. SIN leer la Hoja, SIN reenviar a WH GAS.
  // Le pasamos los totales por cod_barras que el cierre ya tiene (cero riesgo de timing).
  try {
    var _ventas = [];
    Object.keys(totalesPorCodBarras || {}).forEach(function(cod){
      var q = parseFloat(totalesPorCodBarras[cod]) || 0;
      if (q > 0) _ventas.push({ cod_barras: String(cod), cantidad: q });
    });
    if (!_ventas.length) return;
    var _r = _sbRpc('wh', 'crear_pickup_cierre_caja', { p: { id_caja: cajaId, ventas: _ventas } });
    if (!_r || !_r.ok || !_r.data || _r.data.ok === false)
      Logger.log('crear_pickup_cierre_caja FALLO: ' + JSON.stringify((_r && (_r.error || _r.data)) || 'sin respuesta'));
    else
      Logger.log('Pickup → Supabase OK · ' + cajaId + ' · ' + JSON.stringify(_r.data.data));
  } catch(eRPC) { Logger.log('enviarPickupAWH (Supabase) error: ' + eRPC.message); }
  return;

  // ─────────────────────────────────────────────────────────────────────────────
  // CÓDIGO LEGACY (Hoja MOS + forward a WH GAS) — INALCANZABLE tras el return.
  // Reemplazado por el RPC de arriba. Se conserva temporalmente por referencia.
  // ─────────────────────────────────────────────────────────────────────────────
  // El catálogo de productos NO vive en MosExpress. Vive en ProyectoMOS_DB
  // (PRODUCTOS_MASTER + EQUIVALENCIAS), accesible vía Script Property MOS_SS_ID.
  // MosExpress siempre lee desde ahí (ver Catalogo.gs:descargarCatalogo).
  var mosSsId = PropertiesService.getScriptProperties().getProperty('MOS_SS_ID') || '';
  if (!mosSsId) { Logger.log('MOS_SS_ID no configurado — skip pickup'); return; }

  var mosSS;
  try { mosSS = SpreadsheetApp.openById(mosSsId); }
  catch(e) { Logger.log('No se pudo abrir MOS_SS_ID: ' + e.message); return; }

  var sheetProds = mosSS.getSheetByName('PRODUCTOS_MASTER');
  if (!sheetProds) { Logger.log('PRODUCTOS_MASTER no existe en MOS — skip pickup'); return; }

  // 1. Indexar PRODUCTOS_MASTER + EQUIVALENCIAS.
  //    Reglas WH (acordadas con el usuario):
  //    - WH solo maneja CANÓNICOS (factorConversion=1). Las presentaciones
  //      existen en MOS solo para que ME venda packs/cajas, pero al despacho
  //      del almacén siempre se piden unidades del canónico.
  //    - Códigos válidos para escanear en WH: codigoBarra del canónico +
  //      codigoBarra de las EQUIVALENCIAS apuntando al mismo skuBase.
  //      idProducto NO es escaneable (es solo ID de fila).
  //      codigoBarra de presentaciones tampoco se acepta (factor != 1).
  //    - Si la venta ME fue de una presentación con factor F, la cantidad
  //      en el pickup se multiplica por F (ej: 2 packs de 12u → 24 unidades).
  var prodData = sheetProds.getDataRange().getValues();
  var hdrsP    = prodData[0].map(function(h){ return String(h); });
  function _findCol(targets) {
    for (var t = 0; t < targets.length; t++) {
      var idx = hdrsP.indexOf(targets[t]); if (idx >= 0) return idx;
    }
    return -1;
  }
  var iIdP    = _findCol(['idProducto', 'Id_Producto', 'ID_Producto']);
  var iSkuP   = _findCol(['skuBase', 'SKU_Base', 'sku']);
  var iCodP   = _findCol(['codigoBarra', 'Cod_Barras', 'codigo_barra']);
  var iDescP  = _findCol(['descripcion', 'Descripcion', 'nombre']);
  var iFactP  = _findCol(['factorConversion', 'Factor_Conversion', 'factor_conversion']);
  if (iCodP < 0 && iIdP < 0) { Logger.log('Columnas PRODUCTOS_MASTER incompletas — skip pickup'); return; }

  // codAFila[cod_o_idProducto] = { sku, factor, esCanonico, cod, desc }
  // Nos sirve para resolver cualquier cod (canónico o presentación) → su sku + factor.
  var codAFila = {};
  // canonicoPorSku[sku] = { cod, desc, idP }  (la fila con factor=1 del sku)
  var canonicoPorSku = {};
  // equivalentesPorSku[sku] = [codigoBarra, codigoBarra, ...]
  var equivalentesPorSku = {};

  for (var p = 1; p < prodData.length; p++) {
    var idP  = iIdP   >= 0 ? String(prodData[p][iIdP]   || '').trim() : '';
    var sku  = iSkuP  >= 0 ? String(prodData[p][iSkuP]  || '').trim() : '';
    var cod  = iCodP  >= 0 ? String(prodData[p][iCodP]  || '').trim() : '';
    var desc = iDescP >= 0 ? String(prodData[p][iDescP] || '').trim() : '';
    var fac  = iFactP >= 0 ? (parseFloat(String(prodData[p][iFactP] || '1').replace(',', '.')) || 1) : 1;
    var skuFinal = sku || idP;
    if (!skuFinal) continue;
    var esCanonico = (fac === 1);
    var fila = { sku: skuFinal, factor: fac, esCanonico: esCanonico, cod: cod, desc: desc, idP: idP };
    // Indexar por todos los identificadores posibles para resolver al recibir totales
    if (cod) codAFila[cod] = fila;
    if (idP) codAFila[idP] = fila;
    if (esCanonico) {
      // Guardar canónico — preferir el de descripción más larga si hay varios (raro)
      if (!canonicoPorSku[skuFinal] ||
          (desc && desc.length > (canonicoPorSku[skuFinal].desc || '').length)) {
        canonicoPorSku[skuFinal] = { cod: cod, desc: desc, idP: idP };
      }
    }
  }

  // EQUIVALENCIAS — apuntan a un skuBase con factor implícito 1
  var sheetEq = mosSS.getSheetByName('EQUIVALENCIAS');
  if (sheetEq) {
    var eqData = sheetEq.getDataRange().getValues();
    var hdrsE  = eqData[0].map(function(h){ return String(h); });
    var iCodE  = hdrsE.indexOf('codigoBarra') >= 0 ? hdrsE.indexOf('codigoBarra') : hdrsE.indexOf('Cod_Alias');
    var iSkuE  = hdrsE.indexOf('skuBase')    >= 0 ? hdrsE.indexOf('skuBase')    : hdrsE.indexOf('Cod_Barras_Real');
    var iActE  = hdrsE.indexOf('activo');
    if (iCodE >= 0 && iSkuE >= 0) {
      for (var e = 1; e < eqData.length; e++) {
        var ca = String(eqData[e][iCodE] || '').trim();
        var sb = String(eqData[e][iSkuE] || '').trim();
        var act = iActE >= 0 ? String(eqData[e][iActE]) : '';
        if (!ca || !sb) continue;
        if (act && (act === '0' || act.toUpperCase() === 'FALSE' || act.toUpperCase() === 'INACTIVO')) continue;
        if (!codAFila[ca]) codAFila[ca] = { sku: sb, factor: 1, esCanonico: false, esEquivalente: true, cod: ca };
        if (!equivalentesPorSku[sb]) equivalentesPorSku[sb] = [];
        if (equivalentesPorSku[sb].indexOf(ca) < 0) equivalentesPorSku[sb].push(ca);
      }
    }
  }

  // 2. Agrupar totales por skuBase, multiplicando por factor cuando aplica.
  //    Resultado: porSku[sku].solicitado en UNIDADES del canónico.
  var porSku = {};
  Object.keys(totalesPorCodBarras).forEach(function(cod) {
    var fila = codAFila[cod];
    var qty  = parseFloat(totalesPorCodBarras[cod]) || 0;
    if (qty <= 0) return;
    if (!fila) {
      // Cod desconocido en el catálogo — fallback defensivo: usarlo como sku
      // (mejor mostrarlo en el pickup que perderlo silencioso)
      if (!porSku[cod]) porSku[cod] = { solicitado: 0 };
      porSku[cod].solicitado += qty;
      return;
    }
    var sku = fila.sku;
    if (!porSku[sku]) porSku[sku] = { solicitado: 0 };
    // Multiplicar por factor — el pickup en WH siempre habla en unidades del canónico
    porSku[sku].solicitado += qty * fila.factor;
  });

  // 3. Construir items con SOLO codigoBarra del canónico + equivalentes.
  //    NO se incluyen idProducto ni codigoBarra de presentaciones.
  var items = [];
  Object.keys(porSku).forEach(function(sku) {
    if (porSku[sku].solicitado <= 0) return;
    var canonico    = canonicoPorSku[sku];
    var equivCods   = equivalentesPorSku[sku] || [];
    var codigosValidos = [];
    if (canonico && canonico.cod) codigosValidos.push(canonico.cod);
    equivCods.forEach(function(c){ if (codigosValidos.indexOf(c) < 0) codigosValidos.push(c); });
    // Si no encontramos canónico ni equivalentes, fallback al primer código que
    // resolvió este sku (ej: el catálogo tiene la fila pero sin canónico explícito).
    if (codigosValidos.length === 0) {
      // Buscar cualquier cod que apunte a este sku
      Object.keys(codAFila).forEach(function(c){
        if (codAFila[c].sku === sku && codAFila[c].cod && codigosValidos.indexOf(codAFila[c].cod) < 0) {
          codigosValidos.push(codAFila[c].cod);
        }
      });
    }
    items.push({
      skuBase:           sku,
      nombre:            (canonico && canonico.desc) ? canonico.desc : sku,
      solicitado:        porSku[sku].solicitado,
      despachado:        0,
      codigosOriginales: codigosValidos
    });
  });
  if (!items.length) return;

  // 4. POST a WH — directo si WH_GAS_URL está, sino vía MOS bridge
  var props   = PropertiesService.getScriptProperties();
  var whUrl   = props.getProperty('WH_GAS_URL') || '';
  var mosUrl  = props.getProperty('MOS_WEB_APP_URL') || '';
  var payload = {
    action:    'recibirPickupDeME',
    idGuiaME:  idGuia,
    idCaja:    cajaId,
    idZona:    zona,
    cajero:    vendedor,
    items:     items
  };

  function _postJson(url) {
    var res = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'text/plain',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
      followRedirects: true
    });
    return JSON.parse(res.getContentText());
  }

  // Intentar 2 caminos. Si AMBOS fallan, persistir en cola para reintento horario.
  var entregado = false;
  var ultimoError = '';
  if (whUrl) {
    try {
      var r1 = _postJson(whUrl);
      if (r1 && r1.ok !== false) { entregado = true; Logger.log('Pickup → WH directo OK · idGuia=' + idGuia); }
      else { ultimoError = 'WH directo: ' + (r1 && r1.error || 'sin ok'); Logger.log(ultimoError); }
    } catch(e1) { ultimoError = 'WH directo: ' + e1.message; Logger.log(ultimoError); }
  }
  if (!entregado && mosUrl) {
    var mosPayload = Object.assign({}, payload, { action: 'forwardWHPickup' });
    try {
      var r2 = UrlFetchApp.fetch(mosUrl, {
        method: 'post', contentType: 'text/plain',
        payload: JSON.stringify(mosPayload),
        muteHttpExceptions: true, followRedirects: true
      });
      var jr2 = JSON.parse(r2.getContentText() || '{}');
      if (jr2.ok !== false) { entregado = true; Logger.log('Pickup → MOS bridge OK · idGuia=' + idGuia); }
      else { ultimoError = 'MOS bridge: ' + (jr2.error || 'sin ok'); Logger.log(ultimoError); }
    } catch(e2) { ultimoError = 'MOS bridge: ' + e2.message; Logger.log(ultimoError); }
  }
  if (!entregado) {
    if (!whUrl && !mosUrl) ultimoError = 'sin WH_GAS_URL ni MOS_WEB_APP_URL';
    _persistirPickupPendienteEnvio(payload, ultimoError);
  }
}

// ── Cola de pickups que NO se pudieron enviar a WH ────────────
// Se guardan en hoja PICKUPS_PENDIENTES_ENVIO para reintento por trigger.
// Cols: idGuiaME, payload (JSON), intentos, ultimoIntento, ultimoError, estado.
function _getColaPickupsPendientes() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('PICKUPS_PENDIENTES_ENVIO');
  if (!sheet) {
    sheet = ss.insertSheet('PICKUPS_PENDIENTES_ENVIO');
    sheet.appendRow(['idGuiaME','payload','intentos','ultimoIntento','ultimoError','estado']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  }
  return sheet;
}

function _persistirPickupPendienteEnvio(payload, error) {
  try {
    var sheet = _getColaPickupsPendientes();
    // Idempotencia: si ya existe esa idGuiaME en cola, solo incrementa intentos
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(payload.idGuiaME)) {
        sheet.getRange(i + 1, 3).setValue((parseInt(data[i][2]) || 0) + 1);
        sheet.getRange(i + 1, 4).setValue(new Date());
        sheet.getRange(i + 1, 5).setValue(error || '');
        return;
      }
    }
    sheet.appendRow([
      String(payload.idGuiaME), JSON.stringify(payload),
      1, new Date(), error || '', 'PENDIENTE'
    ]);
    Logger.log('Pickup persistido en cola · idGuiaME=' + payload.idGuiaME);
  } catch(e) {
    Logger.log('Falló persistencia cola pickup: ' + e.message);
  }
}

// Trigger horario: lee PENDIENTES, reintenta envío, marca ENVIADO si éxito.
// Se rinde tras 8 intentos para no spammear.
function reintentarPickupsPendientes() {
  var sheet = _getColaPickupsPendientes();
  if (sheet.getLastRow() < 2) return { ok: true, mensaje: 'Cola vacía' };
  var data = sheet.getDataRange().getValues();
  var props  = PropertiesService.getScriptProperties();
  var whUrl  = props.getProperty('WH_GAS_URL') || '';
  var mosUrl = props.getProperty('MOS_WEB_APP_URL') || '';
  var enviados = 0, intentados = 0;

  for (var i = 1; i < data.length; i++) {
    var estado   = String(data[i][5] || '');
    var intentos = parseInt(data[i][2]) || 0;
    if (estado !== 'PENDIENTE') continue;
    if (intentos >= 8) {
      sheet.getRange(i + 1, 6).setValue('ABANDONADO');
      continue;
    }
    var payload; try { payload = JSON.parse(data[i][1]); } catch(_){ continue; }
    intentados++;
    var ok = false, err = '';
    if (whUrl) {
      try {
        var r1 = UrlFetchApp.fetch(whUrl, {
          method:'post', contentType:'text/plain',
          payload: JSON.stringify(payload), muteHttpExceptions:true, followRedirects:true
        });
        var j1 = JSON.parse(r1.getContentText() || '{}');
        if (j1.ok !== false) ok = true; else err = 'WH: ' + (j1.error || 'sin ok');
      } catch(e1) { err = 'WH: ' + e1.message; }
    }
    if (!ok && mosUrl) {
      try {
        var mp = Object.assign({}, payload, { action: 'forwardWHPickup' });
        var r2 = UrlFetchApp.fetch(mosUrl, {
          method:'post', contentType:'text/plain',
          payload: JSON.stringify(mp), muteHttpExceptions:true, followRedirects:true
        });
        var j2 = JSON.parse(r2.getContentText() || '{}');
        if (j2.ok !== false) ok = true; else err = 'MOS: ' + (j2.error || 'sin ok');
      } catch(e2) { err = 'MOS: ' + e2.message; }
    }
    sheet.getRange(i + 1, 3).setValue(intentos + 1);
    sheet.getRange(i + 1, 4).setValue(new Date());
    if (ok) {
      sheet.getRange(i + 1, 6).setValue('ENVIADO');
      enviados++;
    } else {
      sheet.getRange(i + 1, 5).setValue(err);
    }
  }
  return { ok: true, intentados: intentados, enviados: enviados };
}

// Ejecutar 1 vez desde editor para activar reintentos automáticos cada 5min.
function setupPickupRetryTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t){
    if (t.getHandlerFunction() === 'reintentarPickupsPendientes') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('reintentarPickupsPendientes').timeBased().everyMinutes(5).create();
  return { ok: true, mensaje: 'Trigger 5min reintentarPickupsPendientes creado' };
}

// ── Hook anulación: avisar a WH que descuente del pickup ───────
// Llamado desde anularVentaIndividual. Notifica a WH para que descuente del pickup origen.
// No bloquea — solo loggea si falla.
// [delete-safe] FUENTE PRIMARIA: me.venta_reposicion_datos (id_caja + totales_por_cod de la
// venta, sumados por Cod_Barras||SKU). Así el aviso a WH funciona aunque el Sheet ya no exista.
// FALLBACK Sheet cuando la RPC no está disponible (gate OFF / falla).
// FIX: el path Sheet leía mal las columnas de VENTAS_DETALLE (col1=SKU como código, col2=Nombre
// como cantidad → cantidad casi siempre NaN→0 → no-op silencioso). Layout REAL (0-idx):
// 0 ID_Venta | 1 SKU | 2 Nombre | 3 Cantidad | 4 Precio | 5 Subtotal | 6 Cod_Barras | ...
// Ahora usa col 6 (Cod_Barras, fallback col 1 SKU) y col 3 (Cantidad), idéntico al descuento del cierre.
function notificarAnulacionPickupAWH(idVenta) {
  if (!idVenta) return;
  try {
    var idCaja = '';
    var porCod = {};

    // ── FUENTE PRIMARIA: Supabase (delete-safe) ──
    var sb = _meVentaReposicionDatos(idVenta);
    if (sb) {
      idCaja = String(sb.id_caja || '').trim();
      if (!idCaja) { Logger.log('notificarAnulacionPickup: caja no encontrada (Supabase) para ' + idVenta); return; }
      var tpc = sb.totales_por_cod || {};
      Object.keys(tpc).forEach(function (cod) {
        var q = parseFloat(tpc[cod]) || 0;
        if (cod && q > 0) porCod[String(cod)] = q;
      });
    } else {
      // ── FALLBACK: Sheet ──
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var shVC = ss.getSheetByName('VENTAS_CABECERA');
      var shVD = ss.getSheetByName('VENTAS_DETALLE');
      if (!shVC || !shVD) return;
      var vcData = shVC.getDataRange().getValues();
      for (var v = 1; v < vcData.length; v++) {
        if (String(vcData[v][0]) === String(idVenta)) { idCaja = String(vcData[v][10] || ''); break; }
      }
      if (!idCaja) { Logger.log('notificarAnulacionPickup: caja no encontrada para ' + idVenta); return; }

      // Acumular cantidades por código (Cod_Barras col 6, fallback SKU col 1; Cantidad col 3).
      var vdData = shVD.getDataRange().getValues();
      for (var d = 1; d < vdData.length; d++) {
        if (String(vdData[d][0]) !== String(idVenta)) continue;
        var cod = String(vdData[d][6] || vdData[d][1] || '').trim();
        var qty = parseFloat(vdData[d][3]) || 0;
        if (!cod || qty <= 0) continue;
        porCod[cod] = (porCod[cod] || 0) + qty;
      }
    }

    var itemsAnulados = Object.keys(porCod).map(function(c){
      return { codigoBarra: c, cantidad: porCod[c] };
    });
    if (!itemsAnulados.length) return;

    var props  = PropertiesService.getScriptProperties();
    var whUrl  = props.getProperty('WH_GAS_URL') || '';
    var mosUrl = props.getProperty('MOS_WEB_APP_URL') || '';
    var payload = {
      action: 'pickupDescontarVenta',
      idCaja: idCaja, itemsAnulados: itemsAnulados
    };

    if (whUrl) {
      try {
        UrlFetchApp.fetch(whUrl, {
          method:'post', contentType:'text/plain',
          payload: JSON.stringify(payload), muteHttpExceptions:true
        });
        Logger.log('Anulación → WH OK · venta=' + idVenta);
        return;
      } catch(e1) { Logger.log('Anulación → WH falló: ' + e1.message); }
    }
    if (mosUrl) {
      try {
        // Bridge vía MOS: usa la action genérica forwardWHPickup pero con otro action interno
        var bridge = { action: 'forwardWHAction', whAction: 'pickupDescontarVenta',
                       idCaja: idCaja, itemsAnulados: itemsAnulados };
        UrlFetchApp.fetch(mosUrl, {
          method:'post', contentType:'text/plain',
          payload: JSON.stringify(bridge), muteHttpExceptions:true
        });
        Logger.log('Anulación → MOS bridge OK · venta=' + idVenta);
      } catch(e2) { Logger.log('Anulación → MOS bridge falló: ' + e2.message); }
    }
  } catch(e) { Logger.log('notificarAnulacionPickup: ' + e.message); }
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
// LIMPIEZA MASIVA: detecta TODAS las cajas con guías SALIDA_VENTAS duplicadas
// y las limpia en una sola ejecución. Conserva la guía más antigua de cada
// caja, elimina las demás, y revierte el stock descontado de más.
//
// USO MANUAL: ejecutar desde Apps Script editor → ver Logs
// O Web App: POST { tipoEvento: 'LIMPIAR_TODAS_DUPLICADAS' }
// ════════════════════════════════════════════════════════════════════════
function limpiarTodasGuiasDuplicadas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetGC    = ss.getSheetByName('GUIAS_CABECERA');
  var sheetGD    = ss.getSheetByName('GUIAS_DETALLE');
  var sheetStock = ss.getSheetByName('STOCK_ZONAS');
  if (!sheetGC || !sheetGD || !sheetStock) return { ok: false, error: 'Hojas no encontradas' };

  // Leer TODO una sola vez (evita 5x re-lecturas en versión anterior)
  var gcData    = sheetGC.getDataRange().getValues();
  var gdData    = sheetGD.getDataRange().getValues();
  var stockData = sheetStock.getDataRange().getValues();
  var stockHdr  = stockData[0];

  // 1. Agrupar guías SALIDA_VENTAS por observación (cajaId)
  var grupos = {}; // obs → array de {idGuia, rowSheet, fecha, zona}
  for (var i = 1; i < gcData.length; i++) {
    if (String(gcData[i][4]) !== 'SALIDA_VENTAS') continue;
    var obs = String(gcData[i][5] || '').trim();
    if (!obs) continue;
    if (!grupos[obs]) grupos[obs] = [];
    grupos[obs].push({
      idGuia:   String(gcData[i][0]),
      rowSheet: i + 1,
      fecha:    gcData[i][1],
      zona:     String(gcData[i][3]),
      obs:      obs
    });
  }

  // 2. Identificar duplicados y marcar las que se conservan vs las que se eliminan
  var idsAEliminarSet = {}; // idGuia → true
  var grupoStats = []; // para logging
  var zonaPorIdGuia = {}; // idGuia → zona (para reverter stock)

  Object.keys(grupos).forEach(function(obs) {
    var lista = grupos[obs];
    if (lista.length <= 1) return; // sin duplicados
    // Conservar la más antigua
    lista.sort(function(a, b){ return new Date(a.fecha) - new Date(b.fecha); });
    var conservada = lista[0];
    var aEliminar = lista.slice(1);
    aEliminar.forEach(function(g) {
      idsAEliminarSet[g.idGuia] = true;
      zonaPorIdGuia[g.idGuia] = g.zona;
    });
    grupoStats.push({
      obs: obs,
      total: lista.length,
      conservada: conservada.idGuia,
      eliminadas: aEliminar.map(function(g){ return g.idGuia; })
    });
  });

  if (grupoStats.length === 0) {
    return { ok: true, mensaje: 'No hay duplicados que limpiar.' };
  }

  Logger.log('=== LIMPIEZA MASIVA: ' + grupoStats.length + ' cajas con duplicados ===');
  grupoStats.forEach(function(s) {
    Logger.log('  - "' + s.obs + '" → conservar ' + s.conservada + ', eliminar ' + s.eliminadas.length);
  });

  // 3. Calcular stock a revertir en una sola pasada de GUIAS_DETALLE
  var revertStockPorZona = {}; // "cod|zona" → cantidad a sumar
  var detalleRowsAEliminar = []; // filas de GD a borrar
  for (var j = 1; j < gdData.length; j++) {
    var idG = String(gdData[j][0]);
    if (!idsAEliminarSet[idG]) continue;
    var cod = String(gdData[j][1]);
    var cant = parseFloat(gdData[j][2]) || 0;
    var zonaG = zonaPorIdGuia[idG] || '';
    var key = cod + '|' + zonaG;
    revertStockPorZona[key] = (revertStockPorZona[key] || 0) + cant;
    detalleRowsAEliminar.push(j + 1);
  }

  // 4. Aplicar reversión al stock en memoria
  var stockMap = {};
  for (var s = 1; s < stockData.length; s++) {
    stockMap[String(stockData[s][0]) + '|' + String(stockData[s][1])] = s;
  }
  var productosRevertidos = 0;
  Object.keys(revertStockPorZona).forEach(function(key) {
    var idx = stockMap[key];
    if (idx !== undefined) {
      stockData[idx][2] = (parseFloat(stockData[idx][2]) || 0) + revertStockPorZona[key];
      productosRevertidos++;
    }
  });

  // 5. Escribir stock UNA vez (batch)
  if (stockData.length > 1) {
    sheetStock.getRange(2, 1, stockData.length - 1, stockHdr.length).setValues(stockData.slice(1));
  }

  // 6. Eliminar filas de GUIAS_DETALLE (ordenadas descendente para no descuadrar índices)
  detalleRowsAEliminar.sort(function(a, b){ return b - a; });
  detalleRowsAEliminar.forEach(function(r){ sheetGD.deleteRow(r); });

  // 7. Eliminar filas de GUIAS_CABECERA (ordenadas descendente)
  var cabRowsAEliminar = [];
  for (var ii = 1; ii < gcData.length; ii++) {
    if (String(gcData[ii][4]) === 'SALIDA_VENTAS' && idsAEliminarSet[String(gcData[ii][0])]) {
      cabRowsAEliminar.push(ii + 1);
    }
  }
  cabRowsAEliminar.sort(function(a, b){ return b - a; });
  cabRowsAEliminar.forEach(function(r){ sheetGC.deleteRow(r); });

  var totalEliminadas = Object.keys(idsAEliminarSet).length;
  Logger.log('=== TOTAL: ' + totalEliminadas + ' guías eliminadas, ' + productosRevertidos + ' productos revertidos en stock ===');

  return {
    ok: true,
    cajasLimpiadas: grupoStats.length,
    totalGuiasEliminadas: totalEliminadas,
    totalProductosRevertidos: productosRevertidos,
    detalles: grupoStats
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
  // [cutover guías ME] FUENTE_DATOS=supabase & key 'guias' no apagada → lee de me.guias_cabecera
  // (RPC me.zona_guias_listar). Las guías nuevas se escriben directo a Supabase (sync de guias_* apagado);
  // la Hoja queda como espejo de respaldo. Mismo gate/cache/fallback que getStockZonas. SHAPE idéntico a la Hoja.
  if (_fuenteDatos('guias') === 'supabase') {
    try {
      var cache = CacheService.getScriptCache();
      var ckey = ('SB_GUIAS_LIST_' + zona).slice(0, 240);
      var hit = cache.get(ckey);
      if (hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
      var r = _sbRpc('me', 'zona_guias_listar', { zona: String(zona) });
      if (r.ok && r.data && r.data.ok && Array.isArray(r.data.guias)) {
        var json = JSON.stringify({ status: 'success', guias: r.data.guias });
        try { cache.put(ckey, json, _FLIP_CACHE_SEG); } catch (eC) {}
        return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
      }
    } catch (e) { /* cae a la Hoja */ }
  }
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
  // [cutover guías ME] supabase → me.guias_detalle (RPC me.zona_guia_detalle). SHAPE [{cod_barras,cantidad}].
  if (_fuenteDatos('guias') === 'supabase') {
    try {
      var cache = CacheService.getScriptCache();
      var ckey = ('SB_GUIAS_DET_' + idGuia).slice(0, 240);
      var hit = cache.get(ckey);
      if (hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
      var r = _sbRpc('me', 'zona_guia_detalle', { idGuia: String(idGuia) });
      if (r.ok && r.data && r.data.ok && Array.isArray(r.data.items)) {
        var json = JSON.stringify({ status: 'success', items: r.data.items });
        try { cache.put(ckey, json, _FLIP_CACHE_SEG); } catch (eC) {}
        return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
      }
    } catch (e) { /* cae a la Hoja */ }
  }
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
  // [cutover guías ME] supabase → me.guias_cabecera tipo ENTRADA_TRASLADO (RPC me.zona_traslados_entrantes).
  // 'desde' es epoch ms (como hoy). SHAPE [{id_guia,fecha,origen,observacion}].
  if (_fuenteDatos('guias') === 'supabase') {
    try {
      var cache = CacheService.getScriptCache();
      var ckey = ('SB_GUIAS_TRAS_' + zona + '_' + (desde || '')).slice(0, 240);
      var hit = cache.get(ckey);
      if (hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
      var r = _sbRpc('me', 'zona_traslados_entrantes', { zona: String(zona), desde: (desde != null ? String(desde) : null) });
      if (r.ok && r.data && r.data.ok && Array.isArray(r.data.traslados)) {
        var json = JSON.stringify({ status: 'success', traslados: r.data.traslados });
        try { cache.put(ckey, json, _FLIP_CACHE_SEG); } catch (eC) {}
        return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
      }
    } catch (e) { /* cae a la Hoja */ }
  }
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
  // [cutover stock ME] FUENTE_DATOS = supabase → lee el saldo operativo desde
  // me.stock_zonas (RPC me.zona_stock), porque la Hoja STOCK_ZONAS quedó CONGELADA
  // (ME_SYNC_OFF_TABLAS apagó su sync; las escrituras de stock van directo a Supabase).
  // Mismo gate/patrón que estado_cajas/ventas_hoy_zona; ante CUALQUIER fallo cae a la Hoja.
  // SHAPE idéntico al de la Hoja: { status:'success', stock:[{Cod_Barras, Zona_ID, Cantidad}, ...] }.
  if (_fuenteDatos('stock_zonas') === 'supabase') {
    try {
      var cache = CacheService.getScriptCache();
      var hit = cache.get('SB_STOCK_ZONAS');
      if (hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
      var r = _sbRpc('me', 'zona_stock', {});
      if (r.ok && r.data && r.data.ok && Array.isArray(r.data.stock)) {
        var json = JSON.stringify({ status: 'success', stock: r.data.stock });
        try { cache.put('SB_STOCK_ZONAS', json, _FLIP_CACHE_SEG); } catch (eC) {}  // >100KB → no cachea, sin romper
        return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
      }
    } catch (e) { /* cae a la Hoja */ }
  }
  // Sheets: default y fallback
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
  // ════════════════════════════════════════════════════════════════════════
  // [CUTOVER] MODELO ABIERTA (espejo WH): la guía MANUAL nace 'ABIERTA' (solo metadata,
  // cantidad_aplicada=0) y NO aplica stock al crear. El stock lo aplica el CIERRE
  // (me.cerrar_guia_zona_idempotente) UNA vez. Gateado por _meGuiasCicloAbierta()
  // (= flag mos.config ME_GUIAS_CICLO_ABIERTA && escritura directa ON).
  // SALIDA_VENTAS y ENTRADA_TRASLADO quedan FUERA (siguen el camino legacy de abajo).
  // ════════════════════════════════════════════════════════════════════════
  if (_meGuiasCicloAbierta() && _esGuiaManualCicloAbierta(data.tipo)) {
    return registrarGuiaAbierta(data);
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetCab   = ss.getSheetByName("GUIAS_CABECERA");
  var sheetDet   = ss.getSheetByName("GUIAS_DETALLE");
  var sheetStock = ss.getSheetByName("STOCK_ZONAS");
  // [delete-safe] Con escritura de stock directa ON, la persistencia REAL (saldo + cabecera + detalle)
  // va por me.zona_registrar_guia / _meRegistrarGuiaMetaDirecto (Supabase, fuente de verdad). Las hojas
  // son SOLO espejo de respaldo → su ausencia NO debe romper la operación. Sin directo (legacy) las hojas
  // siguen siendo obligatorias (la RPC no aplica el saldo).
  var _stockDir = _meStockDirecto();
  if (!_stockDir) {
    if (!sheetCab)   return generarRespuestaError("Pestaña GUIAS_CABECERA no encontrada.");
    if (!sheetDet)   return generarRespuestaError("Pestaña GUIAS_DETALLE no encontrada.");
    if (!sheetStock) return generarRespuestaError("Pestaña STOCK_ZONAS no encontrada.");
  }

  // Tipos: SALIDA_JEFA | SALIDA_MOVIMIENTO | SALIDA_VENTAS
  //        ENTRADA_ALMACEN | ENTRADA_TRASLADO | ENTRADA_LIBRE
  var tipo        = data.tipo;
  var esSalida    = (tipo.indexOf('SALIDA') === 0);
  var signo       = esSalida ? -1 : 1;
  var zonaDestino = String(data.zona_destino || '');

  // IDEMPOTENCY: respetar el id del cliente si vino (retry de red NO crea guía nueva ni dobla el saldo;
  //   la RPC zona_registrar_guia dedupea por refId 'GUIA:<idGuia>:<cod>' → reintentar con el MISMO idGuia
  //   = kardex dedup = no re-suma). Si no vino, generamos (legacy). El front debe mandarlo.
  var idGuia = String(data.idGuia || data.id_guia || data.localId || data.local_id || '').trim() || ("G-" + new Date().getTime());
  if (sheetCab) {
    try {
      sheetCab.appendRow([idGuia, new Date(), data.vendedor, data.zona, tipo,
        data.observacion || '', zonaDestino, 'CONFIRMADO']);
    } catch (eCab) { Logger.log('[registrarGuia] Sheet cabecera write: ' + eCab.message); }
  }

  // SALIDA_MOVIMIENTO → genera ENTRADA_TRASLADO automática en zona destino (id reservado arriba para la RPC).
  //   Idempotente: derivado del idGuia (estable entre reintentos) → '<idGuia>-IN', no del timestamp.
  var idGuiaEntrada = (tipo === 'SALIDA_MOVIMIENTO' && zonaDestino)
    ? (String(data.idGuiaEntrada || '').trim() || (idGuia + "-IN")) : null;

  // [cutover] directo ON → la mutación del SALDO (origen + destino del traslado) va por RPC atómica
  //   me.zona_registrar_guia (delta firmado por tipo + kardex, idempotente por idGuia+cod). NO tocamos
  //   el cantidad de la Hoja (evita doble conteo con el sync). Si falla → fallback al RMW de la Hoja.
  var guiaDirectoOK = false;
  if (_meStockDirecto()) {
    var rG = _meRegistrarGuiaDirecto(idGuia, data.zona, tipo, data.items, data.vendedor, idGuiaEntrada, zonaDestino);
    guiaDirectoOK = !!(rG && rG.ok);
    if (!guiaDirectoOK) {
      // Money-safety: con el sync de stock_zonas/guias_* apagado, el fallback a la Hoja (abajo)
      // YA NO se propaga a Supabase → drift silencioso. Encolamos para REINTENTO IDEMPOTENTE
      // (la RPC dedupea por idGuia+codBarra en su kardex → reintentar NUNCA duplica el movimiento).
      // NO lanzamos excepción: la guía YA quedó persistida (cabecera+detalle) y no debe romperse.
      Logger.log('registrarGuia: RPC directo falló — encolando + fallback Hoja para guía ' + idGuia);
      try {
        _persistirStockPendiente('guia', idGuia, {
          idGuia: idGuia, zona: data.zona, tipo: tipo, items: data.items,
          usuario: data.vendedor, idGuiaEntrada: idGuiaEntrada, zonaDestino: zonaDestino
        }, (rG && rG.error) || 'rpc');
      } catch (ePG) { Logger.log('Encolar guía pendiente: ' + ePG.message); }
    }
  }

  var stockResult = [];
  (data.items || []).forEach(function(item) {
    var cb = String(item.cod_barras);
    if (sheetDet) {  // [delete-safe] espejo de respaldo; ausente → la guía igual quedó en Supabase
      try {
        var nextDet = sheetDet.getLastRow() + 1;
        sheetDet.getRange(nextDet, 2).setNumberFormat('@STRING@');
        sheetDet.getRange(nextDet, 1, 1, 3).setValues([[idGuia, cb, item.cantidad]]);
      } catch (eDet) { Logger.log('[registrarGuia] Sheet detalle write: ' + eDet.message); }
    }
    if (!guiaDirectoOK && sheetStock) {
      var nuevaCant = actualizarStockFila(sheetStock, cb, data.zona, signo * item.cantidad);
      stockResult.push({ cod_barras: cb, cantidad: nuevaCant });
    } else {
      stockResult.push({ cod_barras: cb });   // saldo lo lleva Supabase (lectura directa lo refleja)
    }
  });

  if (idGuiaEntrada) {
    if (sheetCab) {
      try {
        sheetCab.appendRow([idGuiaEntrada, new Date(), data.vendedor, zonaDestino, 'ENTRADA_TRASLADO',
          'Traslado desde ' + data.zona + ' — Guía origen: ' + idGuia, data.zona, 'CONFIRMADO']);
      } catch (eCabE) { Logger.log('[registrarGuia] Sheet cabecera entrada write: ' + eCabE.message); }
    }
    (data.items || []).forEach(function(item) {
      var cb = String(item.cod_barras);
      if (sheetDet) {
        try {
          var nextDetE = sheetDet.getLastRow() + 1;
          sheetDet.getRange(nextDetE, 2).setNumberFormat('@STRING@');
          sheetDet.getRange(nextDetE, 1, 1, 3).setValues([[idGuiaEntrada, cb, item.cantidad]]);
        } catch (eDetE) { Logger.log('[registrarGuia] Sheet detalle entrada write: ' + eDetE.message); }
      }
      if (!guiaDirectoOK && sheetStock) actualizarStockFila(sheetStock, cb, zonaDestino, item.cantidad);
    });
  }

  // METADATA → Supabase (cabecera+detalle de la guía + la ENTRADA_TRASLADO espejo si aplica).
  //   SOLO si la escritura directa está ON. El saldo ya lo aplicó me.zona_registrar_guia arriba →
  //   esto es METADATA ONLY (no re-aplica stock, SIN doble conteo). La Hoja queda como espejo (appendRow intacto).
  //   Idempotente por idGuia. Si falla → cola ME_STOCK_PENDIENTE ('guia_meta'); NUNCA rompe la operación.
  if (_meStockDirecto()) {
    var metaItems = (data.items || []).map(function(it){ return { codBarra: String(it.cod_barras), cantidad: parseFloat(it.cantidad) || 0 }; });
    var rMeta = _meRegistrarGuiaMetaDirecto({
      idGuia: idGuia, zona: data.zona, tipo: tipo, vendedor: data.vendedor,
      observacion: data.observacion || '', zonaDestino: zonaDestino, items: metaItems
    });
    if (!rMeta || !rMeta.ok) {
      Logger.log('registrarGuia: meta Supabase falló — encolando para guía ' + idGuia);
      try { _persistirStockPendiente('guia_meta', idGuia, {
        idGuia: idGuia, zona: data.zona, tipo: tipo, vendedor: data.vendedor,
        observacion: data.observacion || '', zonaDestino: zonaDestino, items: metaItems
      }, (rMeta && rMeta.error) || 'rpc'); } catch(eM) { Logger.log('Encolar guia_meta: ' + eM.message); }
    }
    // ENTRADA_TRASLADO espejo (zona destino) → su propia metadata, para que zona_traslados_entrantes la vea.
    if (idGuiaEntrada) {
      var rMetaE = _meRegistrarGuiaMetaDirecto({
        idGuia: idGuiaEntrada, zona: zonaDestino, tipo: 'ENTRADA_TRASLADO', vendedor: data.vendedor,
        observacion: 'Traslado desde ' + data.zona + ' — Guía origen: ' + idGuia,
        zonaDestino: data.zona, items: metaItems
      });
      if (!rMetaE || !rMetaE.ok) {
        Logger.log('registrarGuia: meta entrada-espejo falló — encolando para ' + idGuiaEntrada);
        try { _persistirStockPendiente('guia_meta', idGuiaEntrada, {
          idGuia: idGuiaEntrada, zona: zonaDestino, tipo: 'ENTRADA_TRASLADO', vendedor: data.vendedor,
          observacion: 'Traslado desde ' + data.zona + ' — Guía origen: ' + idGuia,
          zonaDestino: data.zona, items: metaItems
        }, (rMetaE && rMetaE.error) || 'rpc'); } catch(eME) { Logger.log('Encolar guia_meta entrada: ' + eME.message); }
      }
    }
  }

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success', idGuia: idGuia, idGuiaEntrada: idGuiaEntrada, stock: stockResult
  })).setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════════════════
// [CUTOVER] registrarGuiaAbierta — crea una guía MANUAL en estado 'ABIERTA' (modelo WH).
// ────────────────────────────────────────────────────────────────────────
// · NO toca stock al crear (NO llama me.zona_registrar_guia). Solo escribe metadata
//   (cabecera+detalle) en Supabase vía me.zona_guia_registrar_meta con estado='ABIERTA'
//   → la RPC pone cantidad_aplicada=0 (fix D del SQL 204) → el cierre aplicará el total UNA vez.
// · IDEMPOTENCY KEY: el cliente puede mandar data.idGuia (o data.localId). El server lo RESPETA
//   (no genera otro idGuia por POST) → reintento de red NO crea guía nueva ni dobla. La RPC meta
//   es idempotente por idGuia (on conflict (id_guia) do update; detalle se re-escribe igual).
//   Si el cliente NO manda id, se genera uno DETERMINISTA-por-payload no es posible → generamos
//   "G-" + timestamp (el front DEBE mandar idGuia para idempotencia real; ver tanda de frontend).
// · ESPEJO DE TRASLADO: SALIDA_MOVIMIENTO con destino → además se escribe la metadata de la
//   ENTRADA_TRASLADO espejo (zona destino), estado CONFIRMADO (visibilidad en trasladosEntrantes),
//   cantidad_aplicada=cantidad → si alguna vez se cerrara, delta 0 (NO re-suma). El IN real al
//   destino lo aplica el CIERRE del SALIDA_MOVIMIENTO origen (espejo en cerrar_guia_zona_idempotente).
// · La Hoja se escribe como espejo de respaldo (best-effort, estado ABIERTA) si existe.
// NUNCA aplica stock aquí. Devuelve {status:'success', idGuia, idGuiaEntrada, estado:'ABIERTA'}.
// ════════════════════════════════════════════════════════════════════════
function registrarGuiaAbierta(data) {
  var tipo        = String(data.tipo || '').toUpperCase();
  var zonaDestino = String(data.zona_destino || '');
  // IDEMPOTENCY: respetar el id del cliente si vino; si no, generar (el front debe mandarlo).
  var idGuia = String(data.idGuia || data.id_guia || data.localId || data.local_id || '').trim() || ("G-" + new Date().getTime());
  var idGuiaEntrada = (tipo === 'SALIDA_MOVIMIENTO' && zonaDestino)
    ? (String(data.idGuiaEntrada || '').trim() || (idGuia + "-IN")) : null;

  var items = (data.items || []).map(function (it) {
    return { codBarra: String(it.cod_barras || it.codBarra || ''), cantidad: parseFloat(it.cantidad) || 0 };
  }).filter(function (it) { return it.codBarra && it.cantidad > 0; });

  // 1) METADATA ABIERTA en Supabase (fuente de verdad). Idempotente por idGuia.
  var rMeta = _meRegistrarGuiaMetaDirecto({
    idGuia: idGuia, zona: data.zona, tipo: tipo, vendedor: data.vendedor,
    observacion: data.observacion || '', zonaDestino: zonaDestino, estado: 'ABIERTA', items: items
  });
  if (!rMeta || !rMeta.ok) {
    // Money-safety: encolar la metadata para reintento idempotente (NO aplica stock → seguro reintentar).
    Logger.log('registrarGuiaAbierta: meta ABIERTA falló — encolando para guía ' + idGuia);
    try { _persistirStockPendiente('guia_meta', idGuia, {
      idGuia: idGuia, zona: data.zona, tipo: tipo, vendedor: data.vendedor,
      observacion: data.observacion || '', zonaDestino: zonaDestino, estado: 'ABIERTA', items: items
    }, (rMeta && rMeta.error) || 'rpc'); } catch (eM) { Logger.log('Encolar guia_meta ABIERTA: ' + eM.message); }
  }

  // 2) ESPEJO de traslado (solo visibilidad). CONFIRMADO + aplicada=cantidad → nunca re-suma al cerrar.
  if (idGuiaEntrada) {
    var rMetaE = _meRegistrarGuiaMetaDirecto({
      idGuia: idGuiaEntrada, zona: zonaDestino, tipo: 'ENTRADA_TRASLADO', vendedor: data.vendedor,
      observacion: 'Traslado desde ' + data.zona + ' — Guía origen: ' + idGuia,
      zonaDestino: data.zona, estado: 'CONFIRMADO', items: items
    });
    if (!rMetaE || !rMetaE.ok) {
      Logger.log('registrarGuiaAbierta: meta espejo falló — encolando para ' + idGuiaEntrada);
      try { _persistirStockPendiente('guia_meta', idGuiaEntrada, {
        idGuia: idGuiaEntrada, zona: zonaDestino, tipo: 'ENTRADA_TRASLADO', vendedor: data.vendedor,
        observacion: 'Traslado desde ' + data.zona + ' — Guía origen: ' + idGuia,
        zonaDestino: data.zona, estado: 'CONFIRMADO', items: items
      }, (rMetaE && rMetaE.error) || 'rpc'); } catch (eME) { Logger.log('Encolar espejo: ' + eME.message); }
    }
  }

  // 3) Espejo de respaldo a la Hoja (best-effort, estado ABIERTA). Ausente → no rompe.
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetCab = ss.getSheetByName("GUIAS_CABECERA");
    var sheetDet = ss.getSheetByName("GUIAS_DETALLE");
    if (sheetCab) {
      sheetCab.appendRow([idGuia, new Date(), data.vendedor, data.zona, tipo,
        data.observacion || '', zonaDestino, 'ABIERTA']);
      if (idGuiaEntrada) sheetCab.appendRow([idGuiaEntrada, new Date(), data.vendedor, zonaDestino, 'ENTRADA_TRASLADO',
        'Traslado desde ' + data.zona + ' — Guía origen: ' + idGuia, data.zona, 'CONFIRMADO']);
    }
    if (sheetDet) {
      (data.items || []).forEach(function (item) {
        var cb = String(item.cod_barras || item.codBarra || '');
        var nextDet = sheetDet.getLastRow() + 1;
        sheetDet.getRange(nextDet, 2).setNumberFormat('@STRING@');
        sheetDet.getRange(nextDet, 1, 1, 3).setValues([[idGuia, cb, item.cantidad]]);
        if (idGuiaEntrada) {
          var nextDetE = sheetDet.getLastRow() + 1;
          sheetDet.getRange(nextDetE, 2).setNumberFormat('@STRING@');
          sheetDet.getRange(nextDetE, 1, 1, 3).setValues([[idGuiaEntrada, cb, item.cantidad]]);
        }
      });
    }
  } catch (eSheet) { Logger.log('registrarGuiaAbierta: espejo Hoja falló (no bloquea): ' + eSheet.message); }

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success', idGuia: idGuia, idGuiaEntrada: idGuiaEntrada, estado: 'ABIERTA'
  })).setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════════════════
// [CUTOVER] cerrarGuia — CIERRA una guía de zona ABIERTA → aplica stock UNA vez.
// ────────────────────────────────────────────────────────────────────────
// Endpoint mínimo para esta tanda (el frontend completo hold-to-confirm es la siguiente).
// Llama me.cerrar_guia_zona_idempotente (vía wrapper): OUT origen + IN espejo destino si traslado,
// delta = cantidad − cantidad_aplicada, idempotente (recerrar/retry/doble-tap = delta 0, no dobla).
// Espeja el estado a la Hoja (best-effort). NUNCA lanza.
// ════════════════════════════════════════════════════════════════════════
function cerrarGuia(data) {
  var idGuia = String((data && (data.idGuia || data.id_guia)) || '').trim();
  if (!idGuia) return generarRespuestaError("Requiere idGuia para cerrar.");

  var r = _meCerrarGuiaDirecto(idGuia);
  if (!r || !r.ok) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error', mensaje: 'No se pudo cerrar la guía: ' + ((r && r.error) || 'rpc'), idGuia: idGuia
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // Espejo a la Hoja: marcar estado CERRADA (best-effort). Col 8 (índice 7) = Estado.
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetCab = ss.getSheetByName("GUIAS_CABECERA");
    if (sheetCab) {
      var gc = sheetCab.getDataRange().getValues();
      for (var i = 1; i < gc.length; i++) {
        if (String(gc[i][0]) === idGuia) { sheetCab.getRange(i + 1, 8).setValue('CERRADA'); break; }
      }
    }
  } catch (eS) { Logger.log('cerrarGuia: espejo Hoja estado falló (no bloquea): ' + eS.message); }

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success', idGuia: idGuia, estado: 'CERRADA',
    lineasAplicadas: r.lineasAplicadas, lineasSaltadas: r.lineasSaltadas
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
  var auditoriasParaSupa = [];   // [fix 🔴#4] espejo idempotente a me.auditorias del camino directo
  (data.items || []).forEach(function(item) {
    var cb      = String(item.cod_barras);
    var cantReal = parseFloat(item.cantReal) || 0;
    // [fix 🟠#3] cantSistema del payload es CACHE del front; lo usamos como semilla, pero si el ajuste
    //   directo corre, lo REEMPLAZAMOS por el stockAntes REAL que devuelve la RPC (verdad de me.stock_zonas).
    var cantSis = parseFloat(item.cantSistema) || 0;

    // Stock: establecer cantidad DIRECTAMENTE al valor real auditado (SET absoluto).
    //   [cutover] directo ON → me.zona_ajustar_stock (set + log + kardex AUDITORIA re-anclado), idempotente por
    //   localId (estable por auditoría+código → re-enviar la misma auditoría no re-ancla raro). Si falla → Hoja.
    var auditDirectoOK = false;
    if (_meStockDirecto()) {
      var localAj = idAudit + ':' + cb;   // estable por auditoría+código
      var rAj = _meAjustarStockDirecto(data.zona, cb, cantReal, usuario, localAj, 'AUDITORIA');
      auditDirectoOK = !!(rAj && rAj.ok);
      // [fix 🟠#3] usar el stockAntes REAL de la RPC para la diferencia (no el cache del front).
      //   En dedup la RPC no devuelve stockAntes → conservamos cantSis del payload (la fila ya existe igual).
      if (auditDirectoOK && rAj.data && rAj.data.stockAntes != null) {
        cantSis = parseFloat(rAj.data.stockAntes) || 0;
      }
      if (!auditDirectoOK) {
        // Money-safety: sync de stock_zonas apagado → el fallback a la Hoja ya no llega a
        // Supabase. Encolamos el ajuste (SET absoluto) para REINTENTO IDEMPOTENTE: la RPC dedupea
        // por localId → reintentar re-ancla al MISMO valor auditado, nunca duplica ni acumula.
        Logger.log('registrarAuditoria: RPC ajuste directo falló — encolando + fallback Hoja para ' + cb + '@' + data.zona);
        try {
          _persistirStockPendiente('ajuste', localAj, {
            zona: data.zona, codBarra: cb, nuevo: cantReal, usuario: usuario, localId: localAj, origen: 'AUDITORIA'
          }, (rAj && rAj.error) || 'rpc');
        } catch (ePA) { Logger.log('Encolar ajuste pendiente: ' + ePA.message); }
      }
    }
    if (!auditDirectoOK) {
      _actualizarStockAuditoria(sheetStock, cb, data.zona, cantReal, usuario, ahoraStr);
    }

    var diff = cantReal - cantSis;
    var key  = usuario + '|' + data.zona + '|' + cb;

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

    // [fix 🔴#4] espejo directo a me.auditorias (no depende del sync de la Hoja). Usa cantSis REAL.
    auditoriasParaSupa.push({
      id_auditoria: idAudit, fecha: ahoraStr, vendedor: usuario, zona_id: String(data.zona),
      cod_barras: cb, cant_sistema: cantSis, cant_real: cantReal, diferencia: diff
    });
  });

  // [fix 🔴#4] upsert idempotente a me.auditorias (PK id_auditoria,cod_barras). Best-effort: no rompe el flujo.
  //   Solo cuando el camino directo está activo (mismo gate que el ajuste de stock); si falla, el sync de la
  //   Hoja sigue siendo respaldo. Re-aplicar la MISMA auditoría no duplica (on conflict por la PK).
  if (_meStockDirecto() && auditoriasParaSupa.length) {
    try {
      var rAud = _sbUpsert('me.auditorias', auditoriasParaSupa, 'id_auditoria,cod_barras', false);
      if (!rAud || !rAud.ok) Logger.log('[auditoria-directo] upsert me.auditorias falló: ' + ((rAud && rAud.error) || '?'));
    } catch (eAud) { Logger.log('[auditoria-directo] EXCEPCIÓN me.auditorias: ' + eAud.message); }
  }

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
