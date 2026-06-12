// ============================================================
// Fase2Auth.gs — mint-token: GAS emite un JWT scoped para que la PWA hable DIRECTO a Supabase.
// ============================================================
// El JWT SECRET (Supabase → Settings → API → JWT Secret) vive SOLO en GAS (Script Property
// SUPABASE_JWT_SECRET), NUNCA en el navegador. La PWA pide un token corto (exp 5min) con el claim
// 'zonas' que sale del binding admin-only mos.dispositivo_zonas. La RLS de Supabase deriva la zona
// de ESE claim (no de params del cliente → no falsificable). Re-mint en heartbeat.
// HS256 mintado en GAS es seguro: el secreto no sale de GAS; el navegador solo recibe el token corto.
// (Upgrade futuro: firma asimétrica RS256 vía Edge Function — ver MIGRACION_FASE2_PLAN.md C4.)

// [Fase 2] Espeja a Sheets una venta NV YA creada directo en Supabase (por crear_venta_directa), para que
// el cierre/SUNAT/reportes —que leen Sheets— sigan cuadrando. Idempotente por Ref_Local (no re-escribe).
// NO mintea correlativo (ya viene del RPC), NO dual-writea (ya está en Supabase), NO imprime. Lo llama la
// PWA async tras la venta directa. Recibe el ventaBase + idVenta + correlativo del RPC.
function mirrorVentaASheets(data){
  data = data || {};
  var h    = data.header || {};
  var auth = data.auth   || {};
  var pos  = data.pos_config || data.pos || {};   // [fix CRITICO 20x] ventaBase lleva la caja en pos_config
  var ref  = String((data.data_sync && data.data_sync.last_sync) || data.ref_local || '').trim();
  if(!ref) return { ok:false, error:'ref_local requerido' };
  var idVenta = String(data.idVenta || '');
  var correlativo = String(data.correlativo || '');
  if(!idVenta || !correlativo) return { ok:false, error:'idVenta/correlativo requeridos (vienen del RPC directo)' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetCab = ss.getSheetByName('VENTAS_CABECERA');
  if(!sheetCab) return { ok:false, error:'VENTAS_CABECERA no existe' };

  // [fix C2/ALTO 20x] LockService alrededor de scan+append: sin esto, dos MIRROR_VENTA concurrentes con el
  // mismo ref_local podrían ambos pasar el dedup y escribir DOS filas (el cierre cuenta doble). Sheets no
  // tiene constraint único que lo frene → el lock es la única defensa del lado Sheets.
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch(e){ return { ok:false, error:'MIRROR_OCUPADO: no se pudo tomar el lock (reintentar)' }; }
  try {
    // idempotencia: ¿ya espejada? dedup por Ref_Local (col 14), escaneo de 1 columna (liviano)
    var lastRow = sheetCab.getLastRow();
    if(lastRow >= 2){
      var refCol = sheetCab.getRange(2, 14, lastRow - 1, 1).getValues();
      for(var i=0;i<refCol.length;i++){ if(String(refCol[i][0]) === ref) return { ok:true, dedup:true, idVenta:idVenta, correlativo:correlativo }; }
    }

    var tipoDocCliente = parseInt((h.cliente && h.cliente.tipo) || 0, 10);
    sheetCab.appendRow([
      idVenta, new Date(), auth.vendedor || '', auth.estacion || '',
      (h.cliente && h.cliente.doc) || '', (h.cliente && h.cliente.nombre) || '',
      (h.total != null ? h.total : 0), h.tipoDoc || 'NOTA_DE_VENTA', h.metodo || 'EFECTIVO',
      correlativo, pos.cajaId || '', auth.deviceId || '', 'COMPLETADO',
      ref, String(h.obs || ''), tipoDocCliente, '', '', ''
    ]);
    // detalle DENTRO del lock → cabecera+detalle atómicos (sin ventana de fila huérfana)
    var items = data.items || [];
    if(items.length){
      var sheetDet = ss.getSheetByName('VENTAS_DETALLE');
      if(sheetDet){
        var rows = items.map(function(it){
          var vu = parseFloat(it.valor_unitario) || Math.round(parseFloat(it.precio||0)/1.18*100)/100;
          return [ idVenta, it.sku, it.nombre, it.cantidad, it.precio, it.subtotal,
                   String(it.codBarras || it.cod_barras || ''), Math.round(vu*100)/100,
                   parseInt(it.tipo_igv || 1, 10), String(it.unidad_de_medida || it.unidad_medida || 'NIU') ];
        });
        var lr = sheetDet.getLastRow();
        sheetDet.getRange(lr+1, 7, rows.length, 1).setNumberFormat('@STRING@');
        sheetDet.getRange(lr+1, 2, rows.length, 1).setNumberFormat('@STRING@');
        sheetDet.getRange(lr+1, 1, rows.length, rows[0].length).setValues(rows);
      }
    }
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }

  // jornada del vendedor (cacheada, barata) — parity con procesarVenta
  try { _registrarJornadaEnMOS(String(auth.vendedor || '')); } catch(_){}
  return { ok:true, dedup:false, idVenta:idVenta, correlativo:correlativo };
}

// [Fase 2 · contrato #2] Reconciliación Supabase→Sheets: red de seguridad para cuando el mirror falla.
// Busca ventas NV de HOY en me.ventas cuyo ref_local NO esté en VENTAS_CABECERA y las espeja (vía
// mirrorVentaASheets, idempotente). Así, aunque un MIRROR_VENTA se pierda, el cierre nunca sub-cuenta.
// Pensado para correr periódico (trigger 5-10min) o al iniciar el cierre. Aditivo: NO toca el flujo de venta.
function reconciliarDirectasSheets(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetCab = ss.getSheetByName('VENTAS_CABECERA');
  if(!sheetCab) return { ok:false, error:'VENTAS_CABECERA no existe' };

  // ref_locals ya en Sheets (set, col 14)
  var enSheets = {};
  var lastRow = sheetCab.getLastRow();
  if(lastRow >= 2){
    var refCol = sheetCab.getRange(2, 14, lastRow - 1, 1).getValues();
    for(var i=0;i<refCol.length;i++){ var rl = String(refCol[i][0]||''); if(rl) enSheets[rl] = 1; }
  }

  // ventas NV de hoy (Lima) en Supabase
  var tz = Session.getScriptTimeZone();
  var hoy0 = Utilities.formatDate(new Date(), 'America/Lima', "yyyy-MM-dd'T'00:00:00XXX");
  var r = _sbSelect('me.ventas', { filters: { fecha:'gte.'+hoy0, tipo_doc:'eq.NOTA_DE_VENTA' } });
  if(!r.ok) return { ok:false, error:'no se pudo leer me.ventas: '+(r.error||'') };

  var faltantes = (r.data||[]).filter(function(v){ var rl=String(v.ref_local||''); return rl && !enSheets[rl]; });
  var espejadas = 0, errores = [];
  faltantes.forEach(function(v){
    try {
      // traer detalle de esa venta
      var det = _sbSelect('me.ventas_detalle', { filters: { id_venta:'eq.'+String(v.id_venta) } });
      var items = (det.ok ? (det.data||[]) : []).map(function(d){
        return { sku:d.sku, nombre:d.nombre, cantidad:d.cantidad, precio:d.precio, subtotal:d.subtotal,
                 cod_barras:d.cod_barras, valor_unitario:d.valor_unitario, tipo_igv:d.tipo_igv, unidad_medida:d.unidad_medida };
      });
      var payload = {
        idVenta: v.id_venta, correlativo: v.correlativo,
        data_sync: { last_sync: v.ref_local },
        header: { total: v.total, tipoDoc: v.tipo_doc, metodo: v.forma_pago, obs: v.obs,
                  cliente: { doc: v.cliente_doc, nombre: v.cliente_nombre, tipo: v.tipo_doc_cliente } },
        auth: { vendedor: v.vendedor, estacion: v.estacion, deviceId: v.dispositivo_id },
        pos:  { cajaId: v.id_caja },
        items: items
      };
      var m = mirrorVentaASheets(payload);   // idempotente (dedup por ref_local) + LockService
      if(m.ok && !m.dedup) espejadas++;
      else if(!m.ok) errores.push(v.ref_local + ': ' + (m.error||''));
    } catch(e){ errores.push(String(v.ref_local) + ': ' + (e && e.message)); }
  });
  var out = { ok:true, ventasHoyEnSupabase:(r.data||[]).length, faltabanEnSheets:faltantes.length, espejadas:espejadas, errores:errores };
  Logger.log('[reconciliarDirectasSheets] ' + JSON.stringify(out));
  return out;
}

// [Fase 2 · fix ALTO 20x] Registra el trigger de reconciliación (cada 10min) — sino el backstop queda huérfano.
// Correr UNA vez en el editor antes de habilitar la escritura directa. Idempotente.
function instalarTriggerReconciliacionDirectas(){
  ScriptApp.getProjectTriggers().forEach(function(t){
    if(t.getHandlerFunction()==='reconciliarDirectasSheets') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('reconciliarDirectasSheets').timeBased().everyMinutes(10).create();
  Logger.log('✅ Trigger reconciliarDirectasSheets cada 10min instalado (backstop del mirror)');
  return { ok:true };
}

function _b64url_(bytes){ return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, ''); }
function _b64urlStr_(str){ return _b64url_(Utilities.newBlob(str).getBytes()); }

// Emite un JWT 'authenticated' scoped por zona para un dispositivo. Lo llama la PWA al iniciar + en heartbeat.
function mintSupabaseToken(deviceId){
  var idd = String(deviceId || '').trim();
  if(!idd) return { ok:false, error:'deviceId requerido' };
  var secret = PropertiesService.getScriptProperties().getProperty('SUPABASE_JWT_SECRET');
  if(!secret) return { ok:false, error:'falta SUPABASE_JWT_SECRET en Script Properties (Supabase → Settings → API → JWT Secret)' };

  // [autorización por UUID, SIN zona — los dispositivos/empleados ROTAN entre zonas] Valida que el UUID
  // esté REGISTRADO + ACTIVO + app=mosExpress (igual que verificarDispositivo). La zona la decide el turno
  // (qué caja abre), no el dispositivo. Fail-closed: dispositivo no registrado/activo → no token.
  var r = _sbSelect('mos.dispositivos', { filters: { id_dispositivo:'eq.'+idd, app:'eq.mosExpress', estado:'eq.ACTIVO' } });
  if(!r.ok) return { ok:false, error:'no se pudo verificar dispositivo: '+(r.error||'') };
  if(!(r.data || []).length) return { ok:false, error:'dispositivo no registrado/activo para mosExpress' };

  var now = Math.floor(Date.now()/1000);
  var header  = { alg:'HS256', typ:'JWT' };
  var payload = {
    iss:'supabase', role:'authenticated', aud:'authenticated', sub:idd,
    app:'mosExpress',   // SIN zona: el dispositivo rota, la zona la pone el turno/caja
    iat: now, exp: now + 300   // 5 minutos (corto a propósito; re-mint en heartbeat)
  };
  var signingInput = _b64urlStr_(JSON.stringify(header)) + '.' + _b64urlStr_(JSON.stringify(payload));
  var sig = Utilities.computeHmacSha256Signature(signingInput, secret);
  var token = signingInput + '.' + _b64url_(sig);
  return { ok:true, token:token, app:'mosExpress', exp:payload.exp };
}

// Wrapper de prueba para el editor (sin args): mintea para el 1er dispositivo con binding y muestra el token.
function probarMintToken(){
  var secret = PropertiesService.getScriptProperties().getProperty('SUPABASE_JWT_SECRET');
  if(!secret){ Logger.log('❌ FALTA cargar SUPABASE_JWT_SECRET en Propiedades del script (Supabase → Settings → API → JWT Secret)'); return; }
  Logger.log('✅ SUPABASE_JWT_SECRET presente ('+secret.length+' chars)');
  var r = _sbSelect('mos.dispositivos', { filters:{ app:'eq.mosExpress', estado:'eq.ACTIVO' }, limit:1 });
  if(!r.ok || !(r.data||[]).length){ Logger.log('sin dispositivos mosExpress ACTIVOS en mos.dispositivos'); return; }
  var dev = String(r.data[0].id_dispositivo);
  var out = mintSupabaseToken(dev);
  Logger.log('mint para dispositivo '+dev+' → '+JSON.stringify({ok:out.ok, app:out.app, error:out.error}));
  if(out.ok){
    var parts = out.token.split('.');
    var payJson = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[1])).getDataAsString();
    Logger.log('payload del JWT: '+payJson);
    Logger.log('TOKEN COMPLETO (pegámelo para probarlo): '+out.token);
  }
}
