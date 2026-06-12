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
    // [CPE-directo] campos NF (cols 17-19): vacíos para NV; para boleta/factura llevan el resultado de NubeFact
    // (estado/hash/enlace) que viene en data.nf_* → así la fila de Sheets de la boleta no queda sin la data fiscal.
    sheetCab.appendRow([
      idVenta, new Date(), auth.vendedor || '', auth.estacion || '',
      (h.cliente && h.cliente.doc) || '', (h.cliente && h.cliente.nombre) || '',
      (h.total != null ? h.total : 0), h.tipoDoc || 'NOTA_DE_VENTA', h.metodo || 'EFECTIVO',
      correlativo, pos.cajaId || '', auth.deviceId || '', 'COMPLETADO',
      ref, String(h.obs || ''), tipoDocCliente,
      String(data.nf_estado || ''), String(data.nf_hash || ''), String(data.nf_enlace || '')
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

// [Fase 2] Espeja a Sheets un movimiento de caja YA creado directo en Supabase (por crear_movimiento_directo),
// para que el cierre —que lee MOVIMIENTOS_EXTRA de Sheets— cuadre. Idempotente por ID_Extra (col 1).
// LockService alrededor de scan+append (sin constraint único en Sheets, el lock es la única defensa contra
// doble fila). Comportamiento IDÉNTICO al handler GAS real (registrarExtraCajaConLog): igual que él, NO dispara
// _chequearAlertaEfectivo (no introducir divergencia atada al flag). Lo llama la PWA async tras el mov directo.
function mirrorMovimientoASheets(data){
  data = data || {};
  var id = String(data.id_extra || '').trim();
  if(!id) return { ok:false, error:'id_extra requerido' };
  var idCaja = String(data.id_caja || '');
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch(e){ return { ok:false, error:'MIRROR_MOV_OCUPADO: no se pudo tomar el lock (reintentar)' }; }
  var dedup = false;
  try {
    var sheet = ss.getSheetByName('MOVIMIENTOS_EXTRA');
    if(!sheet){
      sheet = ss.insertSheet('MOVIMIENTOS_EXTRA');
      sheet.appendRow(['ID_Extra','ID_Caja','Timestamp','Tipo','Monto','Concepto','Obs','Registrado_Por']);
    }
    // idempotencia: ¿ya espejado? dedup por ID_Extra (col 1), escaneo de 1 columna (liviano)
    var lastRow = sheet.getLastRow();
    if(lastRow >= 2){
      var idCol = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for(var i=0;i<idCol.length;i++){ if(String(idCol[i][0]) === id){ dedup = true; break; } }
    }
    if(!dedup){
      sheet.appendRow([
        id, idCaja, new Date(), String(data.tipo || 'EGRESO'),
        parseFloat(data.monto) || 0, String(data.concepto || ''),
        String(data.obs || ''), String(data.registrado_por || '')
      ]);
      SpreadsheetApp.flush();
    }
  } finally { lock.releaseLock(); }

  return { ok:true, dedup:dedup, id_extra:id };
}

// [Fase 2 · contrato #2] Reconciliación Supabase→Sheets: red de seguridad para cuando el mirror falla.
// Busca ventas NV de HOY en me.ventas cuyo ref_local NO esté en VENTAS_CABECERA y las espeja (vía
// mirrorVentaASheets, idempotente). Así, aunque un MIRROR_VENTA se pierda, el cierre nunca sub-cuenta.
// Pensado para correr periódico (trigger 5-10min) o al iniciar el cierre. Aditivo: NO toca el flujo de venta.
function reconciliarDirectasSheets(){
  // [guard-flag] El backstop solo tiene sentido cuando hay escritura directa (CORRELATIVO_SOURCE=supabase).
  // Si está en 'sheets', no hay ventas directas que reconciliar → evita GETs desperdiciados si el trigger
  // quedó instalado antes de tiempo.
  try { if(_fuenteCorrelativo() !== 'supabase') return { ok:true, skip:true, motivo:'CORRELATIVO_SOURCE!=supabase' }; } catch(_){}
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
  // limit/order deterministas: si PostgREST tuviera db-max-rows y hubiera más NV/día que el tope, evitamos
  // un truncado SILENCIOSO (la red de seguridad NO debe sub-cubrir sin avisar). 2000 NV/día es inverosímil
  // para este negocio; si algún día se alcanza, lo logueamos en vez de truncar a ciegas.
  var _LIM_RECON = 2000;
  // incluye NV + CPE (boleta/factura): cuando el CPE-directo esté vivo, sus ventas también se reconcilian.
  var r = _sbSelect('me.ventas', { filters: { fecha:'gte.'+hoy0, tipo_doc:'in.(NOTA_DE_VENTA,BOLETA,FACTURA)' }, order:'fecha.asc', limit:_LIM_RECON });
  if(!r.ok) return { ok:false, error:'no se pudo leer me.ventas: '+(r.error||'') };
  if((r.data||[]).length >= _LIM_RECON) Logger.log('⚠️ [reconciliarDirectasSheets] me.ventas alcanzó el límite '+_LIM_RECON+' — posible truncado, paginar.');

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
        items: items,
        // [CPE-directo] propagar el resultado NF (boleta/factura) al mirror para no perder la data fiscal en Sheets
        nf_estado: v.nf_estado || '', nf_hash: v.nf_hash || '', nf_enlace: v.nf_enlace || ''
      };
      var m = mirrorVentaASheets(payload);   // idempotente (dedup por ref_local) + LockService
      if(m.ok && !m.dedup) espejadas++;
      else if(!m.ok) errores.push(v.ref_local + ': ' + (m.error||''));
    } catch(e){ errores.push(String(v.ref_local) + ': ' + (e && e.message)); }
  });
  // ── Pasada de MOVIMIENTOS (mismo backstop): movimientos de hoy en me.movimientos_extra sin fila en
  // MOVIMIENTOS_EXTRA de Sheets → espejarlos (idempotente por id_extra). Igual que ventas: si un MIRROR_MOV
  // se perdió, el cierre nunca sub-cuenta un egreso/ingreso.
  var movEspejados = 0, movErrores = [], movFaltaban = 0, movEnSupa = 0;
  try {
    var enSheetsMov = {};
    var shMov = ss.getSheetByName('MOVIMIENTOS_EXTRA');
    if(shMov){
      var lrM = shMov.getLastRow();
      if(lrM >= 2){
        var idColM = shMov.getRange(2, 1, lrM - 1, 1).getValues();
        for(var k=0;k<idColM.length;k++){ var idm=String(idColM[k][0]||''); if(idm) enSheetsMov[idm]=1; }
      }
    }
    var rm = _sbSelect('me.movimientos_extra', { filters: { ts:'gte.'+hoy0 }, order:'ts.asc', limit:_LIM_RECON });
    if(rm.ok){
      movEnSupa = (rm.data||[]).length;
      if(movEnSupa >= _LIM_RECON) Logger.log('⚠️ [reconciliarDirectasSheets] me.movimientos_extra alcanzó el límite '+_LIM_RECON+' — posible truncado, paginar.');
      var movFaltantes = (rm.data||[]).filter(function(m){ var i=String(m.id_extra||''); return i && !enSheetsMov[i]; });
      movFaltaban = movFaltantes.length;
      movFaltantes.forEach(function(m){
        try {
          var mm = mirrorMovimientoASheets({
            id_extra: m.id_extra, id_caja: m.id_caja, tipo: m.tipo, monto: m.monto,
            concepto: m.concepto, obs: m.obs, registrado_por: m.registrado_por
          });
          if(mm.ok && !mm.dedup) movEspejados++;
          else if(!mm.ok) movErrores.push(String(m.id_extra)+': '+(mm.error||''));
        } catch(e){ movErrores.push(String(m.id_extra)+': '+(e && e.message)); }
      });
    } else {
      movErrores.push('no se pudo leer me.movimientos_extra: '+(rm.error||''));
    }
  } catch(eMov){ movErrores.push('pasada movimientos: '+(eMov && eMov.message)); }

  var out = { ok:true, ventasHoyEnSupabase:(r.data||[]).length, faltabanEnSheets:faltantes.length,
              espejadas:espejadas, errores:errores,
              movHoyEnSupabase:movEnSupa, movFaltaban:movFaltaban, movEspejados:movEspejados, movErrores:movErrores };
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

// [gate-horario-ALTO] Chequeo de ventana horaria a NIVEL APP (mosExpress) usando la MISMA config central
// (HORARIOS_APPS de MOS) y la MISMA matemática cruza-medianoche que resolverHorarioPersonal. Reusa la hoja
// para no duplicar la fuente de verdad. FAIL-OPEN: ante cualquier error, config ausente, día sin configurar
// u hora inválida → devuelve true (permitir). Razón: es una app de dinero; jamás bloquear el minteo (y por
// ende las ventas) por un hipo del servicio de horario. Solo rechaza cuando hay una ventana EXPLÍCITA y
// estamos fuera de ella. Los horarios CUSTOM por operador siguen los hace cumplir el overlay del frontend
// (que sí tiene el contexto del operador logueado; el mint solo conoce el dispositivo).
var _HOR_DIAS_ME = ['lun','mar','mie','jue','vie','sab','dom'];
function _parseHoraME(s){
  var m = String(s||'').match(/^(\d{1,2}):(\d{2})$/);
  if(!m) return null;
  var hh = parseInt(m[1],10), mm = parseInt(m[2],10);
  if(hh<0||hh>23||mm<0||mm>59) return null;
  return hh + (mm/60);
}
function _horarioAppPermitidoME(mosSS){
  try {
    // Cache 120s: la ventana no cambia entre heartbeats; amortiza el read cross-spreadsheet.
    var cache = CacheService.getScriptCache();
    var ck = 'HOR_APP_ME_PERMITIDO';
    var cached = cache.get(ck);
    if(cached === '1') return true;
    if(cached === '0') return false;

    var sh = mosSS && mosSS.getSheetByName('HORARIOS_APPS');
    if(!sh) return true; // sin hoja de horarios → fail-open
    var rows = sh.getDataRange().getValues();
    if(!rows || rows.length < 2) return true;
    var hdr = rows[0];
    var iApp = hdr.indexOf('app'), iHor = hdr.indexOf('horarioJson');
    if(iApp < 0 || iHor < 0) return true;
    var horJson = null;
    for(var r=1; r<rows.length; r++){
      if(String(rows[r][iApp]) === 'mosExpress'){ horJson = rows[r][iHor]; break; }
    }
    if(!horJson) return true; // sin config para mosExpress → fail-open
    var hor = {};
    try { hor = JSON.parse(horJson) || {}; } catch(_){ return true; } // JSON corrupto → fail-open

    var tz = Session.getScriptTimeZone();
    var ahora = new Date();
    var diaIdx = parseInt(Utilities.formatDate(ahora, tz, 'u'), 10); // 1=lun..7=dom
    var diaKey = _HOR_DIAS_ME[Math.max(0, Math.min(6, diaIdx - 1))];
    var cd = hor[diaKey] || {};
    var permitido;
    if(!cd.activo){
      permitido = false; // día explícitamente cerrado
    } else {
      var apert = _parseHoraME(cd.apertura), cierre = _parseHoraME(cd.cierre);
      if(apert === null || cierre === null){
        permitido = true; // hora inválida → fail-open (igual que el motor central)
      } else {
        var horaActual = parseInt(Utilities.formatDate(ahora, tz, 'H'), 10);
        var minActual  = parseInt(Utilities.formatDate(ahora, tz, 'm'), 10);
        var horaDec = horaActual + (minActual/60);
        if(cierre > apert)      permitido = (horaDec >= apert && horaDec < cierre);     // 07:00-19:00
        else if(cierre < apert) permitido = (horaDec >= apert || horaDec < cierre);     // 14:00-02:00 cruza 00:00
        else                    permitido = false;                                      // apert === cierre
      }
    }
    cache.put(ck, permitido ? '1' : '0', 120);
    return permitido;
  } catch(e){
    return true; // cualquier error → fail-open (nunca bloquear el minteo)
  }
}

// Emite un JWT 'authenticated' scoped por zona para un dispositivo. Lo llama la PWA al iniciar + en heartbeat.
function mintSupabaseToken(deviceId){
  var idd = String(deviceId || '').trim();
  if(!idd) return { ok:false, error:'deviceId requerido' };
  var secret = PropertiesService.getScriptProperties().getProperty('SUPABASE_JWT_SECRET');
  if(!secret) return { ok:false, error:'falta SUPABASE_JWT_SECRET en Script Properties (Supabase → Settings → API → JWT Secret)' };

  // [autorización por UUID, SIN zona — los dispositivos/empleados ROTAN] Valida contra la hoja DISPOSITIVOS
  // VIVA de MOS (autoritativa, igual que verificarDispositivo), NO la sombra (que puede estar atrasada y
  // rechazar un dispositivo válido — fix del ALTO 'revocación stale' del 20×). El token está cacheado ~5min
  // → el openById se amortiza. Fail-closed: no registrado/ACTIVO → no token.
  var mosSsId = PropertiesService.getScriptProperties().getProperty('MOS_SS_ID');
  if(!mosSsId) return { ok:false, error:'falta MOS_SS_ID' };
  var mosSS, dispSheet;
  try { mosSS = SpreadsheetApp.openById(mosSsId); dispSheet = mosSS.getSheetByName('DISPOSITIVOS'); }
  catch(e){ return { ok:false, error:'no se pudo abrir DISPOSITIVOS de MOS: '+(e&&e.message) }; }
  if(!dispSheet) return { ok:false, error:'DISPOSITIVOS no disponible' };
  var datos = obtenerDatosHojaComoJSON(dispSheet), devOk = false;
  for(var di=0; di<datos.length; di++){
    var dd = datos[di];
    var idMatch  = (String(dd.ID_Dispositivo) === idd || String(dd.idDispositivo) === idd);
    var appMatch = (!dd.App || dd.App === 'mosExpress');
    var actMatch = (dd.Estado === 'ACTIVO' || dd.estado === 'ACTIVO' || dd.activo === 1 || dd.activo === '1');
    if(idMatch && appMatch && actMatch){ devOk = true; break; }
  }
  // Estado=ACTIVO ya cubre el bloqueo por UUID (bloquearDispositivosDeUsuario pone Estado='INACTIVO').
  if(!devOk) return { ok:false, error:'dispositivo no registrado/activo para mosExpress' };

  // [gate-horario-ALTO] Defense-in-depth: no emitir token fuera de la ventana horaria de la app (config
  // central, fail-open). El overlay del frontend ya bloquea; esto cierra el bypass directo a las RPC.
  if(!_horarioAppPermitidoME(mosSS)) return { ok:false, error:'fuera de horario operativo' };

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
