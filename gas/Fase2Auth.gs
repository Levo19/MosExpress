// ============================================================
// Fase2Auth.gs — mint-token: GAS emite un JWT scoped para que la PWA hable DIRECTO a Supabase.
// ============================================================
// El JWT SECRET (Supabase → Settings → API → JWT Secret) vive SOLO en GAS (Script Property
// SUPABASE_JWT_SECRET), NUNCA en el navegador. La PWA pide un token corto (exp 5min) con el claim
// 'zonas' que sale del binding admin-only mos.dispositivo_zonas. La RLS de Supabase deriva la zona
// de ESE claim (no de params del cliente → no falsificable). Re-mint en heartbeat.
// HS256 mintado en GAS es seguro: el secreto no sale de GAS; el navegador solo recibe el token corto.
// (Upgrade futuro: firma asimétrica RS256 vía Edge Function — ver MIGRACION_FASE2_PLAN.md C4.)

function _b64url_(bytes){ return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, ''); }
function _b64urlStr_(str){ return _b64url_(Utilities.newBlob(str).getBytes()); }

// Emite un JWT 'authenticated' scoped por zona para un dispositivo. Lo llama la PWA al iniciar + en heartbeat.
function mintSupabaseToken(deviceId){
  var idd = String(deviceId || '').trim();
  if(!idd) return { ok:false, error:'deviceId requerido' };
  var secret = PropertiesService.getScriptProperties().getProperty('SUPABASE_JWT_SECRET');
  if(!secret) return { ok:false, error:'falta SUPABASE_JWT_SECRET en Script Properties (Supabase → Settings → API → JWT Secret)' };

  // Zonas autoritativas desde el binding admin-only (mos.dispositivo_zonas). Fail-closed: sin zona → no token.
  var r = _sbSelect('mos.dispositivo_zonas', { id_dispositivo:'eq.'+idd, activo:'eq.true' });
  if(!r.ok) return { ok:false, error:'no se pudo leer binding dispositivo->zona: '+(r.error||'') };
  var zonas = (r.data || []).map(function(x){ return String(x.id_zona); }).filter(Boolean);
  if(!zonas.length) return { ok:false, error:'dispositivo sin zona asignada — el admin debe asignarlo en mos.dispositivo_zonas' };

  var now = Math.floor(Date.now()/1000);
  var header  = { alg:'HS256', typ:'JWT' };
  var payload = {
    iss:'supabase', role:'authenticated', aud:'authenticated', sub:idd,
    zonas: zonas, app:'mosExpress',
    iat: now, exp: now + 300   // 5 minutos (corto a propósito; re-mint en heartbeat)
  };
  var signingInput = _b64urlStr_(JSON.stringify(header)) + '.' + _b64urlStr_(JSON.stringify(payload));
  var sig = Utilities.computeHmacSha256Signature(signingInput, secret);
  var token = signingInput + '.' + _b64url_(sig);
  return { ok:true, token:token, zonas:zonas, exp:payload.exp };
}

// Wrapper de prueba para el editor (sin args): mintea para el 1er dispositivo con binding y muestra el token.
function _testMintToken(){
  var r = _sbSelect('mos.dispositivo_zonas', { activo:'eq.true' });
  if(!r.ok || !(r.data||[]).length){ Logger.log('sin dispositivos con binding'); return; }
  var dev = String(r.data[0].id_dispositivo);
  var out = mintSupabaseToken(dev);
  Logger.log('mint para '+dev+' → '+JSON.stringify({ok:out.ok, zonas:out.zonas, error:out.error}));
  if(out.ok){
    // decodificar payload para verificar (solo log, el token NO se imprime entero en prod)
    var parts = out.token.split('.');
    var payJson = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[1])).getDataAsString();
    Logger.log('payload: '+payJson);
    Logger.log('token (primeros 40): '+out.token.substring(0,40)+'...');
  }
}
