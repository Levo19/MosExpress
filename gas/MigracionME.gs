/**
 * ============================================================
 * MIGRACIÓN SUPABASE — FASE 1 · Backfill de MosExpress (esquema me)
 * ============================================================
 * Vive en el GAS de MosExpress. Requiere:
 *   - Supabase.gs (helper _sb) copiado aquí.
 *   - Script Properties: SUPABASE_URL, SUPABASE_SERVICE_KEY (legacy JWT eyJ…).
 *   - Haber corrido 01_schema_compartido.sql y 02_schema_me.sql en Supabase.
 *
 * Características:
 *   - REANUDABLE: checkpoint por tabla en Script Properties (límite 6 min GAS).
 *   - `linea` DETERMINISTA para ventas_detalle / guias_detalle (orden de hoja).
 *   - Idempotente: upsert por clave natural (on_conflict). ventas_fantasma = insert-only guardado.
 *   - dryRun para validar headers sin escribir.
 *
 * Uso (desde el editor):
 *   dryRunME()            // valida headers, no escribe
 *   backfillME()          // backfill real (re-correr hasta que todo diga ok:true)
 *   verificarCuadreME()   // compara conteos sheet vs supabase
 *   resetCheckpointsME()  // borra checkpoints (para reempezar limpio)
 */

// ---------- conversores defensivos ----------
function _meText(v){ return (v==null||v==='')?null:String(v); }
function _meNum(v){ if(v==null||v==='')return null; if(typeof v==='number')return isNaN(v)?null:v; var n=parseFloat(String(v).replace(',','.')); return isNaN(n)?null:n; }
function _meInt(v){ var n=_meNum(v); return n==null?null:Math.round(n); }
function _meDate(v){ if(v==null||v==='')return null;
  // date-only STRING → new Date lo lee como UTC y al formatear en Lima cae al día anterior. Anclar a medianoche Lima.
  if(!(v instanceof Date) && /^\d{4}-\d{2}-\d{2}$/.test(String(v).trim())) v=String(v).trim()+'T00:00:00-05:00';
  var d=(v instanceof Date)?v:new Date(v); if(isNaN(d.getTime()))return null; return Utilities.formatDate(d,'America/Lima',"yyyy-MM-dd'T'HH:mm:ssXXX"); }
function _meJson(v){ if(v==null||v==='')return null; if(typeof v==='object')return v; try{var p=JSON.parse(String(v)); return (p&&typeof p==='object')?p:null;}catch(e){return null;} }

function _meVal(raw,t){
  if(t==='text')return _meText(raw);
  if(t==='num') return _meNum(raw);
  if(t==='int') return _meInt(raw);
  if(t==='date')return _meDate(raw);
  if(t==='json')return _meJson(raw);
  return _meText(raw);
}
function _meRow(obj,spec){ var r={}; for(var i=0;i<spec.length;i++){ r[spec[i][0]]=_meVal(obj[spec[i][1]],spec[i][2]); } return r; }

function _meSheetRows(name){
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if(!sh) throw new Error('Hoja no encontrada: '+name);
  return obtenerDatosHojaComoJSON(sh);   // helper existente de ME (Code.gs)
}

// ---------- specs por tabla ----------
// flags: onConflict (pg cols), keyHeader (header para contar/filtrar), lineaBy (pg col), insertOnly, post
var _ME_SPECS = {
  ventas: { sheet:'VENTAS_CABECERA', onConflict:'id_venta', keyHeader:'ID_Venta', big:true, spec:[
    ['id_venta','ID_Venta','text'],['fecha','Fecha','date'],['vendedor','Vendedor','text'],
    ['estacion','Estacion','text'],['cliente_doc','Cliente_Doc','text'],['cliente_nombre','Cliente_Nombre','text'],
    ['total','Total','num'],['tipo_doc','Tipo_Doc','text'],['forma_pago','FormaPago','text'],
    ['correlativo','Correlativo','text'],['id_caja','ID_Caja','text'],['dispositivo_id','ID_Dispositivo','text'],
    ['estado_envio','Estado_Envio','text'],['ref_local','Ref_Local','text'],['obs','Obs','text'],
    ['tipo_doc_cliente','Tipo_Doc_Cliente','int'],['nf_estado','NF_Estado','text'],['nf_hash','NF_Hash','text'],
    ['nf_enlace','NF_Enlace','text'],['historial_cambios','historialCambios','json']
  ], post:function(r,o){
    // LAYOUT VIEJO: el pago iba embebido en Tipo_Doc "NOTA_DE_VENTA (EFECTIVO)" y NO existía
    // la columna FormaPago → todo lo posterior está corrido 1. Se detecta por el "(" en Tipo_Doc.
    var td=String(o['Tipo_Doc']||'');
    if(td.indexOf('(')>=0){
      r.tipo_doc       = td.split('(')[0].trim() || null;
      var m=td.match(/\(([^)]+)\)/); r.forma_pago = m ? m[1].trim() : null;
      r.correlativo    = _meText(o['FormaPago']);      // en viejo: correlativo
      r.id_caja        = _meText(o['Correlativo']);    // en viejo: id_caja
      r.dispositivo_id = _meText(o['ID_Caja']);        // en viejo: device
      r.estado_envio   = _meText(o['ID_Dispositivo']); // en viejo: estado
      r.ref_local=null; r.obs=null; r.tipo_doc_cliente=null;
      r.nf_estado=null; r.nf_hash=null; r.nf_enlace=null; r.historial_cambios=null;
    }
    return r;
  }},
  ventas_detalle: { sheet:'VENTAS_DETALLE', onConflict:'id_venta,linea', keyHeader:'ID_Venta', big:true, lineaBy:'id_venta', spec:[
    ['id_venta','ID_Venta','text'],['sku','SKU','text'],['nombre','Nombre','text'],['cantidad','Cantidad','num'],
    ['precio','Precio','num'],['subtotal','Subtotal','num'],['cod_barras','Cod_Barras','text'],
    ['valor_unitario','Valor_Unitario','num'],['tipo_igv','Tipo_IGV','int'],['unidad_medida','Unidad_Medida','text']
  ]},
  cajas: { sheet:'CAJAS', onConflict:'id_caja', keyHeader:'ID_Caja', spec:[
    ['id_caja','ID_Caja','text'],['vendedor','Vendedor','text'],['estacion','Estacion','text'],
    ['fecha_apertura','Fecha_Apertura','date'],['monto_inicial','Monto_Inicial','num'],['estado','Estado','text'],
    ['monto_final','Monto_Final','num'],['fecha_cierre','Fecha_cierre','date'],['zona_id','Zona_ID','text'],
    ['printnode_id','PrintNode_ID','text']
  ]},
  movimientos_extra: { sheet:'MOVIMIENTOS_EXTRA', onConflict:'id_extra', keyHeader:'ID_Extra', spec:[
    ['id_extra','ID_Extra','text'],['id_caja','ID_Caja','text'],['ts','Timestamp','date'],['tipo','Tipo','text'],
    ['monto','Monto','num'],['concepto','Concepto','text'],['obs','Obs','text'],['registrado_por','Registrado_Por','text'],
    ['historial_cambios','historialCambios','json']
  ]},
  clientes_frecuentes: { sheet:'CLIENTES_FRECUENTES', onConflict:'documento', keyHeader:'Documento', spec:[
    ['documento','Documento','text'],['nombre','Nombre_RazonSocial','text'],['tipo_doc','Tipo','text'],
    ['fecha_registro','Fecha','date'],['direccion','Direccion','text'],['historial_cambios','historialCambios','json']
  ] },
  guias_cabecera: { sheet:'GUIAS_CABECERA', onConflict:'id_guia', keyHeader:'ID_Guia', spec:[
    ['id_guia','ID_Guia','text'],['fecha','Fecha','date'],['vendedor','Vendedor','text'],['zona_id','Zona_ID','text'],
    ['tipo','Tipo','text'],['observacion','Observacion','text'],['zona_destino','Zona_Destino','text'],['estado','Estado','text']
  ]},
  guias_detalle: { sheet:'GUIAS_DETALLE', onConflict:'id_guia,linea', keyHeader:'ID_Guia', lineaBy:'id_guia', spec:[
    ['id_guia','ID_Guia','text'],['cod_barras','Cod_Barras','text'],['cantidad','Cantidad','num']
  ]},
  correlativos: { sheet:'CORRELATIVOS', onConflict:'serie', keyHeader:'Serie', spec:[
    ['serie','Serie','text'],['siguiente','Siguiente','int']
  ]},
  reservas_correlativos: { sheet:'RESERVAS_CORRELATIVOS', onConflict:'id_reserva', keyHeader:'idReserva', spec:[
    ['id_reserva','idReserva','text'],['serie','serie','text'],['numero','numero','int'],['vendedor','vendedor','text'],
    ['dispositivo_id','deviceId','text'],['reservado_at','reservadoAt','date'],['estado','estado','text'],
    ['usado_at','usadoAt','date'],['id_venta','idVenta','text']
  ]},
  creditos_cobro_asignado: { sheet:'CREDITOS_COBRO_ASIGNADO', onConflict:'id_cobro', keyHeader:'ID_Cobro', spec:[
    ['id_cobro','ID_Cobro','text'],['id_venta','ID_Venta','text'],['caja_destino','Caja_Destino','text'],
    ['vendedor_dest','Vendedor_Dest','text'],['metodo_sug','Metodo_Sug','text'],['estado','Estado','text'],
    ['admin_asignador','Admin_Asignador','text'],['fecha_asig','Fecha_Asig','date'],['fecha_res','Fecha_Res','date'],
    ['razon','Razon','text'],['id_caja_origen','ID_Caja_Origen','text'],['monto','Monto','num'],
    ['cliente_nombre','Cliente_Nombre','text'],['correlativo','Correlativo','text'],['fecha_vencimiento','Fecha_Vencimiento','date'],
    ['horas_ttl','Horas_TTL','int'],['mensaje_admin','Mensaje_Admin','text'],['reasignaciones','Reasignaciones','int']
  ]},
  ventas_fantasma: { sheet:'VENTAS_FANTASMA', insertOnly:true, keyHeader:'ts', spec:[
    ['ts','ts','date'],['vendedor','vendedor','text'],['zona_id','zona','text'],['estacion','estacion','text'],
    ['dispositivo_id','deviceId','text'],['monto','monto','num'],['metodo','metodo','text'],['tipo_doc','tipoDoc','text'],
    ['doc_cliente','docCliente','text'],['nombre_cliente','nombreCliente','text'],['correlativo_local','correlativoLocal','text'],
    ['caja_id_enviada','cajaIdEnviada','text'],['motivo','motivo','text'],['mensaje','mensaje','text'],
    ['estado_revision','estado_revision','text'],['revisado_por','revisadoPor','text'],['fecha_revision','fechaRevision','date'],
    ['accion_tomada','accionTomada','text'],['payload_json','payload_json','json']
  ]},
  auditorias: { sheet:'AUDITORIAS', onConflict:'id_auditoria,cod_barras', keyHeader:'ID_Auditoria', spec:[
    ['id_auditoria','ID_Auditoria','text'],['fecha','Fecha','date'],['vendedor','Vendedor','text'],['zona_id','Zona_ID','text'],
    ['cod_barras','Cod_Barras','text'],['cant_sistema','Cant_Sistema','num'],['cant_real','Cant_Real','num'],['diferencia','Diferencia','num']
  ]},
  caja_alertas_efectivo: { sheet:'CAJA_ALERTAS_EFECTIVO', onConflict:'id_caja', keyHeader:'idCaja', spec:[
    ['id_caja','idCaja','text'],['bandera','bandera','text'],['monto_ultimo','montoUltimo','num'],['fecha_actualizada','fechaActualizada','date']
  ]},
  pickups_pendientes_envio: { sheet:'PICKUPS_PENDIENTES_ENVIO', onConflict:'id_guia_me', keyHeader:'idGuiaME', spec:[
    ['id_guia_me','idGuiaME','text'],['payload','payload','json'],['intentos','intentos','int'],
    ['ultimo_intento','ultimoIntento','date'],['ultimo_error','ultimoError','text'],['estado','estado','text']
  ]},
  stock_zonas: { sheet:'STOCK_ZONAS', onConflict:'cod_barras,zona_id', keyHeader:'Cod_Barras', spec:[
    ['cod_barras','Cod_Barras','text'],['zona_id','Zona_ID','text'],['cantidad','Cantidad','num'],
    ['usuario','Usuario','text'],['fecha_ultimo_registro','Fecha_Ultimo_Registro','date']
  ]},
  radio_config: { sheet:'RadioConfig', onConflict:'tipo,key', keyHeader:'Tipo', spec:[
    ['tipo','Tipo','text'],['key','Key','text'],['valor','Valor','text']
  ]}
};

var _ME_TIME_BUDGET = 4.5*60*1000;   // < 6 min límite GAS
var _ME_BATCH = 100;

/** Construye las filas pg de una tabla (mapeo + linea + dedupe + filtro). */
function _meBuildRows(tabla){
  var cfg=_ME_SPECS[tabla];
  var objs=_meSheetRows(cfg.sheet);
  var rows=objs.map(function(o){ var r=_meRow(o,cfg.spec); if(cfg.post) r=cfg.post(r,o); return r; });

  if(cfg.lineaBy){ // linea determinista por grupo (orden de hoja)
    var cnt={};
    rows.forEach(function(r){ var k=String(r[cfg.lineaBy]); cnt[k]=(cnt[k]||0)+1; r.linea=cnt[k]; });
    rows=rows.filter(function(r){ return r[cfg.lineaBy]!=null && r[cfg.lineaBy]!==''; });
  } else if(!cfg.insertOnly){ // pk simple O COMPUESTO: filtra sin pk + dedupe (gana el último)
    var pkCols=String(cfg.onConflict).split(',').map(function(c){ return c.trim(); });
    rows=rows.filter(function(r){ return pkCols.every(function(c){ return r[c]!=null && r[c]!==''; }); });
    var seen={}; rows.forEach(function(r){ var k=pkCols.map(function(c){ return String(r[c]); }).join('||'); seen[k]=r; });
    rows=Object.keys(seen).map(function(k){ return seen[k]; });
  }
  return rows;
}

// [ventas-directo / dual-write] Espeja UNA cabecera de venta a me.ventas EN TIEMPO REAL,
// reusando el MISMO mapeo del batch (_meRow + _ME_SPECS.ventas) → fila byte-idéntica a la que
// produciría el sync. Upsert por id_venta (PK) = idempotente (si el batch ya la subió, actualiza).
// `o` = objeto keyed por cabeceras de hoja (ID_Venta, Fecha, Vendedor, ...). Best-effort: el caller
// lo envuelve en try/catch; NUNCA debe romper la venta (Sheets sigue siendo la fuente de verdad).
// Solo cabecera (Fase A): el detalle sigue por sync batch (≤15min, tolerado para COGS).
function _dualWriteVentaME(o){
  var cfg=_ME_SPECS.ventas;
  var row=_meRow(o, cfg.spec);
  if(cfg.post) row=cfg.post(row, o);
  if(row.id_venta==null || row.id_venta===''){ Logger.log('[dualWrite venta] sin id_venta — omitido'); return {ok:false, error:'sin id_venta'}; }
  // 1 SOLO intento (maxRetry:1, sin backoff): es el hot path de la venta → no colgar al cajero
  // si Supabase está degradado. Si falla, el sync batch (≤15min) reconcilia. Idempotente por id_venta.
  var r=_sb('POST', 'me.ventas', { data:[row], upsert:true, onConflict:cfg.onConflict, maxRetry:1 });
  if(!r.ok) Logger.log('[dualWrite venta] '+row.id_venta+' upsert falló: HTTP '+(r.code)+' '+(r.error||''));
  return r;
}

// [ventas-directo / dual-write detalle] Espeja las líneas de UNA venta a me.ventas_detalle en tiempo real.
// Reusa _meRow + _ME_SPECS.ventas_detalle; linea = orden de items (idéntico a como el sheet las escribe →
// el batch asigna la misma linea por orden de hoja). Upsert por (id_venta,linea) = idempotente. 1 intento,
// best-effort (el caller envuelve en try/catch; nunca rompe la venta). Replica el cálculo de valor_unitario
// del sheet (Ventas.gs). Solo cabecera+detalle en tiempo real; el resto sigue por sync batch.
function _dualWriteDetalleME(idVenta, items){
  if(!items || !items.length) return {ok:true, vacio:true};
  var cfg=_ME_SPECS.ventas_detalle;
  var idv=String(idVenta||'').trim();
  if(!idv) return {ok:false, error:'sin id_venta'};
  var rows=items.map(function(item, idx){
    var vu = parseFloat(item.valor_unitario) || Math.round(parseFloat(item.precio||0)/1.18*100)/100;
    var o = {
      ID_Venta: idv, SKU: String(item.sku||'').trim(), Nombre: item.nombre,
      Cantidad: item.cantidad, Precio: item.precio, Subtotal: item.subtotal,
      Cod_Barras: String(item.codBarras||'').trim(),
      Valor_Unitario: Math.round(vu*100)/100, Tipo_IGV: parseInt(item.tipo_igv||1,10),
      Unidad_Medida: String(item.unidad_de_medida||'NIU')
    };
    var r=_meRow(o, cfg.spec); r.linea = idx+1; return r;
  });
  var res=_sb('POST','me.ventas_detalle',{ data:rows, upsert:true, onConflict:cfg.onConflict, maxRetry:1 });
  if(!res.ok) Logger.log('[dualWrite detalle] '+idv+' x'+rows.length+' falló: HTTP '+(res.code)+' '+(res.error||''));
  return res;
}

// [cajas-directo / dual-write] Espeja UNA caja a me.cajas en tiempo real (apertura Y cierre), reusando el
// mapeo del batch (_meRow + _ME_SPECS.cajas). Upsert por id_caja → la apertura inserta, el cierre actualiza
// la MISMA fila. 1 intento, best-effort (el caller envuelve en try/catch; Sheets=verdad, el batch reconcilia).
// `o` = objeto keyed por cabeceras de hoja (ID_Caja, Vendedor, ..., Fecha_cierre, Zona_ID, PrintNode_ID).
function _dualWriteCajaME(o){
  var cfg=_ME_SPECS.cajas;
  var row=_meRow(o, cfg.spec);
  if(cfg.post) row=cfg.post(row, o);
  if(row.id_caja==null || row.id_caja===''){ Logger.log('[dualWrite caja] sin id_caja — omitido'); return {ok:false, error:'sin id_caja'}; }
  var r=_sb('POST','me.cajas',{ data:[row], upsert:true, onConflict:cfg.onConflict, maxRetry:1 });
  if(!r.ok) Logger.log('[dualWrite caja] '+row.id_caja+' upsert falló: HTTP '+(r.code)+' '+(r.error||''));
  return r;
}

// [movimientos-directo / dual-write] Espeja UN movimiento de caja a me.movimientos_extra en tiempo real
// (ingreso/egreso). Reusa el mapeo del batch (_meRow + _ME_SPECS.movimientos_extra). Upsert por id_extra
// = idempotente. 1 intento, best-effort (Sheets=verdad, el batch reconcilia). `o` keyed por cabeceras
// (ID_Extra, ID_Caja, Timestamp, Tipo, Monto, Concepto, Obs, Registrado_Por).
function _dualWriteMovExtraME(o){
  var cfg=_ME_SPECS.movimientos_extra;
  var row=_meRow(o, cfg.spec);
  if(cfg.post) row=cfg.post(row, o);
  if(row.id_extra==null || row.id_extra===''){ Logger.log('[dualWrite movExtra] sin id_extra — omitido'); return {ok:false, error:'sin id_extra'}; }
  var r=_sb('POST','me.movimientos_extra',{ data:[row], upsert:true, onConflict:cfg.onConflict, maxRetry:1 });
  if(!r.ok) Logger.log('[dualWrite movExtra] '+row.id_extra+' upsert falló: HTTP '+(r.code)+' '+(r.error||''));
  return r;
}

// [ventas-directo / patch] Actualiza campos puntuales de ventas YA existentes en me.ventas en tiempo real
// (ej. forma_pago='ANULADO' al anular). PATCH parcial (NO upsert → no choca con NOT NULL ni inserta filas
// incompletas; si la venta aún no está en la sombra es no-op y el batch la sube luego con el valor nuevo).
// idVenta: string (filtro eq.) o array (in.(...) → 1 sola llamada para anulación masiva). 1 intento, best-effort.
function _dualWriteVentaPatchME(idVenta, patch){
  var ids = Array.isArray(idVenta)
    ? idVenta.map(function(x){ return String(x||'').trim(); }).filter(Boolean)
    : [String(idVenta||'').trim()].filter(Boolean);
  if(!ids.length || !patch) return {ok:false, error:'sin id_venta/patch'};
  var filtro = (ids.length===1) ? ('eq.'+ids[0]) : ('in.('+ids.join(',')+')');
  var r=_sb('PATCH','me.ventas',{ data:patch, filters:{ id_venta:filtro }, maxRetry:1 });
  if(!r.ok) Logger.log('[dualWrite venta patch] '+ids.length+' id(s) falló: HTTP '+(r.code)+' '+(r.error||''));
  return r;
}

// [creditos-directo / dual-write] Espeja UN cobro asignado a me.creditos_cobro_asignado en tiempo real
// (al CREAR la asignación). Reusa el mapeo del batch (_meRow + _ME_SPECS.creditos_cobro_asignado). Upsert
// por id_cobro. 1 intento, best-effort. Las transiciones de estado (cobrado/rechazado/etc.) siguen por
// batch+dirty-sync (tabla chica, sync completo ≤15min). `o` keyed por cabeceras (ID_Cobro, ID_Venta, ...).
function _dualWriteCobroME(o){
  var cfg=_ME_SPECS.creditos_cobro_asignado;
  var row=_meRow(o, cfg.spec);
  if(cfg.post) row=cfg.post(row, o);
  if(row.id_cobro==null || row.id_cobro===''){ Logger.log('[dualWrite cobro] sin id_cobro — omitido'); return {ok:false, error:'sin id_cobro'}; }
  var r=_sb('POST','me.creditos_cobro_asignado',{ data:[row], upsert:true, onConflict:cfg.onConflict, maxRetry:1 });
  if(!r.ok) Logger.log('[dualWrite cobro] '+row.id_cobro+' upsert falló: HTTP '+(r.code)+' '+(r.error||''));
  return r;
}

// [creditos-directo / patch] Actualiza el estado de un cobro YA existente en me.creditos_cobro_asignado en
// tiempo real (transiciones cobrado/rechazado/expirado/cancelado/reasignado). PATCH parcial por id_cobro
// (NO upsert; si no está en la sombra, no-op + batch lo sube). Formatea fecha_* a ISO Lima como el batch.
// 1 intento, best-effort. Con esto cobros_en_vuelo/creditos_pendientes quedan real-time (flipeables).
function _dualWriteCobroPatchME(idCobro, patch){
  var idc=String(idCobro||'').trim();
  if(!idc || !patch) return {ok:false, error:'sin id_cobro/patch'};
  var p={}; Object.keys(patch).forEach(function(k){ p[k]=(k.indexOf('fecha')===0 && patch[k]) ? _meDate(patch[k]) : patch[k]; });
  var r=_sb('PATCH','me.creditos_cobro_asignado',{ data:p, filters:{ id_cobro:'eq.'+idc }, maxRetry:1 });
  if(!r.ok) Logger.log('[dualWrite cobro patch] '+idc+' falló: HTTP '+(r.code)+' '+(r.error||''));
  return r;
}

// [Fase B] Resuelve Ref_Local DUPLICADOS en VENTAS_CABECERA (ventas dobles pre-C9). Conserva el Ref_Local
// de la PRIMERA fila de cada grupo y BLANQUEA el de las siguientes → preserva la fila/venta y su correlativo,
// solo libera la clave para poder crear el índice único parcial en me.ventas.ref_local. dryRun por defecto;
// `resolverDupsRefLocalME(true)` aplica. Devuelve los cambios (id_venta+ref_local) para trazabilidad.
function resolverDupsRefLocalME(aplicar){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sh=ss.getSheetByName('VENTAS_CABECERA');
  if(!sh) return {ok:false, error:'sin hoja VENTAS_CABECERA'};
  var d=sh.getDataRange().getValues(), head=d[0];
  var iRL=head.indexOf('Ref_Local'), iID=head.indexOf('ID_Venta'), iCorr=head.indexOf('Correlativo');
  if(iRL<0) return {ok:false, error:'sin columna Ref_Local'};
  var visto={}, cambios=[];
  for(var r=1;r<d.length;r++){
    var rl=String(d[r][iRL]||'').trim();
    if(!rl) continue;
    if(visto[rl]) cambios.push({fila:r+1, id_venta:String(d[r][iID]||''), correlativo:(iCorr>=0?String(d[r][iCorr]||''):''), ref_local:rl});
    else visto[rl]=true;
  }
  if(aplicar){
    cambios.forEach(function(c){
      // defensivo: reconfirmar que la fila sigue siendo esa venta (que no se corrió por un append/edición
      // concurrente) ANTES de blanquear — evita blanquear el Ref_Local de la venta equivocada.
      var idEnFila = (iID>=0) ? String(sh.getRange(c.fila, iID+1).getValue()||'') : c.id_venta;
      if(idEnFila === c.id_venta){ sh.getRange(c.fila, iRL+1).setValue(''); c.aplicado=true; }
      else { c.aplicado=false; c.nota='fila corrida (id no calza: '+idEnFila+') — OMITIDO'; Logger.log('[resolverDups] fila '+c.fila+' OMITIDA: id no calza ('+idEnFila+' != '+c.id_venta+')'); }
    });
    SpreadsheetApp.flush();
  }
  Logger.log((aplicar?'✅ APLICADO':'🔎 DRY-RUN (corré aplicarResolverDupsRefLocalME para aplicar)')+' — dups Ref_Local en hoja: '+cambios.length+'\n'+JSON.stringify(cambios,null,2));
  return {ok:true, aplicado:!!aplicar, dups:cambios.length, cambios:cambios};
}
// Wrapper para correr desde el desplegable del editor (que solo invoca funciones sin argumentos).
function aplicarResolverDupsRefLocalME(){ return resolverDupsRefLocalME(true); }

/** Backfill principal. opts: {dryRun, soloTabla} */
function migrarME(opts){
  opts=opts||{};
  var props=PropertiesService.getScriptProperties();
  var t0=Date.now();
  // orden explícito (cabeceras antes que detalles para una futura FK)
  var _ORDEN=['ventas','ventas_detalle','cajas','movimientos_extra','clientes_frecuentes',
    'guias_cabecera','guias_detalle','correlativos','reservas_correlativos','creditos_cobro_asignado',
    'ventas_fantasma','auditorias','caja_alertas_efectivo','pickups_pendientes_envio','stock_zonas','radio_config'];
  var tablas=opts.soloTabla?[opts.soloTabla]:_ORDEN;
  var resumen={};

  for(var ti=0; ti<tablas.length; ti++){
    var tabla=tablas[ti], cfg=_ME_SPECS[tabla];
    if(!cfg){ resumen[tabla]={error:'spec desconocida'}; continue; }
    try{
      // hoja inexistente (ej. VENTAS_FANTASMA antes del 1er rechazo) → saltar limpio
      if(!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg.sheet)){ resumen[tabla]={saltado:'hoja no existe: '+cfg.sheet}; continue; }
      // saltar tablas ya completadas (reanudación multi-tabla eficiente; evita re-leer)
      if(!opts.dryRun && !opts.soloTabla && props.getProperty('MEBF_DONE_'+tabla)==='1'){
        resumen[tabla]={saltado:'ya completada (resetCheckpointsME para rehacer)'}; continue;
      }
      var rows=_meBuildRows(tabla);

      if(opts.dryRun){ resumen[tabla]={dryRun:true, filasValidas:rows.length, muestra:rows[0]||null}; continue; }

      // insert-only (ventas_fantasma): si NO hay checkpoint activo y ya hay filas → completado antes, saltar.
      // Si hay checkpoint activo → cae al bucle normal y REANUDA (no salta el resto).
      if(cfg.insertOnly && props.getProperty('MEBF_'+tabla)==null){
        var n=_sbCount('me.'+tabla,null);
        if(n>0){ resumen[tabla]={insertOnly:true, saltado:true, yaEnSupabase:n}; continue; }
      }

      var ckKey='MEBF_'+tabla;
      var start=parseInt(props.getProperty(ckKey)||'0',10);
      var errores=[], upserted=0, corto=false;
      for(var i=start; i<rows.length; i+=_ME_BATCH){
        if(Date.now()-t0 > _ME_TIME_BUDGET){
          props.setProperty(ckKey,String(i));
          resumen[tabla]={incompleto:true, desde:i, total:rows.length, nota:'re-corre backfillME para continuar'};
          Logger.log(JSON.stringify(resumen,null,2));
          return resumen;
        }
        var lote=rows.slice(i,i+_ME_BATCH);
        if(JSON.stringify(lote).length>10000000){ errores.push('lote '+i+': payload muy grande, omitido'); props.setProperty(ckKey,String(i+_ME_BATCH)); continue; }
        var r=cfg.insertOnly ? _sbInsert('me.'+tabla,lote) : _sbUpsert('me.'+tabla,lote,cfg.onConflict);
        if(r.ok){ upserted+=lote.length; props.setProperty(ckKey,String(i+_ME_BATCH)); }   // checkpoint SOLO en éxito
        else { errores.push('lote '+i+': HTTP '+r.code+' '+(r.error||'')); corto=true; break; }  // no avanza → reintenta este lote al re-correr
      }
      if(errores.length===0){ props.deleteProperty(ckKey); props.setProperty('MEBF_DONE_'+tabla,'1'); }
      resumen[tabla]={filas:rows.length, upserted:upserted, errores:errores, ok:errores.length===0, incompleto:corto};
    }catch(e){ resumen[tabla]={error:String(e&&e.message||e)}; }
  }
  Logger.log(JSON.stringify(resumen,null,2));
  return resumen;
}

/** Compara conteos sheet vs supabase. */
function verificarCuadreME(){
  var out={};
  Object.keys(_ME_SPECS).forEach(function(tabla){
    var cfg=_ME_SPECS[tabla], nSheet=-1;
    try{
      var objs=_meSheetRows(cfg.sheet);
      nSheet=objs.filter(function(o){ return o[cfg.keyHeader]!=null && o[cfg.keyHeader]!==''; }).length;
    }catch(e){ nSheet=-1; }
    var nPg=_sbCount('me.'+tabla,null);
    out[tabla]={sheet:nSheet, supabase:nPg, cuadra:(nSheet===nPg)};
  });
  Logger.log(JSON.stringify(out,null,2));
  return out;
}

/** Diagnóstico de conexión para ME (lee me.correlativos; _sbPing de MOS apunta a mos.*). */
function _sbPingME(){
  var out={ok:false, pasos:[]};
  try{ var cfg=_sbCfg_(); out.pasos.push('✓ Credenciales presentes ('+cfg.url+')'); }
  catch(e){ out.error=String(e.message); out.pasos.push('✗ '+out.error); Logger.log(JSON.stringify(out,null,2)); return out; }
  var t0=new Date().getTime();
  var r=_sbSelect('me.correlativos',{select:'serie',limit:1});
  out.latencia_ms=new Date().getTime()-t0;
  if(r.ok){ out.ok=true; out.pasos.push('✓ GET me.correlativos OK ('+out.latencia_ms+' ms, HTTP '+r.code+') — vacío [] es normal antes del backfill'); }
  else{ out.pasos.push('✗ GET me.correlativos falló: HTTP '+r.code+' — '+(r.error||'')); out.pasos.push('  Revisa: esquema me expuesto · 02_schema_me.sql corrido · service_role key'); }
  Logger.log(JSON.stringify(out,null,2)); return out;
}

/** Vuelca la FILA 1 (headers reales) de cada hoja de ME. Diagnóstico para alinear el backfill. */
function dumpHeadersME(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var nombres=['VENTAS_CABECERA','VENTAS_DETALLE','CAJAS','MOVIMIENTOS_EXTRA','CLIENTES_FRECUENTES',
    'GUIAS_CABECERA','GUIAS_DETALLE','CORRELATIVOS','RESERVAS_CORRELATIVOS','CREDITOS_COBRO_ASIGNADO',
    'VENTAS_FANTASMA','AUDITORIAS','CAJA_ALERTAS_EFECTIVO','PICKUPS_PENDIENTES_ENVIO','STOCK_ZONAS','RadioConfig'];
  var out={};
  nombres.forEach(function(n){
    var sh=ss.getSheetByName(n);
    if(!sh){ out[n]='(NO EXISTE)'; return; }
    var lc=sh.getLastColumn();
    if(lc<1){ out[n]='(vacía)'; return; }
    out[n]=sh.getRange(1,1,1,lc).getValues()[0];
  });
  Logger.log(JSON.stringify(out,null,2));
  return out;
}

/** Inspecciona una hoja: encabezados + primeras y últimas filas CRUDAS (para ver evolución de layout). */
function inspeccionarME(nombre, n){
  n=n||3;
  var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre);
  if(!sh){ var e={error:'NO EXISTE: '+nombre}; Logger.log(JSON.stringify(e)); return e; }
  var lc=sh.getLastColumn(), lr=sh.getLastRow();
  var out={ hoja:nombre, columnas:lc, filas:(lr-1),
    headers: sh.getRange(1,1,1,lc).getValues()[0],
    primeras: lr>1 ? sh.getRange(2,1,Math.min(n,lr-1),lc).getValues() : [],
    ultimas:  lr>1 ? sh.getRange(Math.max(2,lr-n+1),1,Math.min(n,lr-1),lc).getValues() : []
  };
  Logger.log(JSON.stringify(out,null,2));
  return out;
}

/** Un solo clic: encabezados + 2 filas viejas y 2 nuevas de las tablas clave. */
function inspeccionarTodoME(){
  var tablas=['VENTAS_CABECERA','CAJAS','CREDITOS_COBRO_ASIGNADO','CLIENTES_FRECUENTES','MOVIMIENTOS_EXTRA','RESERVAS_CORRELATIVOS','STOCK_ZONAS','RadioConfig'];
  var out={};
  tablas.forEach(function(n){
    var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(n);
    if(!sh){ out[n]='(NO EXISTE)'; return; }
    var lc=sh.getLastColumn(), lr=sh.getLastRow();
    out[n]={ columnas:lc, filas:(lr-1),
      headers: lc>0 ? sh.getRange(1,1,1,lc).getValues()[0] : [],
      primeras: lr>1 ? sh.getRange(2,1,Math.min(2,lr-1),lc).getValues() : [],
      ultimas:  lr>2 ? sh.getRange(lr-1,1,2,lc).getValues() : []
    };
  });
  Logger.log(JSON.stringify(out,null,2));
  return out;
}

// ============================================================
// FASE 1.C — Doble escritura vía SYNC INCREMENTAL en segundo plano.
// No toca los endpoints de venta (cero latencia/riesgo al cajero).
// Re-upsertea (idempotente) las filas RECIENTES de cada tabla; las nuevas
// ventas/cajas se agregan al final → quedan dentro de la "cola" sincronizada.
// ============================================================
var _ME_SYNC_TAILS = {   // cuántas filas recientes re-sincronizar (tablas chicas = todas)
  ventas:500, ventas_detalle:1500, cajas:80, movimientos_extra:300,
  clientes_frecuentes:99999, guias_cabecera:300, guias_detalle:1200,
  correlativos:99999, reservas_correlativos:500, creditos_cobro_asignado:99999,
  auditorias:500, caja_alertas_efectivo:99999, pickups_pendientes_envio:99999,
  stock_zonas:99999, radio_config:99999
};

// [fix C2] Cola de filas EDITADAS (cobrar/anular/editar venta vieja cambia FormaPago in-place).
// auditarLog (chokepoint de toda edición auditada) marca la pk; _syncMEImpl re-sincroniza las que
// caen fuera del tail → getFinanzasRango (flip) ya no sub-cuenta el ingreso cobrado hasta las 3am.
var _ME_AUD_TO_SPEC = { 'VENTAS_CABECERA':'ventas', 'MOVIMIENTOS_EXTRA':'movimientos_extra', 'CLIENTES_FRECUENTES':'clientes_frecuentes' };
function _meDirtyGet(spec){ try{ var a=JSON.parse(PropertiesService.getScriptProperties().getProperty('ME_DIRTY_'+spec)||'[]'); return Array.isArray(a)?a:[]; }catch(e){ return []; } }
function _meMarcarDirtySync(auditTabla, pk){
  try{
    var spec=_ME_AUD_TO_SPEC[auditTabla]; if(!spec || pk==null || String(pk)==='') return;
    var arr=_meDirtyGet(spec), s=String(pk);
    if(arr.indexOf(s)<0) arr.push(s);
    if(arr.length>500) arr=arr.slice(arr.length-500);
    PropertiesService.getScriptProperties().setProperty('ME_DIRTY_'+spec, JSON.stringify(arr));
  }catch(e){}
}
function _meDirtyRemove(spec, quitar){
  try{
    if(!quitar||!quitar.length) return;
    var rm={}; quitar.forEach(function(k){ rm[String(k)]=1; });
    var nuevo=_meDirtyGet(spec).filter(function(k){ return !rm[String(k)]; });
    PropertiesService.getScriptProperties().setProperty('ME_DIRTY_'+spec, JSON.stringify(nuevo));
  }catch(e){}
}

function _syncMEImpl(full){
  var resumen={};
  Object.keys(_ME_SPECS).forEach(function(tabla){
    var cfg=_ME_SPECS[tabla];
    // [correlativo Supabase] si Postgres es el minter, NO re-sincronizar el contador desde Sheets
    // (lo revertiría a un valor viejo → duplicado/hueco SUNAT). El espejo Sheets lo mantiene el write-back.
    if(tabla==='correlativos' && _fuenteCorrelativo()==='supabase'){ resumen[tabla]={skipped:'minter=supabase'}; return; }
    if(cfg.insertOnly){
      // append-only (auditoría de rechazos): la hoja SOLO crece → sincronizar la cola nueva.
      // sb_n = filas ya en Supabase actúa de checkpoint natural; insertar rows.slice(sb_n) es
      // exactamente lo que falta. Idempotente (un re-run ya no re-inserta) y sin duplicar.
      try{
        if(!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg.sheet)){ return; }
        var rowsIO=_meBuildRows(tabla);
        var sbN=_sbCount('me.'+tabla);
        if(sbN<0){ resumen[tabla]={sync:0, de:0, insertOnly:true, errores:['_sbCount falló; no se insertó (evita duplicar)']}; return; }
        if(rowsIO.length>sbN){
          var nuevas=rowsIO.slice(sbN), errIO=[], upIO=0;
          for(var j=0;j<nuevas.length;j+=100){
            var loteIO=nuevas.slice(j,j+100);
            var rIO=_sbInsert('me.'+tabla,loteIO);
            if(rIO.ok) upIO+=loteIO.length; else errIO.push('lote '+j+': HTTP '+rIO.code+' '+(rIO.error||''));
          }
          resumen[tabla]={sync:upIO, de:nuevas.length, insertOnly:true, errores:errIO};
        } else {
          resumen[tabla]={sync:0, de:0, insertOnly:true, errores:[]};
        }
      }catch(e){ resumen[tabla]={error:String(e&&e.message||e)}; }
      return;
    }
    try{
      if(!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg.sheet)){ return; }
      var rows=_meBuildRows(tabla);
      var tail=_ME_SYNC_TAILS[tabla]||300;
      var slice = (full || rows.length<=tail) ? rows : rows.slice(rows.length-tail);
      // [fix C2] sumar filas EDITADAS que quedaron fuera del tail (marcadas por auditarLog).
      var dirtyProc=_meDirtyGet(tabla);
      if(dirtyProc.length && !full && rows.length>tail){
        var inSlice={}; slice.forEach(function(r){ inSlice[String(r[cfg.onConflict])]=1; });
        var extra=rows.filter(function(r){ var k=String(r[cfg.onConflict]); return dirtyProc.indexOf(k)>=0 && !inSlice[k]; });
        if(extra.length) slice=slice.concat(extra);
      }
      var err=[], up=0;
      for(var i=0;i<slice.length;i+=100){
        var lote=slice.slice(i,i+100);
        var r=_sbUpsert('me.'+tabla,lote,cfg.onConflict);
        if(r.ok) up+=lote.length; else err.push('lote '+i+': HTTP '+r.code+' '+(r.error||''));
      }
      if(!err.length && dirtyProc.length) _meDirtyRemove(tabla, dirtyProc);  // ya sincronizadas (tail o extra); preserva nuevas
      resumen[tabla]={sync:up, de:slice.length, errores:err};
    }catch(e){ resumen[tabla]={error:String(e&&e.message||e)}; }
  });
  Logger.log(JSON.stringify(resumen,null,2));
  return resumen;
}
function syncMEReciente(){ return _syncMEImpl(false); }  // 15 min: solo cola reciente (barato)
function syncMECompleto(){ var r=_syncMEImpl(true); try{ reconciliarDiarioME(); }catch(e){ Logger.log('recon ME falló: '+e); } return r; }   // recon pegada al sync nocturno (sin trigger extra)

/** Instala (idempotente) AMBOS triggers: incremental 15 min + completo nocturno (3am). Ejecutar 1 vez. */
function instalarTriggersSyncME(){
  ScriptApp.getProjectTriggers().forEach(function(t){
    var h=t.getHandlerFunction(); if(h==='syncMEReciente'||h==='syncMECompleto') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncMEReciente').timeBased().everyMinutes(15).create();
  ScriptApp.newTrigger('syncMECompleto').timeBased().everyDays(1).atHour(3).create();
  Logger.log('Triggers instalados: syncMEReciente (15min) + syncMECompleto (3am)');
  return {ok:true};
}
function desinstalarTriggersSyncME(){
  var n=0; ScriptApp.getProjectTriggers().forEach(function(t){
    var h=t.getHandlerFunction(); if(h==='syncMEReciente'||h==='syncMECompleto'){ ScriptApp.deleteTrigger(t); n++; }
  });
  return {ok:true, eliminados:n};
}

// ---------- wrappers para el editor ----------
function dryRunME(){ return migrarME({dryRun:true}); }
function backfillME(){ return migrarME(); }
function backfillStockZonas(){ return migrarME({soloTabla:'stock_zonas'}); }   // re-hace solo esta (soloTabla ignora DONE)
function backfillRadio(){ return migrarME({soloTabla:'radio_config'}); }
function backfillAuditorias(){ return migrarME({soloTabla:'auditorias'}); }    // tras el ALTER de PK compuesta
function resetCheckpointsME(){
  var props=PropertiesService.getScriptProperties();
  var n=0; Object.keys(_ME_SPECS).forEach(function(t){
    ['MEBF_'+t,'MEBF_DONE_'+t].forEach(function(k){ if(props.getProperty(k)!=null){ props.deleteProperty(k); n++; } });
  });
  Logger.log('Checkpoints/flags borrados: '+n); return {ok:true, borrados:n};
}

// ============================================================
// RECONCILIACIÓN v2 — drift dashboard (conteo + SUMA de columnas clave)
// Detecta drift de VALORES (ediciones/anulaciones) que el solo conteo no ve. 100% lectura.
// ============================================================
var _ME_SUMCOLS = {
  ventas:['total'], ventas_detalle:['subtotal'], cajas:['monto_final'], movimientos_extra:['monto'],
  clientes_frecuentes:[], guias_cabecera:[], guias_detalle:['cantidad'], correlativos:['siguiente'],
  reservas_correlativos:['numero'], creditos_cobro_asignado:['monto'], ventas_fantasma:['monto'],
  auditorias:['diferencia'], caja_alertas_efectivo:['monto_ultimo'], pickups_pendientes_envio:[],
  stock_zonas:['cantidad'], radio_config:[]
};

/** Suma columnas de una tabla de Supabase, paginando ordenado por PK (estable). */
function _sbSumCols(schemaTable, cols, order){
  var sums={}; cols.forEach(function(c){ sums[c]=0; });
  var n=0, offset=0, PAGE=1000;
  while(true){
    var r=_sbSelect(schemaTable,{select:cols.join(',')||order.split(',')[0], order:order, limit:PAGE, offset:offset});
    if(!r.ok) return {error:'HTTP '+r.code+' '+(r.error||'')};
    var rows=r.data||[];
    rows.forEach(function(row){ cols.forEach(function(c){ var num=parseFloat(row[c]); if(!isNaN(num)) sums[c]+=num; }); });   // numeric puede venir como string desde PostgREST
    n+=rows.length;
    if(rows.length<PAGE) break;
    offset+=PAGE;
  }
  return {n:n, sums:sums};
}

function reconciliarME(){
  var out={}, problemas=0;
  Object.keys(_ME_SPECS).forEach(function(tabla){
    var cfg=_ME_SPECS[tabla], cols=_ME_SUMCOLS[tabla]||[], info={};
    try{
      var existe=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg.sheet);
      var rows=existe?_meBuildRows(tabla):[];
      info.sheet_n=rows.length;
      var ss={}; cols.forEach(function(c){ ss[c]=0; });
      rows.forEach(function(r){ cols.forEach(function(c){ var v=r[c]; if(typeof v==='number'&&!isNaN(v)) ss[c]+=v; }); });
      var sb=_sbSumCols('me.'+tabla, cols, cfg.onConflict||'id');
      if(sb.error){ info.error=sb.error; out[tabla]=info; problemas++; return; }
      info.sb_n=sb.n;
      info.n_ok=(info.sheet_n===info.sb_n);
      var sumOk=true; info.sums={};
      cols.forEach(function(c){ var a=ss[c]||0, b=sb.sums[c]||0, ok=Math.abs(a-b)<0.01; if(!ok)sumOk=false;
        info.sums[c]={sheet:Math.round(a*100)/100, sb:Math.round(b*100)/100, ok:ok}; });
      info.ok=info.n_ok && sumOk;
      if(!info.ok) problemas++;
    }catch(e){ info.error=String(e&&e.message||e); problemas++; }
    out[tabla]=info;
  });
  out._resumen={problemas:problemas, veredicto: problemas===0?'✓ SIN DRIFT':'⚠ revisar '+problemas+' tabla(s)'};
  Logger.log(JSON.stringify(out,null,2));
  return out;
}

/** Corre reconciliarME y registra una fila en la hoja RECON_LOG (lo dispara el trigger diario). */
function reconciliarDiarioME(){
  var res=reconciliarME(), r=res._resumen||{};
  var probs={}; Object.keys(res).forEach(function(k){ if(k!=='_resumen' && res[k] && res[k].ok===false) probs[k]=res[k]; });
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sh=ss.getSheetByName('RECON_LOG') || ss.insertSheet('RECON_LOG');
  if(sh.getLastRow()===0) sh.appendRow(['fecha','app','problemas','veredicto','tablas_con_drift']);
  sh.appendRow([Utilities.formatDate(new Date(),'America/Lima','yyyy-MM-dd HH:mm'),'ME', r.problemas||0, r.veredicto||'', JSON.stringify(probs).slice(0,45000)]);
  return res;
}
/** La recon ahora va PEGADA a syncMECompleto (sin trigger propio, por el límite de 20 triggers).
 *  Esta función solo LIMPIA un trigger de recon separado si lo instalaste antes. */
function desinstalarTriggerReconME(){
  var n=0; ScriptApp.getProjectTriggers().forEach(function(t){ if(t.getHandlerFunction()==='reconciliarDiarioME'){ ScriptApp.deleteTrigger(t); n++; } });
  Logger.log('Triggers recon separados eliminados: '+n+' (la recon corre dentro de syncMECompleto)'); return {ok:true, eliminados:n};
}

/** Busca valores absurdos (código de barras tecleado en campo de cantidad) en hojas de ME. Solo lectura.
 *  Reporta hoja · FILA exacta (1-based para ubicarla en el Sheet) · columna · valor · contexto. */
function buscarBasuraME(umbral){
  umbral = umbral || 1000000;   // ninguna cantidad/diferencia real de este negocio supera 1 millón
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var checks=[
    {hoja:'GUIAS_DETALLE', cols:['Cantidad']},
    {hoja:'AUDITORIAS',    cols:['Diferencia','Cant_Sistema','Cant_Real']},
    {hoja:'VENTAS_DETALLE',cols:['Cantidad']},
    {hoja:'STOCK_ZONAS',   cols:['Cantidad']}
  ];
  var out={umbral:umbral, total:0, hallazgos:[]};
  checks.forEach(function(chk){
    var sh=ss.getSheetByName(chk.hoja); if(!sh) return;
    var data=sh.getDataRange().getValues(), hdr=data[0];
    var idxs=chk.cols.map(function(c){ return {col:c, i:hdr.indexOf(c)}; }).filter(function(x){ return x.i>=0; });
    for(var r=1;r<data.length;r++){
      idxs.forEach(function(x){
        var v=data[r][x.i], n=(typeof v==='number')?v:parseFloat(v);
        if(!isNaN(n) && Math.abs(n)>umbral){
          out.hallazgos.push({ hoja:chk.hoja, fila:(r+1), col:x.col, valor:n, contexto:data[r].slice(0,6) });
        }
      });
    }
  });
  out.total=out.hallazgos.length;
  Logger.log(JSON.stringify(out,null,2));
  return out;
}

var _BASURA_GUIA='G-1778022126489', _BASURA_COD='7755019000123', _BASURA_AUD='A-1780784928153', _BASURA_UMBRAL=1000000;

/** Muestra el contexto (cabecera + TODAS las líneas de la guía + la auditoría) para decidir la cantidad real. Solo lectura. */
function verContextoBasuraME(){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), out={};
  function filaObj(h,r){ var o={}; h.forEach(function(k,j){ o[k]=r[j]; }); return o; }
  var gc=ss.getSheetByName('GUIAS_CABECERA');
  if(gc){ var d=gc.getDataRange().getValues(), h=d[0], i=h.indexOf('ID_Guia');
    out.guia_cabecera=null; for(var a=1;a<d.length;a++){ if(String(d[a][i])===_BASURA_GUIA){ out.guia_cabecera=filaObj(h,d[a]); break; } } }
  var gd=ss.getSheetByName('GUIAS_DETALLE');
  if(gd){ var d2=gd.getDataRange().getValues(), h2=d2[0], i2=h2.indexOf('ID_Guia');
    out.guia_lineas=[]; for(var b=1;b<d2.length;b++){ if(String(d2[b][i2])===_BASURA_GUIA){ var o=filaObj(h2,d2[b]); o._fila=b+1; out.guia_lineas.push(o); } } }
  var au=ss.getSheetByName('AUDITORIAS');
  if(au){ var d3=au.getDataRange().getValues(), h3=d3[0], i3=h3.indexOf('ID_Auditoria');
    out.auditoria=null; for(var c=1;c<d3.length;c++){ if(String(d3[c][i3])===_BASURA_AUD){ var o3=filaObj(h3,d3[c]); o3._fila=c+1; out.auditoria=o3; break; } } }
  Logger.log(JSON.stringify(out,null,2));
  return out;
}

/** Corrige la basura de ME de forma SEGURA. Respalda los valores viejos en CORRECCIONES_LOG.
 *  USO:  corregirBasuraME(33)   ← reemplaza 33 por la CANTIDAD REAL despachada en esa línea de la guía.
 *  La auditoría se neutraliza honestamente: Cant_Sistema = Cant_Real, Diferencia = 0. */
function corregirBasuraME(cantidadGuiaReal){
  if(cantidadGuiaReal==null || isNaN(parseFloat(cantidadGuiaReal)))
    return {ok:false, error:'Falta la cantidad real. Ej: corregirBasuraME(33)  ← pon el valor correcto (corre antes verContextoBasuraME para decidirlo)'};
  cantidadGuiaReal=parseFloat(cantidadGuiaReal);
  var ss=SpreadsheetApp.getActiveSpreadsheet(), res={cambios:[]};

  // 1) GUIAS_DETALLE: la línea guía+producto con cantidad basura
  var gd=ss.getSheetByName('GUIAS_DETALLE'), d=gd.getDataRange().getValues(), h=d[0];
  var iG=h.indexOf('ID_Guia'), iC=h.indexOf('Cod_Barras'), iQ=h.indexOf('Cantidad'), fixGD=false;
  for(var i=1;i<d.length;i++){
    if(String(d[i][iG])===_BASURA_GUIA && String(d[i][iC])===_BASURA_COD && Math.abs(parseFloat(d[i][iQ]))>_BASURA_UMBRAL){
      res.cambios.push({hoja:'GUIAS_DETALLE', fila:i+1, col:'Cantidad', viejo:d[i][iQ], nuevo:cantidadGuiaReal});
      gd.getRange(i+1, iQ+1).setValue(cantidadGuiaReal); fixGD=true; break;
    }
  }
  if(!fixGD) res.cambios.push({hoja:'GUIAS_DETALLE', nota:'no se halló la basura (¿ya corregida?)'});

  // 2) AUDITORIAS: neutralizar (Cant_Sistema = Cant_Real, Diferencia = 0)
  var au=ss.getSheetByName('AUDITORIAS'), d2=au.getDataRange().getValues(), h2=d2[0];
  var iA=h2.indexOf('ID_Auditoria'), iCS=h2.indexOf('Cant_Sistema'), iCR=h2.indexOf('Cant_Real'), iD=h2.indexOf('Diferencia'), fixAU=false;
  for(var k=1;k<d2.length;k++){
    if(String(d2[k][iA])===_BASURA_AUD && Math.abs(parseFloat(d2[k][iCS]))>_BASURA_UMBRAL){
      var cantReal=parseFloat(d2[k][iCR]); if(isNaN(cantReal)) cantReal=0;
      res.cambios.push({hoja:'AUDITORIAS', fila:k+1, col:'Cant_Sistema', viejo:d2[k][iCS], nuevo:cantReal});
      res.cambios.push({hoja:'AUDITORIAS', fila:k+1, col:'Diferencia', viejo:d2[k][iD], nuevo:0});
      au.getRange(k+1, iCS+1).setValue(cantReal);
      au.getRange(k+1, iD+1).setValue(0); fixAU=true; break;
    }
  }
  if(!fixAU) res.cambios.push({hoja:'AUDITORIAS', nota:'no se halló la basura (¿ya corregida?)'});

  // Respaldo persistente (para revertir si hiciera falta)
  var log=ss.getSheetByName('CORRECCIONES_LOG') || ss.insertSheet('CORRECCIONES_LOG');
  if(log.getLastRow()===0) log.appendRow(['fecha','accion','cambios_json']);
  log.appendRow([Utilities.formatDate(new Date(),'America/Lima','yyyy-MM-dd HH:mm'),'corregirBasuraME', JSON.stringify(res.cambios)]);

  res.ok=true;
  res.nota='Respaldado en CORRECCIONES_LOG. El sync nocturno propagará la corrección a Supabase; corre reconciliarME tras el próximo sync para ver las sumas ya sanas.';
  Logger.log(JSON.stringify(res,null,2));
  return res;
}

/** ATAJO para el botón ▶ Ejecutar (GAS no deja pasar argumentos al Run).
 *  Cantidad real DEDUCIDA = 25 (el "25" + el código de barras 7751271034081 de la línea siguiente). */
function corregirBasuraME_RUN(){ return corregirBasuraME(25); }

// ============================================================
// FASE 1.D (canary) — comparador de PARIDAD de estadoCajas: Sheets vs Supabase.
// 100% shadow: llama a la función de producción y a la RPC me.estado_cajas(), las
// compara (por idCaja, tolerante a orden de claves y a ruido float) y mide el speedup.
// NO toca el endpoint. Requiere 06_fase1d_estado_cajas.sql corrido.
// ============================================================
function _numEq(a,b){ var na=parseFloat(a), nb=parseFloat(b); if(!isNaN(na)&&!isNaN(nb)) return Math.abs(na-nb)<0.01; return String(a)===String(b); }
function _cajasById(arr){ var m={}; (arr||[]).forEach(function(c){ m[String(c.idCaja)]=c; }); return m; }
function _diffCaja(grupo,id,a,b,diffs){
  ['vendedor','estacion','zona','estado','fechaApertura','fechaCierre','montoInicial','montoFinal',
   'totalVentas','tickets','efectivo','otros','anulados','sinCobrar','entradas','salidas','efectivoEsperado','diferencia'
  ].forEach(function(f){
    var va=a[f], vb=b[f];
    if((va===null)!==(vb===null)){ diffs.push(grupo+' '+id+'.'+f+': sheets='+va+' sb='+vb); return; }
    if(va===null) return;
    if(typeof va==='number' || typeof vb==='number'){ if(!_numEq(va,vb)) diffs.push(grupo+' '+id+'.'+f+': sheets='+va+' sb='+vb); }
    else if(String(va)!==String(vb)){ diffs.push(grupo+' '+id+'.'+f+': sheets="'+va+'" sb="'+vb+'"'); }
  });
  ['byMetodo','byDoc'].forEach(function(f){
    var oa=a[f]||{}, ob=b[f]||{}, keys={};
    Object.keys(oa).forEach(function(k){keys[k]=1;}); Object.keys(ob).forEach(function(k){keys[k]=1;});
    Object.keys(keys).forEach(function(k){ if(!_numEq(oa[k]||0, ob[k]||0)) diffs.push(grupo+' '+id+'.'+f+'['+k+']: sheets='+oa[k]+' sb='+ob[k]); });
  });
}

// separa cajas con ID válido (map) de las chatarra sin ID (el backfill las excluye por PK vacía)
function _splitCajas(arr){ var m={}, junk=[]; (arr||[]).forEach(function(c){ var id=String(c.idCaja||'').trim();
  if(!id) junk.push({estado:c.estado, vendedor:c.vendedor, fechaApertura:c.fechaApertura, totalVentas:c.totalVentas});
  else m[id]=c; }); return {m:m, junk:junk}; }

function compararEstadoCajasME(){
  // 1) Sheets (producción, sin tocarla — solo parseamos su salida)
  var t0=Date.now(); var sheetsObj=JSON.parse(estadoCajas().getContent()); var tSheets=Date.now()-t0;
  // 2) Supabase (agregación server-side)
  var t1=Date.now(); var r=_sbRpc('me','estado_cajas',{}); var tSb=Date.now()-t1;
  if(!r.ok){ var e={ok:false, error:'RPC me.estado_cajas falló: HTTP '+r.code+' — '+(r.error||''), nota:'¿corriste 06_fase1d_estado_cajas.sql?'}; Logger.log(JSON.stringify(e,null,2)); return e; }
  var sbObj=r.data;

  var diffs=[], junk=[];
  // cajas válidas (sin chatarra), por idCaja, orden-independiente
  ['abiertas','cerradas'].forEach(function(grupo){
    var S=_splitCajas(sheetsObj[grupo]), B=_splitCajas(sbObj[grupo]);
    S.junk.forEach(function(j){ j._grupo=grupo; junk.push(j); });
    var ids={}; Object.keys(S.m).forEach(function(k){ids[k]=1;}); Object.keys(B.m).forEach(function(k){ids[k]=1;});
    Object.keys(ids).forEach(function(id){
      if(!S.m[id]){ diffs.push(grupo+' '+id+': falta en SHEETS'); return; }
      if(!B.m[id]){ diffs.push(grupo+' '+id+': falta en SUPABASE'); return; }
      _diffCaja(grupo,id,S.m[id],B.m[id],diffs);
    });
  });
  // KPIs: counts ajustados por chatarra (que solo infla el conteo, aporta 0 a montos); montos estrictos
  var jAb=_splitCajas(sheetsObj.abiertas).junk.length, jCe=_splitCajas(sheetsObj.cerradas).junk.length;
  if(!_numEq(sheetsObj.kpis.cajasAbiertas - jAb, sbObj.kpis.cajasAbiertas)) diffs.push('kpis.cajasAbiertas(ajust): sheets='+(sheetsObj.kpis.cajasAbiertas-jAb)+' sb='+sbObj.kpis.cajasAbiertas);
  if(!_numEq(sheetsObj.kpis.cajasCerradas - jCe, sbObj.kpis.cajasCerradas)) diffs.push('kpis.cajasCerradas(ajust): sheets='+(sheetsObj.kpis.cajasCerradas-jCe)+' sb='+sbObj.kpis.cajasCerradas);
  ['totalDia','ticketsDia','anuladosDia','sinCobrarDia'].forEach(function(k){
    if(!_numEq(sheetsObj.kpis[k], sbObj.kpis[k])) diffs.push('kpis.'+k+': sheets='+sheetsObj.kpis[k]+' sb='+sbObj.kpis[k]);
  });

  var out={
    ok: diffs.length===0,
    veredicto: diffs.length===0 ? '✓ PARIDAD EXACTA (sobre datos válidos) — listo para flip' : '⚠ '+diffs.length+' diferencias REALES (revisar)',
    velocidad:{ sheets_ms:tSheets, supabase_ms:tSb, speedup: (tSheets&&tSb)?(Math.round(tSheets/tSb*10)/10+'x'):'n/a' },
    cajas_validas:{ sheets:(sheetsObj.abiertas.length+sheetsObj.cerradas.length-jAb-jCe), supabase:(sbObj.abiertas.length+sbObj.cerradas.length) },
    cajas_sin_id_sheets:{ total:junk.length, nota:'chatarra sin ID_Caja en hoja CAJAS — la migración las excluye (PK no vacía). Al flip desaparecen del UI; conviene limpiarlas.', filas:junk.slice(0,10) },
    diferencias: diffs.slice(0,40)
  };
  Logger.log(JSON.stringify(out,null,2));
  return out;
}

/** Borra de la hoja CAJAS las filas TOTALMENTE vacías (ID_Caja vacío + toda la fila vacía).
 *  Seguro: solo toca filas 100% en blanco. Respalda en CORRECCIONES_LOG. Borra de abajo hacia arriba. */
function limpiarCajasSinIdME(){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName('CAJAS');
  if(!sh) return {ok:false, error:'CAJAS no encontrada'};
  var data=sh.getDataRange().getValues(), h=data[0], iId=h.indexOf('ID_Caja');
  var aBorrar=[];
  for(var i=1;i<data.length;i++){
    var idVacio   = String(data[i][iId]||'').trim()==='';
    var todoVacio = data[i].every(function(v){ return v===''||v===null||v===undefined; });
    if(idVacio && todoVacio) aBorrar.push(i+1);   // 1-based
  }
  if(!aBorrar.length){ var z={ok:true, borradas:0, nota:'no hay filas 100% vacías'}; Logger.log(JSON.stringify(z)); return z; }
  var log=ss.getSheetByName('CORRECCIONES_LOG') || ss.insertSheet('CORRECCIONES_LOG');
  if(log.getLastRow()===0) log.appendRow(['fecha','accion','cambios_json']);
  log.appendRow([Utilities.formatDate(new Date(),'America/Lima','yyyy-MM-dd HH:mm'),'limpiarCajasSinIdME', JSON.stringify({filas_borradas:aBorrar})]);
  aBorrar.sort(function(a,b){ return b-a; }).forEach(function(f){ sh.deleteRow(f); });
  var res={ok:true, borradas:aBorrar.length, filas:aBorrar, nota:'corre compararEstadoCajasME para confirmar 53=53 sin ajuste'};
  Logger.log(JSON.stringify(res,null,2)); return res;
}

// ---------- Canary #2: getCobrosEnVueloAdmin (Sheets vs me.cobros_en_vuelo()) ----------
function _dateEqSec(a,b){   // compara fechas a granularidad de SEGUNDO (la migración truncó ms)
  var sa=String(a||''), sb=String(b||'');
  if(sa===''||sb==='') return sa===sb;
  var ta=new Date(sa).getTime(), tb=new Date(sb).getTime();
  if(isNaN(ta)||isNaN(tb)) return sa===sb;
  return Math.floor(ta/1000)===Math.floor(tb/1000);
}
function _diffCobro(label,a,b,diffs){
  var dateF={fechaAsig:1,fechaRes:1,fechaVencimiento:1}, numF={monto:1,horasTTL:1,reasignaciones:1};
  ['idVenta','cajaDestino','vendedorDest','metodoSug','estado','adminAsig','fechaAsig','fechaRes',
   'razon','monto','cliente','correlativo','fechaVencimiento','horasTTL','mensajeAdmin','reasignaciones'
  ].forEach(function(f){
    var va=a[f], vb=b[f];
    if(dateF[f]){ if(!_dateEqSec(va,vb)) diffs.push(label+'.'+f+': sheets="'+va+'" sb="'+vb+'"'); }
    else if(numF[f]){ if(!_numEq(va,vb)) diffs.push(label+'.'+f+': sheets='+va+' sb='+vb); }
    else if(String(va)!==String(vb)) diffs.push(label+'.'+f+': sheets="'+va+'" sb="'+vb+'"');
  });
}
function compararCobrosEnVueloME(){
  var t0=Date.now(); var sh=JSON.parse(getCobrosEnVueloAdmin().getContent()); var tS=Date.now()-t0;
  var t1=Date.now(); var r=_sbRpc('me','cobros_en_vuelo',{}); var tB=Date.now()-t1;
  if(!r.ok){ var e={ok:false, error:'RPC me.cobros_en_vuelo falló: HTTP '+r.code+' — '+(r.error||''), nota:'¿corriste 07_fase1d_cobros.sql?'}; Logger.log(JSON.stringify(e,null,2)); return e; }
  var sb=r.data, diffs=[];
  function byId(arr){ var m={}; (arr||[]).forEach(function(c){ m[String(c.idCobro)]=c; }); return m; }
  ['enVuelo','recientes'].forEach(function(grupo){
    var ms=byId(sh[grupo]), mb=byId(sb[grupo]), ids={};
    Object.keys(ms).forEach(function(k){ids[k]=1;}); Object.keys(mb).forEach(function(k){ids[k]=1;});
    Object.keys(ids).forEach(function(id){
      if(!ms[id]){ diffs.push(grupo+' '+id+': falta en SHEETS'); return; }
      if(!mb[id]){ diffs.push(grupo+' '+id+': falta en SUPABASE'); return; }
      _diffCobro(grupo+' '+id, ms[id], mb[id], diffs);
    });
  });
  var out={ ok:diffs.length===0,
    veredicto: diffs.length===0 ? '✓ PARIDAD EXACTA — listo para flip' : '⚠ '+diffs.length+' diferencias',
    velocidad:{ sheets_ms:tS, supabase_ms:tB, speedup:(tS&&tB)?(Math.round(tS/tB*10)/10+'x'):'n/a' },
    conteos:{ enVuelo:{sheets:(sh.enVuelo||[]).length, sb:(sb.enVuelo||[]).length},
              recientes:{sheets:(sh.recientes||[]).length, sb:(sb.recientes||[]).length} },
    diferencias: diffs.slice(0,40) };
  Logger.log(JSON.stringify(out,null,2)); return out;
}

// ============================================================
// FASE 1.D — FLIP con feature flag FUENTE_DATOS (Script Property).
//   default 'sheets'  → comportamiento ACTUAL, cero cambio en producción.
//   'supabase'        → lee de Supabase (RPC); ante CUALQUIER fallo cae a Sheets.
// El router (Code.gs) llama a estadoCajasFlip()/getCobrosEnVueloAdminFlip().
// Encender:  activarSupabaseME()   ·   Apagar (rollback):  desactivarSupabaseME()
// ============================================================
function _fuenteDatos(key){
  try{
    var p=PropertiesService.getScriptProperties();
    if(String(p.getProperty('FUENTE_DATOS')||'sheets').toLowerCase()!=='supabase') return 'sheets';
    // kill-switch GRANULAR: FUENTE_DATOS_OFF = CSV de endpoints a forzar a Sheets aunque el master esté en supabase
    var off=String(p.getProperty('FUENTE_DATOS_OFF')||'').toLowerCase();
    if(key && off){ var arr=off.split(',').map(function(s){return s.trim();}); if(arr.indexOf(String(key).toLowerCase())>=0) return 'sheets'; }
    return 'supabase';
  }catch(e){ return 'sheets'; }
}

var _FLIP_CACHE_SEG = 15;   // coalesce polls → protege la cuota UrlFetchApp de ME al flipear

function estadoCajasFlip(){
  if(_fuenteDatos('estado_cajas')==='supabase'){
    try{
      var cache=CacheService.getScriptCache();
      var hit=cache.get('SB_ESTADO_CAJAS');
      if(hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
      var r=_sbRpc('me','estado_cajas',{});
      if(r.ok && r.data && r.data.kpis && Array.isArray(r.data.abiertas) && Array.isArray(r.data.cerradas)){
        var json=JSON.stringify({
          status:'success',
          generadoEn: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
          kpis: r.data.kpis, abiertas: r.data.abiertas, cerradas: r.data.cerradas
        });
        try{ cache.put('SB_ESTADO_CAJAS', json, _FLIP_CACHE_SEG); }catch(eC){}   // >100KB → no cachea, sin romper
        return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
      }
    }catch(e){ /* cae a Sheets */ }
  }
  return estadoCajas();   // Sheets: default y fallback
}

function getCobrosEnVueloAdminFlip(){
  if(_fuenteDatos('cobros_en_vuelo')==='supabase'){
    try{
      var cache=CacheService.getScriptCache();
      var hit=cache.get('SB_COBROS_EN_VUELO');
      if(hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
      var r=_sbRpc('me','cobros_en_vuelo',{});
      if(r.ok && r.data && Array.isArray(r.data.enVuelo) && Array.isArray(r.data.recientes)){
        var json=JSON.stringify({ status:'success', enVuelo:r.data.enVuelo, recientes:r.data.recientes });
        try{ cache.put('SB_COBROS_EN_VUELO', json, _FLIP_CACHE_SEG); }catch(eC){}
        return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
      }
    }catch(e){ /* cae a Sheets */ }
  }
  return getCobrosEnVueloAdmin();   // Sheets: default y fallback
}

function getCreditosPendientesFlip(diasAtras){
  if(_fuenteDatos('creditos_pendientes')==='supabase'){
    try{
      var dias=parseInt(diasAtras,10)||30;
      var cache=CacheService.getScriptCache(), ckey='SB_CREDITOS_'+dias;   // cache por-parámetro
      var hit=cache.get(ckey);
      if(hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
      var r=_sbRpc('me','creditos_pendientes',{dias_atras:dias});
      if(r.ok && r.data && Array.isArray(r.data.grupos)){
        var json=JSON.stringify({ status:'success', grupos:r.data.grupos, totalAcumulado:r.data.totalAcumulado, totalTickets:r.data.totalTickets });
        try{ cache.put(ckey, json, _FLIP_CACHE_SEG); }catch(eC){}
        return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
      }
    }catch(e){ /* cae a Sheets */ }
  }
  return getCreditosPendientes(diasAtras);   // Sheets: default y fallback
}

function ventasHoyZonaFlip(prefijos, desde){
  if(_fuenteDatos('ventas_hoy_zona')==='supabase'){
    try{
      var cache=CacheService.getScriptCache(), ckey=('SB_VENTASZONA_'+(prefijos||'')+'_'+(desde||'')).slice(0,240);
      var hit=cache.get(ckey);
      if(hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
      var r=_sbRpc('me','ventas_hoy_zona',{prefijos_str:(prefijos||null), desde_str:(desde||null)});
      if(r.ok && r.data && Array.isArray(r.data.ventas)){
        var json=JSON.stringify({ status:'success', ventas:r.data.ventas });
        try{ cache.put(ckey, json, _FLIP_CACHE_SEG); }catch(eC){}
        return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
      }
    }catch(e){ /* cae a Sheets */ }
  }
  return ventasHoyZona(prefijos, desde);   // Sheets: default y fallback
}

// ---- Controles (correr desde el editor) ----
function activarSupabaseME(){ PropertiesService.getScriptProperties().setProperty('FUENTE_DATOS','supabase'); Logger.log('✅ FUENTE_DATOS = supabase — los 4 reads (estado_cajas/cobros/creditos/ventas_zona) leen de Supabase (fallback a Sheets si falla)'); return {ok:true, fuente:'supabase'}; }
function desactivarSupabaseME(){
  PropertiesService.getScriptProperties().setProperty('FUENTE_DATOS','sheets');
  try{ CacheService.getScriptCache().removeAll(['SB_ESTADO_CAJAS','SB_COBROS_EN_VUELO']); }catch(e){}  // higiene (las de creditos/ventas tienen sufijo variable y expiran en 15s)
  Logger.log('↩️ FUENTE_DATOS = sheets — rollback instantáneo y completo de los 4 a Sheets'); return {ok:true, fuente:'sheets'};
}
function estadoFuenteDatosME(){ var p=PropertiesService.getScriptProperties(); var out={master:String(p.getProperty('FUENTE_DATOS')||'sheets'), off:String(p.getProperty('FUENTE_DATOS_OFF')||'')}; Logger.log(JSON.stringify(out)); return out; }
// kill-switch GRANULAR: apaga/prende UN endpoint sin tocar los otros (endpoints: estado_cajas, cobros_en_vuelo, creditos_pendientes, ventas_hoy_zona)
function desactivarUnoME(endpoint){ var p=PropertiesService.getScriptProperties(); var off=(p.getProperty('FUENTE_DATOS_OFF')||'').split(',').map(function(s){return s.trim();}).filter(Boolean); if(off.indexOf(endpoint)<0) off.push(endpoint); p.setProperty('FUENTE_DATOS_OFF', off.join(',')); Logger.log('🔻 '+endpoint+' forzado a Sheets. OFF=['+off.join(',')+']'); return {ok:true, off:off}; }
function reactivarUnoME(endpoint){ var p=PropertiesService.getScriptProperties(); var off=(p.getProperty('FUENTE_DATOS_OFF')||'').split(',').map(function(s){return s.trim();}).filter(Boolean).filter(function(e){return e!==endpoint;}); p.setProperty('FUENTE_DATOS_OFF', off.join(',')); Logger.log('🔼 '+endpoint+' reactivado a Supabase. OFF=['+off.join(',')+']'); return {ok:true, off:off}; }

// [reads-reflip] UN clic: re-flipea SOLO ventas_hoy_zona a Supabase (la lectura del bug del cajero, que
// ahora es real-time porque ventas tiene dual-write). Deja estado_cajas/cobros/creditos en Sheets (sus
// fuentes aún no son 100% real-time). Rollback total = desactivarSupabaseME(). Granular = reactivarUnoME('estado_cajas') etc.
function flipSoloVentasZona(){
  desactivarUnoME('estado_cajas');
  desactivarUnoME('cobros_en_vuelo');
  desactivarUnoME('creditos_pendientes');
  var r = activarSupabaseME();
  var off = PropertiesService.getScriptProperties().getProperty('FUENTE_DATOS_OFF') || '';
  Logger.log('✅ SOLO ventas_hoy_zona lee de Supabase (real-time). En Sheets: ['+off+']. Rollback: desactivarSupabaseME()');
  return { ok:true, supabase:['ventas_hoy_zona'], sheets: off.split(',') };
}

// [reads-reflip] UN clic: SUMA estado_cajas a Supabase (sus fuentes ventas+movimientos+cajas ya son
// real-time por dual-write). Mantiene cobros/creditos en Sheets. Es display (el cierre recomputa de Sheets).
// Requiere haber corrido flipSoloVentasZona antes (o que FUENTE_DATOS ya sea supabase). Rollback granular:
// desactivarUnoME('estado_cajas'); total: desactivarSupabaseME().
function flipSumarEstadoCajas(){
  if(_fuenteDatos('ventas_hoy_zona')!=='supabase'){ activarSupabaseME(); desactivarUnoME('cobros_en_vuelo'); desactivarUnoME('creditos_pendientes'); }
  reactivarUnoME('estado_cajas');
  var off = PropertiesService.getScriptProperties().getProperty('FUENTE_DATOS_OFF') || '';
  Logger.log('✅ Supabase: ventas_hoy_zona + estado_cajas (real-time). En Sheets: ['+off+']. Rollback: desactivarUnoME(\'estado_cajas\') o desactivarSupabaseME()');
  return { ok:true, supabase:['ventas_hoy_zona','estado_cajas'], sheets: off.split(',').filter(Boolean) };
}

// ---------- Canary #3: getCreditosPendientes (Sheets vs me.creditos_pendientes()) ----------
function compararCreditosPendientesME(){ return _compararCreditos(30); }
function _compararCreditos(dias){
  var t0=Date.now(); var sh=JSON.parse(getCreditosPendientes(dias).getContent()); var tS=Date.now()-t0;
  var t1=Date.now(); var r=_sbRpc('me','creditos_pendientes',{dias_atras:dias}); var tB=Date.now()-t1;
  if(!r.ok){ var e={ok:false, error:'RPC me.creditos_pendientes falló: HTTP '+r.code+' — '+(r.error||''), nota:'¿corriste 08_fase1d_creditos.sql?'}; Logger.log(JSON.stringify(e,null,2)); return e; }
  var sb=r.data, diffs=[];
  if(!_numEq(sh.totalAcumulado, sb.totalAcumulado)) diffs.push('totalAcumulado: sheets='+sh.totalAcumulado+' sb='+sb.totalAcumulado);
  if(!_numEq(sh.totalTickets, sb.totalTickets)) diffs.push('totalTickets: sheets='+sh.totalTickets+' sb='+sb.totalTickets);
  function byKey(arr,k){ var m={}; (arr||[]).forEach(function(x){ m[String(x[k])]=x; }); return m; }
  var gs=byKey(sh.grupos,'fecha'), gb=byKey(sb.grupos,'fecha'), fechas={};
  Object.keys(gs).forEach(function(k){fechas[k]=1;}); Object.keys(gb).forEach(function(k){fechas[k]=1;});
  Object.keys(fechas).forEach(function(f){
    if(!gs[f]){ diffs.push('grupo '+f+': falta en SHEETS'); return; }
    if(!gb[f]){ diffs.push('grupo '+f+': falta en SUPABASE'); return; }
    if(!_numEq(gs[f].total, gb[f].total)) diffs.push('grupo '+f+'.total: sheets='+gs[f].total+' sb='+gb[f].total);
    if(!_numEq(gs[f].cuenta, gb[f].cuenta)) diffs.push('grupo '+f+'.cuenta: sheets='+gs[f].cuenta+' sb='+gb[f].cuenta);
    var ts=byKey(gs[f].tickets,'idVenta'), tb=byKey(gb[f].tickets,'idVenta'), ids={};
    Object.keys(ts).forEach(function(k){ids[k]=1;}); Object.keys(tb).forEach(function(k){ids[k]=1;});
    Object.keys(ids).forEach(function(id){
      if(!ts[id]){ diffs.push(f+' ticket '+id+': falta en SHEETS'); return; }
      if(!tb[id]){ diffs.push(f+' ticket '+id+': falta en SUPABASE'); return; }
      _diffTicketCredito(f+' '+id, ts[id], tb[id], diffs);
    });
  });
  var out={ ok:diffs.length===0,
    veredicto: diffs.length===0 ? '✓ PARIDAD EXACTA — listo para flip' : '⚠ '+diffs.length+' diferencias',
    velocidad:{ sheets_ms:tS, supabase_ms:tB, speedup:(tS&&tB)?(Math.round(tS/tB*10)/10+'x'):'n/a' },
    conteos:{ grupos:{sheets:(sh.grupos||[]).length, sb:(sb.grupos||[]).length}, totalTickets:{sheets:sh.totalTickets, sb:sb.totalTickets} },
    diferencias: diffs.slice(0,50) };
  Logger.log(JSON.stringify(out,null,2)); return out;
}
function _diffTicketCredito(label,a,b,diffs){
  ['correlativo','cliente','clienteDoc','vendedor','formaPago','obs','idCaja','fechaISO'].forEach(function(f){
    if(String(a[f])!==String(b[f])) diffs.push(label+'.'+f+': sheets="'+a[f]+'" sb="'+b[f]+'"');
  });
  if(!_numEq(a.total,b.total))           diffs.push(label+'.total: sheets='+a.total+' sb='+b.total);
  if(!_numEq(a.itemsCount,b.itemsCount)) diffs.push(label+'.itemsCount: sheets='+a.itemsCount+' sb='+b.itemsCount);
  var aa=a.asignado, ba=b.asignado;
  if((aa==null)!==(ba==null)) diffs.push(label+'.asignado: sheets='+JSON.stringify(aa)+' sb='+JSON.stringify(ba));
  else if(aa!=null){
    ['idCobro','cajaDestino','vendedorDest'].forEach(function(f){ if(String(aa[f])!==String(ba[f])) diffs.push(label+'.asignado.'+f+': sheets="'+aa[f]+'" sb="'+ba[f]+'"'); });
    if(!_dateEqSec(aa.fechaAsig,ba.fechaAsig)) diffs.push(label+'.asignado.fechaAsig: sheets="'+aa.fechaAsig+'" sb="'+ba.fechaAsig+'"');
  }
  var ia=a.items||[], ib=b.items||[];
  if(ia.length!==ib.length) diffs.push(label+'.items.length: sheets='+ia.length+' sb='+ib.length);
  else for(var i=0;i<ia.length;i++){
    if(String(ia[i].nombre)!==String(ib[i].nombre)) diffs.push(label+'.items['+i+'].nombre: sheets="'+ia[i].nombre+'" sb="'+ib[i].nombre+'"');
    if(!_numEq(ia[i].cantidad,ib[i].cantidad))       diffs.push(label+'.items['+i+'].cantidad: sheets='+ia[i].cantidad+' sb='+ib[i].cantidad);
    if(!_numEq(ia[i].subtotal,ib[i].subtotal))       diffs.push(label+'.items['+i+'].subtotal: sheets='+ia[i].subtotal+' sb='+ib[i].subtotal);
  }
}

// ---------- Canary #4: ventasHoyZona (Sheets vs me.ventas_hoy_zona()) ----------
function compararVentasHoyZonaME(){
  var desde30 = new Date(Date.now() - 30*86400000).toISOString();   // corte FIJO → mismo string a ambos lados (sin 2-relojes)
  var escenarios = [
    {n:'hoy (sin desde, sin zona)',     pref:null, desde:null},
    {n:'30d desde-fijo (sin zona)',     pref:null, desde:desde30}
  ];
  var salida={ ok:true, escenarios:[] };
  escenarios.forEach(function(esc){
    var t0=Date.now(); var sh=JSON.parse(ventasHoyZona(esc.pref, esc.desde).getContent()); var tS=Date.now()-t0;
    var t1=Date.now(); var r=_sbRpc('me','ventas_hoy_zona',{prefijos_str:esc.pref, desde_str:esc.desde}); var tB=Date.now()-t1;
    var res={escenario:esc.n};
    if(!r.ok){ res.error='RPC falló: HTTP '+r.code+' — '+(r.error||''); res.nota='¿corriste 09_fase1d_ventas_zona.sql?'; salida.ok=false; salida.escenarios.push(res); return; }
    var sb=r.data, diffs=[];
    function byId(arr){ var m={}; (arr||[]).forEach(function(x){ m[String(x.id_venta)]=x; }); return m; }
    var ms=byId(sh.ventas), mb=byId(sb.ventas), ids={};
    Object.keys(ms).forEach(function(k){ids[k]=1;}); Object.keys(mb).forEach(function(k){ids[k]=1;});
    Object.keys(ids).forEach(function(id){
      if(!ms[id]){ diffs.push(id+': falta en SHEETS'); return; }
      if(!mb[id]){ diffs.push(id+': falta en SUPABASE'); return; }
      _diffVentaZona(id, ms[id], mb[id], diffs);
    });
    res.ok=diffs.length===0;
    res.ventas={sheets:(sh.ventas||[]).length, sb:(sb.ventas||[]).length};
    res.velocidad={sheets_ms:tS, supabase_ms:tB, speedup:(tS&&tB)?(Math.round(tS/tB*10)/10+'x'):'n/a'};
    res.diferencias=diffs.slice(0,30);
    if(!res.ok) salida.ok=false;
    salida.escenarios.push(res);
  });
  salida.veredicto = salida.ok ? '✓ PARIDAD EXACTA en todos los escenarios — listo para flip' : '⚠ revisar diferencias';
  Logger.log(JSON.stringify(salida,null,2)); return salida;
}
function _diffVentaZona(label,a,b,diffs){
  ['vendedor','estacion','cliente_doc','cliente_nombre','tipo_doc','forma_pago','correlativo','id_caja','id_dispositivo','status','ref_local','obs'].forEach(function(f){
    if(String(a[f])!==String(b[f])) diffs.push(label+'.'+f+': sheets="'+a[f]+'" sb="'+b[f]+'"');
  });
  if(!_numEq(a.total,b.total))      diffs.push(label+'.total: sheets='+a.total+' sb='+b.total);
  if(!_dateEqSec(a.fecha,b.fecha))  diffs.push(label+'.fecha: sheets="'+a.fecha+'" sb="'+b.fecha+'"');
}

/** Muestra las filas de VENTAS_CABECERA con ID_Venta vacío (las 3 que GAS incluye y Supabase no). Solo lectura.
 *  Reporta fila, si está 100% vacía, y las celdas con contenido (para decidir si limpiar o investigar). */
function verVentasSinIdME(){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName('VENTAS_CABECERA');
  if(!sh) return {ok:false, error:'VENTAS_CABECERA no encontrada'};
  var data=sh.getDataRange().getValues(), h=data[0], iId=h.indexOf('ID_Venta');
  var out={total:0, filas:[]};
  for(var i=1;i<data.length;i++){
    if(String(data[i][iId]||'').trim()===''){
      var vacio=data[i].every(function(v){ return v===''||v===null||v===undefined; });
      var o={ fila:i+1, todoVacio:vacio, contenido:{} };
      h.forEach(function(k,j){ var v=data[i][j]; if(v!==''&&v!==null&&v!==undefined) o.contenido[k]=(v instanceof Date)?v.toISOString():v; });
      out.filas.push(o);
    }
  }
  out.total=out.filas.length;
  Logger.log(JSON.stringify(out,null,2));
  return out;
}

/** Borra de VENTAS_CABECERA las filas TOTALMENTE vacías (ID_Venta vacío + toda la fila vacía).
 *  Seguro: SOLO filas 100% en blanco (nunca una venta con datos). Respalda en CORRECCIONES_LOG. Borra de abajo hacia arriba. */
function limpiarVentasSinIdME(){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName('VENTAS_CABECERA');
  if(!sh) return {ok:false, error:'VENTAS_CABECERA no encontrada'};
  var data=sh.getDataRange().getValues(), h=data[0], iId=h.indexOf('ID_Venta');
  var aBorrar=[];
  for(var i=1;i<data.length;i++){
    var idVacio   = String(data[i][iId]||'').trim()==='';
    var todoVacio = data[i].every(function(v){ return v===''||v===null||v===undefined; });
    if(idVacio && todoVacio) aBorrar.push(i+1);   // 1-based; SOLO si la fila entera está vacía
  }
  if(!aBorrar.length){ var z={ok:true, borradas:0, nota:'no hay filas 100% vacías'}; Logger.log(JSON.stringify(z)); return z; }
  var log=ss.getSheetByName('CORRECCIONES_LOG') || ss.insertSheet('CORRECCIONES_LOG');
  if(log.getLastRow()===0) log.appendRow(['fecha','accion','cambios_json']);
  log.appendRow([Utilities.formatDate(new Date(),'America/Lima','yyyy-MM-dd HH:mm'),'limpiarVentasSinIdME', JSON.stringify({filas_borradas:aBorrar})]);
  aBorrar.sort(function(a,b){ return b-a; }).forEach(function(f){ sh.deleteRow(f); });
  var res={ok:true, borradas:aBorrar.length, filas:aBorrar, nota:'corre compararVentasHoyZonaME para confirmar PARIDAD EXACTA'};
  Logger.log(JSON.stringify(res,null,2)); return res;
}
