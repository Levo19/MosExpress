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
function _meDate(v){ if(v==null||v==='')return null; var d=(v instanceof Date)?v:new Date(v); if(isNaN(d.getTime()))return null; return Utilities.formatDate(d,'America/Lima',"yyyy-MM-dd'T'HH:mm:ssXXX"); }
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
  ]},
  ventas_detalle: { sheet:'VENTAS_DETALLE', onConflict:'id_venta,linea', keyHeader:'ID_Venta', big:true, lineaBy:'id_venta', spec:[
    ['id_venta','ID_Venta','text'],['sku','SKU','text'],['nombre','Nombre','text'],['cantidad','Cantidad','num'],
    ['precio','Precio','num'],['subtotal','Subtotal','num'],['cod_barras','Cod_Barras','text'],
    ['valor_unitario','Valor_Unitario','num'],['tipo_igv','Tipo_IGV','int'],['unidad_medida','Unidad_Medida','text']
  ]},
  cajas: { sheet:'CAJAS', onConflict:'id_caja', keyHeader:'ID_Caja', spec:[
    ['id_caja','ID_Caja','text'],['vendedor','Vendedor','text'],['estacion','Estacion','text'],
    ['fecha_apertura','Fecha_Apertura','date'],['monto_inicial','Monto_Inicial','num'],['estado','Estado','text'],
    ['monto_final','Monto_Final','num'],['fecha_cierre','Fecha_Cierre','date'],['zona_id','Zona_ID','text'],
    ['printnode_id','PrintNode_ID','text']
  ]},
  movimientos_extra: { sheet:'MOVIMIENTOS_EXTRA', onConflict:'id_extra', keyHeader:'ID_Extra', spec:[
    ['id_extra','ID_Extra','text'],['id_caja','ID_Caja','text'],['ts','Timestamp','date'],['tipo','Tipo','text'],
    ['monto','Monto','num'],['concepto','Concepto','text'],['obs','Obs','text'],['registrado_por','Registrado_Por','text'],
    ['historial_cambios','historialCambios','json']
  ]},
  clientes_frecuentes: { sheet:'CLIENTES_FRECUENTES', onConflict:'documento', keyHeader:'Documento', spec:[
    ['documento','Documento','text'],['nombre','Nombre','text'],['tipo_doc','Tipo_Doc','int'],
    ['fecha_registro','Fecha_Registro','date'],['direccion','Direccion','text'],['historial_cambios','historialCambios','json']
  ], post:function(r,o){ if(r.nombre==null) r.nombre=_meText(o['Nombre_RazonSocial']); return r; } },
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
  auditorias: { sheet:'AUDITORIAS', onConflict:'id_auditoria', keyHeader:'ID_Auditoria', spec:[
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
  } else if(!cfg.insertOnly){ // pk simple: filtra sin pk + dedupe (gana el último)
    var pk=cfg.onConflict;
    rows=rows.filter(function(r){ return r[pk]!=null && r[pk]!==''; });
    var seen={}; rows.forEach(function(r){ seen[String(r[pk])]=r; });
    rows=Object.keys(seen).map(function(k){ return seen[k]; });
  }
  return rows;
}

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

// ---------- wrappers para el editor ----------
function dryRunME(){ return migrarME({dryRun:true}); }
function backfillME(){ return migrarME(); }
function resetCheckpointsME(){
  var props=PropertiesService.getScriptProperties();
  var n=0; Object.keys(_ME_SPECS).forEach(function(t){
    ['MEBF_'+t,'MEBF_DONE_'+t].forEach(function(k){ if(props.getProperty(k)!=null){ props.deleteProperty(k); n++; } });
  });
  Logger.log('Checkpoints/flags borrados: '+n); return {ok:true, borrados:n};
}
