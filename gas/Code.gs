// ============================================================
// MosExpress — Code.gs
// Router principal del Web App (GAS).
// Desplegar como Web App: Execute as Me, Anyone with link.
//
// Archivos del proyecto:
//   Code.gs       ← este archivo (router + helpers)
//   Catalogo.gs   ← descargarCatalogo, verificarDispositivo, consultarCliente
//   Ventas.gs     ← procesarVenta, correlativo, ventasHoyZona, detalleVenta
//   Caja.gs       ← apertura/cierre, cobros, anulaciones, créditos, extras
//   Guias.gs      ← guías, stock por zona, traslados, auditorías
//   Impresion.gs  ← procesarImpresion, imprimirTicketInternamente, ESC/POS
//   NubeFact.gs   ← emitirNubeFact (SUNAT CPE)
//   Radio.gs      ← radioConfig, topProductosHoy, setupRadioSheet (pantalla TV)
// ============================================================

function doGet(e) {
  var accion = e.parameter.accion;

  // ── Reporte HTML de cierre de caja (devuelve HTML, no JSON — intencional) ──
  if (accion === 'ver_cierre') return getCierreHtml(e.parameter.id_caja);

  // Todas las demás acciones siempre devuelven JSON (igual que doPost)
  try {
    if (accion === 'descargar')             return descargarCatalogo();
    if (accion === 'verificar_dispositivo') return verificarDispositivo(e.parameter.id);
    if (accion === 'ventas_hoy_zona')       return ventasHoyZonaFlip(e.parameter.prefijos, e.parameter.desde);  // [Fase 1.D] flag FUENTE_DATOS (default sheets)
    if (accion === 'detalle_venta')         return detalleVenta(e.parameter.id_venta);
    // [Fase 2 · gate lectura directa] paridad Sheets↔Supabase (solo lectura) — GET ?accion=verificar_paridad_lectura&dias=3
    if (accion === 'verificar_paridad_lectura') return ContentService.createTextOutput(JSON.stringify(verificarParidadLectura(e.parameter.dias))).setMimeType(ContentService.MimeType.JSON);
    if (accion === 'stock_zonas')           return getStockZonas();
    if (accion === 'lista_auditoria')       return getListaAuditoria(e.parameter.zona, e.parameter.usuario);
    if (accion === 'cajero_activo')         return cajeroActivo(e.parameter.zona);
    if (accion === 'caja_activa_zona')      return ContentService.createTextOutput(JSON.stringify(getCajaActivaZona({ zona: e.parameter.zona }))).setMimeType(ContentService.MimeType.JSON);
    if (accion === 'consultarCaja')         return consultarCaja(e.parameter.idCaja);
    if (accion === 'verificarHorarioME')    return verificarHorarioME(e.parameter);
    // [v2.5.58] Pre-reserva de correlativo (NV) — anti-LOCAL
    if (accion === 'reservarCorrelativo')   return reservarCorrelativo(e.parameter);
    if (accion === 'cancelarReservaCorrelativo') return cancelarReservaCorrelativo(e.parameter);
    // [v2.5.51] Retomar caja cuando localStorage se perdió pero la caja sigue ABIERTA
    if (accion === 'retomar_caja_device')   return retomarCajaPorDeviceId(e.parameter.deviceId);
    // [v2.5.55] Proxy GET para recuperar device state desde MOS (DEVICE_STATE)
    if (accion === 'getDeviceStateProxy')   return getDeviceStateProxy(e.parameter.deviceId);
    // [v2.5.33] Wizard moderno — estado bulk de impresoras + ping + cajeros activos
    if (accion === 'estado_impresoras')     return estadoImpresoras(e.parameter.ids);
    if (accion === 'cajeros_activos_todos') return cajerosActivosTodos();
    if (accion === 'ping')                  return ContentService.createTextOutput(JSON.stringify({ status: 'success', ok: true, ts: Date.now() })).setMimeType(ContentService.MimeType.JSON);
    // [v2.5.47] Endpoints para el Centro Tributario en MOS
    if (accion === 'tributario_ventas_mes')        return tributarioVentasMes(e.parameter.mes, e.parameter.anio);
    if (accion === 'tributario_cpe_mes')           return tributarioCPEMes(e.parameter.mes, e.parameter.anio);
    if (accion === 'tributario_reintentar_cpe')    return tributarioReintentarCPE(e.parameter.idVenta);
    if (accion === 'tributario_limpiar_huerfanas') return limpiarVentasHuerfanas();
    if (accion === 'tributario_reconciliar')       return reconciliarCPEsPendientes(35);
    if (accion === 'listar_guias')          return listarGuias(e.parameter.zona);
    if (accion === 'detalle_guia')          return detalleGuia(e.parameter.id_guia);
    if (accion === 'traslados_entrantes')   return trasladosEntrantes(e.parameter.zona, e.parameter.desde);
    if (accion === 'consultar_cliente')     return consultarCliente(e.parameter.doc, e.parameter.tipoDoc);
    // [v2.5.59] Diagnóstico APISPeru (verifica token + saldo + respuesta)
    if (accion === 'test_apisperu')         return testApiSperu();
    if (accion === 'test_apisperu_doc')     return testApiSperuDoc(e.parameter.doc);
    if (accion === 'extras_caja')           return getExtrasCaja(e.parameter.cajaId);
    // [v40.3] Sistema de cobro asignado de créditos (MOS ↔ ME)
    if (accion === 'creditos_pendientes')      return getCreditosPendientesFlip(e.parameter.diasAtras);  // [Fase 1.D] flag FUENTE_DATOS (default sheets)
    if (accion === 'cobros_asignados_cajero')  return getCobrosAsignadosCajero(e.parameter.cajaId);
    // [v2.5.28] Panel "Cobros en vuelo" para MOS Admin
    if (accion === 'cobros_en_vuelo_admin')    return getCobrosEnVueloAdminFlip();  // [Fase 1.D] flag FUENTE_DATOS (default sheets)
    if (accion === 'estado_cajas')          return estadoCajasFlip();              // [Fase 1.D] flag FUENTE_DATOS (default sheets)
    if (accion === 'historial_venta')       return getHistorialEndpoint('VENTAS_CABECERA', e.parameter.idVenta);
    if (accion === 'historial_extra')       return getHistorialEndpoint('MOVIMIENTOS_EXTRA', e.parameter.idExtra);
    if (accion === 'historial_cliente')     return getHistorialEndpoint('CLIENTES_FRECUENTES', e.parameter.doc);
    if (accion === 'radio_config')          return radioConfig();
    if (accion === 'top_productos_hoy')     return topProductosHoy();
    if (accion === 'radio_productos')       return radioProductos();
    return generarRespuestaError('Acción no válida: ' + accion);
  } catch(err) {
    return generarRespuestaError('Error interno [' + accion + ']: ' + err.message);
  }
}

// ── Estado completo de cajas con analítica en tiempo real ──────
// Devuelve TODAS las cajas (abiertas + cerradas del día) con
// totales de ventas calculados desde VENTAS_CABECERA.
function estadoCajas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cajasSheet  = ss.getSheetByName('CAJAS');
  var ventasSheet = ss.getSheetByName('VENTAS_CABECERA');
  var extrasSheet = ss.getSheetByName('MOVIMIENTOS_EXTRA');

  if (!cajasSheet) return generarRespuestaError('CAJAS no encontrada');

  var tz     = Session.getScriptTimeZone();
  var hoy    = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  // Límite: cerradas de los últimos 30 días
  var limite = new Date();
  limite.setDate(limite.getDate() - 30);

  // ── Agregar ventas por caja ───────────────────────────────────
  // VENTAS_CABECERA cols: 0=ID_Venta 6=Total 7=TipoDoc 8=FormaPago 10=ID_Caja 12=Estado
  var ventasPorCaja = {};  // { idCaja: { total, tickets, efectivo, otros, anulados, sinCobrar, byMetodo, byDoc } }
  if (ventasSheet) {
    var vd = ventasSheet.getDataRange().getValues();
    for (var v = 1; v < vd.length; v++) {
      var idCaja  = String(vd[v][10] || '');
      var metodo  = String(vd[v][8]  || 'EFECTIVO');
      var tipoDoc = String(vd[v][7]  || 'NOTA_DE_VENTA');
      var total   = parseFloat(vd[v][6]) || 0;
      // [MED14 500x-2] Anulado se deriva de FormaPago (col8) por PREFIJO, no de col12/Estado_Envio (que
      // SIEMPRE es 'COMPLETADO' — la rama era código muerto y contaba las anuladas como venta, inflando
      // el arqueo). Prefijo cubre 'ANULADO' y 'ANULADO_CONVERSION'.
      var anulado = metodo.toUpperCase().indexOf('ANULADO') === 0;
      if (!idCaja) continue;

      if (!ventasPorCaja[idCaja]) {
        ventasPorCaja[idCaja] = { total:0, tickets:0, efectivo:0, otros:0,
                                  anulados:0, sinCobrar:0, byMetodo:{}, byDoc:{} };
      }
      var vc = ventasPorCaja[idCaja];

      if (anulado) {
        vc.anulados++;
      } else if (metodo === 'POR_COBRAR') {
        vc.sinCobrar++;
        vc.tickets++;
      } else {
        vc.total   += total;
        vc.tickets++;
        if (metodo === 'EFECTIVO') {
          vc.efectivo += total;
        } else if (metodo.indexOf('MIXTO') === 0) {
          var _efeM = metodo.match(/EFE:([\d.]+)/i);
          var _virM = metodo.match(/VIR:([\d.]+)/i);
          var _efe  = _efeM ? parseFloat(_efeM[1]) : 0;
          var _vir  = _virM ? parseFloat(_virM[1]) : total - _efe;
          vc.efectivo += _efe;
          vc.otros    += _vir;
        } else {
          vc.otros += total;
        }
        vc.byMetodo[metodo] = (vc.byMetodo[metodo] || 0) + total;
        vc.byDoc[tipoDoc]   = (vc.byDoc[tipoDoc]   || 0) + total;
      }
    }
  }

  // ── Agregar extras por caja ───────────────────────────────────
  var extrasPorCaja = {};
  if (extrasSheet) {
    var ed = extrasSheet.getDataRange().getValues();
    for (var e2 = 1; e2 < ed.length; e2++) {
      var ec   = String(ed[e2][1] || '');
      var tipo = String(ed[e2][3] || 'EGRESO');
      var mto  = parseFloat(ed[e2][4]) || 0;
      if (!ec) continue;
      if (!extrasPorCaja[ec]) extrasPorCaja[ec] = { entradas:0, salidas:0, entradasVirtual:0, salidasVirtual:0 };
      if      (tipo === 'INGRESO')         extrasPorCaja[ec].entradas        += mto;
      else if (tipo === 'INGRESO_VIRTUAL') extrasPorCaja[ec].entradasVirtual += mto;
      else if (tipo === 'EGRESO')          extrasPorCaja[ec].salidas         += mto;
      else if (tipo === 'EGRESO_VIRTUAL')  extrasPorCaja[ec].salidasVirtual  += mto;
    }
  }

  // ── Construir resultado ───────────────────────────────────────
  var cajasData = cajasSheet.getDataRange().getValues();
  var abiertas  = [];
  var cerradas  = [];

  for (var i = 1; i < cajasData.length; i++) {
    var row    = cajasData[i];
    var idC    = String(row[0] || '');
    var estado2 = String(row[5] || '');
    var fApert = row[3] instanceof Date ? row[3] : (row[3] ? new Date(row[3]) : null);
    var fCierr = row[7] instanceof Date ? row[7] : (row[7] ? new Date(row[7]) : null);

    // Cerradas: incluir últimos 30 días
    if (estado2 === 'CERRADA' && fCierr && fCierr < limite) continue;
    // Cerradas sin fecha de cierre: omitir
    if (estado2 === 'CERRADA' && !fCierr) continue;

    var vc2   = ventasPorCaja[idC] || { total:0, tickets:0, efectivo:0, otros:0, anulados:0, sinCobrar:0, byMetodo:{}, byDoc:{} };
    var ext   = extrasPorCaja[idC] || { entradas:0, salidas:0 };
    var mInicial = parseFloat(row[4]) || 0;
    var mFinal   = parseFloat(row[6]) || 0;
    var efectivoEsp = mInicial + vc2.efectivo + ext.entradas - ext.salidas;
    var diferencia  = estado2 === 'CERRADA' ? (mFinal - efectivoEsp) : null;

    var obj = {
      idCaja:        idC,
      vendedor:      String(row[1] || ''),
      estacion:      String(row[2] || ''),
      zona:          String(row[8] || ''),
      estado:        estado2,
      fechaApertura: fApert ? Utilities.formatDate(fApert, tz, 'yyyy-MM-dd HH:mm') : '',
      fechaCierre:   fCierr ? Utilities.formatDate(fCierr, tz, 'yyyy-MM-dd HH:mm') : '',
      montoInicial:  mInicial,
      montoFinal:    mFinal,
      // analítica en tiempo real
      totalVentas:   Math.round(vc2.total * 100) / 100,
      tickets:       vc2.tickets,
      efectivo:      Math.round(vc2.efectivo * 100) / 100,
      otros:         Math.round(vc2.otros * 100) / 100,
      anulados:      vc2.anulados,
      sinCobrar:     vc2.sinCobrar,
      byMetodo:      vc2.byMetodo,
      byDoc:         vc2.byDoc,
      entradas:      ext.entradas,
      salidas:       ext.salidas,
      efectivoEsperado: Math.round(efectivoEsp * 100) / 100,
      diferencia:    diferencia !== null ? Math.round(diferencia * 100) / 100 : null
    };

    if (estado2 === 'ABIERTA') abiertas.push(obj);
    else                       cerradas.push(obj);
  }

  cerradas.reverse(); // más recientes primero

  // ── KPIs globales del día ─────────────────────────────────────
  var todasHoy = abiertas.concat(cerradas);
  var kpis = {
    cajasAbiertas:  abiertas.length,
    cajasCerradas:  cerradas.length,
    totalDia:       todasHoy.reduce(function(a,c){ return a + c.totalVentas; }, 0),
    ticketsDia:     todasHoy.reduce(function(a,c){ return a + c.tickets; }, 0),
    anuladosDia:    todasHoy.reduce(function(a,c){ return a + c.anulados; }, 0),
    sinCobrarDia:   todasHoy.reduce(function(a,c){ return a + c.sinCobrar; }, 0)
  };

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    generadoEn: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss'),
    kpis:    kpis,
    abiertas: abiertas,
    cerradas: cerradas
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    if (data.tipoEvento === 'APERTURA_CAJA')      return procesarAperturaCaja(data);
    if (data.tipoEvento === 'CIERRE_CAJA')         return procesarCierreCaja(data);
    if (data.tipoEvento === 'CIERRE_CAJA_FORZADO') return cerrarCajaForzado(data);
    // [v2.5.52] Retoma de caja por deviceId con autorización admin
    if (data.tipoEvento === 'CONFIRMAR_RETOMA_CAJA') return confirmarRetomaCaja(data);
    // [v2.5.55] Sync de DEVICE_STATE a MOS (ME actúa como proxy con MOS_WEB_APP_URL)
    if (data.tipoEvento === 'SYNC_DEVICE_STATE') return syncDeviceStateProxy(data);
    // [v2.5.58] Pre-reserva correlativo vía POST (acepta header)
    if (data.tipoEvento === 'RESERVAR_CORRELATIVO') return reservarCorrelativo(data);
    if (data.tipoEvento === 'CANCELAR_RESERVA_CORRELATIVO') return cancelarReservaCorrelativo(data);
    // [Fase 2] La PWA pide un JWT scoped por zona para hablar DIRECTO a Supabase (mint-token).
    if (data.tipoEvento === 'MINT_TOKEN') {
      var _mt = mintSupabaseToken(data.deviceId || data.id || '');
      return ContentService.createTextOutput(JSON.stringify(_mt)).setMimeType(ContentService.MimeType.JSON);
    }
    // [Fase 2] Espejo a Sheets de una venta NV creada directo en Supabase (cierre/SUNAT cuadran).
    if (data.tipoEvento === 'MIRROR_VENTA') {
      var _mv = mirrorVentaASheets(data);
      return ContentService.createTextOutput(JSON.stringify(_mv)).setMimeType(ContentService.MimeType.JSON);
    }
    // [Fase 2] Espejo a Sheets de un movimiento de caja creado directo en Supabase (cierre cuadra + alerta).
    if (data.tipoEvento === 'MIRROR_MOV') {
      var _mm = mirrorMovimientoASheets(data);
      return ContentService.createTextOutput(JSON.stringify(_mm)).setMimeType(ContentService.MimeType.JSON);
    }
    if (data.tipoEvento === 'CAMBIO_IMPRESORA_CAJA') return cambiarImpresoraCaja(data);
    if (data.tipoEvento === 'LIMPIAR_DUPLICADOS') {
      var _r = limpiarGuiasDuplicadasCaja(data.cajaId);
      return ContentService.createTextOutput(JSON.stringify(_r)).setMimeType(ContentService.MimeType.JSON);
    }
    if (data.tipoEvento === 'DIAGNOSTICAR_SALIDA_VENTAS') {
      var _d = diagnosticarSalidaVentas();
      return ContentService.createTextOutput(JSON.stringify(_d)).setMimeType(ContentService.MimeType.JSON);
    }
    if (data.tipoEvento === 'LIMPIAR_TODAS_DUPLICADAS') {
      var _t = limpiarTodasGuiasDuplicadas();
      return ContentService.createTextOutput(JSON.stringify(_t)).setMimeType(ContentService.MimeType.JSON);
    }
    if (data.tipoEvento === 'COBRAR_VENTA')        return cobrarVentaExistente(data);
    if (data.tipoEvento === 'ANULACION_MASIVA')    return anulacionMasiva(data);
    if (data.tipoEvento === 'CREDITAR_VENTA')      return creditarVenta(data);
    if (data.tipoEvento === 'EXTRA_CAJA')          return registrarExtraCajaConLog(data);
    if (data.tipoEvento === 'ANULACION')           return anularVentaIndividual(data);
    // Endpoints nuevos de edición posterior
    if (data.tipoEvento === 'COBRAR_CREDITO_CON_EXTRA')  return cobrarCreditoConExtra(data);
    if (data.tipoEvento === 'EDITAR_FORMA_PAGO_VENTA')   return editarFormaPagoVenta(data);
    if (data.tipoEvento === 'EDITAR_CLIENTE_VENTA')      return editarClienteVenta(data);
    if (data.tipoEvento === 'CONVERTIR_NV_A_CPE')        return convertirNVaCPE(data);
    if (data.tipoEvento === 'BAJA_CPE')                  return bajaCPEVenta(data);
    if (data.tipoEvento === 'REGISTRAR_GUIA')      return registrarGuia(data);
    if (data.tipoEvento === 'CERRAR_GUIA')         return cerrarGuia(data);   // [CUTOVER] cierre aplica stock una vez
    if (data.tipoEvento === 'REGISTRAR_AUDITORIA') return registrarAuditoria(data);
    // [v40.3] Cobro asignado de créditos
    if (data.tipoEvento === 'ASIGNAR_COBRO_CAJERO')    return asignarCobroACajero(data);
    if (data.tipoEvento === 'CONFIRMAR_COBRO_ASIGNADO') return confirmarCobroAsignado(data);
    if (data.tipoEvento === 'RECHAZAR_COBRO_ASIGNADO')  return rechazarCobroAsignado(data);
    // [v2.5.28] Admin puede cancelar/reasignar antes que el cajero procese
    if (data.tipoEvento === 'CANCELAR_COBRO_ASIGNADO')  return cancelarCobroAsignado(data);
    if (data.tipoEvento === 'REASIGNAR_COBRO_ASIGNADO') return reasignarCobroAsignado(data);
    if (data.accion === 'imprimir')                return procesarImpresion(data);
    // [v2.5.16] Verifica estado de impresora en PrintNode antes de imprimir
    if (data.accion === 'verificarImpresora')      return verificarEstadoImpresora(data);
    // [Mensajería] Relay del PUSH de un mensaje a los destinatarios presentes.
    // La RPC me.enviar_mensaje ya persistió el mensaje y resolvió la lista; acá GAS
    // dispara el FCM real (vía MOS enviarPushUsuario) por cada destinatario.
    if (data.action === 'msg_push_destinatarios')  return msgPushDestinatarios(data);

    // Default: registrar venta
    var response = procesarVenta(data);
    // [v2.5.45] Propagar error de payload inválido del guard anti-huérfanas
    if (response && response.error === 'PAYLOAD_INVALIDO') {
      return ContentService.createTextOutput(JSON.stringify({
        status: 'error',
        error: 'PAYLOAD_INVALIDO',
        mensaje: response.mensaje,
        campos_faltantes: response.campos_faltantes
      })).setMimeType(ContentService.MimeType.JSON);
    }
    // [v2.7.4] Propagar CUALQUIER error que devuelva procesarVenta — antes
    // solo PAYLOAD_INVALIDO. Esto causaba el bug de "ventas fantasma":
    // procesarVenta rechazaba con NO_CAJA_ACTIVA_EN_ZONA (idVenta=null) pero
    // doPost devolvía status=success → frontend marcaba como sincronizada
    // → ticket impreso pero sin fila en VENTAS_CABECERA.
    if (response && response.error && !response.idVenta) {
      // Registrar como VENTA_FANTASMA para revisión del admin
      try { _registrarVentaFantasma(data, response); } catch(eF) { Logger.log('[fantasma] error: ' + eF.message); }
      return ContentService.createTextOutput(JSON.stringify({
        status:  'error',
        error:   response.error,
        mensaje: response.mensaje || ('Venta rechazada: ' + response.error),
        idVenta: null,
        zona:    response.zona || ''
      })).setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({
      status:         "success",
      idVenta:        response.idVenta,
      correlativo:    response.correlativo,
      printDispatched:response.printDispatched,
      dedupVenta:     response.dedupVenta === true,   // [v40] flag anti ticket-doble
      nfEstado:       response.nfEstado || 'NA',
      nfHash:         response.nfHash   || '',
      nfEnlace:       response.nfEnlace || '',
      nfQrString:     response.nfQrString || '',
      mensaje:        "Venta procesada con éxito"
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return generarRespuestaError(error.toString());
  }
}

// ── Helpers compartidos (accesibles desde todos los módulos) ──
function generarRespuestaError(msg) {
  return ContentService.createTextOutput(JSON.stringify({
    status: "error", mensaje: msg
  })).setMimeType(ContentService.MimeType.JSON);
}

// Columnas que deben tratarse siempre como texto (nunca como número).
// [v2.7.2] Agregados campos de documento — los DNIs peruanos pueden empezar
// con cero (ej. 08123456). Si Sheets los infiere como Number, el cero se
// PIERDE para siempre. Estos nombres deben coincidir EXACTO con los headers
// de cada hoja (case-sensitive). NO afecta los datos existentes — solo cómo
// se LEEN: ahora llegan al frontend como string en vez de Number.
var COLUMNAS_TEXTO = [
  'Cod_Barras', 'Cod_Barras_Real', 'SKU_Base', 'SKU',
  'ID_Dispositivo', 'ID_Venta', 'ID_Caja', 'ID_Guia',
  // [v2.7.2] Documentos identificación (DNI puede empezar con 0)
  'Documento', 'Documento_Cliente', 'doc', 'docCliente',
  'DNI', 'RUC', 'numero_documento', 'Numero_Documento',
  'Cliente_Doc', 'Cliente_Documento'
];

function obtenerDatosHojaComoJSON(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  var result  = [];
  for (var i = 1; i < data.length; i++) {
    var rowData = {};
    for (var j = 0; j < headers.length; j++) {
      var header = String(headers[j]).trim();
      var val    = data[i][j];
      if (COLUMNAS_TEXTO.indexOf(header) !== -1) {
        val = (val === '' || val === null || val === undefined) ? '' : String(val).trim();
      }
      rowData[header] = val;
    }
    result.push(rowData);
  }
  return result;
}

// [v2.5.55] Proxy a MOS para sincronizar DEVICE_STATE. ME ya tiene
// MOS_WEB_APP_URL configurado en Script Properties — no es necesario
// exponer esa URL al cliente frontend.
function syncDeviceStateProxy(data) {
  if (!data || !data.deviceId) return generarRespuestaError('deviceId requerido');
  var url = PropertiesService.getScriptProperties().getProperty('MOS_WEB_APP_URL');
  if (!url) return generarRespuestaError('MOS_WEB_APP_URL no configurado');
  try {
    var resp = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'text/plain',
      payload: JSON.stringify({
        action:      'syncDeviceState',
        deviceId:    String(data.deviceId),
        app:         String(data.app || 'ME'),
        config:      data.config || null,
        cajaActiva:  data.cajaActiva || null,
        fechaSesion: String(data.fechaSesion || '')
      }),
      followRedirects: true,
      muteHttpExceptions: true
    });
    var txt = resp.getContentText();
    try {
      var parsed = JSON.parse(txt);
      return ContentService.createTextOutput(JSON.stringify({ status: 'success', proxied: true, mosResponse: parsed })).setMimeType(ContentService.MimeType.JSON);
    } catch(_) {
      return generarRespuestaError('MOS respondió no-JSON: ' + txt.substring(0, 150));
    }
  } catch(e) {
    return generarRespuestaError('Proxy MOS error: ' + (e && e.message || e));
  }
}

// [v2.5.55] Proxy GET para recuperar el snapshot remoto del device.
function getDeviceStateProxy(deviceId) {
  if (!deviceId) return generarRespuestaError('deviceId requerido');
  var url = PropertiesService.getScriptProperties().getProperty('MOS_WEB_APP_URL');
  if (!url) return generarRespuestaError('MOS_WEB_APP_URL no configurado');
  try {
    var resp = UrlFetchApp.fetch(url + '?action=getDeviceState&deviceId=' + encodeURIComponent(String(deviceId)), {
      method: 'get',
      followRedirects: true,
      muteHttpExceptions: true
    });
    var txt = resp.getContentText();
    try {
      var parsed = JSON.parse(txt);
      // Aplanar: el cliente espera data en el top-level
      var snap = parsed && parsed.data ? parsed.data : parsed;
      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        data: snap
      })).setMimeType(ContentService.MimeType.JSON);
    } catch(_) {
      return generarRespuestaError('MOS respondió no-JSON: ' + txt.substring(0, 150));
    }
  } catch(e) {
    return generarRespuestaError('Proxy MOS error: ' + (e && e.message || e));
  }
}

// ── Notificar a ProyectoMOS vía push (requiere MOS_WEB_APP_URL en Script Properties) ──
// Solo a MASTER/ADMIN. Auto-excluye al sender si es admin.
function _notificarMOS(titulo, cuerpo, excluirUsuario, idNotif) {
  var url = PropertiesService.getScriptProperties().getProperty('MOS_WEB_APP_URL');
  if (!url) { Logger.log('[Push] MOS_WEB_APP_URL no configurada'); return; }
  try {
    var resp = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        action: 'enviarPushNotif',
        titulo: titulo,
        cuerpo: cuerpo,
        soloRolesAdmin: true,
        excluirUsuario: excluirUsuario || null,
        idNotif: idNotif || ''
      }),
      muteHttpExceptions: true
    });
    Logger.log('[Push→MOS] HTTP ' + resp.getResponseCode() + ' | ' + resp.getContentText().substring(0, 120));
  } catch(e) { Logger.log('[Push→MOS] excepcion: ' + e.message); }
}

// ── [Mensajería] Disparar el PUSH FCM de un mensaje a los destinatarios ──────
// El FCM real lo hace MOS (Push.gs:enviarPushUsuario, que tiene la server key).
// ME NO conoce la key: solo reenvía. enviarPushUsuario matchea por NOMBRE (lower)
// contra PUSH_TOKENS, por eso usamos d.nombre como 'usuario'.
//
// Entrada (POST con action='msg_push_destinatarios'):
//   {
//     destinatarios: [ { id_personal, nombre, push_token } ],  // de me.enviar_mensaje
//     titulo, cuerpo,
//     prioridad?  ('normal'|'alta'),
//     mensajeId?  (para el deep-link in-app)
//   }
// Devuelve { status, data:{ enviados, fallidos, total, detalle } }.
// El IN-APP NO depende de esto: sale del polling de me.mis_mensajes. Esto es
// SOLO el aviso push en vivo; tolerante a fallos (un destinatario sin token en
// PUSH_TOKENS no rompe a los demás).
function msgPushDestinatarios(data) {
  var url = PropertiesService.getScriptProperties().getProperty('MOS_WEB_APP_URL');
  if (!url) return generarRespuestaError('MOS_WEB_APP_URL no configurado');

  var dests = (data && data.destinatarios) || [];
  if (!Array.isArray(dests) || dests.length === 0) {
    // Sin destinatarios presentes no es error: el mensaje quedó persistido y se
    // verá in-app cuando entren. Respondemos ok con 0 enviados.
    return ContentService.createTextOutput(JSON.stringify({
      status: 'success', data: { enviados: 0, fallidos: 0, total: 0, detalle: [] }
    })).setMimeType(ContentService.MimeType.JSON);
  }

  var titulo = String((data && data.titulo) || 'Mensaje');
  var cuerpo = String((data && data.cuerpo) || '');
  if ((data && data.prioridad) === 'alta' && titulo.indexOf('⚠') !== 0) titulo = '⚠ ' + titulo;

  var enviados = 0, fallidos = 0, detalle = [];
  for (var i = 0; i < dests.length; i++) {
    var d = dests[i] || {};
    var nombre = String(d.nombre || '').trim();
    if (!nombre) { fallidos++; detalle.push({ id_personal: d.id_personal || '', ok: false, error: 'sin nombre' }); continue; }
    try {
      var resp = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          action:  'enviarPushUsuario',
          usuario: nombre,
          titulo:  titulo,
          cuerpo:  cuerpo,
          idNotif: 'ME_MENSAJE',
          extra:   { mensajeId: (data && data.mensajeId) || null, id_personal: d.id_personal || '' }
        }),
        muteHttpExceptions: true
      });
      var ok = resp.getResponseCode() === 200;
      if (ok) enviados++; else fallidos++;
      detalle.push({ id_personal: d.id_personal || '', usuario: nombre, ok: ok, http: resp.getResponseCode() });
    } catch (ePush) {
      fallidos++;
      detalle.push({ id_personal: d.id_personal || '', usuario: nombre, ok: false, error: ePush.message });
    }
  }
  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    data: { enviados: enviados, fallidos: fallidos, total: dests.length, detalle: detalle }
  })).setMimeType(ContentService.MimeType.JSON);
}
