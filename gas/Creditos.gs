// ============================================================
// MosExpress — Creditos.gs
// Sistema de cobro asignado de créditos:
//   - asignarCobroACajero: admin/master asigna un crédito a una caja
//   - confirmarCobroAsignado: cajero confirma el cobro recibido
//   - rechazarCobroAsignado: cajero rechaza (cliente no llegó, etc)
//   - getCobrosAsignadosCajero: lista cobros pendientes de una caja
//   - getCreditosPendientes: lista todos los créditos sin cobrar
//   - escalarCobrosVencidos: trigger 5min detecta > 1h sin resolver
//
// Hoja CREDITOS_COBRO_ASIGNADO (se autocreaa):
//   ID_Cobro | ID_Venta | Caja_Destino | Vendedor_Dest | Metodo_Sug |
//   Estado | Admin_Asignador | Fecha_Asig | Fecha_Res | Razon |
//   ID_Caja_Origen | Monto | Cliente_Nombre | Correlativo
//
// Estados: ASIGNADO | COBRADO | RECHAZADO | TIMEOUT
// ============================================================

// [v2.5.27] Columnas nuevas: Fecha_Vencimiento + Horas_TTL.
// Permite vencimiento configurable (1h default · 2h · 4h · 6h) y la
// columna explícita evita recalcular Fecha_Asig + N horas en cada poll.
var _CREDITO_COBRO_HEADERS = [
  'ID_Cobro','ID_Venta','Caja_Destino','Vendedor_Dest','Metodo_Sug',
  'Estado','Admin_Asignador','Fecha_Asig','Fecha_Res','Razon',
  'ID_Caja_Origen','Monto','Cliente_Nombre','Correlativo',
  'Fecha_Vencimiento','Horas_TTL',
  // [v2.5.28] Mensaje opcional del admin al cajero + tracking de reasignaciones
  'Mensaje_Admin','Reasignaciones'
];

// ════════════════════════════════════════════════════════════
// [Lote1-A · fix C1+C2+C3] Lock GLOBAL del flujo de cobro de créditos.
// Sin esto, dos confirmaciones simultáneas del mismo cobro creaban DOS
// INGRESOS (doble cobro), y el trigger escalarCobrosVencidos podía revertir
// a CREDITO una venta MIENTRAS el cajero la cobraba (re-asignable = re-cobrable).
// Reentrante vía _credLockHeld (confirmarCobroAsignado llama internamente a
// cobrarCreditoConExtra; reasignar llama a asignar). Patrón _conLock de WH.
// REGLA: dentro del lock SOLO operaciones de Sheets (rápidas). Los UrlFetch
// pesados (push a MOS, reimpresión PrintNode) van FUERA — es el MISMO
// ScriptLock que usa el correlativo de ventas y no debemos retenerlo segundos.
// ════════════════════════════════════════════════════════════
var _credLockHeld = false;
function _conLockCred(fn, onBusy) {
  if (_credLockHeld) return fn();           // reentrante dentro de la misma ejecución
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch (e) { return onBusy ? onBusy() : { ok: false, error: 'Sistema ocupado' }; }
  _credLockHeld = true;
  try { return fn(); }
  finally { _credLockHeld = false; try { lock.releaseLock(); } catch(_){} }
}

function _getHojaCobrosAsignados() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('CREDITOS_COBRO_ASIGNADO');
  if (!sheet) {
    sheet = ss.insertSheet('CREDITOS_COBRO_ASIGNADO');
    sheet.appendRow(_CREDITO_COBRO_HEADERS);
    sheet.setFrozenRows(1);
    return sheet;
  }
  // [v2.5.27] Migrar headers — agregar columnas nuevas si faltan
  var firstRow = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), _CREDITO_COBRO_HEADERS.length)).getValues()[0];
  var current = firstRow.map(function(h){ return String(h || '').trim(); });
  var faltan = _CREDITO_COBRO_HEADERS.filter(function(h) { return current.indexOf(h) === -1; });
  if (faltan.length > 0) {
    var startCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, startCol, 1, faltan.length).setValues([faltan]);
    sheet.getRange(1, startCol, 1, faltan.length)
         .setFontWeight('bold').setBackground('#1f2937').setFontColor('#fff');
  }
  return sheet;
}

// ============================================================
// asignarCobroACajero — admin/master desde MOS asigna un crédito
// ============================================================
// Recibe via POST desde MOS (con clave admin validada).
// Crea row en CREDITOS_COBRO_ASIGNADO + push al cajero destino.
//
// payload:
// { tipoEvento: 'ASIGNAR_COBRO_CAJERO',
//   idVenta, cajaDestino, metodoSugerido ('EFECTIVO'|'VIRTUAL'|'MIXTO'),
//   adminAuth: { nombre, rol, via:'PIN_8DIG' } }
function asignarCobroACajero(data) {
  if (!data.idVenta)         return generarRespuestaError('idVenta requerido');
  if (!data.cajaDestino)     return generarRespuestaError('cajaDestino requerida');
  // [v2.5.39] metodoSugerido ya no es obligatorio — el cajero elige al cobrar
  if (!data.adminAuth || !data.adminAuth.nombre) {
    return generarRespuestaError('adminAuth requerido (esta acción requiere admin/master)');
  }
  // Normalizar método (puede venir vacío = "cajero decide")
  var metodoSugStr = String(data.metodoSugerido || '').toUpperCase().trim();

  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var ventas = ss.getSheetByName('VENTAS_CABECERA');
  var cajas  = ss.getSheetByName('CAJAS');
  if (!ventas || !cajas) return generarRespuestaError('Hojas requeridas no encontradas');

  // 1. Validar que la venta exista, sea CREDITO o POR_COBRAR, no anulada
  var fv = ventas.getDataRange().getValues();
  var ventaData = null;
  for (var i = fv.length - 1; i > 0; i--) {
    if (String(fv[i][0]) === String(data.idVenta)) {
      ventaData = {
        formaPago:   String(fv[i][8] || ''),
        correlativo: String(fv[i][9] || ''),
        cliente:     String(fv[i][5] || ''),
        total:       parseFloat(fv[i][6]) || 0,
        cajaOriginal:String(fv[i][10] || '')
      };
      break;
    }
  }
  if (!ventaData) return generarRespuestaError('Venta ' + data.idVenta + ' no encontrada');
  var fpUpper = ventaData.formaPago.toUpperCase();
  if (fpUpper !== 'CREDITO' && fpUpper !== 'POR_COBRAR') {
    return generarRespuestaError('Venta no está pendiente (estado: ' + ventaData.formaPago + ')');
  }

  // 2. Validar caja destino ABIERTA y obtener vendedor (cajero)
  var fc = cajas.getDataRange().getValues();
  var cajaInfo = null;
  for (var j = fc.length - 1; j > 0; j--) {
    if (String(fc[j][0]) === String(data.cajaDestino)) {
      cajaInfo = {
        vendedor: String(fc[j][1] || ''),
        estado:   String(fc[j][5] || '')
      };
      break;
    }
  }
  if (!cajaInfo || cajaInfo.estado !== 'ABIERTA') {
    return generarRespuestaError('Caja destino no está abierta');
  }

  // 3. Verificar que no haya OTRA asignación ASIGNADO para esta venta (idempotencia)
  var hoja = _getHojaCobrosAsignados();
  var fa = hoja.getDataRange().getValues();
  for (var k = 1; k < fa.length; k++) {
    if (String(fa[k][1]) === String(data.idVenta) && String(fa[k][5]) === 'ASIGNADO') {
      return generarRespuestaError('Ya hay un cobro ASIGNADO para esta venta (' + String(fa[k][0]) + ')');
    }
  }

  // 4. [v2.5.27] Calcular Fecha_Vencimiento según horasTTL (default 1h)
  var horasTTL = parseInt(data.horasTTL, 10);
  if (![1, 2, 4, 6].includes(horasTTL)) horasTTL = 1; // sanitizar
  var ahora = new Date();
  var fechaVencimiento = new Date(ahora.getTime() + horasTTL * 60 * 60 * 1000);

  // 5. Crear row de asignación
  var idCobro = 'CB-' + ahora.getTime();
  // [v2.5.28] Mensaje opcional del admin al cajero (max 140 chars)
  var mensajeAdmin = String(data.mensajeAdmin || '').substring(0, 140).trim();
  hoja.appendRow([
    idCobro, data.idVenta, data.cajaDestino, cajaInfo.vendedor,
    metodoSugStr, // [v2.5.39] puede ser '' (cajero decide al cobrar)
    'ASIGNADO',
    String(data.adminAuth.nombre || '').replace(/^admin:/i, ''),
    ahora, '', '',
    ventaData.cajaOriginal, ventaData.total, ventaData.cliente, ventaData.correlativo,
    fechaVencimiento, horasTTL,
    mensajeAdmin, 0  // reasignaciones=0 al inicio
  ]);
  SpreadsheetApp.flush();

  // [creditos-directo] espejo del cobro recién asignado a Supabase en tiempo real (best-effort).
  // Las transiciones de estado (cobrado/rechazado/expirado/cancelado/reasignado) siguen por batch+dirty-sync.
  try {
    _dualWriteCobroME({
      ID_Cobro: idCobro, ID_Venta: data.idVenta, Caja_Destino: data.cajaDestino,
      Vendedor_Dest: cajaInfo.vendedor, Metodo_Sug: metodoSugStr, Estado: 'ASIGNADO',
      Admin_Asignador: String(data.adminAuth.nombre || '').replace(/^admin:/i, ''),
      Fecha_Asig: ahora, Fecha_Res: '', Razon: '',
      ID_Caja_Origen: ventaData.cajaOriginal, Monto: ventaData.total,
      Cliente_Nombre: ventaData.cliente, Correlativo: ventaData.correlativo,
      Fecha_Vencimiento: fechaVencimiento, Horas_TTL: horasTTL,
      Mensaje_Admin: mensajeAdmin, Reasignaciones: 0
    });
  } catch (eDW) { Logger.log('[dualWrite cobro] ' + (eDW && eDW.message)); }

  // 5. Push al cajero destino via MOS
  try {
    var url = PropertiesService.getScriptProperties().getProperty('MOS_WEB_APP_URL');
    if (url) {
      var titulo = '💳 Cobro pendiente · ' + (ventaData.cliente || 'cliente');
      // [v2.5.39] No mencionar método si está vacío (cajero decide al cobrar)
      var cuerpo = (data.adminAuth.nombre || 'Admin').replace(/^admin:/i, '') +
                   ' te asignó un crédito de S/ ' + ventaData.total.toFixed(2) +
                   (metodoSugStr ? ' (' + metodoSugStr + ')' : '') + '. Tocá para cobrar.';
      UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          action: 'enviarPushUsuario',
          usuario: cajaInfo.vendedor,
          titulo: titulo,
          cuerpo: cuerpo,
          idNotif: 'CREDITO_COBRO_PENDIENTE',
          // Datos para que el cliente PWA pueda abrir el modal directo
          extra: { idCobro: idCobro, idVenta: data.idVenta, total: ventaData.total }
        }),
        muteHttpExceptions: true
      });
    }
  } catch(ePush) { Logger.log('Push asignar cobro: ' + ePush.message); }

  // 6. Log auditoría
  try {
    auditarLog('CREDITOS_COBRO_ASIGNADO', idCobro, {
      usuario: 'MOS-Admin', rol: 'ADMIN',
      source: 'MOS_ASIGNAR_COBRO',
      accion: 'crear',
      autorizadoPor: { nombre: data.adminAuth.nombre, rol: data.adminAuth.rol || 'ADMIN', via: 'PIN_8DIG' },
      ref: { idVenta: data.idVenta, cajaDestino: data.cajaDestino, vendedor: cajaInfo.vendedor, monto: ventaData.total },
      motivo: ''
    });
  } catch(eLog) { Logger.log('Log asignar: ' + eLog.message); }

  // [v2.5.27] Auto-instalar trigger horario (idempotente)
  try { _ensureTriggerEscalarCobros(); } catch(_){}

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    idCobro: idCobro,
    cajeroDestino: cajaInfo.vendedor,
    horasTTL: horasTTL,
    fechaVencimiento: fechaVencimiento.toISOString(),
    mensaje: 'Cobro asignado a ' + cajaInfo.vendedor + ' · vence en ' + horasTTL + 'h'
  })).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// confirmarCobroAsignado — el cajero recibe el dinero y confirma
// ============================================================
// Internamente llama a cobrarCreditoConExtra (que ya maneja todo:
// MOVIMIENTOS_EXTRA + cambio formaPago + audit). Después marca el
// row de CREDITOS_COBRO_ASIGNADO como COBRADO.
//
// payload:
// { tipoEvento: 'CONFIRMAR_COBRO_ASIGNADO',
//   idCobro, metodoFinal ('EFECTIVO'|'VIRTUAL'|'MIXTO (...)'),
//   montoEfectivo, montoVirtual (si MIXTO),
//   auth: { vendedor, esCajero, deviceId } }
function confirmarCobroAsignado(data) {
  if (!data.idCobro)     return generarRespuestaError('idCobro requerido');
  if (!data.metodoFinal) return generarRespuestaError('metodoFinal requerido');

  // [Lote1-A · fix C2] SECCIÓN CRÍTICA bajo el lock global de cobros:
  // leer estado → cobrar (reentrante) → marcar COBRADO, todo atómico.
  // Antes la validación 'ASIGNADO' y la marca 'COBRADO' estaban separadas por
  // todo el flujo del cobro → dos confirmaciones concurrentes del mismo idCobro
  // ambas veían ASIGNADO → doble cobro (TOCTOU). El push al admin y la
  // reimpresión (UrlFetch lentos) quedan FUERA del lock.
  var res = _conLockCred(function() {
    var hoja = _getHojaCobrosAsignados();
    var fa = hoja.getDataRange().getValues();
    var cRow = -1, cobroData = null;
    for (var i = 1; i < fa.length; i++) {
      if (String(fa[i][0]) === String(data.idCobro)) {
        cRow = i + 1;
        cobroData = {
          idVenta:    String(fa[i][1]),
          cajaDest:   String(fa[i][2]),
          estado:     String(fa[i][5]),
          adminAsig:  String(fa[i][6]),
          monto:      parseFloat(fa[i][11]) || 0,
          cliente:    String(fa[i][12])
        };
        break;
      }
    }
    if (cRow < 2) return { error: 'Cobro ' + data.idCobro + ' no encontrado' };
    if (cobroData.estado !== 'ASIGNADO') {
      return { error: 'El cobro no está en estado ASIGNADO (actual: ' + cobroData.estado + ')' };
    }

    // Llamar internamente al endpoint existente cobrarCreditoConExtra
    // (reentrante: _credLockHeld=true → no re-toma el lock)
    var cobroResp = cobrarCreditoConExtra({
      idVenta:       cobroData.idVenta,
      cajaReceptora: cobroData.cajaDest,
      metodo:        data.metodoFinal,
      montoEfectivo: data.montoEfectivo,
      montoVirtual:  data.montoVirtual,
      obs:           'Cobro asignado ' + data.idCobro,
      auth:          data.auth || {},
      adminAuth:     null   // ya autorizado al asignar, el cajero solo confirma
    });

    // Detectar error del cobrarCreditoConExtra (ContentService no se puede leer directo)
    var cobroResult;
    try {
      cobroResult = JSON.parse(cobroResp.getContent());
    } catch(eP) {
      return { error: 'Error parsing cobro response' };
    }
    if (cobroResult.status !== 'success') {
      return { error: 'Error procesando cobro: ' + (cobroResult.mensaje || '') };
    }

    // Marcar row CREDITOS_COBRO_ASIGNADO como COBRADO
    hoja.getRange(cRow, 6).setValue('COBRADO');                            // Estado
    hoja.getRange(cRow, 9).setValue(new Date());                            // Fecha_Res
    hoja.getRange(cRow, 5).setValue(String(data.metodoFinal).toUpperCase());// Metodo final
    SpreadsheetApp.flush();
    return { ok: true, cobroData: cobroData };
  }, function() {
    return { error: 'Sistema ocupado procesando otro cobro — reintenta en unos segundos' };
  });

  if (res.error) return generarRespuestaError(res.error);
  var cobroData = res.cobroData;

  // ── Desde aquí, FUERA del lock (best-effort, no afectan el dinero ya registrado) ──
  // [creditos-directo] espejo de la transición a Supabase en tiempo real (best-effort)
  try { _dualWriteCobroPatchME(data.idCobro, { estado:'COBRADO', fecha_res:new Date(), metodo_sug:String(data.metodoFinal).toUpperCase() }); } catch(_dw){}

  // Push de vuelta a admin (cierre del ciclo)
  try {
    var url = PropertiesService.getScriptProperties().getProperty('MOS_WEB_APP_URL');
    if (url && cobroData.adminAsig) {
      UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          action: 'enviarPushUsuario',
          usuario: cobroData.adminAsig,
          titulo: '✅ Cobro confirmado · ' + cobroData.cliente,
          cuerpo: 'S/ ' + cobroData.monto.toFixed(2) + ' cobrado en ' + cobroData.cajaDest +
                  ' (' + String(data.metodoFinal).toUpperCase() + ')',
          idNotif: 'CREDITO_COBRO_CONFIRMADO'
        }),
        muteHttpExceptions: true
      });
    }
  } catch(_){}

  // Reimprimir ticket con sello "PAGADO" si la caja tiene printer
  try {
    var pos = _resolverPrinterCaja(cobroData.cajaDest);
    if (pos && pos.printerId) {
      _reimprimirTicketConSello(cobroData.idVenta, pos.printerId, {
        metodoFinal: data.metodoFinal,
        cajaDest:    cobroData.cajaDest,
        cajero:      (data.auth && data.auth.vendedor) || '',
        adminAsig:   cobroData.adminAsig
      });
    }
  } catch(eImp) { Logger.log('Reimpresion sello: ' + eImp.message); }

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    mensaje: 'Cobro confirmado',
    idCobro: data.idCobro,
    idVenta: cobroData.idVenta,
    monto:   cobroData.monto
  })).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// rechazarCobroAsignado — el cajero rechaza con razón
// ============================================================
// payload: { tipoEvento:'RECHAZAR_COBRO_ASIGNADO', idCobro, razon, auth }
function rechazarCobroAsignado(data) {
  if (!data.idCobro) return generarRespuestaError('idCobro requerido');
  if (!data.razon)   return generarRespuestaError('razon es obligatoria');

  // [Lote1-A] Validación + transición bajo el MISMO lock que confirmar/escalar:
  // un rechazo concurrente con una confirmación ya no puede pisar el estado.
  var res = _conLockCred(function() {
    var hoja = _getHojaCobrosAsignados();
    var fa = hoja.getDataRange().getValues();
    var cRow = -1, cobroData = null;
    for (var i = 1; i < fa.length; i++) {
      if (String(fa[i][0]) === String(data.idCobro)) {
        cRow = i + 1;
        cobroData = {
          estado:    String(fa[i][5]),
          adminAsig: String(fa[i][6]),
          cliente:   String(fa[i][12]),
          monto:     parseFloat(fa[i][11]) || 0,
          cajaDest:  String(fa[i][2])
        };
        break;
      }
    }
    if (cRow < 2) return { error: 'Cobro no encontrado' };
    if (cobroData.estado !== 'ASIGNADO') {
      return { error: 'Solo se puede rechazar un cobro ASIGNADO' };
    }

    hoja.getRange(cRow, 6).setValue('RECHAZADO');
    hoja.getRange(cRow, 9).setValue(new Date());
    hoja.getRange(cRow, 10).setValue(String(data.razon).substring(0, 250));
    SpreadsheetApp.flush();
    return { ok: true, cobroData: cobroData };
  }, function() {
    return { error: 'Sistema ocupado — reintenta en unos segundos' };
  });
  if (res.error) return generarRespuestaError(res.error);
  var cobroData = res.cobroData;

  // ── Fuera del lock (best-effort) ──
  // [creditos-directo] espejo de la transición a Supabase en tiempo real (best-effort)
  try { _dualWriteCobroPatchME(data.idCobro, { estado:'RECHAZADO', fecha_res:new Date(), razon:String(data.razon).substring(0, 250) }); } catch(_dw){}

  // Push al admin que asignó (para reaccionar)
  try {
    var url = PropertiesService.getScriptProperties().getProperty('MOS_WEB_APP_URL');
    if (url && cobroData.adminAsig) {
      UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          action: 'enviarPushUsuario',
          usuario: cobroData.adminAsig,
          titulo: '⚠ Cobro rechazado · ' + cobroData.cliente,
          cuerpo: 'S/ ' + cobroData.monto.toFixed(2) + ' rechazado: ' + String(data.razon).substring(0, 100),
          idNotif: 'CREDITO_COBRO_RECHAZADO'
        }),
        muteHttpExceptions: true
      });
    }
  } catch(_){}

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success', mensaje: 'Cobro rechazado'
  })).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// getCobrosAsignadosCajero — el cajero pregunta qué le asignaron
// ============================================================
// GET ?accion=cobros_asignados_cajero&cajaId=CAJA-X
// Devuelve los cobros con estado ASIGNADO para esa caja.
function getCobrosAsignadosCajero(cajaId) {
  if (!cajaId) return generarRespuestaError('cajaId requerido');

  // [delete-safe] Supabase primero: me.creditos_cobro_asignado (poblada por _dualWriteCobroME /
  // _dualWriteCobroPatchME) filtrada por caja_destino + estado ASIGNADO. Enriquecemos items del
  // ticket original desde me.ventas_detalle y el vendedor desde me.ventas. Sheet fallback abajo.
  if (typeof _meLecturaCierreDirecta === 'function' && _meLecturaCierreDirecta()) {
    try {
      var rSB = _sb('GET', 'me.creditos_cobro_asignado', {
        select: 'id_cobro,id_venta,caja_destino,vendedor_dest,metodo_sug,admin_asignador,fecha_asig,monto,cliente_nombre,correlativo,fecha_vencimiento,horas_ttl,mensaje_admin',
        filters: { caja_destino: 'eq.' + String(cajaId), estado: 'eq.ASIGNADO' },
        order: 'fecha_vencimiento.asc',
        maxRetry: 1
      });
      if (rSB && rSB.ok && Array.isArray(rSB.data)) {
        var resSB = rSB.data.map(function(r) {
          var idv = String(r.id_venta || '');
          var itemsOrig = [], vendOrig = '';
          try {
            var rDet = _sb('GET', 'me.ventas_detalle', {
              select: 'nombre,cantidad,precio,subtotal',
              filters: { id_venta: 'eq.' + idv }, order: 'linea.asc', limit: 20, maxRetry: 1
            });
            if (rDet && rDet.ok && Array.isArray(rDet.data)) {
              itemsOrig = rDet.data.map(function(d){
                var cant = parseFloat(d.cantidad) || 0, prec = parseFloat(d.precio) || 0;
                var sub = parseFloat(d.subtotal) || (cant * prec);
                return { nombre: String(d.nombre || ''), cantidad: cant, precio: prec, subtotal: sub };
              });
            }
          } catch (_eD) {}
          try {
            var rVen = _sb('GET', 'me.ventas', { select: 'vendedor', filters: { id_venta: 'eq.' + idv }, limit: 1, maxRetry: 1 });
            if (rVen && rVen.ok && rVen.data && rVen.data.length) vendOrig = String(rVen.data[0].vendedor || '');
          } catch (_eV) {}
          var fVenc = r.fecha_vencimiento ? String(r.fecha_vencimiento) : '';
          var fAsig = r.fecha_asig ? String(r.fecha_asig) : '';
          if (!fVenc && fAsig) {
            var ttl = parseInt(r.horas_ttl, 10) || 1;
            fVenc = new Date(new Date(fAsig).getTime() + ttl * 3600000).toISOString();
          }
          return {
            idCobro:          String(r.id_cobro || ''),
            idVenta:          idv,
            cajaDestino:      String(r.caja_destino || ''),
            vendedorDest:     String(r.vendedor_dest || ''),
            metodoSug:        String(r.metodo_sug || ''),
            adminAsig:        String(r.admin_asignador || ''),
            fechaAsig:        fAsig,
            fechaVencimiento: fVenc,
            horasTTL:         parseInt(r.horas_ttl, 10) || 1,
            monto:            parseFloat(r.monto) || 0,
            cliente:          String(r.cliente_nombre || ''),
            correlativo:      String(r.correlativo || ''),
            mensajeAdmin:     String(r.mensaje_admin || ''),
            itemsOriginal:    itemsOrig,
            vendedorOriginal: vendOrig
          };
        });
        return ContentService.createTextOutput(JSON.stringify({
          status: 'success', cobros: resSB
        })).setMimeType(ContentService.MimeType.JSON);
      }
    } catch (eSB) { Logger.log('[getCobrosAsignadosCajero] Supabase: ' + (eSB && eSB.message)); }
  }

  var hoja = _getHojaCobrosAsignados();
  var fa = hoja.getDataRange().getValues();
  var result = [];
  // [v2.5.51] Pre-cargar items del ticket original para imprimirlos como
  // copia adjunta al aviso de cobro asignado (cliente pregunta "por qué tanto").
  // Construimos un map idVenta → { items, vendedor } leyendo una sola vez.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var itemsPorVenta = {};
  var vendedorPorVenta = {};
  try {
    var detSh = ss.getSheetByName('VENTAS_DETALLE');
    if (detSh) {
      var fd = detSh.getDataRange().getValues();
      if (fd.length >= 2) {
        var hdrsD = fd[0].map(function(h){ return String(h).trim(); });
        var iIdD   = hdrsD.indexOf('ID_Venta');
        var iNomD  = hdrsD.indexOf('Nombre');
        var iCantD = hdrsD.indexOf('Cantidad');
        var iPrecD = hdrsD.indexOf('Precio');
        var iSubD  = hdrsD.indexOf('Subtotal');
        if (iIdD   < 0) iIdD   = 0;
        if (iNomD  < 0) iNomD  = 2;
        if (iCantD < 0) iCantD = 3;
        if (iPrecD < 0) iPrecD = 4;
        if (iSubD  < 0) iSubD  = 5;
        for (var dd = 1; dd < fd.length; dd++) {
          var idVD = String(fd[dd][iIdD]);
          if (!idVD) continue;
          if (!itemsPorVenta[idVD]) itemsPorVenta[idVD] = [];
          if (itemsPorVenta[idVD].length >= 20) continue;
          var cantD = parseFloat(fd[dd][iCantD]) || 0;
          var subD  = parseFloat(fd[dd][iSubD])  || 0;
          var precD = parseFloat(fd[dd][iPrecD]) || 0;
          if (!subD) subD = cantD * precD;
          itemsPorVenta[idVD].push({
            nombre:   String(fd[dd][iNomD] || ''),
            cantidad: cantD,
            precio:   precD,
            subtotal: subD
          });
        }
      }
    }
  } catch(eDet) { Logger.log('items detalle: ' + eDet.message); }
  // Cargar vendedor original desde VENTAS_CABECERA
  try {
    var venSh = ss.getSheetByName('VENTAS_CABECERA');
    if (venSh) {
      var fv = venSh.getDataRange().getValues();
      var hdrsV = fv[0].map(function(h){ return String(h).trim(); });
      var iIdV = hdrsV.indexOf('ID_Venta');
      var iVndV = hdrsV.indexOf('Vendedor');
      if (iIdV < 0) iIdV = 0;
      if (iVndV < 0) iVndV = 2;
      for (var vv = 1; vv < fv.length; vv++) {
        vendedorPorVenta[String(fv[vv][iIdV])] = String(fv[vv][iVndV] || '');
      }
    }
  } catch(eVen) { Logger.log('vendedor cabec: ' + eVen.message); }
  // [v2.5.27] Indexar columnas por header para soportar columnas nuevas
  var hdrs = fa[0].map(function(h){ return String(h || '').trim(); });
  var iIdCobro   = hdrs.indexOf('ID_Cobro');
  var iIdVenta   = hdrs.indexOf('ID_Venta');
  var iCajaDest  = hdrs.indexOf('Caja_Destino');
  var iVendDest  = hdrs.indexOf('Vendedor_Dest');
  var iMetodo    = hdrs.indexOf('Metodo_Sug');
  var iEstado    = hdrs.indexOf('Estado');
  var iAdminAsig = hdrs.indexOf('Admin_Asignador');
  var iFAsig     = hdrs.indexOf('Fecha_Asig');
  var iMonto     = hdrs.indexOf('Monto');
  var iCliente   = hdrs.indexOf('Cliente_Nombre');
  var iCorrel    = hdrs.indexOf('Correlativo');
  var iVenc      = hdrs.indexOf('Fecha_Vencimiento');
  var iHorasTTL  = hdrs.indexOf('Horas_TTL');
  var iMsgAdmin  = hdrs.indexOf('Mensaje_Admin');
  for (var i = 1; i < fa.length; i++) {
    if (String(fa[i][iCajaDest]) !== String(cajaId)) continue;
    if (String(fa[i][iEstado]) !== 'ASIGNADO')      continue;
    // Calcular Fecha_Vencimiento — si no está set, fallback a Fecha_Asig + horasTTL (o 1h)
    var fVencISO = '';
    if (iVenc >= 0 && fa[i][iVenc]) {
      fVencISO = fa[i][iVenc] instanceof Date ? fa[i][iVenc].toISOString() : new Date(fa[i][iVenc]).toISOString();
    } else if (iFAsig >= 0 && fa[i][iFAsig]) {
      var fAsigMs = fa[i][iFAsig] instanceof Date ? fa[i][iFAsig].getTime() : new Date(fa[i][iFAsig]).getTime();
      var ttl = (iHorasTTL >= 0 ? (parseInt(fa[i][iHorasTTL], 10) || 1) : 1);
      fVencISO = new Date(fAsigMs + ttl * 60 * 60 * 1000).toISOString();
    }
    result.push({
      idCobro:           String(fa[i][iIdCobro]),
      idVenta:           String(fa[i][iIdVenta]),
      cajaDestino:       String(fa[i][iCajaDest]),
      vendedorDest:      String(fa[i][iVendDest]),
      metodoSug:         String(fa[i][iMetodo]),
      adminAsig:         String(fa[i][iAdminAsig]),
      fechaAsig:         fa[i][iFAsig] instanceof Date ? fa[i][iFAsig].toISOString() : String(fa[i][iFAsig] || ''),
      fechaVencimiento:  fVencISO,
      horasTTL:          iHorasTTL >= 0 ? (parseInt(fa[i][iHorasTTL], 10) || 1) : 1,
      monto:             parseFloat(fa[i][iMonto]) || 0,
      cliente:           String(fa[i][iCliente]),
      correlativo:       String(fa[i][iCorrel]),
      // [v2.5.28] Mensaje opcional del admin
      mensajeAdmin:      iMsgAdmin >= 0 ? String(fa[i][iMsgAdmin] || '') : '',
      // [v2.5.51] Copia del ticket original para reimprimirlo abajo del aviso
      itemsOriginal:     itemsPorVenta[String(fa[i][iIdVenta])] || [],
      vendedorOriginal:  vendedorPorVenta[String(fa[i][iIdVenta])] || ''
    });
  }
  return ContentService.createTextOutput(JSON.stringify({
    status: 'success', cobros: result
  })).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// getCreditosPendientes — para mostrar en MOS las cards
// ============================================================
// GET ?accion=creditos_pendientes&diasAtras=30
// Lista ventas con FormaPago=CREDITO o POR_COBRAR de los últimos N días,
// agrupadas por día. Excluye las que ya tienen una asignación COBRADA.
function getCreditosPendientes(diasAtras) {
  var dias = parseInt(diasAtras, 10) || 30;
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var ventas = ss.getSheetByName('VENTAS_CABECERA');
  if (!ventas) return generarRespuestaError('VENTAS_CABECERA no encontrada');

  // Set de idVenta YA cobradas (estado COBRADO en CREDITOS_COBRO_ASIGNADO)
  var cobradas = {};
  var asigSet  = {};   // {idVenta: idCobro} si está ASIGNADO
  try {
    var hoja = _getHojaCobrosAsignados();
    var fa = hoja.getDataRange().getValues();
    for (var k = 1; k < fa.length; k++) {
      var estK = String(fa[k][5]);
      if (estK === 'COBRADO')  cobradas[String(fa[k][1])] = true;
      if (estK === 'ASIGNADO') asigSet[String(fa[k][1])]  = {
        idCobro: String(fa[k][0]),
        cajaDestino: String(fa[k][2]),
        vendedorDest: String(fa[k][3]),
        fechaAsig: fa[k][7]
      };
    }
  } catch(_){}

  var tz = Session.getScriptTimeZone();
  var hoy = new Date();
  var limite = new Date();
  limite.setDate(hoy.getDate() - dias);

  // [v41.2] Cargar VENTAS_DETALLE una sola vez para mapear items por idVenta.
  // Resolvemos los índices de columna leyendo la cabecera real (no asumir orden).
  var itemsPorVenta = {};
  try {
    var detSh = ss.getSheetByName('VENTAS_DETALLE');
    if (detSh) {
      var fd = detSh.getDataRange().getValues();
      if (fd.length >= 2) {
        var hdrs = fd[0].map(function(h){ return String(h).trim(); });
        var iId   = hdrs.indexOf('ID_Venta');
        var iNom  = hdrs.indexOf('Nombre');
        var iCant = hdrs.indexOf('Cantidad');
        var iPrec = hdrs.indexOf('Precio');
        var iSub  = hdrs.indexOf('Subtotal');
        // Fallbacks defensivos si el header está en otro idioma/case
        if (iId   < 0) iId   = 0;
        if (iNom  < 0) iNom  = 2;
        if (iCant < 0) iCant = 3;
        if (iPrec < 0) iPrec = 4;
        if (iSub  < 0) iSub  = 5;
        for (var d = 1; d < fd.length; d++) {
          var idVD = String(fd[d][iId]);
          if (!idVD) continue;
          if (!itemsPorVenta[idVD]) itemsPorVenta[idVD] = [];
          if (itemsPorVenta[idVD].length >= 12) continue; // tope generoso por ticket
          var cant = parseFloat(fd[d][iCant]) || 0;
          var sub  = parseFloat(fd[d][iSub])  || 0;
          // Si Subtotal está vacío pero hay cantidad × precio, calcular
          if (!sub) sub = cant * (parseFloat(fd[d][iPrec]) || 0);
          itemsPorVenta[idVD].push({
            nombre:   String(fd[d][iNom] || ''),
            cantidad: cant,
            subtotal: sub
          });
        }
      }
    }
  } catch(_){}

  var fv = ventas.getDataRange().getValues();
  // Agrupar por día
  var porDia = {};
  for (var i = 1; i < fv.length; i++) {
    var fp = String(fv[i][8] || '').toUpperCase();
    // [v40.5] Solo CREDITO entra a la baraja. POR_COBRAR es del flow del
    // turno del vendedor (no es un crédito formal otorgado al cliente).
    if (fp !== 'CREDITO') continue;
    var idV = String(fv[i][0]);
    if (cobradas[idV]) continue;

    var fecha = fv[i][1] instanceof Date ? fv[i][1] : new Date(fv[i][1]);
    if (isNaN(fecha.getTime())) continue;
    if (fecha < limite) continue;

    var diaKey = Utilities.formatDate(fecha, tz, 'yyyy-MM-dd');
    if (!porDia[diaKey]) porDia[diaKey] = [];
    var itemsTicket = itemsPorVenta[idV] || [];
    porDia[diaKey].push({
      idVenta:        idV,
      correlativo:    String(fv[i][9] || ''),
      cliente:        String(fv[i][5] || ''),
      clienteDoc:     String(fv[i][4] || ''),
      vendedor:       String(fv[i][2] || ''),
      total:          parseFloat(fv[i][6]) || 0,
      formaPago:      String(fv[i][8] || ''),
      obs:            String(fv[i][14] || ''),
      idCaja:         String(fv[i][10] || ''),
      fechaISO:       Utilities.formatDate(fecha, tz, 'yyyy-MM-dd HH:mm:ss'),
      asignado:       asigSet[idV] || null,  // si ya está siendo cobrado
      items:          itemsTicket,           // [v41] top 5 items para mostrar resumen
      itemsCount:     itemsTicket.length
    });
  }

  // Construir array ordenado de días (más reciente primero)
  var dias_ = Object.keys(porDia).sort(function(a,b){ return b.localeCompare(a); });
  var grupos = dias_.map(function(d) {
    return {
      fecha:   d,
      tickets: porDia[d],
      total:   porDia[d].reduce(function(s,t){ return s + t.total; }, 0),
      cuenta:  porDia[d].length
    };
  });
  var totalAcum = grupos.reduce(function(s,g){ return s + g.total; }, 0);

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    grupos: grupos,
    totalAcumulado: totalAcum,
    totalTickets: grupos.reduce(function(s,g){ return s + g.cuenta; }, 0)
  })).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// escalarCobrosVencidos — trigger 5min
// ============================================================
// Marca como TIMEOUT los cobros ASIGNADO > 1 hora y dispara push
// de escalación a admin+master en MOS.
// [v2.5.27] Procesador de vencimientos — debe correr via trigger horario.
// Compara Fecha_Vencimiento contra ahora. Si venció:
//   1. Estado: EXPIRADO (semánticamente "vencido sin cobrarse")
//   2. VENTAS_CABECERA.formaPago = 'CREDITO' (vuelve al pool original)
//   3. VENTAS_CABECERA.cobrado = false
//   4. Push push al admin asignador: "expiró, re-asignar?"
// Ya NO usa el TTL hardcodeado 1h — lee Fecha_Vencimiento de cada row.
function escalarCobrosVencidos() {
  // [Lote1-A · fix C3] TODO el escaneo bajo el MISMO lock que confirmar/cobrar.
  // Antes el trigger podía leer la venta aún en CREDITO (el cajero no llegaba al
  // setValue) → marcaba EXPIRADO y revertía la venta MIENTRAS se cobraba →
  // INGRESO registrado + venta re-asignable = doble cobro. Con el lock, el
  // trigger espera a que el cobro termine y su guard ve el FormaPago final.
  // Los push (UrlFetch lentos) se recolectan y se envían FUERA del lock.
  var pushes = [];
  var out = _conLockCred(function() {
    return _escalarCobrosVencidosCore(pushes);
  }, function() {
    Logger.log('[escalarCobros] lock ocupado — el próximo tick (5min) reintenta');
    return { ok: false, error: 'lock ocupado' };
  });
  // ── Push fuera del lock ──
  if (pushes.length) {
    var url = PropertiesService.getScriptProperties().getProperty('MOS_WEB_APP_URL');
    if (url) {
      pushes.forEach(function(p) {
        try {
          UrlFetchApp.fetch(url, {
            method: 'post', contentType: 'application/json',
            payload: JSON.stringify(p), muteHttpExceptions: true
          });
        } catch(_){}
      });
    }
  }
  return out;
}
function _escalarCobrosVencidosCore(pushes) {
  var hoja = _getHojaCobrosAsignados();
  var fa = hoja.getDataRange().getValues();
  if (fa.length < 2) return { ok: true, vencidos: 0 };
  // Indexar columnas por header (resiliente a reordenamiento)
  var hdrs = fa[0].map(function(h){ return String(h || '').trim(); });
  var iEstado    = hdrs.indexOf('Estado');
  var iAsig      = hdrs.indexOf('Fecha_Asig');
  var iRes       = hdrs.indexOf('Fecha_Res');
  var iRazon     = hdrs.indexOf('Razon');
  var iIdVenta   = hdrs.indexOf('ID_Venta');
  var iMonto     = hdrs.indexOf('Monto');
  var iCliente   = hdrs.indexOf('Cliente_Nombre');
  var iVendedor  = hdrs.indexOf('Vendedor_Dest');
  var iAdminAsig = hdrs.indexOf('Admin_Asignador');
  var iVenc      = hdrs.indexOf('Fecha_Vencimiento');
  var iHorasTTL  = hdrs.indexOf('Horas_TTL');
  var ahora = new Date().getTime();
  var UNA_HORA = 60 * 60 * 1000;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ventasSh = ss.getSheetByName('VENTAS_CABECERA');
  var n = 0;

  for (var i = 1; i < fa.length; i++) {
    if (String(fa[i][iEstado]) !== 'ASIGNADO') continue;
    // Determinar la fecha de vencimiento:
    // 1. Si Fecha_Vencimiento está set → usarla
    // 2. Sino fallback Fecha_Asig + (Horas_TTL || 1h)
    var fVencMs;
    var fVenc = iVenc >= 0 ? fa[i][iVenc] : null;
    if (fVenc) {
      fVencMs = fVenc instanceof Date ? fVenc.getTime() : new Date(fVenc).getTime();
    } else {
      var fAsigMs = fa[i][iAsig] instanceof Date ? fa[i][iAsig].getTime() : new Date(fa[i][iAsig]).getTime();
      if (isNaN(fAsigMs)) continue;
      var ttlMs = (iHorasTTL >= 0 ? (parseInt(fa[i][iHorasTTL], 10) || 1) : 1) * UNA_HORA;
      fVencMs = fAsigMs + ttlMs;
    }
    if (isNaN(fVencMs) || ahora < fVencMs) continue;

    // [v2.5.49] GUARD CRÍTICO — ANTES de marcar EXPIRADO, verificar que la
    // venta original NO haya sido cobrada por otra vía. Si ya está en
    // EFECTIVO/VIRTUAL/MIXTO, el cajero la cobró → NO expirar (sería falso
    // positivo) y NO revertir VENTAS_CABECERA (perdería la trazabilidad
    // del cobro real). En vez, sincronizar el row a COBRADO.
    var idVentaCheck = String(fa[i][iIdVenta]);
    var fpVentaActual = '';
    var fpCheckConfiable = false;   // [Lote1-A · fix A3] solo expirar si el guard pudo leer FormaPago REAL
    try {
      if (ventasSh) {
        var vdCheck = ventasSh.getDataRange().getValues();
        var vHdrsCheck = vdCheck[0].map(function(h){ return String(h || '').trim(); });
        var iVId    = vHdrsCheck.indexOf('ID_Venta') >= 0 ? vHdrsCheck.indexOf('ID_Venta')
                    : vHdrsCheck.indexOf('ID') >= 0 ? vHdrsCheck.indexOf('ID') : 0;
        // [fix A3] SIN fallback numérico hardcodeado (antes `: 8`): si algún día se
        // inserta una columna antes de FormaPago, el guard leería la col equivocada
        // y EXPIRARÍA cobros YA pagados en silencio. Sin header → guard no confiable
        // → NO expirar (fail-safe; mejor un cobro vencido tarde que revertir uno pagado).
        var iVFpCk  = vHdrsCheck.indexOf('FormaPago') >= 0 ? vHdrsCheck.indexOf('FormaPago')
                    : vHdrsCheck.indexOf('Forma_Pago');
        if (iVFpCk >= 0) {
          fpCheckConfiable = true;   // header OK → la lectura es confiable (aunque la venta no exista)
          for (var kk = vdCheck.length - 1; kk > 0; kk--) {
            if (String(vdCheck[kk][iVId]) === idVentaCheck) {
              fpVentaActual = String(vdCheck[kk][iVFpCk] || '').toUpperCase();
              break;
            }
          }
        } else {
          Logger.log('[escalarCobros] header FormaPago NO encontrado — guard no confiable, no se expira nada');
        }
      }
    } catch(eCheck) { fpCheckConfiable = false; Logger.log('Check fp venta: ' + eCheck.message); }
    // [fix A3] Guard no confiable (header ausente o error de lectura) → saltar este
    // cobro sin expirarlo. El próximo tick lo reintenta con datos sanos.
    if (!fpCheckConfiable) continue;

    // Si la venta YA fue cobrada (no es CREDITO/POR_COBRAR), sincronizar
    // el row a COBRADO en vez de expirarlo. Sin push falso al admin.
    if (fpVentaActual && fpVentaActual !== 'CREDITO' && fpVentaActual !== 'POR_COBRAR') {
      hoja.getRange(i + 1, iEstado + 1).setValue('COBRADO');
      hoja.getRange(i + 1, iRes + 1).setValue(new Date());
      hoja.getRange(i + 1, iRazon + 1).setValue('Cobrado fuera del flujo · auto-reconciliado');
      try { _dualWriteCobroPatchME(String(fa[i][0]), { estado:'COBRADO', fecha_res:new Date(), razon:'Cobrado fuera del flujo · auto-reconciliado' }); } catch(_dw){}  // [creditos-directo]
      Logger.log('[escalarCobros] cobro ' + fa[i][0] + ' YA cobrado (' + fpVentaActual +
                 ') · marcado COBRADO en vez de EXPIRADO');
      continue; // no contar como vencido, no mandar push
    }

    // VENCIDO — marcar y restaurar
    hoja.getRange(i + 1, iEstado + 1).setValue('EXPIRADO');
    hoja.getRange(i + 1, iRes + 1).setValue(new Date());
    hoja.getRange(i + 1, iRazon + 1).setValue('Vencido sin cobrarse · cliente no llegó');
    try { _dualWriteCobroPatchME(String(fa[i][0]), { estado:'EXPIRADO', fecha_res:new Date(), razon:'Vencido sin cobrarse · cliente no llegó' }); } catch(_dw){}  // [creditos-directo]

    // [v2.5.27] Restaurar VENTAS_CABECERA → vuelve a estado CREDITO
    // Cols esperadas en VENTAS_CABECERA: 0=ID, 7=cobrado?, 8=formaPago
    // (puede variar — buscar por header)
    try {
      if (ventasSh) {
        var vdAll = ventasSh.getDataRange().getValues();
        var vHdrs = vdAll[0].map(function(h){ return String(h || '').trim(); });
        var iVId   = vHdrs.indexOf('ID') >= 0 ? vHdrs.indexOf('ID') : 0;
        var iVCob  = vHdrs.indexOf('Cobrado') >= 0 ? vHdrs.indexOf('Cobrado') : vHdrs.indexOf('cobrado');
        var iVFp   = vHdrs.indexOf('FormaPago') >= 0 ? vHdrs.indexOf('FormaPago') : vHdrs.indexOf('Forma_Pago');
        var idV = String(fa[i][iIdVenta]);
        for (var k = vdAll.length - 1; k > 0; k--) {
          if (String(vdAll[k][iVId]) === idV) {
            if (iVFp >= 0)  ventasSh.getRange(k + 1, iVFp + 1).setValue('CREDITO');
            if (iVCob >= 0) ventasSh.getRange(k + 1, iVCob + 1).setValue(false);
            try { _meMarcarDirtySync('VENTAS_CABECERA', idV); } catch(_e){}   // [fix C2-gap] re-sync el revert a CREDITO (≤15min)
            // [Lote3-C · M2-GAS] PATCH inmediato a la sombra (antes solo dirty-sync ≤15min):
            // con lecturas flipeadas, la venta expirada seguía "cobrada" en Supabase esa ventana.
            try { _dualWriteVentaPatchME(idV, { forma_pago: 'CREDITO' }); } catch(_dwR){}
            break;
          }
        }
      }
    } catch(eR) { Logger.log('Restore venta a CREDITO: ' + eR.message); }

    n++;

    // [Lote1-A] Push al admin asignador — RECOLECTADO; se envía FUERA del lock
    // (UrlFetch lento no debe retener el ScriptLock que usan las ventas).
    pushes.push({
      action: 'enviarPushUsuario',
      usuario: String(fa[i][iAdminAsig] || ''),
      titulo: '⏰ Cobro expirado · ' + String(fa[i][iCliente] || 'cliente'),
      cuerpo: 'S/ ' + parseFloat(fa[i][iMonto] || 0).toFixed(2) +
              ' asignado a ' + String(fa[i][iVendedor]) + ' venció sin cobrarse. ' +
              'Volvió a CRÉDITO. Re-asignar?',
      idNotif: 'CREDITO_COBRO_EXPIRADO',
      extra: { idVenta: String(fa[i][iIdVenta]) }
    });
  }
  Logger.log('escalarCobrosVencidos · vencidos: ' + n);
  return { ok: true, vencidos: n };
}

// ============================================================
// configurarTriggerEscalacion — ejecutar UNA vez desde editor GAS
// ============================================================
function configurarTriggerEscalacionCobros() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'escalarCobrosVencidos') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('escalarCobrosVencidos').timeBased().everyMinutes(5).create();
  Logger.log('Trigger escalarCobrosVencidos configurado · cada 5 min');
  return { ok: true };
}

// ============================================================
// Helpers internos
// ============================================================

// Devuelve { printerId } de la caja consultando ZONAS_CONFIG (estación de la caja).
function _resolverPrinterCaja(cajaId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var cajas = ss.getSheetByName('CAJAS');
    if (!cajas) return null;
    var fc = cajas.getDataRange().getValues();
    var estacionNombre = '';
    for (var i = fc.length - 1; i > 0; i--) {
      if (String(fc[i][0]) === String(cajaId)) { estacionNombre = String(fc[i][2] || ''); break; }
    }
    if (!estacionNombre) return null;

    var zonas = ss.getSheetByName('ZONAS_CONFIG');
    if (!zonas) return null;
    var fz = zonas.getDataRange().getValues();
    var hdrs = fz[0].map(function(h){ return String(h).trim(); });
    var iEst = hdrs.indexOf('Estacion_Nombre');
    var iPid = hdrs.indexOf('PrintNode_ID');
    if (iEst < 0 || iPid < 0) return null;
    for (var j = 1; j < fz.length; j++) {
      if (String(fz[j][iEst]) === estacionNombre) {
        return { printerId: fz[j][iPid] };
      }
    }
  } catch(e) { Logger.log('_resolverPrinterCaja: ' + e.message); }
  return null;
}

// Reimprime el ticket original con un sello "PAGADO · COBRO DIFERIDO" arriba.
// Reutiliza imprimirTicketInternamente con un payload reconstruido + flag esPagoDiferido.
function _reimprimirTicketConSello(idVenta, printerId, ctx) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ventas = ss.getSheetByName('VENTAS_CABECERA');
  var detSheet = ss.getSheetByName('VENTAS_DETALLE');
  if (!ventas || !detSheet) return false;

  var fv = ventas.getDataRange().getValues();
  var venta = null;
  for (var i = fv.length - 1; i > 0; i--) {
    if (String(fv[i][0]) === String(idVenta)) {
      venta = {
        correlativo: String(fv[i][9]),
        tipoDoc:     String(fv[i][7]),
        cliente: {
          doc:    String(fv[i][4] || ''),
          nombre: String(fv[i][5] || ''),
          tipo:   parseInt(fv[i][15] || 0, 10)
        },
        metodo:      ctx.metodoFinal,
        total:       parseFloat(fv[i][6]) || 0,
        vendedor:    String(fv[i][2] || '')
      };
      break;
    }
  }
  if (!venta) return false;

  // Reconstruir items
  var fd = detSheet.getDataRange().getValues();
  var items = [];
  for (var k = 1; k < fd.length; k++) {
    if (String(fd[k][0]) !== String(idVenta)) continue;
    items.push({
      sku:      String(fd[k][1] || ''),
      nombre:   String(fd[k][2] || ''),
      cantidad: parseFloat(fd[k][3]) || 0,
      precio:   parseFloat(fd[k][4]) || 0,
      subtotal: parseFloat(fd[k][5]) || 0
    });
  }

  var data = {
    auth:   { vendedor: ctx.cajero || venta.vendedor },
    header: {
      tipoDoc: venta.tipoDoc,
      total:   venta.total,
      metodo:  venta.metodo,
      cliente: venta.cliente
    },
    items: items,
    // [v40.3] Flag para que imprimirTicketInternamente agregue el sello
    esPagoDiferido: true,
    pagoDiferido: {
      cajaCobro:  ctx.cajaDest,
      cajeroCobro: ctx.cajero,
      adminAsig:  ctx.adminAsig,
      fechaCobro: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
    }
  };

  return imprimirTicketInternamente(data, venta.correlativo, printerId, null);
}

// ============================================================
// [v2.5.27] Auto-instalación del trigger horario que procesa
// vencimientos. Se llama desde asignarCobroACajero — si nadie
// asigna no se instala, pero al primer asignamiento queda activo.
// ============================================================
function _ensureTriggerEscalarCobros() {
  try {
    var existe = ScriptApp.getProjectTriggers().some(function(t) {
      return t.getHandlerFunction() === 'escalarCobrosVencidos';
    });
    if (!existe) {
      ScriptApp.newTrigger('escalarCobrosVencidos').timeBased().everyMinutes(5).create();
      Logger.log('[Trigger] escalarCobrosVencidos auto-instalado · cada 5min');
    }
  } catch(e) { Logger.log('[Trigger escalarCobrosVencidos] auto fallo: ' + e.message); }
}

// Setup público — llamable desde Apps Script editor manualmente
function setupTriggerEscalarCobros() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'escalarCobrosVencidos') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('escalarCobrosVencidos').timeBased().everyMinutes(5).create();
  return { ok: true, mensaje: 'Trigger escalarCobrosVencidos creado · cada 5min' };
}

// ============================================================
// [v2.5.28] cancelarCobroAsignado — admin retira el cobro antes
// que el cajero lo procese. Solo válido si está en ASIGNADO.
// El ticket vuelve a CREDITO inmediatamente.
// ============================================================
function cancelarCobroAsignado(data) {
  if (!data.idCobro) return generarRespuestaError('idCobro requerido');
  if (!data.adminAuth || !data.adminAuth.nombre) {
    return generarRespuestaError('adminAuth requerido');
  }
  // [Lote1-A] Validación + transición + restore de la venta bajo el lock global:
  // cancelar concurrente con confirmar ya no puede revertir una venta cobrada.
  var res = _conLockCred(function() {
    var hoja = _getHojaCobrosAsignados();
    var fa = hoja.getDataRange().getValues();
    var hdrs = fa[0].map(function(h){ return String(h || '').trim(); });
    var iIdCobro = hdrs.indexOf('ID_Cobro');
    var iIdVenta = hdrs.indexOf('ID_Venta');
    var iEstado  = hdrs.indexOf('Estado');
    var iRes     = hdrs.indexOf('Fecha_Res');
    var iRazon   = hdrs.indexOf('Razon');
    var iVendDest= hdrs.indexOf('Vendedor_Dest');
    var iCliente = hdrs.indexOf('Cliente_Nombre');
    var iMonto   = hdrs.indexOf('Monto');
    for (var i = 1; i < fa.length; i++) {
      if (String(fa[i][iIdCobro]) !== String(data.idCobro)) continue;
      if (String(fa[i][iEstado]) !== 'ASIGNADO') {
        return { error: 'Solo se puede cancelar mientras está ASIGNADO (actual: ' + fa[i][iEstado] + ')' };
      }
      var idVenta = String(fa[i][iIdVenta]);
      // Marcar como CANCELADO_ADMIN
      hoja.getRange(i + 1, iEstado + 1).setValue('CANCELADO_ADMIN');
      hoja.getRange(i + 1, iRes + 1).setValue(new Date());
      hoja.getRange(i + 1, iRazon + 1).setValue('Cancelado por admin: ' + (data.razon || 'sin razón'));
      // Restaurar VENTAS_CABECERA → CREDITO
      try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var ventasSh = ss.getSheetByName('VENTAS_CABECERA');
        if (ventasSh) {
          var vdAll = ventasSh.getDataRange().getValues();
          var vHdrs = vdAll[0].map(function(h){ return String(h || '').trim(); });
          var iVId   = vHdrs.indexOf('ID') >= 0 ? vHdrs.indexOf('ID') : 0;
          var iVCob  = vHdrs.indexOf('Cobrado') >= 0 ? vHdrs.indexOf('Cobrado') : vHdrs.indexOf('cobrado');
          var iVFp   = vHdrs.indexOf('FormaPago') >= 0 ? vHdrs.indexOf('FormaPago') : vHdrs.indexOf('Forma_Pago');
          for (var k = vdAll.length - 1; k > 0; k--) {
            if (String(vdAll[k][iVId]) === idVenta) {
              if (iVFp >= 0)  ventasSh.getRange(k + 1, iVFp + 1).setValue('CREDITO');
              if (iVCob >= 0) ventasSh.getRange(k + 1, iVCob + 1).setValue(false);
              try { _meMarcarDirtySync('VENTAS_CABECERA', idVenta); } catch(_e){}   // [fix C2-gap] re-sync el revert a CREDITO (≤15min)
              try { _dualWriteVentaPatchME(idVenta, { forma_pago: 'CREDITO' }); } catch(_dwR){}   // [Lote3-C · M2-GAS] PATCH inmediato
              break;
            }
          }
        }
      } catch(eR) { Logger.log('Cancelar - restore venta: ' + eR.message); }
      SpreadsheetApp.flush();
      return { ok: true, vendDest: String(fa[i][iVendDest] || ''),
               cliente: String(fa[i][iCliente] || 'cliente'),
               monto: parseFloat(fa[i][iMonto] || 0) };
    }
    return { error: 'Cobro ' + data.idCobro + ' no encontrado' };
  }, function() {
    return { error: 'Sistema ocupado — reintenta en unos segundos' };
  });
  if (res.error) return generarRespuestaError(res.error);

  // ── Fuera del lock (best-effort) ──
  try { _dualWriteCobroPatchME(data.idCobro, { estado:'CANCELADO_ADMIN', fecha_res:new Date(), razon:'Cancelado por admin: ' + (data.razon || 'sin razón') }); } catch(_dw){}  // [creditos-directo]
  // Push al cajero destino (avisar que ya no debe cobrar)
  try {
    var url = PropertiesService.getScriptProperties().getProperty('MOS_WEB_APP_URL');
    if (url) {
      UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          action: 'enviarPushUsuario',
          usuario: res.vendDest,
          titulo: '⊘ Cobro cancelado · ' + res.cliente,
          cuerpo: 'S/ ' + res.monto.toFixed(2) +
                  ' fue cancelado por ' + (data.adminAuth.nombre || 'admin') +
                  '. Ya no debes cobrarlo.',
          idNotif: 'CREDITO_COBRO_CANCELADO'
        }),
        muteHttpExceptions: true
      });
    }
  } catch(_){}

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success', mensaje: 'Cobro cancelado y ticket retornado a CRÉDITO'
  })).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// [v2.5.28] reasignarCobroAsignado — admin cambia caja destino sin
// cancelar y crear de nuevo. Cancela el actual + crea uno nuevo
// con misma idVenta. Incrementa contador de reasignaciones.
// ============================================================
function reasignarCobroAsignado(data) {
  if (!data.idCobro)         return generarRespuestaError('idCobro requerido');
  if (!data.cajaDestino)     return generarRespuestaError('cajaDestino requerida');
  if (!data.adminAuth || !data.adminAuth.nombre) return generarRespuestaError('adminAuth requerido');
  var hoja = _getHojaCobrosAsignados();
  var fa = hoja.getDataRange().getValues();
  var hdrs = fa[0].map(function(h){ return String(h || '').trim(); });
  var iIdCobro = hdrs.indexOf('ID_Cobro');
  var iIdVenta = hdrs.indexOf('ID_Venta');
  var iEstado  = hdrs.indexOf('Estado');
  var iMetodo  = hdrs.indexOf('Metodo_Sug');
  var iHorasTTL= hdrs.indexOf('Horas_TTL');
  var iReasig  = hdrs.indexOf('Reasignaciones');
  var idVenta = '', metodoSug = 'EFECTIVO', horasTTL = 1, reasignaciones = 0;
  var found = false;
  for (var i = 1; i < fa.length; i++) {
    if (String(fa[i][iIdCobro]) !== String(data.idCobro)) continue;
    if (String(fa[i][iEstado]) !== 'ASIGNADO') {
      return generarRespuestaError('Solo se puede reasignar mientras está ASIGNADO (actual: ' + fa[i][iEstado] + ')');
    }
    idVenta = String(fa[i][iIdVenta]);
    metodoSug = String(fa[i][iMetodo] || 'EFECTIVO');
    horasTTL = iHorasTTL >= 0 ? (parseInt(fa[i][iHorasTTL], 10) || 1) : 1;
    reasignaciones = iReasig >= 0 ? (parseInt(fa[i][iReasig], 10) || 0) : 0;
    // Marcar viejo como REASIGNADO (no CANCELADO_ADMIN ni nada raro)
    hoja.getRange(i + 1, iEstado + 1).setValue('REASIGNADO');
    var iRes = hdrs.indexOf('Fecha_Res');
    var iRazon = hdrs.indexOf('Razon');
    if (iRes >= 0)  hoja.getRange(i + 1, iRes + 1).setValue(new Date());
    if (iRazon >= 0) hoja.getRange(i + 1, iRazon + 1).setValue('Reasignado a ' + data.cajaDestino);
    try { _dualWriteCobroPatchME(data.idCobro, { estado:'REASIGNADO', fecha_res:new Date(), razon:'Reasignado a ' + data.cajaDestino }); } catch(_dw){}  // [creditos-directo]
    found = true;
    break;
  }
  if (!found) return generarRespuestaError('Cobro ' + data.idCobro + ' no encontrado');
  // Crear nuevo cobro
  data.adminAuth = data.adminAuth;
  data.metodoSugerido = data.metodoSugerido || metodoSug;
  data.horasTTL       = data.horasTTL       || horasTTL;
  data.idVenta        = idVenta;
  data.mensajeAdmin   = data.mensajeAdmin || ('Reasignación #' + (reasignaciones + 1));
  var nuevoResp = asignarCobroACajero(data);
  return nuevoResp;
}

// ============================================================
// [v2.5.28] getCobrosEnVueloAdmin — devuelve TODOS los cobros
// activos para el panel "Cobros en vuelo" del admin en MOS.
// Incluye ASIGNADO + últimos COBRADOS/EXPIRADOS/CANCELADOS (5).
// ============================================================
function getCobrosEnVueloAdmin() {
  var hoja = _getHojaCobrosAsignados();
  var fa = hoja.getDataRange().getValues();
  if (fa.length < 2) return ContentService.createTextOutput(JSON.stringify({
    status: 'success', enVuelo: [], recientes: []
  })).setMimeType(ContentService.MimeType.JSON);
  var hdrs = fa[0].map(function(h){ return String(h || '').trim(); });
  var H = {};
  hdrs.forEach(function(h, idx) { H[h] = idx; });
  var enVuelo = [];
  var recientes = [];
  for (var i = 1; i < fa.length; i++) {
    var row = fa[i];
    var estado = String(row[H.Estado] || '');
    var item = {
      idCobro:          String(row[H.ID_Cobro]),
      idVenta:          String(row[H.ID_Venta]),
      cajaDestino:      String(row[H.Caja_Destino]),
      vendedorDest:     String(row[H.Vendedor_Dest]),
      metodoSug:        String(row[H.Metodo_Sug]),
      estado:           estado,
      adminAsig:        String(row[H.Admin_Asignador]),
      fechaAsig:        row[H.Fecha_Asig] instanceof Date ? row[H.Fecha_Asig].toISOString() : String(row[H.Fecha_Asig] || ''),
      fechaRes:         row[H.Fecha_Res] instanceof Date ? row[H.Fecha_Res].toISOString() : String(row[H.Fecha_Res] || ''),
      razon:            String(row[H.Razon] || ''),
      monto:            parseFloat(row[H.Monto]) || 0,
      cliente:          String(row[H.Cliente_Nombre] || ''),
      correlativo:      String(row[H.Correlativo] || ''),
      fechaVencimiento: row[H.Fecha_Vencimiento] instanceof Date ? row[H.Fecha_Vencimiento].toISOString() : String(row[H.Fecha_Vencimiento] || ''),
      horasTTL:         parseInt(row[H.Horas_TTL], 10) || 1,
      mensajeAdmin:     String(row[H.Mensaje_Admin] || ''),
      reasignaciones:   parseInt(row[H.Reasignaciones], 10) || 0
    };
    if (estado === 'ASIGNADO') enVuelo.push(item);
    else recientes.push(item);
  }
  // Ordenar recientes por fecha_res desc, tomar últimos 10
  recientes.sort(function(a, b) {
    var ta = new Date(a.fechaRes).getTime() || 0;
    var tb = new Date(b.fechaRes).getTime() || 0;
    return tb - ta;
  });
  recientes = recientes.slice(0, 10);
  // Ordenar enVuelo por fecha_vencimiento (los que vencen primero, primero)
  enVuelo.sort(function(a, b) {
    var ta = new Date(a.fechaVencimiento).getTime() || 0;
    var tb = new Date(b.fechaVencimiento).getTime() || 0;
    return ta - tb;
  });
  return ContentService.createTextOutput(JSON.stringify({
    status: 'success', enVuelo: enVuelo, recientes: recientes
  })).setMimeType(ContentService.MimeType.JSON);
}
