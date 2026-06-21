// ============================================================
// MosExpress — EditarVenta.gs
// Endpoints de edición posterior de tickets:
//   - cobrarCreditoConExtra: cobra un crédito creando MOVIMIENTOS_EXTRA
//   - editarFormaPagoVenta: cambia formaPago sin afectar caja (corrección)
//   - editarClienteVenta: cambia cliente (solo NOTA_DE_VENTA)
//   - convertirNVaCPE: emite CPE retroactivo desde una nota de venta
//   - bajaCPEVenta: comunica baja a SUNAT vía NubeFact
//   - registrarExtraCajaConLog: extra con validaciones + log
// Todos los endpoints requieren autenticación (data.auth.vendedor)
// y opcionalmente data.adminAuth si el cajero usó PIN admin.
// ============================================================

// ── Whitelist de tipos válidos para MOVIMIENTOS_EXTRA ──
var _TIPOS_EXTRA_VALIDOS = ['INGRESO','EGRESO','INGRESO_VIRTUAL','EGRESO_VIRTUAL'];

// ── Conceptos típicos (no obligatorio, solo guía) ──
var _CONCEPTOS_EXTRA = [
  'Retiro jefe','Pago proveedor','Abono deuda','Retorno efectivo',
  'Otro ingreso','Gasto operativo'
];

// ============================================================
// COBRAR CRÉDITO CON CREACIÓN AUTOMÁTICA DE EXTRA
// ============================================================
// Convierte un ticket POR_COBRAR/CREDITO en cobrado, creando un
// INGRESO en MOVIMIENTOS_EXTRA en la caja receptora elegida, y
// vinculando ambos registros en el historial.
//
// payload esperado:
// {
//   tipoEvento: 'COBRAR_CREDITO_CON_EXTRA',
//   idVenta:    'V-...',
//   cajaReceptora: 'CAJA-...',
//   metodo:     'EFECTIVO' | 'VIRTUAL' | 'MIXTO (VIR:X/EFE:Y)',
//   montoEfectivo: opcional, parte efectivo si MIXTO,
//   montoVirtual:  opcional, parte virtual si MIXTO,
//   obs:        observación opcional,
//   auth:       { vendedor, rol, deviceId },
//   adminAuth:  { nombre, rol, via } // si requirió PIN
// }
function cobrarCreditoConExtra(data) {
  // [Lote1-A · fix C1] TODO el flujo (validar venta pendiente → crear extras →
  // cambiar FormaPago) bajo el lock global de cobros. Sin esto, dos requests
  // simultáneos (doble tap / replay) leían ambos FormaPago='CREDITO', ambos
  // pasaban la validación y ambos creaban INGRESO → doble ingreso en caja.
  return _conLockCred(function() { return _cobrarCreditoConExtraImpl(data); },
    function() { return generarRespuestaError('Sistema ocupado procesando otro cobro — reintenta en unos segundos'); });
}
function _cobrarCreditoConExtraImpl(data) {
  if (!data.idVenta)       return generarRespuestaError('idVenta requerido');
  if (!data.cajaReceptora) return generarRespuestaError('cajaReceptora requerida');
  if (!data.metodo)        return generarRespuestaError('metodo requerido');

  var idV = String(data.idVenta);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ventas = ss.getSheetByName('VENTAS_CABECERA');
  var cajas  = ss.getSheetByName('CAJAS');

  // ── Validar que la venta exista y esté pendiente de cobro ── [delete-safe] Supabase primero.
  var vRow = -1, ventaPrev = null;
  if (typeof _meLecturaCierreDirecta === 'function' && _meLecturaCierreDirecta()) {
    try {
      var rVc = _sb('GET', 'me.ventas', { select: 'forma_pago,cliente_nombre,total,id_caja', filters: { id_venta: 'eq.' + idV }, limit: 1, maxRetry: 1 });
      if (rVc && rVc.ok && Array.isArray(rVc.data) && rVc.data.length) {
        var v0 = rVc.data[0];
        ventaPrev = { formaPagoActual: String(v0.forma_pago || ''), cliente: String(v0.cliente_nombre || ''), total: parseFloat(v0.total) || 0, cajaOriginal: String(v0.id_caja || '') };
      }
    } catch (eVc) { Logger.log('[cobrarCredito] leer venta Supabase: ' + eVc.message); }
  }
  if (!ventaPrev && ventas) {
    var fv = ventas.getDataRange().getValues();
    for (var i = fv.length - 1; i > 0; i--) {
      if (String(fv[i][0]) === idV) {
        vRow = i + 1;
        ventaPrev = { formaPagoActual: String(fv[i][8] || ''), cliente: String(fv[i][5] || ''), total: parseFloat(fv[i][6]) || 0, cajaOriginal: String(fv[i][10] || '') };
        break;
      }
    }
  }
  if (!ventaPrev) return generarRespuestaError('Venta ' + idV + ' no encontrada');

  var fpActual = ventaPrev.formaPagoActual.toUpperCase();
  if (fpActual === 'ANULADO') {
    return generarRespuestaError('La venta está anulada');
  }
  if (fpActual !== 'CREDITO' && fpActual !== 'POR_COBRAR') {
    return generarRespuestaError(
      'La venta no está pendiente de cobro (estado actual: ' + ventaPrev.formaPagoActual + ')'
    );
  }

  // ── Validar caja receptora abierta ── [delete-safe] Supabase primero.
  var cajaAbierta = false, cajeroReceptor = '', resCaja = false;
  var abSBc = (typeof _meCajaAbiertaEnZona === 'function') ? _meCajaAbiertaEnZona(String(data.cajaReceptora), '') : null;
  if (abSBc !== null) {
    resCaja = true; cajaAbierta = (abSBc === true);
    if (cajaAbierta) {
      try { var rCr = _sb('GET', 'me.cajas', { select: 'vendedor', filters: { id_caja: 'eq.' + String(data.cajaReceptora) }, limit: 1, maxRetry: 1 }); if (rCr && rCr.ok && rCr.data && rCr.data.length) cajeroReceptor = String(rCr.data[0].vendedor || ''); } catch(_cr){}
    }
  }
  if (!resCaja) {
    if (!cajas) return generarRespuestaError('Hojas requeridas no encontradas');
    var fc = cajas.getDataRange().getValues();
    for (var j = fc.length - 1; j > 0; j--) {
      if (String(fc[j][0]) === String(data.cajaReceptora)) {
        if (String(fc[j][5]) === 'ABIERTA') { cajaAbierta = true; cajeroReceptor = String(fc[j][1] || ''); }
        break;
      }
    }
  }
  if (!cajaAbierta) {
    return generarRespuestaError('Caja receptora ' + data.cajaReceptora + ' no está abierta');
  }

  var actor = _audExtraerActor(data);
  var conceptoExtra = 'Abono deuda';
  var obsExtra = 'Cobro de crédito ticket ' + data.idVenta + ' · cliente ' + (ventaPrev.cliente || '—');
  if (data.obs) obsExtra += ' · ' + data.obs;

  var monto = parseFloat(ventaPrev.total) || 0;
  var metodoUpper = String(data.metodo).toUpperCase();
  // [Lote1-A · fix A1] Sufijo aleatorio en el id (idempotencia del dual-write).
  var _sufijoEx = function() { return Utilities.getUuid().split('-')[0]; };
  var idExtra = 'EX-' + new Date().getTime() + '-' + _sufijoEx();
  var tsExtraCC = new Date();

  // 1) Construir los movimientos (DATOS) — sin tocar el Sheet aún.
  var extrasCreados = [];
  if (metodoUpper.indexOf('MIXTO') === 0) {
    var efe = parseFloat(data.montoEfectivo) || 0;
    var vir = parseFloat(data.montoVirtual)  || 0;
    if (Math.abs((efe + vir) - monto) > 0.01) {
      return generarRespuestaError(
        'Suma de montoEfectivo + montoVirtual (' + (efe + vir).toFixed(2) +
        ') no coincide con el total del ticket (' + monto.toFixed(2) + ')'
      );
    }
    if (efe > 0) extrasCreados.push({ idExtra: 'EX-' + new Date().getTime() + '-' + _sufijoEx() + '-E', tipo: 'INGRESO', monto: efe });
    if (vir > 0) extrasCreados.push({ idExtra: 'EX-' + new Date().getTime() + '-' + _sufijoEx() + '-V', tipo: 'INGRESO_VIRTUAL', monto: vir });
    idExtra = extrasCreados.map(function(x){ return x.idExtra; }).join('+');
  } else {
    var tipoExtra = metodoUpper === 'EFECTIVO' ? 'INGRESO' : 'INGRESO_VIRTUAL';
    extrasCreados.push({ idExtra: idExtra, tipo: tipoExtra, monto: monto });
  }

  // 2) [delete-safe] SHEET (best-effort espejo) — la persistencia REAL va por _dualWriteMovExtraME (abajo).
  var extras = ss.getSheetByName('MOVIMIENTOS_EXTRA');
  if (extras) {
    try {
      extrasCreados.forEach(function(ex){
        extras.appendRow([ex.idExtra, data.cajaReceptora, tsExtraCC, ex.tipo, ex.monto, conceptoExtra, obsExtra, actor.usuario]);
      });
      SpreadsheetApp.flush();
    } catch (eEX) { Logger.log('[cobrarCredito] Sheet extras write: ' + eEX.message); }
  }

  // ── Actualizar VENTAS_CABECERA: cambiar formaPago al método elegido (best-effort Sheet) ──
  if (ventas) {
    try {
      if (vRow < 2) { var fvU = ventas.getDataRange().getValues(); for (var kk = fvU.length - 1; kk > 0; kk--) { if (String(fvU[kk][0]) === idV) { vRow = kk + 1; break; } } }
      if (vRow >= 2) ventas.getRange(vRow, 9).setValue(data.metodo);
    } catch (eVS) { Logger.log('[cobrarCredito] Sheet venta write: ' + eVS.message); }
  }
  // Si la caja receptora es distinta, NO sobreescribir ID_Caja original (preserva trazabilidad).

  // [Lote1-A · fix A2] PATCH inmediato del cambio CREDITO→pagado a la sombra.
  // Antes solo llegaba por el dirty-sync (≤15min): con lecturas flipeadas a
  // Supabase, el crédito cobrado seguía "pendiente" esa ventana → re-asignable.
  // Best-effort (igual que anulación); el batch lo corrige si falla.
  try { _dualWriteVentaPatchME(data.idVenta, { forma_pago: String(data.metodo) }); } catch(_dwV){}
  // Espejar también los extras creados a la sombra en tiempo real (best-effort;
  // _dualWriteMovExtraME es upsert por id_extra = idempotente, keys por cabecera).
  try {
    extrasCreados.forEach(function(ex) {
      if (typeof _dualWriteMovExtraME === 'function') {
        _dualWriteMovExtraME({
          ID_Extra: ex.idExtra, ID_Caja: data.cajaReceptora, Timestamp: tsExtraCC,
          Tipo: ex.tipo, Monto: ex.monto, Concepto: conceptoExtra, Obs: obsExtra,
          Registrado_Por: actor.usuario
        });
      }
    });
  } catch(_dwM){}

  // ── Log en VENTAS_CABECERA: cobro de crédito con vínculo ──
  auditarLog('VENTAS_CABECERA', data.idVenta, {
    usuario: actor.usuario, rol: actor.rol,
    source: 'ME_COBRO_CREDITO',
    accion: 'cobrar_credito',
    cambios: [{ campo: 'FormaPago', antes: ventaPrev.formaPagoActual, despues: data.metodo }],
    autorizadoPor: actor.autorizadoPor || null,
    ref: {
      idCajaReceptora: data.cajaReceptora,
      cajeroReceptor:  cajeroReceptor,
      idCajaOriginal:  ventaPrev.cajaOriginal,
      extras:          extrasCreados,
      monto:           monto
    },
    motivo: data.obs || ''
  });

  // ── Log en MOVIMIENTOS_EXTRA: cada uno apunta a la venta ──
  extrasCreados.forEach(function(ex) {
    auditarLog('MOVIMIENTOS_EXTRA', ex.idExtra, {
      usuario: actor.usuario, rol: actor.rol,
      source: 'ME_COBRO_CREDITO',
      accion: 'crear',
      autorizadoPor: actor.autorizadoPor || null,
      ref: { idVenta: data.idVenta, monto: ex.monto, tipo: ex.tipo },
      motivo: 'Vinculado al cobro del crédito ' + data.idVenta
    });
  });

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    mensaje: 'Crédito cobrado · S/ ' + monto.toFixed(2) + ' registrado en ' + data.cajaReceptora,
    idVenta: data.idVenta,
    formaPagoNueva: data.metodo,
    extras: extrasCreados
  })).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// EDITAR FORMA DE PAGO (sin afectar caja)
// ============================================================
// Solo cambia el campo FormaPago. NO crea movimientos extra.
// Útil para corregir errores de tipeo (ej: cajero registró EFECTIVO
// pero el cliente realmente pagó con Yape).
//
// payload:
// { tipoEvento:'EDITAR_FORMA_PAGO_VENTA', idVenta, formaPagoNueva,
//   motivo (obligatorio), auth, adminAuth }
function editarFormaPagoVenta(data) {
  if (!data.idVenta)        return generarRespuestaError('idVenta requerido');
  if (!data.formaPagoNueva) return generarRespuestaError('formaPagoNueva requerida');
  if (!data.motivo)         return generarRespuestaError('motivo es obligatorio para auditoría');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('VENTAS_CABECERA');

  var formaAnt = '';
  var encontrada = false;

  // SHEET (best-effort espejo): localizar fila, leer forma anterior, escribir nueva.
  if (sheet) {
    var data2 = sheet.getDataRange().getValues();
    for (var i = data2.length - 1; i > 0; i--) {
      if (String(data2[i][0]) === String(data.idVenta)) {
        formaAnt = String(data2[i][8] || '');
        encontrada = true;
        try { sheet.getRange(i + 1, 9).setValue(data.formaPagoNueva); } catch (eFS) { Logger.log('[editarFormaPago] Sheet write: ' + eFS.message); }
        break;
      }
    }
  }

  // [delete-safe] PATCH durable a me.ventas (fuente de verdad cuando el Sheet ya no existe).
  var sbOK = false;
  try {
    var rPF = _dualWriteVentaPatchME(String(data.idVenta), { forma_pago: data.formaPagoNueva });
    sbOK = !!(rPF && rPF.ok);
    if (sbOK) encontrada = true;  // PATCH por id_venta llegó a Supabase (no-op si la venta no existe allá, pero la lectura es directa)
  } catch (ePF) { Logger.log('[editarFormaPago] patch Supabase: ' + ePF.message); }

  if (!encontrada && !sbOK) return generarRespuestaError('Venta ' + data.idVenta + ' no encontrada');

  var actor = _audExtraerActor(data);
  auditarLog('VENTAS_CABECERA', data.idVenta, {
    usuario: actor.usuario, rol: actor.rol,
    source: 'ME_EDITAR_FORMA_PAGO',
    accion: 'editar_forma_pago',
    cambios: [{ campo:'FormaPago', antes: formaAnt, despues: data.formaPagoNueva }],
    autorizadoPor: actor.autorizadoPor || null,
    motivo: data.motivo
  });
  return ContentService.createTextOutput(JSON.stringify({
    status:'success', mensaje:'Forma de pago actualizada'
  })).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// EDITAR CLIENTE DE UNA VENTA
// ============================================================
// Solo permitido en NOTA_DE_VENTA. Si el ticket tiene CPE emitido
// (NF_Estado='EMITIDO'), rechaza la edición.
//
// payload:
// { tipoEvento:'EDITAR_CLIENTE_VENTA', idVenta,
//   clienteDoc, clienteNombre, clienteDireccion, motivo,
//   auth, adminAuth }
function editarClienteVenta(data) {
  if (!data.idVenta) return generarRespuestaError('idVenta requerido');

  var idV = String(data.idVenta);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('VENTAS_CABECERA');

  // [delete-safe] Leer el estado actual de la venta (tipo_doc, nf_estado, doc/nombre previos):
  // FUENTE PRIMARIA Supabase; FALLBACK Sheet.
  var tipoDoc = '', nfEstado = '', docAnt = '', nomAnt = '', filaIdx = -1, encontrada = false;
  var sbV = null;
  try {
    if (typeof _meLecturaCierreDirecta === 'function' && _meLecturaCierreDirecta()) {
      var rV = _sb('GET', 'me.ventas', { select: 'tipo_doc,nf_estado,cliente_doc,cliente_nombre', filters: { id_venta: 'eq.' + idV }, limit: 1, maxRetry: 1 });
      if (rV && rV.ok && Array.isArray(rV.data) && rV.data.length) sbV = rV.data[0];
    }
  } catch (eRV) { Logger.log('[editarCliente] leer Supabase: ' + eRV.message); }

  if (sbV) {
    encontrada = true;
    tipoDoc = String(sbV.tipo_doc || ''); nfEstado = String(sbV.nf_estado || '');
    docAnt = String(sbV.cliente_doc || ''); nomAnt = String(sbV.cliente_nombre || '');
  } else if (sheet) {
    var fv = sheet.getDataRange().getValues();
    for (var i = fv.length - 1; i > 0; i--) {
      if (String(fv[i][0]) === idV) {
        encontrada = true; filaIdx = i;
        tipoDoc = String(fv[i][7] || ''); nfEstado = String(fv[i][16] || '');
        docAnt = String(fv[i][4] || ''); nomAnt = String(fv[i][5] || '');
        break;
      }
    }
  }
  if (!encontrada) return generarRespuestaError('Venta ' + idV + ' no encontrada');

  // Bloqueo: CPE emitido NO se puede editar (SUNAT)
  if (tipoDoc !== 'NOTA_DE_VENTA' && nfEstado === 'EMITIDO') {
    return generarRespuestaError('CPE emitido (' + tipoDoc + ') no se puede editar. Solicite la baja del CPE primero.');
  }

  var docNew = String(data.clienteDoc || '');
  var nomNew = String(data.clienteNombre || '');
  var dirNew = String(data.clienteDireccion || '');
  var tipoDocCli = docNew.length === 8 ? 1 : (docNew.length === 11 ? 6 : 0);

  // PATCH durable a Supabase.
  try {
    _dualWriteVentaPatchME(idV, { cliente_doc: docNew, cliente_nombre: nomNew, tipo_doc_cliente: tipoDocCli });
  } catch (ePc) { Logger.log('[editarCliente] patch Supabase: ' + ePc.message); }

  // SHEET (best-effort espejo): si tenemos la fila (o la buscamos), escribir.
  if (sheet) {
    try {
      if (filaIdx < 0) {
        var fv2 = sheet.getDataRange().getValues();
        for (var k = fv2.length - 1; k > 0; k--) { if (String(fv2[k][0]) === idV) { filaIdx = k; break; } }
      }
      if (filaIdx > 0) {
        sheet.getRange(filaIdx + 1, 5).setValue(docNew);
        sheet.getRange(filaIdx + 1, 6).setValue(nomNew);
        sheet.getRange(filaIdx + 1, 16).setValue(tipoDocCli);
      }
    } catch (eWS) { Logger.log('[editarCliente] Sheet write: ' + eWS.message); }
  }

  // Actualizar también en clientes frecuentes (Supabase + Sheet, vía verificarYAgregaCliente).
  if (docNew && nomNew) {
    try { verificarYAgregaCliente(docNew, nomNew, tipoDoc, dirNew); } catch(_){}
  }

  var actor = _audExtraerActor(data);
  var cambios = [];
  if (docAnt !== docNew) cambios.push({ campo:'Cliente_Doc',    antes:docAnt, despues:docNew });
  if (nomAnt !== nomNew) cambios.push({ campo:'Cliente_Nombre', antes:nomAnt, despues:nomNew });
  if (cambios.length > 0) {
    auditarLog('VENTAS_CABECERA', idV, {
      usuario: actor.usuario, rol: actor.rol,
      source: 'ME_EDITAR_CLIENTE', accion: 'editar_cliente',
      cambios: cambios, autorizadoPor: actor.autorizadoPor || null, motivo: data.motivo || ''
    });
  }
  return ContentService.createTextOutput(JSON.stringify({
    status:'success', mensaje:'Cliente actualizado', cambios: cambios.length
  })).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// CONVERTIR NV → BOLETA o FACTURA (emisión retroactiva)
// ============================================================
// Permite emitir un CPE oficial cuando el cliente decidió después
// que sí quería comprobante. Crea una nueva venta con CPE y
// referencia la venta original. La NV original queda con
// FormaPago='ANULADO_CONVERSION' y obs apuntando a la nueva.
//
// payload:
// { tipoEvento:'CONVERTIR_NV_A_CPE', idVentaNV,
//   tipoDocNuevo:'BOLETA'|'FACTURA',
//   clienteDoc, clienteNombre, clienteDireccion,
//   serieNueva (ej: 'B001' o 'F001'),
//   auth, adminAuth }
function convertirNVaCPE(data) {
  if (!data.idVentaNV)    return generarRespuestaError('idVentaNV requerido');
  if (!data.tipoDocNuevo) return generarRespuestaError('tipoDocNuevo requerido (BOLETA o FACTURA)');
  if (!data.serieNueva)   return generarRespuestaError('serieNueva requerida');
  if (!data.clienteDoc)   return generarRespuestaError('clienteDoc requerido');
  if (!data.clienteNombre) return generarRespuestaError('clienteNombre requerido');

  var tipoDocNuevo = String(data.tipoDocNuevo).toUpperCase();
  if (tipoDocNuevo !== 'BOLETA' && tipoDocNuevo !== 'FACTURA') {
    return generarRespuestaError('tipoDocNuevo debe ser BOLETA o FACTURA');
  }
  // Validar formato del documento según tipo
  var docLimpio = String(data.clienteDoc).trim();
  if (tipoDocNuevo === 'BOLETA' && !/^\d{8}$/.test(docLimpio)) {
    return generarRespuestaError('BOLETA requiere DNI de 8 dígitos');
  }
  if (tipoDocNuevo === 'FACTURA' && !/^\d{11}$/.test(docLimpio)) {
    return generarRespuestaError('FACTURA requiere RUC de 11 dígitos');
  }

  var idNV = String(data.idVentaNV);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ventas = ss.getSheetByName('VENTAS_CABECERA');
  var detalles = ss.getSheetByName('VENTAS_DETALLE');

  // [delete-safe] Leer NV (cabecera + detalle) — FUENTE PRIMARIA Supabase, FALLBACK Sheet.
  var nvRow = -1, nvData = null, items = [];
  var usadoSupabaseNV = false;
  if (typeof _meLecturaCierreDirecta === 'function' && _meLecturaCierreDirecta()) {
    try {
      var rC = _sb('GET', 'me.ventas', { select: 'id_venta,fecha,vendedor,estacion,total,tipo_doc,forma_pago,id_caja,dispositivo_id', filters: { id_venta: 'eq.' + idNV }, limit: 1, maxRetry: 1 });
      if (rC && rC.ok && Array.isArray(rC.data) && rC.data.length) {
        var c0 = rC.data[0];
        nvData = { idVenta: String(c0.id_venta), fecha: c0.fecha, vendedor: String(c0.vendedor || ''), estacion: String(c0.estacion || ''),
          total: parseFloat(c0.total) || 0, tipoDoc: String(c0.tipo_doc || ''), formaPago: String(c0.forma_pago || ''),
          idCaja: String(c0.id_caja || ''), idDispositivo: String(c0.dispositivo_id || '') };
        var rD = _sb('GET', 'me.ventas_detalle', { select: 'sku,nombre,cantidad,precio,subtotal,cod_barras,valor_unitario,tipo_igv,unidad_medida', filters: { id_venta: 'eq.' + idNV }, order: 'linea.asc', limit: 500, maxRetry: 1 });
        if (rD && rD.ok && Array.isArray(rD.data)) {
          rD.data.forEach(function(d){
            items.push({ sku:String(d.sku||''), nombre:String(d.nombre||''), cantidad:parseFloat(d.cantidad)||1,
              precio:parseFloat(d.precio)||0, subtotal:parseFloat(d.subtotal)||0, codBarras:String(d.cod_barras||''),
              valor_unitario:parseFloat(d.valor_unitario)||0, tipo_igv:parseInt(d.tipo_igv||1,10), unidad_de_medida:String(d.unidad_medida||'NIU') });
          });
        }
        usadoSupabaseNV = true;
      }
    } catch (eNVsb) { Logger.log('[convertirNVaCPE] leer Supabase: ' + eNVsb.message); }
  }

  if (!usadoSupabaseNV) {
    if (!ventas || !detalles) return generarRespuestaError('Hojas de ventas no encontradas');
    var fv = ventas.getDataRange().getValues();
    for (var i = fv.length - 1; i > 0; i--) {
      if (String(fv[i][0]) === idNV) {
        nvRow = i + 1;
        nvData = { idVenta: String(fv[i][0]), fecha: fv[i][1], vendedor: String(fv[i][2] || ''), estacion: String(fv[i][3] || ''),
          total: parseFloat(fv[i][6]) || 0, tipoDoc: String(fv[i][7] || ''), formaPago: String(fv[i][8] || ''),
          idCaja: String(fv[i][10] || ''), idDispositivo: String(fv[i][11] || '') };
        break;
      }
    }
    if (nvRow < 2) return generarRespuestaError('Venta original ' + idNV + ' no encontrada');
    var fd = detalles.getDataRange().getValues();
    for (var j = 1; j < fd.length; j++) {
      if (String(fd[j][0]) === idNV) {
        items.push({ sku:String(fd[j][1]||''), nombre:String(fd[j][2]||''), cantidad:parseFloat(fd[j][3])||1,
          precio:parseFloat(fd[j][4])||0, subtotal:parseFloat(fd[j][5])||0, codBarras:String(fd[j][6]||''),
          valor_unitario:parseFloat(fd[j][7])||0, tipo_igv:parseInt(fd[j][8]||1,10), unidad_de_medida:String(fd[j][9]||'NIU') });
      }
    }
  }

  if (!nvData) return generarRespuestaError('Venta original ' + idNV + ' no encontrada');
  if (nvData.tipoDoc !== 'NOTA_DE_VENTA') {
    return generarRespuestaError('Solo se pueden convertir NOTA_DE_VENTA. Esta es ' + nvData.tipoDoc);
  }
  if (nvData.formaPago === 'ANULADO_CONVERSION' || nvData.formaPago === 'ANULADO') {
    return generarRespuestaError('La venta original ya fue anulada/convertida');
  }
  if (!items.length) return generarRespuestaError('La venta original no tiene items');

  // Construir payload para procesarVenta: emite CPE retroactivo
  var actor = _audExtraerActor(data);
  var newPayload = {
    auth: { vendedor: actor.usuario, esCajero: nvData.formaPago !== 'POR_COBRAR', deviceId: nvData.idDispositivo, estacion: nvData.estacion },
    pos_config: { cajaId: nvData.idCaja, serieActual: data.serieNueva, print_request: false },
    header: {
      tipoDoc: tipoDocNuevo,
      total: nvData.total,
      metodo: nvData.formaPago,
      cliente: {
        doc: docLimpio,
        nombre: String(data.clienteNombre),
        tipo: tipoDocNuevo === 'BOLETA' ? 1 : 6,
        direccion: String(data.clienteDireccion || '')
      },
      obs: 'Conversión retroactiva de ' + data.idVentaNV
    },
    items: items,
    data_sync: { last_sync: 'CONVERT-' + data.idVentaNV }
  };

  var resultado = procesarVenta(newPayload);
  // Marcar la NV original como anulada por conversión.
  // [delete-safe] PATCH durable a Supabase (forma_pago + obs); Sheet best-effort.
  try {
    _dualWriteVentaPatchME(idNV, { forma_pago: 'ANULADO_CONVERSION', obs: 'Convertido a ' + tipoDocNuevo + ' ' + resultado.correlativo });
  } catch (eAnu) { Logger.log('[convertirNVaCPE] patch anular Supabase: ' + eAnu.message); }
  if (ventas) {
    try {
      if (nvRow < 2) { var fvA = ventas.getDataRange().getValues(); for (var kk = fvA.length - 1; kk > 0; kk--) { if (String(fvA[kk][0]) === idNV) { nvRow = kk + 1; break; } } }
      if (nvRow >= 2) {
        ventas.getRange(nvRow, 9).setValue('ANULADO_CONVERSION');
        ventas.getRange(nvRow, 15).setValue('Convertido a ' + tipoDocNuevo + ' ' + resultado.correlativo);
      }
    } catch (eAnS) { Logger.log('[convertirNVaCPE] Sheet anular write: ' + eAnS.message); }
  }

  // ── STOCK (money-safety) — NO se reintegra aquí. Ver análisis abajo. ──
  // El físico de una conversión NV→CPE es EL MISMO (solo cambia el documento), así que
  // el stock debe decrementarse EXACTAMENTE una vez en total a través del par NV+CPE.
  //   · Conversión ANTES del cierre de la caja NV: el cierre filtra la NV (FormaPago
  //     empieza con 'ANULADO' → ahora excluida por /^ANULADO/ en generarGuiaSalidaVentas)
  //     y descuenta SOLO el CPE → neto −1. Correcto. (Lo arregla el Fix de generarGuiaSalidaVentas.)
  //   · Conversión DESPUÉS del cierre: la NV YA descontó −1 en su cierre. El CPE hereda
  //     la MISMA caja (procesarVenta NO reasigna caja cuando esCajero=true, que es el caso
  //     de toda NV pagada) y esa caja ya tiene su guía SALIDA_VENTAS → el CPE NO genera un
  //     segundo descuento (dedup por caja). Neto ya = −1. Correcto SIN tocar nada.
  // ⚠️ Por eso NO llamamos _reponerStockVentaAnulada aquí: reintegrar (+1) sin un segundo
  //    descuento del CPE dejaría el neto en 0 = stock SOBRECONTADO (la mercadería sí salió).
  //    Esto difiere del caso de anulación pura (sin CPE de reemplazo), donde sí se repone.

  // Logs en ambas filas
  auditarLog('VENTAS_CABECERA', data.idVentaNV, {
    usuario: actor.usuario, rol: actor.rol,
    source: 'ME_CONVERTIR_NV_CPE',
    accion: 'anular_por_conversion',
    cambios: [{ campo:'FormaPago', antes: nvData.formaPago, despues:'ANULADO_CONVERSION' }],
    autorizadoPor: actor.autorizadoPor || null,
    ref: { idVentaCPE: resultado.idVenta, correlativoCPE: resultado.correlativo, tipoDoc: tipoDocNuevo }
  });
  auditarLog('VENTAS_CABECERA', resultado.idVenta, {
    usuario: actor.usuario, rol: actor.rol,
    source: 'ME_CONVERTIR_NV_CPE',
    accion: 'crear_desde_nv',
    autorizadoPor: actor.autorizadoPor || null,
    ref: { idVentaNV: data.idVentaNV, correlativoNV: '' },
    motivo: 'Emisión retroactiva desde nota de venta ' + data.idVentaNV
  });

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    mensaje: 'NV convertida a ' + tipoDocNuevo + ' · ' + resultado.correlativo,
    idVentaNuevo: resultado.idVenta,
    correlativoNuevo: resultado.correlativo,
    nfEstado: resultado.nfEstado,
    nfHash: resultado.nfHash,
    nfEnlace: resultado.nfEnlace
  })).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// BAJA DEL CPE (comunicación a SUNAT vía NubeFact)
// ============================================================
// payload:
// { tipoEvento:'BAJA_CPE', idVenta, motivo, auth, adminAuth }
function bajaCPEVenta(data) {
  if (!data.idVenta) return generarRespuestaError('idVenta requerido');
  if (!data.motivo)  return generarRespuestaError('motivo es obligatorio para SUNAT');

  var idV = String(data.idVenta);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('VENTAS_CABECERA');

  // [delete-safe] Leer tipo_doc/nf_estado/correlativo — Supabase primero, Sheet fallback.
  var tipoDoc = '', nfEstado = '', correlativo = '', filaIdx = -1, encontrada = false, sbV = null;
  try {
    if (typeof _meLecturaCierreDirecta === 'function' && _meLecturaCierreDirecta()) {
      var rV = _sb('GET', 'me.ventas', { select: 'tipo_doc,nf_estado,correlativo', filters: { id_venta: 'eq.' + idV }, limit: 1, maxRetry: 1 });
      if (rV && rV.ok && Array.isArray(rV.data) && rV.data.length) sbV = rV.data[0];
    }
  } catch (eRV) { Logger.log('[bajaCPE] leer Supabase: ' + eRV.message); }

  if (sbV) {
    encontrada = true;
    tipoDoc = String(sbV.tipo_doc || ''); nfEstado = String(sbV.nf_estado || ''); correlativo = String(sbV.correlativo || '');
  } else if (sheet) {
    var fv = sheet.getDataRange().getValues();
    for (var i = fv.length - 1; i > 0; i--) {
      if (String(fv[i][0]) === idV) {
        encontrada = true; filaIdx = i;
        tipoDoc = String(fv[i][7] || ''); nfEstado = String(fv[i][16] || ''); correlativo = String(fv[i][9] || '');
        break;
      }
    }
  }
  if (!encontrada) return generarRespuestaError('Venta ' + idV + ' no encontrada');

  if (tipoDoc !== 'BOLETA' && tipoDoc !== 'FACTURA') {
    return generarRespuestaError('Solo se da de baja BOLETA o FACTURA. Esta venta es ' + tipoDoc);
  }
  if (nfEstado !== 'EMITIDO') {
    return generarRespuestaError('CPE no está EMITIDO en SUNAT (estado: ' + nfEstado + ')');
  }
  // partir correlativo "B001-000042" → serie B001, numero 42
  var partes = correlativo.split('-');
  if (partes.length < 2) return generarRespuestaError('Correlativo inválido: ' + correlativo);
  var serie = partes[0];
  var numero = parseInt(partes[partes.length - 1], 10);

  var resp = bajaCPENubeFact(serie, numero, data.motivo, tipoDoc);
  var nuevoEstado = resp.ok ? (resp.aceptada ? 'BAJA_ACEPTADA' : 'BAJA_SOLICITADA') : 'BAJA_ERROR';

  // PATCH durable a Supabase (nf_estado).
  try { _dualWriteVentaPatchME(idV, { nf_estado: nuevoEstado }); } catch (ePc) { Logger.log('[bajaCPE] patch Supabase: ' + ePc.message); }
  // SHEET espejo.
  if (sheet) {
    try {
      if (filaIdx < 0) { var fv2 = sheet.getDataRange().getValues(); for (var k = fv2.length - 1; k > 0; k--) { if (String(fv2[k][0]) === idV) { filaIdx = k; break; } } }
      if (filaIdx > 0) sheet.getRange(filaIdx + 1, 17).setValue(nuevoEstado);
    } catch (eWS) { Logger.log('[bajaCPE] Sheet write: ' + eWS.message); }
  }

  var actor = _audExtraerActor(data);
  auditarLog('VENTAS_CABECERA', idV, {
    usuario: actor.usuario, rol: actor.rol,
    source: 'ME_BAJA_CPE', accion: 'baja_cpe',
    cambios: [{ campo:'NF_Estado', antes:nfEstado, despues:nuevoEstado }],
    autorizadoPor: actor.autorizadoPor || null,
    ref: { serie: serie, numero: numero, tipoDoc: tipoDoc, respuestaNF: resp },
    motivo: data.motivo
  });

  if (!resp.ok) return generarRespuestaError('NubeFact rechazó la baja: ' + (resp.error || 'sin detalle'));

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    mensaje: 'Baja del CPE ' + (resp.aceptada ? 'aceptada por SUNAT' : 'solicitada (esperando SUNAT)'),
    nuevoEstado: nuevoEstado, respuestaNF: resp
  })).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// REGISTRAR EXTRA con validaciones reforzadas + log
// (reemplaza al registrarExtraCaja simple para uso desde frontend)
// ============================================================
function registrarExtraCajaConLog(data) {
  if (!data.cajaId)             return generarRespuestaError('cajaId requerido');
  if (!data.tipo)               return generarRespuestaError('tipo requerido');
  if (_TIPOS_EXTRA_VALIDOS.indexOf(String(data.tipo).toUpperCase()) < 0) {
    return generarRespuestaError('tipo inválido. Válidos: ' + _TIPOS_EXTRA_VALIDOS.join(', '));
  }
  var monto = parseFloat(data.monto);
  if (!monto || monto <= 0) return generarRespuestaError('monto debe ser > 0');
  if (!data.concepto) return generarRespuestaError('concepto requerido');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cajas = ss.getSheetByName('CAJAS');

  // Validar caja abierta — [delete-safe] FUENTE PRIMARIA Supabase, FALLBACK Sheet.
  var cajaAbierta = false, resuelto = false;
  var abSB = (typeof _meCajaAbiertaEnZona === 'function') ? _meCajaAbiertaEnZona(String(data.cajaId), '') : null;  // true|false|null
  if (abSB !== null) { resuelto = true; cajaAbierta = (abSB === true); }
  if (!resuelto) {
    if (!cajas) return generarRespuestaError('CAJAS no encontrada');
    var fc = cajas.getDataRange().getValues();
    for (var j = fc.length - 1; j > 0; j--) {
      if (String(fc[j][0]) === String(data.cajaId)) {
        if (String(fc[j][5]) === 'ABIERTA') cajaAbierta = true;
        break;
      }
    }
  }
  if (!cajaAbierta) {
    return generarRespuestaError('Caja ' + data.cajaId + ' no está abierta. No se permiten extras.');
  }

  var actor = _audExtraerActor(data);
  // [fix 20x idempotencia cruzada Fase 2] usar el idExtra provisto por el cliente (el MISMO que usó el path
  // directo a Supabase) si vino; si no, generar uno (back-compat). Así, si la escritura directa escribió en
  // me.movimientos_extra pero su respuesta se perdió y caemos a este fallback, el id_extra coincide → el batch
  // upserta (no duplica en Supabase) y la reconciliación ve la fila en Sheets → el cierre NO cuenta doble.
  var id = String(data.idExtra || '').trim() || ('EX-' + new Date().getTime());
  var tsExtra = new Date();
  var tipoExtra = String(data.tipo).toUpperCase();

  // [delete-safe] Persistencia REAL en Supabase (upsert por id_extra, idempotente) — el cierre lee de aquí
  // (me.cierre_datos_caja agrega ingresos/egresos por id_caja). Mapeo via _ME_SPECS.movimientos_extra.
  try {
    _dualWriteMovExtraME({
      ID_Extra: id, ID_Caja: String(data.cajaId || ''), Timestamp: tsExtra, Tipo: tipoExtra,
      Monto: monto, Concepto: String(data.concepto), Obs: String(data.obs || ''),
      Registrado_Por: actor.usuario
    });
  } catch (eMX) { Logger.log('[registrarExtraCaja] dualWrite Supabase: ' + (eMX && eMX.message)); }

  // SHEET (best-effort espejo).
  var sheet = ss.getSheetByName('MOVIMIENTOS_EXTRA');
  if (sheet) {
    try {
      sheet.appendRow([
        id, data.cajaId, tsExtra, tipoExtra, monto,
        String(data.concepto), String(data.obs || ''), actor.usuario
      ]);
    } catch (eMS) { Logger.log('[registrarExtraCaja] Sheet write: ' + eMS.message); }
  }

  auditarLog('MOVIMIENTOS_EXTRA', id, {
    usuario: actor.usuario, rol: actor.rol,
    source: 'ME_EXTRA_CAJA',
    accion: 'crear',
    autorizadoPor: actor.autorizadoPor || null,
    ref: { idCaja: data.cajaId, monto: monto, tipo: String(data.tipo).toUpperCase() },
    motivo: String(data.obs || '')
  });

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success', idExtra: id
  })).setMimeType(ContentService.MimeType.JSON);
}
