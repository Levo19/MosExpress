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

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ventas = ss.getSheetByName('VENTAS_CABECERA');
  var cajas  = ss.getSheetByName('CAJAS');
  if (!ventas || !cajas) return generarRespuestaError('Hojas requeridas no encontradas');

  // ── Validar que la venta exista y esté pendiente de cobro ──
  var fv = ventas.getDataRange().getValues();
  var vRow = -1, ventaPrev = null;
  for (var i = fv.length - 1; i > 0; i--) {  // buscar desde el final
    if (String(fv[i][0]) === String(data.idVenta)) {
      vRow = i + 1;
      ventaPrev = {
        formaPagoActual: String(fv[i][8] || ''),
        cliente:         String(fv[i][5] || ''),
        total:           parseFloat(fv[i][6]) || 0,
        cajaOriginal:    String(fv[i][10] || '')
      };
      break;
    }
  }
  if (vRow < 2) return generarRespuestaError('Venta ' + data.idVenta + ' no encontrada');

  var fpActual = ventaPrev.formaPagoActual.toUpperCase();
  if (fpActual !== 'CREDITO' && fpActual !== 'POR_COBRAR') {
    return generarRespuestaError(
      'La venta no está pendiente de cobro (estado actual: ' + ventaPrev.formaPagoActual + ')'
    );
  }
  if (fpActual === 'ANULADO') {
    return generarRespuestaError('La venta está anulada');
  }

  // ── Validar caja receptora abierta ──
  var fc = cajas.getDataRange().getValues();
  var cajaAbierta = false, cajeroReceptor = '';
  for (var j = fc.length - 1; j > 0; j--) {
    if (String(fc[j][0]) === String(data.cajaReceptora)) {
      if (String(fc[j][5]) === 'ABIERTA') {
        cajaAbierta = true;
        cajeroReceptor = String(fc[j][1] || '');
      }
      break;
    }
  }
  if (!cajaAbierta) {
    return generarRespuestaError('Caja receptora ' + data.cajaReceptora + ' no está abierta');
  }

  // ── Crear MOVIMIENTOS_EXTRA con tipo según método ──
  var extras = ss.getSheetByName('MOVIMIENTOS_EXTRA');
  if (!extras) {
    extras = ss.insertSheet('MOVIMIENTOS_EXTRA');
    extras.appendRow(['ID_Extra','ID_Caja','Timestamp','Tipo','Monto','Concepto','Obs','Registrado_Por']);
  }
  var actor = _audExtraerActor(data);
  var conceptoExtra = 'Abono deuda';
  var obsExtra = 'Cobro de crédito ticket ' + data.idVenta + ' · cliente ' + (ventaPrev.cliente || '—');
  if (data.obs) obsExtra += ' · ' + data.obs;

  var monto = parseFloat(ventaPrev.total) || 0;
  var metodoUpper = String(data.metodo).toUpperCase();
  // [Lote1-A · fix A1] Sufijo aleatorio en el id: 'EX-'+ms solo COLISIONABA entre
  // cajas distintas en el mismo milisegundo → como id_extra es la clave de
  // idempotencia del dual-write, el segundo upsert SOBRESCRIBÍA al primero en
  // la sombra Supabase (un movimiento desaparecía). Mismo patrón que 'RES-'.
  var _sufijoEx = function() { return Utilities.getUuid().split('-')[0]; };
  var idExtra = 'EX-' + new Date().getTime() + '-' + _sufijoEx();

  // Caso MIXTO: crear DOS movimientos (uno EFECTIVO + uno VIRTUAL)
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
    if (efe > 0) {
      var idE = 'EX-' + new Date().getTime() + '-' + _sufijoEx() + '-E';
      extras.appendRow([idE, data.cajaReceptora, new Date(), 'INGRESO', efe, conceptoExtra, obsExtra, actor.usuario]);
      extrasCreados.push({ idExtra: idE, tipo: 'INGRESO', monto: efe });
    }
    if (vir > 0) {
      var idV = 'EX-' + new Date().getTime() + '-' + _sufijoEx() + '-V';
      extras.appendRow([idV, data.cajaReceptora, new Date(), 'INGRESO_VIRTUAL', vir, conceptoExtra, obsExtra, actor.usuario]);
      extrasCreados.push({ idExtra: idV, tipo: 'INGRESO_VIRTUAL', monto: vir });
    }
    idExtra = extrasCreados.map(function(x){ return x.idExtra; }).join('+');
  } else {
    var tipoExtra = metodoUpper === 'EFECTIVO' ? 'INGRESO' : 'INGRESO_VIRTUAL';
    extras.appendRow([idExtra, data.cajaReceptora, new Date(), tipoExtra, monto, conceptoExtra, obsExtra, actor.usuario]);
    extrasCreados.push({ idExtra: idExtra, tipo: tipoExtra, monto: monto });
  }
  SpreadsheetApp.flush();

  // ── Actualizar VENTAS_CABECERA: cambiar formaPago al método elegido ──
  ventas.getRange(vRow, 9).setValue(data.metodo);
  // Si la caja receptora es distinta, NO sobreescribir ID_Caja original (preserva trazabilidad).
  // El vínculo entre el cobro y la caja receptora vive en el log + en MOVIMIENTOS_EXTRA.

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
          ID_Extra: ex.idExtra, ID_Caja: data.cajaReceptora, Timestamp: new Date(),
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
  if (!sheet) return generarRespuestaError('VENTAS_CABECERA no encontrada');

  var data2 = sheet.getDataRange().getValues();
  for (var i = data2.length - 1; i > 0; i--) {
    if (String(data2[i][0]) === String(data.idVenta)) {
      var formaAnt = String(data2[i][8] || '');
      sheet.getRange(i + 1, 9).setValue(data.formaPagoNueva);
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
  }
  return generarRespuestaError('Venta ' + data.idVenta + ' no encontrada');
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

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('VENTAS_CABECERA');
  if (!sheet) return generarRespuestaError('VENTAS_CABECERA no encontrada');

  var fv = sheet.getDataRange().getValues();
  for (var i = fv.length - 1; i > 0; i--) {
    if (String(fv[i][0]) === String(data.idVenta)) {
      var tipoDoc  = String(fv[i][7] || '');
      var nfEstado = String(fv[i][16] || '');
      // Bloqueo: CPE emitido NO se puede editar (SUNAT)
      if (tipoDoc !== 'NOTA_DE_VENTA' && nfEstado === 'EMITIDO') {
        return generarRespuestaError(
          'CPE emitido (' + tipoDoc + ') no se puede editar. Solicite la baja del CPE primero.'
        );
      }
      var docAnt = String(fv[i][4] || '');
      var nomAnt = String(fv[i][5] || '');
      var docNew = String(data.clienteDoc || '');
      var nomNew = String(data.clienteNombre || '');
      var dirNew = String(data.clienteDireccion || '');
      // Detectar tipoDocCliente por largo
      var tipoDocCli = docNew.length === 8 ? 1 : (docNew.length === 11 ? 6 : 0);

      sheet.getRange(i + 1, 5).setValue(docNew);
      sheet.getRange(i + 1, 6).setValue(nomNew);
      sheet.getRange(i + 1, 16).setValue(tipoDocCli);

      // Actualizar también en CLIENTES_FRECUENTES si tiene doc nuevo
      if (docNew && nomNew) {
        try { verificarYAgregaCliente(docNew, nomNew, tipoDoc, dirNew); } catch(_){}
      }

      var actor = _audExtraerActor(data);
      var cambios = [];
      if (docAnt !== docNew) cambios.push({ campo:'Cliente_Doc',    antes:docAnt, despues:docNew });
      if (nomAnt !== nomNew) cambios.push({ campo:'Cliente_Nombre', antes:nomAnt, despues:nomNew });
      if (cambios.length > 0) {
        auditarLog('VENTAS_CABECERA', data.idVenta, {
          usuario: actor.usuario, rol: actor.rol,
          source: 'ME_EDITAR_CLIENTE',
          accion: 'editar_cliente',
          cambios: cambios,
          autorizadoPor: actor.autorizadoPor || null,
          motivo: data.motivo || ''
        });
      }
      return ContentService.createTextOutput(JSON.stringify({
        status:'success', mensaje:'Cliente actualizado',
        cambios: cambios.length
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  return generarRespuestaError('Venta ' + data.idVenta + ' no encontrada');
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

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ventas = ss.getSheetByName('VENTAS_CABECERA');
  var detalles = ss.getSheetByName('VENTAS_DETALLE');
  if (!ventas || !detalles) return generarRespuestaError('Hojas de ventas no encontradas');

  // Buscar venta original
  var fv = ventas.getDataRange().getValues();
  var nvRow = -1, nvData = null;
  for (var i = fv.length - 1; i > 0; i--) {
    if (String(fv[i][0]) === String(data.idVentaNV)) {
      nvRow = i + 1;
      nvData = {
        idVenta:     String(fv[i][0]),
        fecha:       fv[i][1],
        vendedor:    String(fv[i][2] || ''),
        estacion:    String(fv[i][3] || ''),
        total:       parseFloat(fv[i][6]) || 0,
        tipoDoc:     String(fv[i][7] || ''),
        formaPago:   String(fv[i][8] || ''),
        idCaja:      String(fv[i][10] || ''),
        idDispositivo: String(fv[i][11] || '')
      };
      break;
    }
  }
  if (nvRow < 2) return generarRespuestaError('Venta original ' + data.idVentaNV + ' no encontrada');
  if (nvData.tipoDoc !== 'NOTA_DE_VENTA') {
    return generarRespuestaError('Solo se pueden convertir NOTA_DE_VENTA. Esta es ' + nvData.tipoDoc);
  }
  if (nvData.formaPago === 'ANULADO_CONVERSION' || nvData.formaPago === 'ANULADO') {
    return generarRespuestaError('La venta original ya fue anulada/convertida');
  }

  // Cargar items del detalle original
  var fd = detalles.getDataRange().getValues();
  var items = [];
  for (var j = 1; j < fd.length; j++) {
    if (String(fd[j][0]) === String(data.idVentaNV)) {
      items.push({
        sku:              String(fd[j][1] || ''),
        nombre:           String(fd[j][2] || ''),
        cantidad:         parseFloat(fd[j][3]) || 1,
        precio:           parseFloat(fd[j][4]) || 0,
        subtotal:         parseFloat(fd[j][5]) || 0,
        codBarras:        String(fd[j][6] || ''),
        valor_unitario:   parseFloat(fd[j][7]) || 0,
        tipo_igv:         parseInt(fd[j][8] || 1, 10),
        unidad_de_medida: String(fd[j][9] || 'NIU')
      });
    }
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
  // Marcar la NV original como anulada por conversión
  ventas.getRange(nvRow, 9).setValue('ANULADO_CONVERSION');
  ventas.getRange(nvRow, 15).setValue('Convertido a ' + tipoDocNuevo + ' ' + resultado.correlativo);

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

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('VENTAS_CABECERA');
  if (!sheet) return generarRespuestaError('VENTAS_CABECERA no encontrada');

  var fv = sheet.getDataRange().getValues();
  for (var i = fv.length - 1; i > 0; i--) {
    if (String(fv[i][0]) === String(data.idVenta)) {
      var tipoDoc  = String(fv[i][7] || '');
      var nfEstado = String(fv[i][16] || '');
      var correlativo = String(fv[i][9] || '');

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
      sheet.getRange(i + 1, 17).setValue(nuevoEstado);

      var actor = _audExtraerActor(data);
      auditarLog('VENTAS_CABECERA', data.idVenta, {
        usuario: actor.usuario, rol: actor.rol,
        source: 'ME_BAJA_CPE',
        accion: 'baja_cpe',
        cambios: [{ campo:'NF_Estado', antes:nfEstado, despues:nuevoEstado }],
        autorizadoPor: actor.autorizadoPor || null,
        ref: { serie: serie, numero: numero, tipoDoc: tipoDoc, respuestaNF: resp },
        motivo: data.motivo
      });

      if (!resp.ok) return generarRespuestaError('NubeFact rechazó la baja: ' + (resp.error || 'sin detalle'));

      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        mensaje: 'Baja del CPE ' + (resp.aceptada ? 'aceptada por SUNAT' : 'solicitada (esperando SUNAT)'),
        nuevoEstado: nuevoEstado,
        respuestaNF: resp
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  return generarRespuestaError('Venta ' + data.idVenta + ' no encontrada');
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
  if (!cajas) return generarRespuestaError('CAJAS no encontrada');

  // Validar caja abierta
  var fc = cajas.getDataRange().getValues();
  var cajaAbierta = false;
  for (var j = fc.length - 1; j > 0; j--) {
    if (String(fc[j][0]) === String(data.cajaId)) {
      if (String(fc[j][5]) === 'ABIERTA') cajaAbierta = true;
      break;
    }
  }
  if (!cajaAbierta) {
    return generarRespuestaError('Caja ' + data.cajaId + ' no está abierta. No se permiten extras.');
  }

  var sheet = ss.getSheetByName('MOVIMIENTOS_EXTRA');
  if (!sheet) {
    sheet = ss.insertSheet('MOVIMIENTOS_EXTRA');
    sheet.appendRow(['ID_Extra','ID_Caja','Timestamp','Tipo','Monto','Concepto','Obs','Registrado_Por']);
  }
  var actor = _audExtraerActor(data);
  // [fix 20x idempotencia cruzada Fase 2] usar el idExtra provisto por el cliente (el MISMO que usó el path
  // directo a Supabase) si vino; si no, generar uno (back-compat). Así, si la escritura directa escribió en
  // me.movimientos_extra pero su respuesta se perdió y caemos a este fallback, el id_extra coincide → el batch
  // upserta (no duplica en Supabase) y la reconciliación ve la fila en Sheets → el cierre NO cuenta doble.
  var id = String(data.idExtra || '').trim() || ('EX-' + new Date().getTime());
  sheet.appendRow([
    id, data.cajaId, new Date(),
    String(data.tipo).toUpperCase(),
    monto,
    String(data.concepto),
    String(data.obs || ''),
    actor.usuario
  ]);

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
