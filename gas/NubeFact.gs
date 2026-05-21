// ============================================================
// MosExpress — NubeFact.gs
// Emisión de CPE (Boleta/Factura) vía NubeFact API.
// Script Properties requeridas: NUBEFACT_TOKEN, NUBEFACT_RUC
// ============================================================

function emitirNubeFact(data, correlativo) {
  var props  = PropertiesService.getScriptProperties();
  var token  = props.getProperty('NUBEFACT_TOKEN');
  var ruc    = props.getProperty('NUBEFACT_RUC');

  if (!token || !ruc) {
    Logger.log('NubeFact: NUBEFACT_TOKEN o NUBEFACT_RUC no configurados en Script Properties.');
    return { ok: false, error: 'NubeFact no configurado' };
  }

  var header  = data.header || {};
  var items   = data.items  || [];
  var tipoDoc = header.tipoDoc;

  // "B001-000000042" → serie=B001, numero=42
  var partes = correlativo.split('-');
  var serie  = partes[0] || '';
  var numero = parseInt(partes[partes.length - 1], 10) || 1;
  var tipoComprobante = (tipoDoc === 'FACTURA') ? 1 : 2;

  var totalGravada   = 0;
  var totalIVAP      = 0;   // base IVAP sin impuesto (arroz pilado 4%)
  var totalImpIVAP   = 0;   // monto 4% IVAP
  var totalExonerada = 0;
  var totalInafecta  = 0;

  var nfItems = items.map(function(item) {
    // Catálogo 07 SUNAT: 1=Gravado(18%) 8=IVAP(4%) 9=Exonerado 10=Gratuito 11+=Inafecto
    var tipoIgv       = parseInt(item.tipo_igv || 1, 10);
    var cantidad      = parseFloat(item.cantidad || 1);
    var valorUnitario = parseFloat(item.valor_unitario || 0);
    var subtotalVU    = Math.round(valorUnitario * cantidad * 100) / 100;
    var precioTotal   = parseFloat(item.subtotal || 0);
    var igvItem;

    if (tipoIgv === 1) {
      igvItem = Math.round((precioTotal - subtotalVU) * 100) / 100;
      totalGravada += subtotalVU;
    } else if (tipoIgv === 8) {
      // IVAP: valor_unitario ya viene sin el 4%
      igvItem = Math.round((precioTotal - subtotalVU) * 100) / 100;
      totalIVAP    += subtotalVU;
      totalImpIVAP += igvItem;
    } else if (tipoIgv === 9 || tipoIgv === 10) {
      igvItem = 0;
      totalExonerada += precioTotal;
    } else {
      // 11=Inafecto, 12=Exportación, etc.
      igvItem = 0;
      totalInafecta += precioTotal;
    }

    return {
      unidad_de_medida:         String(item.unidad_de_medida || 'NIU'),
      codigo:                   String(item.sku || ''),
      codigo_producto_sunat:    String(item.cod_sunat || ''),
      descripcion:              String(item.nombre || ''),
      cantidad:                 cantidad,
      valor_unitario:           Math.round(valorUnitario * 100) / 100,
      precio_unitario:          parseFloat(item.precio || 0),
      descuento:                '',
      subtotal:                 subtotalVU,
      tipo_de_igv:              tipoIgv,
      igv:                      igvItem,
      total:                    precioTotal,
      anticipo_regularizacion:  false,
      anticipo_documento_serie: '',
      anticipo_documento_numero:''
    };
  });

  totalGravada   = Math.round(totalGravada   * 100) / 100;
  totalIVAP      = Math.round(totalIVAP      * 100) / 100;
  totalImpIVAP   = Math.round(totalImpIVAP   * 100) / 100;
  totalExonerada = Math.round(totalExonerada * 100) / 100;
  totalInafecta  = Math.round(totalInafecta  * 100) / 100;
  var totalGeneral = parseFloat(header.total || 0);
  var totalIgv     = Math.round((totalGeneral - totalGravada - totalIVAP - totalExonerada - totalInafecta) * 100) / 100;

  var cliente  = header.cliente || {};
  var fechaHoy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy');

  var payload = {
    operacion:                         'generar_comprobante',
    tipo_de_comprobante:               tipoComprobante,
    serie:                             serie,
    numero:                            numero,
    sunat_transaction:                 1,
    cliente_tipo_de_documento:         parseInt(cliente.tipo || 0, 10),
    cliente_numero_de_documento:       String(cliente.doc   || '0'),
    cliente_denominacion:              String(cliente.nombre || 'CLIENTE ANONIMO'),
    cliente_direccion:                 String(cliente.direccion || ''),
    cliente_email:                     '',
    fecha_de_emision:                  fechaHoy,
    fecha_de_vencimiento:              '',
    moneda:                            1,
    tipo_de_cambio:                    '',
    porcentaje_de_igv:                 18,
    total_gravada:                     totalGravada   > 0 ? totalGravada   : '',
    total_ivap:                        totalIVAP      > 0 ? totalIVAP      : '',
    total_imp_ivap:                    totalImpIVAP   > 0 ? totalImpIVAP   : '',
    total_exonerada:                   totalExonerada > 0 ? totalExonerada : '',
    total_inafecta:                    totalInafecta  > 0 ? totalInafecta  : '',
    total_igv:                         totalIgv       > 0 ? totalIgv       : '',
    total_precio_de_venta:             totalGeneral,
    total_descuentos:                  '',
    total_otros_cargos:                '',
    total:                             totalGeneral,
    detraccion:                        false,
    enviar_automaticamente_a_la_sunat: true,
    enviar_automaticamente_al_cliente: false,
    formato_de_pdf:                    'TICKET',
    items:                             nfItems
  };

  var endpoint = 'https://api.nubefact.com/api/v1/' + ruc + '/' +
                 (tipoDoc === 'FACTURA' ? 'factura' : 'boleta');

  try {
    var resp = UrlFetchApp.fetch(endpoint, {
      method:             'post',
      headers:            { 'Authorization': 'Token ' + token, 'Content-Type': 'application/json' },
      payload:            JSON.stringify(payload),
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    var body = {};
    try { body = JSON.parse(resp.getContentText() || '{}'); } catch(pe) {}

    if (code === 200 || code === 201) {
      // [v2.5.47] Distinguir aceptada_por_sunat:false → RECHAZADO_SUNAT
      // El comprobante existe en NubeFact pero SUNAT lo rechazó (datos malos).
      // Hay que avisar al admin que corrija y reintente.
      if (body.aceptada_por_sunat === false) {
        return {
          ok: false,
          rechazadoPorSunat: true,
          error: 'SUNAT rechazó: ' + (body.sunat_description || body.errors || 'sin detalle'),
          hash:     String(body.codigo_hash            || ''),
          enlace:   String(body.enlace_del_pdf         || ''),
          qrString: String(body.cadena_para_codigo_qr  || ''),
          sunatDescription: String(body.sunat_description || ''),
          enlace_xml: String(body.enlace_del_xml || ''),
          enlace_cdr: String(body.enlace_del_cdr || ''),
          numero_orden_sunat: String(body.numero_de_orden_sunat || '')
        };
      }
      return {
        ok:       true,
        hash:     String(body.codigo_hash            || ''),
        enlace:   String(body.enlace_del_pdf         || ''),
        qrString: String(body.cadena_para_codigo_qr  || ''),
        aceptada: body.aceptada_por_sunat === true,
        // [v2.5.47] Metadata extendida — útil para auditoría SUNAT
        enlace_xml: String(body.enlace_del_xml || ''),
        enlace_cdr: String(body.enlace_del_cdr || ''),
        numero_orden_sunat: String(body.numero_de_orden_sunat || ''),
        sunatDescription: String(body.sunat_description || '')
      };
    }
    // [v2.5.47] Detectar duplicado — NubeFact responde HTTP 400 con
    // texto "ya fue informado" o "duplicado" → consultar el existente
    // y devolver datos como éxito (idempotencia).
    var errMsg = (body.errors || body.message || resp.getContentText() || '').toString();
    var esDuplicado = /ya\s+fue\s+informado|duplicad|comprobante\s+ya\s+existe|already\s+exists/i.test(errMsg);
    if (esDuplicado) {
      Logger.log('NubeFact duplicado detectado para ' + correlativo + ' — consultando existente');
      var consulta = consultarCPENubeFact(serie, numero, tipoDoc);
      if (consulta.ok) {
        return {
          ok: true,
          hash:     consulta.hash,
          enlace:   consulta.enlace,
          qrString: consulta.qrString,
          aceptada: consulta.aceptada,
          enlace_xml: consulta.enlace_xml,
          enlace_cdr: consulta.enlace_cdr,
          numero_orden_sunat: consulta.numero_orden_sunat,
          dedupNubeFact: true // bandera para que el caller sepa que era idempotencia
        };
      }
    }
    Logger.log('NubeFact HTTP ' + code + ': ' + errMsg.substring(0, 300));
    return { ok: false, error: 'HTTP ' + code + ': ' + errMsg.substring(0, 250) };

  } catch (e) {
    Logger.log('NubeFact excepcion: ' + e.toString());
    return { ok: false, error: e.toString() };
  }
}

// ============================================================
// [v2.5.47] CONSULTAR COMPROBANTE — operación 3 de NubeFact.
// Pregunta el estado actual de un CPE ya enviado SIN re-emitirlo.
// Crítico para reconciliación: si una venta quedó en NF_Estado='ERROR'
// porque NubeFact respondió lento o con timeout, pero en realidad el
// comprobante ya está OK en SUNAT, podemos detectarlo y actualizar.
// ============================================================
function consultarCPENubeFact(serie, numero, tipoDoc) {
  var props  = PropertiesService.getScriptProperties();
  var token  = props.getProperty('NUBEFACT_TOKEN');
  var ruc    = props.getProperty('NUBEFACT_RUC');
  if (!token || !ruc) return { ok: false, error: 'NubeFact no configurado' };
  if (!serie || !numero) return { ok: false, error: 'serie y numero requeridos' };

  var tipoComprobante = (String(tipoDoc).toUpperCase() === 'FACTURA') ? 1 : 2;
  var payload = {
    operacion:           'consultar_comprobante',
    tipo_de_comprobante: tipoComprobante,
    serie:               String(serie),
    numero:              parseInt(numero, 10)
  };
  // El endpoint es el mismo que para emitir — NubeFact distingue por `operacion`
  var endpoint = 'https://api.nubefact.com/api/v1/' + ruc + '/' +
                 (tipoComprobante === 1 ? 'factura' : 'boleta');
  try {
    var resp = UrlFetchApp.fetch(endpoint, {
      method:             'post',
      headers:            { 'Authorization': 'Token ' + token, 'Content-Type': 'application/json' },
      payload:            JSON.stringify(payload),
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    var body = {};
    try { body = JSON.parse(resp.getContentText() || '{}'); } catch(pe) {}
    if (code === 200 || code === 201) {
      return {
        ok: true,
        aceptada: body.aceptada_por_sunat === true,
        hash:     String(body.codigo_hash || ''),
        enlace:   String(body.enlace_del_pdf || ''),
        qrString: String(body.cadena_para_codigo_qr || ''),
        enlace_xml: String(body.enlace_del_xml || ''),
        enlace_cdr: String(body.enlace_del_cdr || ''),
        numero_orden_sunat: String(body.numero_de_orden_sunat || ''),
        sunatDescription: String(body.sunat_description || ''),
        body: body
      };
    }
    var errMsg = (body.errors || body.message || resp.getContentText() || '').toString();
    // NubeFact responde HTTP 400 con "no existe" → no se emitió todavía
    var noExiste = /no\s+(existe|encontrado|registrado)/i.test(errMsg);
    if (noExiste) {
      return { ok: false, noExiste: true, error: errMsg.substring(0, 200) };
    }
    return { ok: false, error: 'HTTP ' + code + ': ' + errMsg.substring(0, 200) };
  } catch(e) {
    return { ok: false, error: 'NETWORK: ' + e.toString() };
  }
}

// ============================================================
// BAJA DEL CPE — Comunicación de baja a SUNAT vía NubeFact.
// SUNAT permite anular boletas/facturas dentro de las 72h con
// comunicación de baja (resumen diario). NubeFact lo gestiona.
// ============================================================
function bajaCPENubeFact(serie, numero, motivo, tipoDoc) {
  var props  = PropertiesService.getScriptProperties();
  var token  = props.getProperty('NUBEFACT_TOKEN');
  var ruc    = props.getProperty('NUBEFACT_RUC');

  if (!token || !ruc) {
    return { ok: false, error: 'NubeFact no configurado (token/ruc)' };
  }
  if (!serie || !numero || !motivo) {
    return { ok: false, error: 'Faltan parámetros: serie, numero o motivo' };
  }

  var tipoComprobante = (String(tipoDoc).toUpperCase() === 'FACTURA') ? 1 : 2;

  var payload = {
    operacion:           'generar_anulacion',
    tipo_de_comprobante: tipoComprobante,
    serie:               String(serie),
    numero:              parseInt(numero, 10),
    motivo:              String(motivo).substring(0, 250)  // SUNAT limita el motivo
  };

  var endpoint = 'https://api.nubefact.com/api/v1/' + ruc + '/anular-comprobante';

  try {
    var resp = UrlFetchApp.fetch(endpoint, {
      method:             'post',
      headers:            { 'Authorization': 'Token ' + token, 'Content-Type': 'application/json' },
      payload:            JSON.stringify(payload),
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    var body = {};
    try { body = JSON.parse(resp.getContentText() || '{}'); } catch(pe) {}

    if (code === 200 || code === 201) {
      return {
        ok:        true,
        aceptada:  body.aceptada_por_sunat === true,
        ticketSunat: String(body.numero_ticket_sunat || ''),
        enlace:    String(body.enlace_del_pdf || ''),
        hash:      String(body.codigo_hash || ''),
        body:      body
      };
    }
    var errMsg = (body.errors || body.message || resp.getContentText() || '').toString().substring(0, 250);
    Logger.log('NubeFact BAJA HTTP ' + code + ': ' + errMsg);
    return { ok: false, error: 'HTTP ' + code + ': ' + errMsg };

  } catch (e) {
    Logger.log('NubeFact BAJA excepcion: ' + e.toString());
    return { ok: false, error: e.toString() };
  }
}

// ============================================================
// [v2.5.47] RECONCILIACIÓN DIARIA — corre desde cron 8pm.
// Busca todas las ventas (últimos 35 días) con Tipo_Doc IN
// ('BOLETA','FACTURA') y NF_Estado != 'EMITIDO'. Para cada una llama
// consultarCPENubeFact:
//   - SI aceptada_por_sunat true → actualiza NF_Estado='EMITIDO'
//     + hash/enlace/QR/XML/CDR/num_orden
//   - SI aceptada false → marca 'RECHAZADO_SUNAT' con descripción
//   - SI no existe → intenta re-emitir con mismo correlativo
//   - SI red falla → deja como estaba (próximo cron reintenta)
// Devuelve resumen con cuántas reconcilió.
// ============================================================
function reconciliarCPEsPendientes(diasAtras) {
  diasAtras = parseInt(diasAtras, 10) || 35;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('VENTAS_CABECERA');
  if (!sheet) return { ok: false, error: 'VENTAS_CABECERA no encontrada' };
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { ok: true, total: 0, emitidos: 0, rechazados: 0, sin_cambio: 0 };

  var headers = data[0];
  var idxFecha = headers.indexOf('Fecha');
  var idxTipo  = headers.indexOf('Tipo_Doc');
  var idxCorr  = headers.indexOf('Correlativo');
  var idxNFE   = headers.indexOf('NF_Estado');
  var idxNFH   = headers.indexOf('NF_Hash');
  var idxNFL   = headers.indexOf('NF_Enlace');
  if (idxTipo < 0 || idxCorr < 0 || idxNFE < 0) {
    return { ok: false, error: 'Columnas requeridas no encontradas' };
  }

  var hace = new Date();
  hace.setDate(hace.getDate() - diasAtras);

  var stats = { ok: true, total: 0, emitidos: 0, rechazados: 0, re_emitidos: 0, sin_cambio: 0, errores: 0, detalles: [] };
  for (var i = 1; i < data.length; i++) {
    var fila = i + 1;
    var fecha = data[i][idxFecha];
    var tipo  = String(data[i][idxTipo] || '');
    var corr  = String(data[i][idxCorr] || '');
    var nfe   = String(data[i][idxNFE] || '');

    if (tipo !== 'BOLETA' && tipo !== 'FACTURA') continue;
    if (nfe === 'EMITIDO') continue;
    if (!corr || corr.indexOf('undefined-') === 0) continue;
    if (fecha instanceof Date && fecha < hace) continue;

    var partes = corr.split('-');
    if (partes.length < 2) continue;
    var serie  = partes[0];
    var numero = parseInt(partes[partes.length - 1], 10);
    if (!serie || !numero) continue;
    stats.total++;

    var cons = consultarCPENubeFact(serie, numero, tipo);
    if (cons.ok && cons.aceptada) {
      // Aceptada por SUNAT → actualizar fila
      sheet.getRange(fila, idxNFE + 1).setValue('EMITIDO');
      if (idxNFH >= 0 && cons.hash)   sheet.getRange(fila, idxNFH + 1).setValue(cons.hash);
      if (idxNFL >= 0 && cons.enlace) sheet.getRange(fila, idxNFL + 1).setValue(cons.enlace);
      stats.emitidos++;
      stats.detalles.push({ corr: corr, accion: 'reconciliado_emitido' });
    } else if (cons.ok && !cons.aceptada) {
      sheet.getRange(fila, idxNFE + 1).setValue('RECHAZADO_SUNAT');
      stats.rechazados++;
      stats.detalles.push({ corr: corr, accion: 'rechazado_sunat', motivo: cons.sunatDescription });
    } else if (cons.noExiste) {
      // NubeFact no tiene el comprobante → reintentar emisión
      // Reusa la data original de VENTAS_DETALLE (skip por ahora — es complejo)
      // El admin puede re-emitir manualmente desde MOS panel
      stats.detalles.push({ corr: corr, accion: 'no_existe_en_nubefact_necesita_re_emision_manual' });
      stats.errores++;
    } else {
      stats.sin_cambio++;
    }
    Utilities.sleep(200); // no saturar API
  }
  SpreadsheetApp.flush();
  Logger.log('[reconciliarCPEs] ' + JSON.stringify(stats));
  return stats;
}

// Trigger horario 8pm — instala una vez con _ensureTriggerReconciliarCPE()
function _cronReconciliarCPEs() {
  reconciliarCPEsPendientes(35);
}
function _ensureTriggerReconciliarCPE() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === '_cronReconciliarCPEs') return; // ya existe
  }
  ScriptApp.newTrigger('_cronReconciliarCPEs').timeBased().atHour(20).everyDays(1).create();
  Logger.log('Trigger reconciliarCPEs instalado · 8pm diario');
}
