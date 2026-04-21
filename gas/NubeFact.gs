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
      return {
        ok:       true,
        hash:     String(body.codigo_hash            || ''),
        enlace:   String(body.enlace_del_pdf         || ''),
        qrString: String(body.cadena_para_codigo_qr  || ''),
        aceptada: body.aceptada_por_sunat === true
      };
    }
    var errMsg = (body.errors || body.message || resp.getContentText() || '').toString().substring(0, 200);
    Logger.log('NubeFact HTTP ' + code + ': ' + errMsg);
    return { ok: false, error: 'HTTP ' + code + ': ' + errMsg };

  } catch (e) {
    Logger.log('NubeFact excepcion: ' + e.toString());
    return { ok: false, error: e.toString() };
  }
}
