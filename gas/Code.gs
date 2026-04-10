// ESTRUCTURA DE BASE DE DATOS ESPERADA EN GOOGLE SHEETS
// Pestañas requeridas: PRODUCTO_BASE, PRESENTACIONES, EQUIVALENCIAS, PROMOCIONES,
//                      ZONAS_CONFIG, CLIENTES_FRECUENTES, VENTAS_CABECERA, VENTAS_DETALLE,
//                      CAJAS, DISPOSITIVOS
// Opcional: LOG_IMPRESIONES

// ============================================================
// CLAVE DE PRINTNODE: guardarla en Proyecto > Propiedades del script
// Nombre de la propiedad: PRINTNODE_API_KEY
// Valor: tu_clave_de_printnode
// ============================================================

function doGet(e) {
  var accion = e.parameter.accion;

  if (accion === 'descargar') {
    return descargarCatalogo();
  }

  if (accion === 'verificar_dispositivo') {
    return verificarDispositivo(e.parameter.id);
  }

  if (accion === 'ventas_hoy_zona') {
    return ventasHoyZona(e.parameter.prefijos);
  }

  if (accion === 'detalle_venta') {
    return detalleVenta(e.parameter.id_venta);
  }

  if (accion === 'stock_zonas') {
    return getStockZonas();
  }

  if (accion === 'cajero_activo') {
    return cajeroActivo(e.parameter.zona);
  }

  if (accion === 'listar_guias') {
    return listarGuias(e.parameter.zona);
  }

  if (accion === 'detalle_guia') {
    return detalleGuia(e.parameter.id_guia);
  }

  if (accion === 'traslados_entrantes') {
    return trasladosEntrantes(e.parameter.zona, e.parameter.desde);
  }

  if (accion === 'consultar_cliente') {
    return consultarCliente(e.parameter.doc);
  }

  if (accion === 'extras_caja') {
    return getExtrasCaja(e.parameter.cajaId);
  }

  return ContentService.createTextOutput(JSON.stringify({ error: "Acción no válida" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var rawData = e.postData.contents;
    var data = JSON.parse(rawData);

    if (data.tipoEvento === 'APERTURA_CAJA') {
      return procesarAperturaCaja(data);
    }

    if (data.tipoEvento === 'CIERRE_CAJA') {
      return procesarCierreCaja(data);
    }

    if (data.tipoEvento === 'COBRAR_VENTA') {
      return cobrarVentaExistente(data);
    }

    if (data.tipoEvento === 'ANULACION_MASIVA') {
      return anulacionMasiva(data);
    }

    if (data.tipoEvento === 'CREDITAR_VENTA') {
      return creditarVenta(data);
    }

    if (data.tipoEvento === 'EXTRA_CAJA') {
      return registrarExtraCaja(data);
    }

    if (data.tipoEvento === 'ANULACION') {
      return anularVentaIndividual(data);
    }

    // NUEVO: proxy de impresión — la clave de PrintNode vive aquí, no en el browser
    if (data.accion === 'imprimir') {
      return procesarImpresion(data);
    }

    if (data.tipoEvento === 'REGISTRAR_GUIA') {
      return registrarGuia(data);
    }

    if (data.tipoEvento === 'REGISTRAR_AUDITORIA') {
      return registrarAuditoria(data);
    }

    // Default: registrar venta
    var response = procesarVenta(data);
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      idVenta: response.idVenta,
      correlativo: response.correlativo,
      printDispatched: response.printDispatched,
      nfEstado: response.nfEstado || 'NA',
      nfHash:   response.nfHash   || '',
      nfEnlace: response.nfEnlace || '',
      mensaje: "Venta procesada con éxito"
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return generarRespuestaError(error.toString());
  }
}

// ------ FUNCIONES DE NEGOCIO ------ //

function procesarAperturaCaja(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetCajas = ss.getSheetByName("CAJAS");
  if (!sheetCajas) return generarRespuestaError("Pestaña CAJAS no encontrada.");

  // Verificar cajero único por zona — solo un cajero activo a la vez por zona
  if (data.zona) {
    var filasCheck = sheetCajas.getDataRange().getValues();
    for (var c = 1; c < filasCheck.length; c++) {
      if (String(filasCheck[c][5]) === 'ABIERTA' && String(filasCheck[c][8] || '') === String(data.zona)) {
        return generarRespuestaError(
          "Ya hay un turno activo en " + data.zona + " (cajero: " + filasCheck[c][1] + "). Cierra ese turno primero."
        );
      }
    }
  }

  var fechaActual = new Date();
  var idCaja = "CAJA-" + fechaActual.getTime();

  // Columnas: ID_Caja | Vendedor | Estacion | Fecha_Apertura | Monto_Inicial | Estado | Monto_Final | Fecha_Cierre | Zona_ID
  sheetCajas.appendRow([idCaja, data.vendedor, data.estacion, fechaActual, data.montoInicial, "ABIERTA", "", "", data.zona || '']);

  return ContentService.createTextOutput(JSON.stringify({
    status: "success", idCaja: idCaja, mensaje: "Caja aperturada exitosamente"
  })).setMimeType(ContentService.MimeType.JSON);
}

function procesarCierreCaja(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetCajas = ss.getSheetByName("CAJAS");
  if (!sheetCajas) return generarRespuestaError("Pestaña CAJAS no encontrada.");

  var filas = sheetCajas.getDataRange().getValues();
  var cajaEncontrada = false;
  var cajaVendedor = '';
  var cajaZona = '';

  for (var i = 1; i < filas.length; i++) {
    if (String(filas[i][0]) === String(data.cajaId)) {
      sheetCajas.getRange(i + 1, 6).setValue("CERRADA");
      sheetCajas.getRange(i + 1, 7).setValue(data.montoFinal);
      sheetCajas.getRange(i + 1, 8).setValue(new Date()); // Fecha_Cierre col 8
      cajaVendedor = String(filas[i][1]);
      cajaZona = String(filas[i][8] || ''); // col 9 = Zona_ID
      cajaEncontrada = true;
      break;
    }
  }

  if (!cajaEncontrada) return generarRespuestaError("Caja con ID " + data.cajaId + " no encontrada.");

  // Auto-generar SALIDA_VENTAS en segundo plano (no bloquea la respuesta si falla)
  if (cajaZona) {
    try { generarGuiaSalidaVentas(ss, data.cajaId, cajaVendedor, cajaZona); }
    catch(e) { Logger.log("Error guia ventas: " + e.toString()); }
  }

  return ContentService.createTextOutput(JSON.stringify({
    status: "success", mensaje: "Caja cerrada correctamente"
  })).setMimeType(ContentService.MimeType.JSON);
}

// Agrega registros a GUIAS y descuenta STOCK_ZONAS con todas las ventas del turno
function generarGuiaSalidaVentas(ss, cajaId, vendedor, zona) {
  var sheetVC    = ss.getSheetByName("VENTAS_CABECERA");
  var sheetVD    = ss.getSheetByName("VENTAS_DETALLE");
  var sheetGC    = ss.getSheetByName("GUIAS_CABECERA");
  var sheetGD    = ss.getSheetByName("GUIAS_DETALLE");
  var sheetStock = ss.getSheetByName("STOCK_ZONAS");
  if (!sheetVC || !sheetVD || !sheetGC || !sheetGD || !sheetStock) return;

  // 1. Recopilar IDs de ventas no anuladas de esta caja
  var ventas = sheetVC.getDataRange().getValues();
  var idsVenta = [];
  for (var i = 1; i < ventas.length; i++) {
    if (String(ventas[i][10]) === String(cajaId) && String(ventas[i][8]) !== 'ANULADO') {
      idsVenta.push(String(ventas[i][0]));
    }
  }
  if (!idsVenta.length) return;

  // 2. Sumar cantidades por Cod_Barras
  // Col 6 (índice 6) = Cod_Barras real; col 1 (índice 1) = SKU (fallback para ventas antiguas sin col 7)
  var detalle = sheetVD.getDataRange().getValues();
  var totales = {};
  for (var j = 1; j < detalle.length; j++) {
    if (idsVenta.indexOf(String(detalle[j][0])) === -1) continue;
    var cod = String(detalle[j][6] || detalle[j][1]).trim();
    if (!cod) continue;
    totales[cod] = (totales[cod] || 0) + (parseFloat(detalle[j][3]) || 0);
  }

  var cods = Object.keys(totales);
  if (!cods.length) return;

  // 3. Registrar guía y descontar stock
  var idGuia = "G-VENTAS-" + new Date().getTime();
  // GUIAS_CABECERA: ID_Guia | Fecha | Vendedor | Zona_ID | Tipo | Observacion | Zona_Destino | Estado
  sheetGC.appendRow([idGuia, new Date(), vendedor, zona, 'SALIDA_VENTAS', 'Auto cierre de caja · ' + cajaId, '', 'CONFIRMADO']);
  cods.forEach(function(cod) {
    sheetGD.appendRow([idGuia, String(cod), totales[cod]]);
    actualizarStockFila(sheetStock, cod, zona, -totales[cod]);
  });
}

function procesarVenta(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetCabecera = ss.getSheetByName("VENTAS_CABECERA");
  var sheetDetalle  = ss.getSheetByName("VENTAS_DETALLE");

  if (!sheetCabecera) throw new Error("Pestaña VENTAS_CABECERA no encontrada.");
  if (!sheetDetalle)  throw new Error("Pestaña VENTAS_DETALLE no encontrada.");

  var auth   = data.auth        || {};
  var pos    = data.pos_config  || {};
  var header = data.header      || {};
  var items  = data.items       || [];

  // ── Seguridad de rol: si dice ser cajero pero no tiene caja abierta → POR_COBRAR ──
  // Previene que dispositivos mal configurados registren cobros directos sin caja
  if (auth.esCajero && !pos.cajaId) {
    header.metodo = 'POR_COBRAR';
  }

  var fechaActual = new Date();
  var idVenta = "V-" + fechaActual.getTime();

  // ── Idempotencia: si ya existe una venta con este ref_local, devolver la existente ──
  // Previene duplicados cuando el browser reintenta una venta que GAS ya procesó
  var refLocal = (data.data_sync && data.data_sync.last_sync) ? String(data.data_sync.last_sync) : '';
  if (refLocal) {
    var totalFilas  = sheetCabecera.getLastRow();
    var buscarDesde = Math.max(2, totalFilas - 199); // revisar últimas 200 filas
    var filasBuscar = sheetCabecera.getRange(buscarDesde, 1, totalFilas - buscarDesde + 1, 16).getValues();
    for (var fi = filasBuscar.length - 1; fi >= 0; fi--) {
      if (String(filasBuscar[fi][13]) === refLocal) { // col 14 = Ref_Local (índice 13)
        // Venta ya registrada — retornar idempotente sin insertar nada nuevo
        return { idVenta: String(filasBuscar[fi][0]), correlativo: String(filasBuscar[fi][9]), printDispatched: false };
      }
    }
  }

  // ── Correlativo rápido (hoja CORRELATIVOS, O(1) vs O(n)) ─────────────────
  var correlativoNumero = obtenerSiguienteCorrelativoRapido(ss, pos.serieActual);
  var correlativoFinal  = pos.serieActual + "-" + ("000000" + correlativoNumero).slice(-6);

  // ── VENTAS_CABECERA (19 columnas) ────────────────────────────────────────
  // ID_Venta | Fecha | Vendedor | Estacion | Cliente_Doc | Cliente_Nombre | Total
  // | Tipo_Doc | FormaPago | Correlativo | ID_Caja | ID_Dispositivo | Estado_Envio
  // | Ref_Local | Obs | Tipo_Doc_Cliente | NF_Estado | NF_Hash | NF_Enlace
  var tipoDocCliente = parseInt((header.cliente && header.cliente.tipo) || 0, 10); // 0=sin doc, 1=DNI, 6=RUC
  sheetCabecera.appendRow([
    idVenta, fechaActual, auth.vendedor, auth.estacion,
    (header.cliente && header.cliente.doc)    || '',
    (header.cliente && header.cliente.nombre) || '',
    header.total,
    header.tipoDoc,
    header.metodo || 'EFECTIVO',
    correlativoFinal, pos.cajaId, auth.deviceId, "COMPLETADO",
    refLocal,
    String(header.obs || ''),
    tipoDocCliente,                             // col 16: Tipo_Doc_Cliente (NubeFact)
    '', '', ''                                  // cols 17-19: NF_Estado, NF_Hash, NF_Enlace (se llenan después)
  ]);

  // ── VENTAS_DETALLE — escritura por lote (1 setValues vs N appendRow) ──────
  // ID_Venta | SKU | Nombre | Cantidad | Precio | Subtotal | Cod_Barras
  // | Valor_Unitario | Tipo_IGV | Unidad_Medida
  if (items.length > 0) {
    var detalleRows = items.map(function(item) {
      var valorUnitario = parseFloat(item.valor_unitario) || Math.round(parseFloat(item.precio || 0) / 1.18 * 100) / 100;
      return [
        idVenta,
        item.sku,
        item.nombre,
        item.cantidad,
        item.precio,
        item.subtotal,
        String(item.codBarras || ''),
        Math.round(valorUnitario * 100) / 100,  // col 8: Valor_Unitario sin IGV (NubeFact)
        parseInt(item.tipo_igv || 1, 10),        // col 9: Tipo_IGV 1=Gravado (NubeFact)
        String(item.unidad_de_medida || 'NIU')   // col 10: Unidad_Medida (NubeFact)
      ];
    });
    var lastRow = sheetDetalle.getLastRow();
    sheetDetalle
      .getRange(lastRow + 1, 1, detalleRows.length, detalleRows[0].length)
      .setValues(detalleRows);
  }

  // ── Registrar cliente frecuente ───────────────────────────────────────────
  if (header.tipoDoc !== 'NOTA_DE_VENTA') {
    verificarYAgregaCliente(
      (header.cliente && header.cliente.doc)      || '',
      (header.cliente && header.cliente.nombre)   || '',
      header.tipoDoc,
      (header.cliente && header.cliente.direccion) || ''
    );
  }

  // ── Emitir CPE en NubeFact (solo BOLETA y FACTURA) ───────────────────────
  var nfEstado = 'NA';
  var nfHash   = '';
  var nfEnlace = '';
  var nfResult = null;

  if (header.tipoDoc === 'BOLETA' || header.tipoDoc === 'FACTURA') {
    nfResult = emitirNubeFact(data, correlativoFinal);
    nfEstado  = nfResult.ok ? 'EMITIDO' : 'ERROR';
    nfHash    = nfResult.hash   || '';
    nfEnlace  = nfResult.enlace || '';
    // Actualizar cols 17-19 en la fila recién escrita
    var nfRow = sheetCabecera.getLastRow();
    sheetCabecera.getRange(nfRow, 17, 1, 3).setValues([[nfEstado, nfHash, nfEnlace]]);
    if (!nfResult.ok) {
      Logger.log('NubeFact error venta ' + idVenta + ': ' + (nfResult.error || ''));
    }
  }

  // ── Imprimir en el mismo round-trip — SOLO si el browser lo pide explícitamente ─
  // pos.print_request = true lo envía solo el browser nuevo (index.html v39+).
  // Sin este flag, el browser antiguo (caché vieja) manejaría la impresión por su cuenta
  // y tendríamos ticket doble.
  var printDispatched = false;
  if (pos.print_request === true && pos.printerId) {
    printDispatched = imprimirTicketInternamente(data, correlativoFinal, pos.printerId, nfResult);
  }

  return { idVenta: idVenta, correlativo: correlativoFinal, printDispatched: printDispatched, nfEstado: nfEstado, nfHash: nfHash, nfEnlace: nfEnlace };
}

// NUEVO: proxy de impresión para PrintNode — la clave nunca sale del servidor
function procesarImpresion(data) {
  var printNodeKey = PropertiesService.getScriptProperties().getProperty('PRINTNODE_API_KEY');
  if (!printNodeKey) return generarRespuestaError("PRINTNODE_API_KEY no configurada en Propiedades del script.");
  if (!data.printerId || !data.content) return generarRespuestaError("Faltan datos de impresión (printerId o content).");

  var printerId = parseInt(data.printerId, 10);
  if (isNaN(printerId) || printerId <= 0) {
    return generarRespuestaError("printerId inválido: '" + data.printerId + "'. Verifica el campo PrintNode_ID en la hoja ZONAS_CONFIG.");
  }

  var options = {
    method: 'post',
    headers: {
      'Authorization': 'Basic ' + Utilities.base64Encode(printNodeKey + ':')
    },
    contentType: 'application/json',
    payload: JSON.stringify({
      printerId: printerId,
      title: data.title || 'MOSexpress',
      contentType: 'raw_base64',
      content: data.content,
      source: 'MOSexpress'
    }),
    muteHttpExceptions: true
  };

  try {
    var resp = UrlFetchApp.fetch('https://api.printnode.com/printjobs', options);
    var code = resp.getResponseCode();
    if (code !== 201) {
      return generarRespuestaError("PrintNode respondió " + code + " (printerId=" + printerId + "): " + resp.getContentText());
    }
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      printJobId: resp.getContentText()
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return generarRespuestaError("Error llamando a PrintNode: " + err.toString());
  }
}

// ── Emite un CPE (Boleta o Factura) vía NubeFact API ─────────────────────────
// Script Properties requeridas: NUBEFACT_TOKEN, NUBEFACT_RUC
// Retorna: { ok, hash, enlace, qrString, aceptada, error }
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

  // Extraer serie y número del correlativo (ej: "B001-000000042" → serie=B001, numero=42)
  var partes = correlativo.split('-');
  var serie  = partes[0] || '';
  var numero = parseInt(partes[partes.length - 1], 10) || 1;
  var tipoComprobante = (tipoDoc === 'FACTURA') ? 1 : 2; // 1=Factura, 2=Boleta

  // Calcular totales gravada/exonerada/inafecta/igv
  var totalGravada   = 0;
  var totalExonerada = 0;
  var totalInafecta  = 0;

  var nfItems = items.map(function(item) {
    var tipoIgv       = parseInt(item.tipo_igv || 1, 10);
    var cantidad      = parseFloat(item.cantidad || 1);
    var valorUnitario = parseFloat(item.valor_unitario || 0);
    var subtotalVU    = Math.round(valorUnitario * cantidad * 100) / 100;
    var precioTotal   = parseFloat(item.subtotal || 0);
    var igvItem       = Math.round((precioTotal - subtotalVU) * 100) / 100;

    if (tipoIgv === 1) {
      totalGravada += subtotalVU;
    } else if (tipoIgv === 2) {
      totalExonerada += precioTotal;
      igvItem = 0;
    } else {
      totalInafecta += precioTotal;
      igvItem = 0;
    }

    return {
      unidad_de_medida:       String(item.unidad_de_medida || 'NIU'),
      codigo:                 String(item.sku || ''),
      codigo_producto_sunat:  String(item.cod_sunat || ''),
      descripcion:            String(item.nombre || ''),
      cantidad:               cantidad,
      valor_unitario:         Math.round(valorUnitario * 100) / 100,
      precio_unitario:        parseFloat(item.precio || 0),
      descuento:              '',
      subtotal:               subtotalVU,
      tipo_de_igv:            tipoIgv,
      igv:                    igvItem,
      total:                  precioTotal,
      anticipo_regularizacion:  false,
      anticipo_documento_serie: '',
      anticipo_documento_numero:''
    };
  });

  totalGravada   = Math.round(totalGravada   * 100) / 100;
  totalExonerada = Math.round(totalExonerada * 100) / 100;
  totalInafecta  = Math.round(totalInafecta  * 100) / 100;
  var totalGeneral = parseFloat(header.total || 0);
  var totalIgv     = Math.round((totalGeneral - totalGravada - totalExonerada - totalInafecta) * 100) / 100;

  var cliente    = header.cliente || {};
  var fechaHoy   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy');

  var payload = {
    operacion:                    'generar_comprobante',
    tipo_de_comprobante:          tipoComprobante,
    serie:                        serie,
    numero:                       numero,
    sunat_transaction:            1,
    cliente_tipo_de_documento:    parseInt(cliente.tipo || 0, 10),
    cliente_numero_de_documento:  String(cliente.doc   || '0'),
    cliente_denominacion:         String(cliente.nombre || 'CLIENTE ANONIMO'),
    cliente_direccion:            String(cliente.direccion || ''),
    cliente_email:                '',
    fecha_de_emision:             fechaHoy,
    fecha_de_vencimiento:         '',
    moneda:                       1,
    tipo_de_cambio:               '',
    porcentaje_de_igv:            18,
    total_gravada:                totalGravada   > 0 ? totalGravada   : '',
    total_exonerada:              totalExonerada > 0 ? totalExonerada : '',
    total_inafecta:               totalInafecta  > 0 ? totalInafecta  : '',
    total_igv:                    totalIgv       > 0 ? totalIgv       : '',
    total_precio_de_venta:        totalGeneral,
    total_descuentos:             '',
    total_otros_cargos:           '',
    total:                        totalGeneral,
    detraccion:                   false,
    enviar_automaticamente_a_la_sunat:    true,
    enviar_automaticamente_al_cliente:    false,
    formato_de_pdf:               'TICKET',
    items:                        nfItems
  };

  var endpoint = 'https://api.nubefact.com/api/v1/' + ruc + '/' + (tipoDoc === 'FACTURA' ? 'factura' : 'boleta');

  try {
    var resp = UrlFetchApp.fetch(endpoint, {
      method:            'post',
      headers:           { 'Authorization': 'Token ' + token, 'Content-Type': 'application/json' },
      payload:           JSON.stringify(payload),
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    var body = {};
    try { body = JSON.parse(resp.getContentText() || '{}'); } catch(pe) {}

    if (code === 200 || code === 201) {
      return {
        ok:       true,
        hash:     String(body.codigo_hash              || ''),
        enlace:   String(body.enlace_del_pdf           || ''),
        qrString: String(body.cadena_para_codigo_qr   || ''),
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

// ------ FUNCIONES AUXILIARES ------ //

// CORRECCIÓN: eliminada VENTAS_CABECERA del catálogo — no es necesaria en el browser
// y exponía todo el historial de ventas a cualquier dispositivo autorizado
function descargarCatalogo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetsNames = ["PRODUCTO_BASE", "PRESENTACIONES", "EQUIVALENCIAS", "PROMOCIONES", "ZONAS_CONFIG", "CLIENTES_FRECUENTES", "STOCK_ZONAS"];
  var catalogo = {};

  sheetsNames.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    catalogo[name] = sheet ? obtenerDatosHojaComoJSON(sheet) : [];
  });

  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    data: catalogo
  })).setMimeType(ContentService.MimeType.JSON);
}

// Columnas que deben tratarse siempre como texto (nunca como número)
var COLUMNAS_TEXTO = ['Cod_Barras', 'Cod_Barras_Real', 'SKU_Base', 'SKU', 'ID_Dispositivo', 'ID_Venta', 'ID_Caja', 'ID_Guia'];

function obtenerDatosHojaComoJSON(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  var headers = data[0];
  var result = [];

  for (var i = 1; i < data.length; i++) {
    var rowData = {};
    for (var j = 0; j < headers.length; j++) {
      var header = String(headers[j]).trim();
      var val = data[i][j];
      // Columnas de código siempre como texto para no perder ceros a la izquierda
      if (COLUMNAS_TEXTO.indexOf(header) !== -1) {
        val = val === '' || val === null || val === undefined ? '' : String(val).trim();
      }
      rowData[header] = val;
    }
    result.push(rowData);
  }
  return result;
}

// CORRECCIÓN: parsing más robusto — extrae solo el número final después del último guion
function obtenerSiguienteCorrelativo(sheet, serie) {
  var data = sheet.getDataRange().getValues();
  var prefijo = serie + "-";
  var maxCorrelativo = 0;

  for (var i = 1; i < data.length; i++) {
    var valorSerie = String(data[i][9]); // Columna J: correlativo (col 10, índice 9)
    if (valorSerie.indexOf(prefijo) === 0) {
      var parte = valorSerie.substring(prefijo.length); // todo lo que viene después del prefijo
      var num = parseInt(parte, 10);
      if (!isNaN(num) && num > maxCorrelativo) maxCorrelativo = num;
    }
  }
  return maxCorrelativo + 1;
}

// ── Verifica si un correlativo ya existe en las últimas 100 filas de VENTAS_CABECERA ──
// Bounded O(100) — detecta cuando CORRELATIVOS está desfasado sin escanear toda la hoja.
function correlativoYaExiste(ss, serie, numero) {
  var sheetCab = ss.getSheetByName('VENTAS_CABECERA');
  if (!sheetCab) return false;
  var objetivo = serie + '-' + ('000000' + numero).slice(-6);
  var totalRows = sheetCab.getLastRow();
  if (totalRows < 2) return false;
  var desde = Math.max(2, totalRows - 99);
  var filas = sheetCab.getRange(desde, 10, totalRows - desde + 1, 1).getValues(); // col 10 = correlativo
  for (var i = 0; i < filas.length; i++) {
    if (String(filas[i][0]) === objetivo) return true;
  }
  return false;
}

// ── Correlativo O(1): lee/incrementa una celda en hoja CORRELATIVOS ─────────
// Hoja CORRELATIVOS: encabezados Serie | Siguiente
// Crea la hoja automáticamente si no existe.
// Usa LockService para evitar duplicados en ventas simultáneas.
function obtenerSiguienteCorrelativoRapido(ss, serie) {
  var sheet = ss.getSheetByName('CORRELATIVOS');
  if (!sheet) {
    // Primera vez: fallback lento + creación automática de la hoja
    var sheetCab = ss.getSheetByName('VENTAS_CABECERA');
    var initial  = sheetCab ? obtenerSiguienteCorrelativo(sheetCab, serie) : 1;
    sheet = ss.insertSheet('CORRELATIVOS');
    sheet.appendRow(['Serie', 'Siguiente']);
    sheet.appendRow([serie, initial + 1]);
    return initial;
  }

  var lock = LockService.getScriptLock();
  try { lock.waitLock(6000); } catch (e) {
    // Contención excesiva — tomar el mayor entre CORRELATIVOS y O(n) para no retroceder
    var sheetCabFb = ss.getSheetByName('VENTAS_CABECERA');
    var numON = sheetCabFb ? obtenerSiguienteCorrelativo(sheetCabFb, serie) : 1;
    try {
      var dataFb = sheet.getDataRange().getValues();
      for (var fi = 1; fi < dataFb.length; fi++) {
        if (String(dataFb[fi][0]) === serie) {
          numON = Math.max(numON, parseInt(dataFb[fi][1], 10) || 1);
          break;
        }
      }
    } catch(e2) {}
    return numON;
  }

  try {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === serie) {
        var siguiente = parseInt(data[i][1], 10) || 1;

        // Salvaguarda anti-duplicado: si CORRELATIVOS fue editado manualmente
        // y quedó detrás de la realidad, avanzar hasta el siguiente número libre.
        while (correlativoYaExiste(ss, serie, siguiente)) {
          siguiente++;
        }

        sheet.getRange(i + 1, 2).setValue(siguiente + 1);
        return siguiente;
      }
    }
    // Serie nueva — detectar desde VENTAS_CABECERA para no reiniciar
    var sheetCab2 = ss.getSheetByName('VENTAS_CABECERA');
    var initial2  = sheetCab2 ? obtenerSiguienteCorrelativo(sheetCab2, serie) : 1;
    sheet.appendRow([serie, initial2 + 1]);
    return initial2;
  } finally {
    lock.releaseLock();
  }
}

// ── Normaliza texto para ESC/POS (quita tildes y no-ASCII) ──────────────────
function normalizarTextoGAS(str) {
  return String(str || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')   // eliminar diacríticos
    .replace(/[^\x20-\x7E]/g, '?');    // reemplazar no-ASCII con ?
}

// ── Genera bloque ESC/POS de QR Code (mismo algoritmo que el browser) ───────
function qrESCPOSGas(text) {
  var len = text.length + 3;
  var pL  = len & 0xFF;
  var pH  = (len >> 8) & 0xFF;
  return '\x1d\x28\x6b\x04\x00\x31\x41\x32\x00' +
         '\x1d\x28\x6b\x03\x00\x31\x43\x05' +
         '\x1d\x28\x6b\x03\x00\x31\x45\x31' +
         '\x1d\x28\x6b' + String.fromCharCode(pL) + String.fromCharCode(pH) +
         '\x31\x50\x30' + text +
         '\x1d\x28\x6b\x03\x00\x31\x51\x30' +
         '\n';
}

// ── Construye el ticket ESC/POS en GAS y envía a PrintNode ──────────────────
// Elimina el segundo round-trip browser→GAS→PrintNode.
// nfResult: objeto devuelto por emitirNubeFact (puede ser null para NOTA_DE_VENTA).
function imprimirTicketInternamente(data, correlativo, printerId, nfResult) {
  var printNodeKey = PropertiesService.getScriptProperties().getProperty('PRINTNODE_API_KEY');
  if (!printNodeKey || !printerId) return false;

  var auth   = data.auth   || {};
  var header = data.header || {};
  var items  = data.items  || [];

  var W    = 48;
  var SEP  = new Array(W + 1).join('=') + '\n';
  var SEPd = new Array(W + 1).join('-') + '\n';

  var tipoLabel = header.tipoDoc === 'NOTA_DE_VENTA' ? 'NOTA DE VENTA' :
                  header.tipoDoc === 'BOLETA'         ? 'BOLETA'         :
                  header.tipoDoc === 'FACTURA'        ? 'FACTURA'        :
                  normalizarTextoGAS(header.tipoDoc || '');

  var txt = '\x1b\x40';                                         // reset impresora
  txt += '\x1b\x61\x01';                                        // centrar
  txt += '\x1b\x21\x30MOSexpress\x1b\x21\x00\n';               // logo grande
  txt += '\x1b\x21\x10' + tipoLabel + '\x1b\x21\x00\n';        // tipo doc
  txt += 'Tk: ' + correlativo + '\n';
  txt += SEP;
  txt += '\x1b\x61\x00';                                        // izquierda
  txt += 'FECHA   : ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss') + '\n';

  var clienteNombre = normalizarTextoGAS((header.cliente && header.cliente.nombre) || '');
  var clienteDoc    = (header.cliente && header.cliente.doc) ? String(header.cliente.doc) : '';
  if (clienteNombre) txt += 'CLIENTE : ' + clienteNombre.substring(0, 38) + '\n';
  if (clienteDoc)    txt += 'DOC     : ' + clienteDoc + '\n';
  txt += (auth.esCajero ? 'CAJERO  ' : 'VENDEDOR') + ': ' + normalizarTextoGAS(auth.vendedor || '') + '\n';
  txt += SEP;
  txt += 'CANT  DESCRIPCION                      SUBTOTAL \n';
  txt += SEPd;

  items.forEach(function(item) {
    var nombre   = normalizarTextoGAS(item.nombre || '');
    var m        = nombre.match(/^(.+?)\s+\((.+)\)$/);
    var baseName = m ? m[1] : nombre;
    var empaque  = m ? m[2] : null;
    var desc = baseName.substring(0, 31);
    while (desc.length < 31) desc += ' ';
    var cant = String(item.cantidad || '').substring(0, 4);
    while (cant.length < 5) cant += ' ';
    var sub = parseFloat(item.subtotal || 0).toFixed(2);
    while (sub.length < 10) sub = ' ' + sub;
    txt += cant + ' ' + desc + ' ' + sub + '\n';
    if (empaque) txt += '        ' + empaque.substring(0, 38) + '\n';
  });

  txt += SEPd;
  txt += '\x1b\x61\x02';                                        // derecha
  txt += '\x1b\x21\x10TOTAL: S/ ' + parseFloat(header.total || 0).toFixed(2) + '\x1b\x21\x00\n';
  txt += 'METODO: ' + normalizarTextoGAS(header.metodo || 'EFECTIVO') + '\n';
  txt += '\n\x1b\x61\x01*** GRACIAS POR SU COMPRA ***\n';
  // Si NubeFact devolvió cadena QR SUNAT, usarla; si no, usar el correlativo
  var qrData = (nfResult && nfResult.qrString) ? nfResult.qrString : correlativo;
  txt += qrESCPOSGas(qrData);
  if (nfResult && nfResult.hash) {
    txt += '\x1b\x61\x01';
    txt += normalizarTextoGAS('Hash: ' + nfResult.hash).substring(0, W) + '\n';
  }
  if (nfResult && !nfResult.ok && nfResult.error) {
    txt += '\x1b\x61\x01[CPE pendiente de emision]\n';
  }
  txt += '\n\n\n\n\n\x1d\x56\x00\x1b\x6d\x1b\x69\x1b\x42\x05\x02'; // feed + corte + beep

  // Convertir a bytes explícitos (evita ambigüedad UTF-8 vs Latin-1)
  var bytes = [];
  for (var ci = 0; ci < txt.length; ci++) {
    bytes.push(txt.charCodeAt(ci) & 0xFF);
  }
  var content = Utilities.base64Encode(bytes);

  try {
    var resp = UrlFetchApp.fetch('https://api.printnode.com/printjobs', {
      method: 'post',
      headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(printNodeKey + ':') },
      contentType: 'application/json',
      payload: JSON.stringify({
        printerId:   parseInt(printerId, 10),
        title:       tipoLabel + ' ' + correlativo,
        contentType: 'raw_base64',
        content:     content,
        source:      'MOSexpress-GAS'
      }),
      muteHttpExceptions: true
    });
    return resp.getResponseCode() === 201;
  } catch (e) {
    Logger.log('imprimirTicketInternamente error: ' + e.toString());
    return false;
  }
}

function verificarYAgregaCliente(doc, nombre, tipoDoc, direccion) {
  if (!doc) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("CLIENTES_FRECUENTES");
  if (!sheet) return;
  var data = obtenerDatosHojaComoJSON(sheet);
  for (var i = 0; i < data.length; i++) {
    if (String(data[i].Documento) === String(doc)) return; // ya existe
  }
  // Esquema: Documento | Nombre | Tipo | Fecha | Direccion
  sheet.appendRow([doc, nombre, tipoDoc, new Date(), String(direccion || '')]);
}

// Anula en masa todos los tickets POR_COBRAR que no fueron cobrados al cierre del turno
function anulacionMasiva(data) {
  if (!data.ids || !data.ids.length) return generarRespuestaError("No se enviaron IDs a anular.");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("VENTAS_CABECERA");
  if (!sheet) return generarRespuestaError("VENTAS_CABECERA no encontrada.");

  var filas = sheet.getDataRange().getValues();
  var anulados = 0;

  for (var i = 1; i < filas.length; i++) {
    if (data.ids.indexOf(String(filas[i][0])) !== -1) {
      sheet.getRange(i + 1, 9).setValue('ANULADO'); // col 9: FormaPago = ANULADO
      anulados++;
    }
  }

  return ContentService.createTextOutput(JSON.stringify({
    status: "success", anulados: anulados
  })).setMimeType(ContentService.MimeType.JSON);
}

// Devuelve todas las ventas de hoy que pertenecen a la zona del cajero
// prefijosStr: "NV-001,B-001,F-001" (series de la zona separadas por coma)
function ventasHoyZona(prefijosStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("VENTAS_CABECERA");
  if (!sheet) return generarRespuestaError("VENTAS_CABECERA no encontrada");

  var prefijos = prefijosStr ? prefijosStr.split(',').map(function(p) { return p.trim(); }) : [];
  var data = sheet.getDataRange().getValues();
  var hoy = new Date().toDateString();
  var result = [];

  for (var i = 1; i < data.length; i++) {
    var fechaFila = new Date(data[i][1]).toDateString();
    if (fechaFila !== hoy) continue;

    var correlativo = String(data[i][9]); // col 10 (índice 9)
    if (prefijos.length > 0) {
      var enZona = prefijos.some(function(p) { return correlativo.indexOf(p) === 0; });
      if (!enZona) continue;
    }

    result.push({
      id_venta:       data[i][0],
      fecha:          data[i][1],
      vendedor:       data[i][2],
      estacion:       data[i][3],
      cliente_doc:    String(data[i][4] || ''),
      cliente_nombre: String(data[i][5] || ''),
      total:          parseFloat(data[i][6]) || 0,
      tipo_doc:       String(data[i][7] || ''),   // col 8: Tipo_Doc limpio
      forma_pago:     String(data[i][8] || ''),   // col 9: FormaPago
      correlativo:    correlativo,
      id_caja:        String(data[i][10] || ''),
      id_dispositivo: String(data[i][11] || ''),
      status:         String(data[i][12] || ''),  // col 13: Estado_Envio
      ref_local:      String(data[i][13] || ''), // col 14: Ref_Local (cross-ref QR)
      obs:            String(data[i][14] || '')  // col 15: Obs
    });
  }

  return ContentService.createTextOutput(JSON.stringify({
    status: "success", ventas: result
  })).setMimeType(ContentService.MimeType.JSON);
}

// Devuelve los items de una venta específica desde VENTAS_DETALLE
function detalleVenta(idVenta) {
  if (!idVenta) return generarRespuestaError("id_venta requerido");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("VENTAS_DETALLE");
  if (!sheet) return generarRespuestaError("VENTAS_DETALLE no encontrada");

  var data = sheet.getDataRange().getValues();
  var items = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(idVenta)) {
      items.push({
        sku:      String(data[i][1] || ''),
        nombre:   String(data[i][2] || ''),
        cantidad: parseFloat(data[i][3]) || 0,
        precio:   parseFloat(data[i][4]) || 0,
        subtotal: parseFloat(data[i][5]) || 0
      });
    }
  }
  return ContentService.createTextOutput(JSON.stringify({
    status: "success", items: items
  })).setMimeType(ContentService.MimeType.JSON);
}

// Actualiza el método de pago y estado de una venta existente (flujo COBRAR cajero)
function cobrarVentaExistente(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("VENTAS_CABECERA");
  if (!sheet) return generarRespuestaError("VENTAS_CABECERA no encontrada");

  var filas = sheet.getDataRange().getValues();
  for (var i = 1; i < filas.length; i++) {
    if (String(filas[i][0]) === String(data.idVenta)) {
      sheet.getRange(i + 1, 9).setValue(data.metodo);                          // col 9: FormaPago = método real
      if (data.cajaId) sheet.getRange(i + 1, 11).setValue(String(data.cajaId)); // col 11: ID_Caja (trazabilidad)
      return ContentService.createTextOutput(JSON.stringify({
        status: "success", mensaje: "Venta cobrada correctamente"
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  return generarRespuestaError("Venta con ID " + data.idVenta + " no encontrada.");
}

// Marca una venta como CREDITO y guarda observación (tipoEvento='CREDITAR_VENTA')
function creditarVenta(data) {
  if (!data.idVenta) return generarRespuestaError("idVenta requerido");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("VENTAS_CABECERA");
  if (!sheet) return generarRespuestaError("VENTAS_CABECERA no encontrada");
  var filas = sheet.getDataRange().getValues();
  for (var i = 1; i < filas.length; i++) {
    if (String(filas[i][0]) === String(data.idVenta)) {
      sheet.getRange(i + 1, 9).setValue('CREDITO');        // col 9: FormaPago
      sheet.getRange(i + 1, 15).setValue(String(data.obs || '')); // col 15: Obs
      return ContentService.createTextOutput(JSON.stringify({
        status: "success", mensaje: "Crédito registrado"
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  return generarRespuestaError("Venta " + data.idVenta + " no encontrada.");
}

// Registra un movimiento extra de caja (ingreso o egreso no asociado a venta)
function registrarExtraCaja(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("MOVIMIENTOS_EXTRA");
  if (!sheet) {
    sheet = ss.insertSheet("MOVIMIENTOS_EXTRA");
    sheet.appendRow(["ID_Extra","ID_Caja","Timestamp","Tipo","Monto","Concepto","Obs","Registrado_Por"]);
  }
  var id = "EX-" + new Date().getTime();
  sheet.appendRow([
    id,
    String(data.cajaId || ''),
    new Date(),
    String(data.tipo || 'EGRESO'),
    parseFloat(data.monto) || 0,
    String(data.concepto || ''),
    String(data.obs || ''),
    String(data.registradoPor || '')
  ]);
  return ContentService.createTextOutput(JSON.stringify({
    status: "success", idExtra: id
  })).setMimeType(ContentService.MimeType.JSON);
}

// Devuelve los movimientos extra registrados para una caja
function getExtrasCaja(cajaId) {
  if (!cajaId) return generarRespuestaError("cajaId requerido");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("MOVIMIENTOS_EXTRA");
  if (!sheet) return ContentService.createTextOutput(JSON.stringify({ status: "success", extras: [] })).setMimeType(ContentService.MimeType.JSON);
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(cajaId)) {
      result.push({
        id:            String(data[i][0] || ''),
        tipo:          String(data[i][3] || 'EGRESO'),
        monto:         parseFloat(data[i][4]) || 0,
        concepto:      String(data[i][5] || ''),
        obs:           String(data[i][6] || ''),
        registradoPor: String(data[i][7] || '')
      });
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ status: "success", extras: result })).setMimeType(ContentService.MimeType.JSON);
}

// Anula un ticket individual enviado desde el frontend (tipoEvento='ANULACION')
function anularVentaIndividual(data) {
  if (!data.ventaId) return generarRespuestaError("No se proporcionó ventaId.");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("VENTAS_CABECERA");
  if (!sheet) return generarRespuestaError("VENTAS_CABECERA no encontrada.");

  var filas = sheet.getDataRange().getValues();
  for (var i = 1; i < filas.length; i++) {
    if (String(filas[i][0]) === String(data.ventaId)) {
      sheet.getRange(i + 1, 9).setValue('ANULADO'); // col 9: FormaPago = ANULADO
      return ContentService.createTextOutput(JSON.stringify({
        status: "success", mensaje: "Venta anulada correctamente"
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  return generarRespuestaError("Venta con ID " + data.ventaId + " no encontrada.");
}

function generarRespuestaError(msg) {
  return ContentService.createTextOutput(JSON.stringify({
    status: "error", mensaje: msg
  })).setMimeType(ContentService.MimeType.JSON);
}

// ------ STOCK: GUÍAS Y AUDITORÍA ------ //

function cajeroActivo(zona) {
  if (!zona) return generarRespuestaError("zona requerida");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("CAJAS");
  if (!sheet) return generarRespuestaError("CAJAS no encontrada");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    // Col 6 (índice 5) = Estado, Col 9 (índice 8) = Zona_ID
    if (String(data[i][5]) === 'ABIERTA' && String(data[i][8] || '') === zona) {
      return ContentService.createTextOutput(JSON.stringify({
        status: 'success', activo: true,
        vendedor: String(data[i][1]), idCaja: String(data[i][0]), desde: data[i][3]
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ status: 'success', activo: false }))
    .setMimeType(ContentService.MimeType.JSON);
}

function listarGuias(zona) {
  if (!zona) return generarRespuestaError("zona requerida");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("GUIAS_CABECERA");
  if (!sheet) return ContentService.createTextOutput(JSON.stringify({ status: 'success', guias: [] })).setMimeType(ContentService.MimeType.JSON);
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    // Incluir guías donde Zona_ID === zona O Zona_Destino === zona (traslados entrantes)
    if (String(data[i][3]) === zona || String(data[i][6] || '') === zona) {
      result.push({
        id_guia:      String(data[i][0]),
        fecha:        data[i][1],
        vendedor:     String(data[i][2]),
        zona:         String(data[i][3]),
        tipo:         String(data[i][4]),
        observacion:  String(data[i][5] || ''),
        zona_destino: String(data[i][6] || ''),
        estado:       String(data[i][7] || '')
      });
    }
  }
  result.sort(function(a, b) { return new Date(b.fecha) - new Date(a.fecha); });
  return ContentService.createTextOutput(JSON.stringify({ status: 'success', guias: result }))
    .setMimeType(ContentService.MimeType.JSON);
}

function detalleGuia(idGuia) {
  if (!idGuia) return generarRespuestaError("id_guia requerido");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("GUIAS_DETALLE");
  if (!sheet) return generarRespuestaError("GUIAS_DETALLE no encontrada");
  var data = sheet.getDataRange().getValues();
  var items = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(idGuia)) {
      items.push({ cod_barras: String(data[i][1]), cantidad: parseFloat(data[i][2]) || 0 });
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ status: 'success', items: items }))
    .setMimeType(ContentService.MimeType.JSON);
}

function trasladosEntrantes(zona, desde) {
  if (!zona) return generarRespuestaError("zona requerida");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("GUIAS_CABECERA");
  if (!sheet) return ContentService.createTextOutput(JSON.stringify({ status: 'success', traslados: [] })).setMimeType(ContentService.MimeType.JSON);
  var data = sheet.getDataRange().getValues();
  // Default: últimas 24 horas si no se pasa timestamp
  var desdeDate = desde ? new Date(parseInt(desde)) : new Date(Date.now() - 86400000);
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][4]) !== 'ENTRADA_TRASLADO') continue;
    if (String(data[i][3]) !== zona) continue;
    var fechaGuia = new Date(data[i][1]);
    if (fechaGuia > desdeDate) {
      result.push({
        id_guia:   String(data[i][0]),
        fecha:     data[i][1],
        origen:    String(data[i][6] || ''),
        observacion: String(data[i][5] || '')
      });
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ status: 'success', traslados: result }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getStockZonas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("STOCK_ZONAS");
  if (!sheet) return ContentService.createTextOutput(JSON.stringify({ status: 'success', stock: [] })).setMimeType(ContentService.MimeType.JSON);
  var data = obtenerDatosHojaComoJSON(sheet);
  return ContentService.createTextOutput(JSON.stringify({ status: 'success', stock: data })).setMimeType(ContentService.MimeType.JSON);
}

// Actualiza (o crea) la fila de stock para un código+zona. Devuelve la cantidad resultante.
function actualizarStockFila(sheet, codBarras, zonaId, delta) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(codBarras) && String(data[i][1]) === String(zonaId)) {
      var nuevaCant = (parseFloat(data[i][2]) || 0) + delta;
      sheet.getRange(i + 1, 3).setValue(nuevaCant);
      return nuevaCant;
    }
  }
  // No existe — crear fila nueva
  var cantInicial = Math.max(0, delta);
  sheet.appendRow([String(codBarras), String(zonaId), cantInicial]);
  return cantInicial;
}

function registrarGuia(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetCab   = ss.getSheetByName("GUIAS_CABECERA");
  var sheetDet   = ss.getSheetByName("GUIAS_DETALLE");
  var sheetStock = ss.getSheetByName("STOCK_ZONAS");
  if (!sheetCab)   return generarRespuestaError("Pestaña GUIAS_CABECERA no encontrada.");
  if (!sheetDet)   return generarRespuestaError("Pestaña GUIAS_DETALLE no encontrada.");
  if (!sheetStock) return generarRespuestaError("Pestaña STOCK_ZONAS no encontrada.");

  // Tipos: SALIDA_JEFA | SALIDA_MOVIMIENTO | SALIDA_VENTAS
  //        ENTRADA_ALMACEN | ENTRADA_TRASLADO | ENTRADA_LIBRE
  var tipo        = data.tipo;
  var esSalida    = (tipo.indexOf('SALIDA') === 0);
  var signo       = esSalida ? -1 : 1;
  var zonaDestino = String(data.zona_destino || '');

  var idGuia = "G-" + new Date().getTime();
  // GUIAS_CABECERA: ID_Guia | Fecha | Vendedor | Zona_ID | Tipo | Observacion | Zona_Destino | Estado
  sheetCab.appendRow([idGuia, new Date(), data.vendedor, data.zona, tipo, data.observacion || '', zonaDestino, 'CONFIRMADO']);

  var stockResult = [];
  // Validar que las salidas no excedan el stock disponible
  if (esSalida) {
    var stockData = sheetStock.getDataRange().getValues();
    for (var si = 0; si < (data.items || []).length; si++) {
      var siItem = data.items[si];
      var siCb   = String(siItem.cod_barras);
      var stockActual = 0;
      for (var sr = 1; sr < stockData.length; sr++) {
        if (String(stockData[sr][0]) === siCb && String(stockData[sr][1]) === String(data.zona)) {
          stockActual = parseFloat(stockData[sr][2]) || 0;
          break;
        }
      }
      if (stockActual < siItem.cantidad) {
        return generarRespuestaError('Stock insuficiente para ' + siCb + ': disponible=' + stockActual + ', solicitado=' + siItem.cantidad);
      }
    }
  }
  (data.items || []).forEach(function(item) {
    var cb = String(item.cod_barras);
    sheetDet.appendRow([idGuia, cb, item.cantidad]);
    var nuevaCant = actualizarStockFila(sheetStock, cb, data.zona, signo * item.cantidad);
    stockResult.push({ cod_barras: cb, cantidad: nuevaCant });
  });

  // SALIDA_MOVIMIENTO → ENTRADA_TRASLADO automática en zona destino
  var idGuiaEntrada = null;
  if (tipo === 'SALIDA_MOVIMIENTO' && zonaDestino) {
    idGuiaEntrada = "G-TRA-" + (new Date().getTime() + 1);
    sheetCab.appendRow([idGuiaEntrada, new Date(), data.vendedor, zonaDestino, 'ENTRADA_TRASLADO',
      'Traslado desde ' + data.zona + ' — Guía origen: ' + idGuia, data.zona, 'CONFIRMADO']);
    (data.items || []).forEach(function(item) {
      var cb = String(item.cod_barras);
      sheetDet.appendRow([idGuiaEntrada, cb, item.cantidad]);
      actualizarStockFila(sheetStock, cb, zonaDestino, item.cantidad);
    });
  }

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success', idGuia: idGuia, idGuiaEntrada: idGuiaEntrada, stock: stockResult
  })).setMimeType(ContentService.MimeType.JSON);
}

function registrarAuditoria(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetAudit = ss.getSheetByName("AUDITORIAS");
  var sheetStock = ss.getSheetByName("STOCK_ZONAS");
  if (!sheetAudit) return generarRespuestaError("Pestaña AUDITORIAS no encontrada.");
  if (!sheetStock) return generarRespuestaError("Pestaña STOCK_ZONAS no encontrada.");

  var idAudit = "A-" + new Date().getTime();
  // Columnas AUDITORIAS: ID_Auditoria | Fecha | Vendedor | Zona_ID | Cod_Barras | Cant_Sistema | Cant_Real | Diferencia
  (data.items || []).forEach(function(item) {
    var cb       = String(item.cod_barras);
    var cantReal = parseFloat(item.cantReal) || 0;
    var diff     = cantReal - (parseFloat(item.cantSistema) || 0);
    sheetAudit.appendRow([idAudit, new Date(), data.vendedor, data.zona, cb, item.cantSistema, cantReal, diff]);
    // Actualizar stock a la cantidad real contada (diff puede ser + o -)
    actualizarStockFila(sheetStock, cb, data.zona, diff);
  });

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success', idAuditoria: idAudit
  })).setMimeType(ContentService.MimeType.JSON);
}

function verificarDispositivo(deviceId) {
  if (!deviceId) return generarRespuestaError("ID de dispositivo no proporcionado");

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("DISPOSITIVOS");

  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "success", autorizado: false, mensaje: "Tabla DISPOSITIVOS no encontrada"
    })).setMimeType(ContentService.MimeType.JSON);
  }

  var data = obtenerDatosHojaComoJSON(sheet);
  var autorizado = false;

  for (var i = 0; i < data.length; i++) {
    if (data[i].ID_Dispositivo === deviceId && data[i].Estado === 'ACTIVO') {
      autorizado = true;
      break;
    }
  }

  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    autorizado: autorizado
  })).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// CONSULTA DNI/RUC — vía APISPeru
// Requiere propiedad de script: APISPERU_TOKEN
// Registro en: https://apisperu.com/servicios/dniruc/
// ============================================================
function consultarCliente(doc) {
  if (!doc) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'Documento requerido' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  doc = String(doc).trim();

  // 1. Buscar primero en la tabla local CLIENTES_FRECUENTES
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('CLIENTES_FRECUENTES');
  if (sheet) {
    var rows = sheet.getDataRange().getValues();
    var headers = rows[0].map(function(h) { return String(h).trim(); });
    var docIdx = headers.indexOf('Documento');
    var nomIdx = headers.indexOf('Nombre');
    var dirIdx = headers.indexOf('Direccion'); // puede no existir en tablas antiguas
    if (docIdx >= 0 && nomIdx >= 0) {
      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][docIdx]).trim() === doc) {
          return ContentService.createTextOutput(JSON.stringify({
            status:    'success',
            nombre:    String(rows[i][nomIdx]),
            documento: doc,
            tipo:      doc.length === 11 ? 'RUC' : 'DNI',
            fuente:    'local',
            direccion: dirIdx >= 0 ? String(rows[i][dirIdx] || '') : ''
          })).setMimeType(ContentService.MimeType.JSON);
        }
      }
    }
  }

  // 2. Consultar APISPeru
  var token = PropertiesService.getScriptProperties().getProperty('APISPERU_TOKEN');
  if (!token) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: 'Token no configurado. Agregar APISPERU_TOKEN en Propiedades del script.'
    })).setMimeType(ContentService.MimeType.JSON);
  }

  try {
    var tipo = doc.length === 11 ? 'ruc' : 'dni';
    var url = 'https://dniruc.apisperu.com/api/v1/' + tipo + '/' + doc + '?token=' + token;
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var json = JSON.parse(response.getContentText());

    var nombre    = '';
    var direccion = '';
    if (tipo === 'dni') {
      nombre = [json.nombres, json.apellidoPaterno, json.apellidoMaterno].filter(Boolean).join(' ').trim();
      // RENIEC no expone dirección públicamente — se deja vacía
    } else {
      nombre    = (json.razonSocial || '').trim();
      direccion = (json.direccion   || '').trim(); // APISPeru devuelve dirección fiscal del RUC
    }

    if (!nombre) {
      return ContentService.createTextOutput(JSON.stringify({
        status: 'not_found',
        message: 'No se encontró información para ' + doc
      })).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({
      status:    'success',
      nombre:    nombre,
      documento: doc,
      tipo:      tipo === 'ruc' ? 'RUC' : 'DNI',
      fuente:    'api',
      direccion: direccion
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: 'Error consultando API: ' + e.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
