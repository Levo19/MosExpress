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
// ============================================================

function doGet(e) {
  var accion = e.parameter.accion;

  if (accion === 'descargar')             return descargarCatalogo();
  if (accion === 'verificar_dispositivo') return verificarDispositivo(e.parameter.id);
  if (accion === 'ventas_hoy_zona')       return ventasHoyZona(e.parameter.prefijos);
  if (accion === 'detalle_venta')         return detalleVenta(e.parameter.id_venta);
  if (accion === 'stock_zonas')           return getStockZonas();
  if (accion === 'cajero_activo')         return cajeroActivo(e.parameter.zona);
  if (accion === 'listar_guias')          return listarGuias(e.parameter.zona);
  if (accion === 'detalle_guia')          return detalleGuia(e.parameter.id_guia);
  if (accion === 'traslados_entrantes')   return trasladosEntrantes(e.parameter.zona, e.parameter.desde);
  if (accion === 'consultar_cliente')     return consultarCliente(e.parameter.doc);
  if (accion === 'extras_caja')           return getExtrasCaja(e.parameter.cajaId);

  return generarRespuestaError("Acción no válida");
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    if (data.tipoEvento === 'APERTURA_CAJA')      return procesarAperturaCaja(data);
    if (data.tipoEvento === 'CIERRE_CAJA')         return procesarCierreCaja(data);
    if (data.tipoEvento === 'COBRAR_VENTA')        return cobrarVentaExistente(data);
    if (data.tipoEvento === 'ANULACION_MASIVA')    return anulacionMasiva(data);
    if (data.tipoEvento === 'CREDITAR_VENTA')      return creditarVenta(data);
    if (data.tipoEvento === 'EXTRA_CAJA')          return registrarExtraCaja(data);
    if (data.tipoEvento === 'ANULACION')           return anularVentaIndividual(data);
    if (data.tipoEvento === 'REGISTRAR_GUIA')      return registrarGuia(data);
    if (data.tipoEvento === 'REGISTRAR_AUDITORIA') return registrarAuditoria(data);
    if (data.accion === 'imprimir')                return procesarImpresion(data);

    // Default: registrar venta
    var response = procesarVenta(data);
    return ContentService.createTextOutput(JSON.stringify({
      status:         "success",
      idVenta:        response.idVenta,
      correlativo:    response.correlativo,
      printDispatched:response.printDispatched,
      nfEstado:       response.nfEstado || 'NA',
      nfHash:         response.nfHash   || '',
      nfEnlace:       response.nfEnlace || '',
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

// Columnas que deben tratarse siempre como texto (nunca como número)
var COLUMNAS_TEXTO = [
  'Cod_Barras', 'Cod_Barras_Real', 'SKU_Base', 'SKU',
  'ID_Dispositivo', 'ID_Venta', 'ID_Caja', 'ID_Guia'
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
