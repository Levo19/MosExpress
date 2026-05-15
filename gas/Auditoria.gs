// ============================================================
// MosExpress — Auditoria.gs
// Sistema de log JSON unificado para tickets, extras, clientes.
// Auto-creación de columna `historialCambios` en cada tabla.
// Cada entrada captura: usuario, ts, source, accion, cambios[], autorizadoPor
// ============================================================

// Tablas que llevan historial. Si necesitas habilitar otra, agrégala aquí.
var _AUD_TABLAS_CON_LOG = {
  'VENTAS_CABECERA':    { sheetName: 'VENTAS_CABECERA',    pkCol: 0  /* ID_Venta */ },
  'MOVIMIENTOS_EXTRA':  { sheetName: 'MOVIMIENTOS_EXTRA',  pkCol: 0  /* ID_Extra */ },
  'CLIENTES_FRECUENTES':{ sheetName: 'CLIENTES_FRECUENTES',pkCol: 0  /* Documento */ }
};

// Asegura que exista la columna `historialCambios` al final de la hoja.
// Idempotente: solo crea la columna si no existe. Devuelve el índice (1-based) de la columna.
function _audAsegurarColumna(sheet) {
  if (!sheet) return -1;
  var lastCol = sheet.getLastColumn() || 1;
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim() === 'historialCambios') return i + 1;
  }
  // Crear nueva al final
  var newColIdx = lastCol + 1;
  sheet.getRange(1, newColIdx).setValue('historialCambios');
  // Forzar ancho razonable (Sheets puede comprimirla y dificultar lectura)
  try { sheet.setColumnWidth(newColIdx, 380); } catch(_){}
  return newColIdx;
}

// Append una entrada al array JSON de historial de una fila.
// rowNum: índice 1-based (incluye header en row 1, datos desde row 2).
// entrada: objeto plano que ya cumple el schema.
// Limita a 200 entradas por fila para no inflar la celda (drop más antiguas).
function _audAppend(sheet, rowNum, entrada) {
  if (!sheet || !rowNum || !entrada) return false;
  var col = _audAsegurarColumna(sheet);
  if (col < 1) return false;
  var range = sheet.getRange(rowNum, col);
  var raw = range.getValue();
  var arr = [];
  if (raw) {
    try {
      var p = JSON.parse(String(raw));
      if (Array.isArray(p)) arr = p;
    } catch(_){}
  }
  // Asegurar timestamp ISO si no viene
  if (!entrada.ts) entrada.ts = new Date().toISOString();
  arr.push(entrada);
  // Cap a 200 entradas (drop primeras si excede)
  if (arr.length > 200) arr = arr.slice(arr.length - 200);
  range.setValue(JSON.stringify(arr));
  return true;
}

// Helper de alto nivel: construye una entrada estándar y la appendea.
// Uso desde GAS:
//   auditarLog('VENTAS_CABECERA', idVenta, {
//     usuario: 'Sara', rol: 'CAJERO',
//     source: 'ME_COBRO_CREDITO',
//     accion: 'cobrar_credito',
//     cambios: [{campo:'FormaPago', antes:'CREDITO', despues:'EFECTIVO'}],
//     autorizadoPor: { nombre:'Luis', rol:'MASTER', via:'PIN_8DIG' },
//     ref: { idCajaReceptora:'CAJA-...', idExtra:'EX-...' },
//     motivo: 'cliente pagó al día siguiente'
//   });
function auditarLog(tabla, pk, entrada) {
  try {
    var meta = _AUD_TABLAS_CON_LOG[tabla];
    if (!meta) { Logger.log('auditarLog: tabla no soportada: ' + tabla); return false; }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(meta.sheetName);
    if (!sheet) { Logger.log('auditarLog: hoja no encontrada: ' + meta.sheetName); return false; }
    var rowNum = _audBuscarRowPorPK(sheet, meta.pkCol, pk);
    if (rowNum < 2) { Logger.log('auditarLog: PK ' + pk + ' no encontrado en ' + tabla); return false; }
    return _audAppend(sheet, rowNum, entrada);
  } catch(e) {
    Logger.log('auditarLog excepcion: ' + e.message);
    return false;
  }
}

// Busca el rowNum (1-based) de una fila por su PK. Optimizado para PK en col 0.
function _audBuscarRowPorPK(sheet, pkColIdx, pkValor) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var col = sheet.getRange(2, pkColIdx + 1, lastRow - 1, 1).getValues();
  for (var i = col.length - 1; i >= 0; i--) {  // buscar desde el final (más probable)
    if (String(col[i][0]) === String(pkValor)) return i + 2;
  }
  return -1;
}

// Lee el historial de una fila como array. Devuelve [] si no hay nada.
function obtenerHistorialFila(tabla, pk) {
  try {
    var meta = _AUD_TABLAS_CON_LOG[tabla];
    if (!meta) return [];
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(meta.sheetName);
    if (!sheet) return [];
    var rowNum = _audBuscarRowPorPK(sheet, meta.pkCol, pk);
    if (rowNum < 2) return [];
    var col = _audAsegurarColumna(sheet);
    var raw = sheet.getRange(rowNum, col).getValue();
    if (!raw) return [];
    var p = JSON.parse(String(raw));
    return Array.isArray(p) ? p : [];
  } catch(e) {
    Logger.log('obtenerHistorialFila: ' + e.message);
    return [];
  }
}

// Endpoint GET para que el frontend lea el historial de una venta/extra.
function getHistorialEndpoint(tabla, pk) {
  if (!tabla || !pk) return generarRespuestaError('tabla y pk requeridos');
  var hist = obtenerHistorialFila(tabla, pk);
  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    historial: hist
  })).setMimeType(ContentService.MimeType.JSON);
}

// ── Helper: extrae usuario/admin del payload de cualquier endpoint ──
// Convención del payload (frontend siempre debe mandar esto en cualquier acción):
//   data.auth = { vendedor, esCajero, rol, deviceId }
//   data.adminAuth = { nombre, rol, via:'PIN_8DIG' }   ← solo si la acción usó PIN
// Devuelve un objeto listo para entrada.usuario / entrada.autorizadoPor.
function _audExtraerActor(data) {
  var auth = data.auth || {};
  // [v40.3] Fallback en cascada para capturar el operador real en TODOS los
  // endpoints. Antes solo leía auth.vendedor → cualquier endpoint legacy
  // (registrarExtra, etc.) que mandaba `registradoPor` o `usuario` en raíz
  // caía a "desconocido" en el campo Registrado_Por de MOVIMIENTOS_EXTRA.
  var usuario = String(
    auth.vendedor ||
    auth.nombre ||
    data.registradoPor ||
    data.usuario ||
    data.vendedor ||
    'desconocido'
  );
  var actor = {
    usuario:  usuario,
    rol:      String(auth.rol || (auth.esCajero ? 'CAJERO' : 'VENDEDOR')),
    deviceId: String(auth.deviceId || data.deviceId || '')
  };
  var admin = data.adminAuth || null;
  if (admin && admin.nombre) {
    actor.autorizadoPor = {
      nombre: String(admin.nombre),
      rol:    String(admin.rol || 'ADMIN'),
      via:    String(admin.via || 'PIN_8DIG')
    };
  }
  return actor;
}
