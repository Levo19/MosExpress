// ============================================================
// MosExpress — Catalogo.gs
// Descarga del catálogo al dispositivo + verificación de dispositivo
// + consulta DNI/RUC (APISPeru)
//
// Bridge MOS (Phase 2):
//   Cuando MOS_SS_ID esté configurado en Script Properties,
//   descargarCatalogo() leerá PRODUCTOS_MASTER y ESTACIONES
//   de MOS en lugar de las hojas locales PRODUCTO_BASE,
//   PRESENTACIONES y ZONAS_CONFIG.
//   Mientras esté vacío → tablas locales (sin cambio de comportamiento).
// ============================================================

function descargarCatalogo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Catálogo: tablas locales (Phase 1) ──────────────────────
  // Phase 2: cuando MOS_SS_ID esté activo, _descargarDesdeMOS() reemplazará
  // PRODUCTO_BASE, PRESENTACIONES y ZONAS_CONFIG con datos de PRODUCTOS_MASTER/ESTACIONES.
  var sheetsNames = [
    "PRODUCTO_BASE", "PRESENTACIONES", "EQUIVALENCIAS",
    "PROMOCIONES", "ZONAS_CONFIG", "CLIENTES_FRECUENTES", "STOCK_ZONAS"
  ];
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
  var autorizado = data.some(function(d) {
    return d.ID_Dispositivo === deviceId && d.Estado === 'ACTIVO';
  });

  return ContentService.createTextOutput(JSON.stringify({
    status: "success", autorizado: autorizado
  })).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// Consulta DNI/RUC vía APISPeru
// Script Property requerida: APISPERU_TOKEN
// ============================================================
function consultarCliente(doc) {
  if (!doc) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error', message: 'Documento requerido'
    })).setMimeType(ContentService.MimeType.JSON);
  }
  doc = String(doc).trim();

  // 1. Buscar primero en CLIENTES_FRECUENTES local
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('CLIENTES_FRECUENTES');
  if (sheet) {
    var rows    = sheet.getDataRange().getValues();
    var headers = rows[0].map(function(h) { return String(h).trim(); });
    var docIdx  = headers.indexOf('Documento');
    var nomIdx  = headers.indexOf('Nombre');
    var dirIdx  = headers.indexOf('Direccion');
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
      nombre = [json.nombres, json.apellidoPaterno, json.apellidoMaterno]
                .filter(Boolean).join(' ').trim();
    } else {
      nombre    = (json.razonSocial || '').trim();
      direccion = (json.direccion   || '').trim();
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
