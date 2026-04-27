// ============================================================
// MosExpress — Catalogo.gs
// Descarga del catálogo al dispositivo + verificación de dispositivo
// + consulta DNI/RUC (APISPeru)
// Requiere Script Property: MOS_SS_ID (ID del Spreadsheet de MOS)
// ============================================================

function descargarCatalogo() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var mosSsId = PropertiesService.getScriptProperties().getProperty('MOS_SS_ID') || '';
  var catalogo = {};

  // ── Catálogo de productos desde MOS ────────────────────────
  try {
    var mosSS     = SpreadsheetApp.openById(mosSsId);
    var prodRows  = _obtenerHojaMOS(mosSS, 'PRODUCTOS_MASTER');
    var equivRows = _obtenerHojaMOS(mosSS, 'EQUIVALENCIAS');

    var _normKey = function(v) { return String(v === null || v === undefined ? '' : v).trim(); };
    var _activo  = function(p) { return String(p.estado) !== '0'; };

    // Agrupar por skuBase
    var grupos = {};
    prodRows.forEach(function(p) {
      if (!_activo(p)) return;
      var sku = _normKey(p.skuBase) || _normKey(p.idProducto);
      if (!sku) return;
      if (!grupos[sku]) grupos[sku] = [];
      grupos[sku].push(p);
    });

    var _pf = function(v) { return parseFloat(String(v === null || v === undefined ? '' : v).replace(',', '.')) || 1; };
    Object.keys(grupos).forEach(function(sku) {
      grupos[sku].sort(function(a, b) { return _pf(a.factorConversion) - _pf(b.factorConversion); });
    });

    Logger.log('MOS bridge — filas: ' + prodRows.length + ' | grupos: ' + Object.keys(grupos).length);

    // PRODUCTO_BASE: representante con factor=1 por grupo
    catalogo['PRODUCTO_BASE'] = Object.keys(grupos).map(function(sku) {
      var members = grupos[sku];
      var base = members.find(function(p) { return _pf(p.factorConversion) === 1; }) || members[0];
      return {
        SKU_Base:      sku,
        Nombre:        base.descripcion || '',
        Tipo_IGV:      _convertirTipoIGV(base.Tipo_IGV),
        Unidad_Medida: base.Unidad_Medida || 'NIU',
        Cod_SUNAT:     base.Cod_SUNAT || ''
      };
    });

    // PRESENTACIONES: todos los miembros del grupo
    catalogo['PRESENTACIONES'] = [];
    Object.keys(grupos).forEach(function(sku) {
      grupos[sku].forEach(function(p) {
        catalogo['PRESENTACIONES'].push({
          SKU_Base:     sku,
          SKU:          _normKey(p.idProducto),
          Cod_Barras:   _normKey(p.codigoBarra) || _normKey(p.idProducto),
          Empaque:      p.descripcion || '',
          Precio_Venta: _parsePrice(p.precioVenta),
          Factor:       _pf(p.factorConversion)
        });
      });
    });

    // EQUIVALENCIAS: { Cod_Alias, Cod_Barras_Real }
    catalogo['EQUIVALENCIAS'] = equivRows
      .filter(function(e) { return String(e.activo) === '1'; })
      .map(function(e) {
        return { Cod_Alias: e.codigoBarra, Cod_Barras_Real: e.skuBase };
      });

  } catch(e) {
    Logger.log('MOS bridge ERROR (productos): ' + e.message + ' | stack: ' + e.stack);
    catalogo['PRODUCTO_BASE']  = [];
    catalogo['PRESENTACIONES'] = [];
    catalogo['EQUIVALENCIAS']  = [];
  }

  // ── ZONAS_CONFIG desde MOS ──────────────────────────────────
  try {
    var mosSS2     = SpreadsheetApp.openById(mosSsId);
    var estRows    = _obtenerHojaMOS(mosSS2, 'ESTACIONES');
    var impRows    = _obtenerHojaMOS(mosSS2, 'IMPRESORAS');
    var seriesRows = _obtenerHojaMOS(mosSS2, 'SERIES_DOCUMENTALES');

    var printMap = {};
    impRows.forEach(function(imp) {
      if (String(imp.activo) === '0') return;
      var origenImp = String(imp.appOrigen || '').toLowerCase();
      if (origenImp && origenImp !== 'mosexpress') return;
      var tipo = String(imp.tipo || '').toUpperCase();
      if (tipo && tipo !== 'TICKET') return;
      var est = String(imp.idEstacion || '').trim();
      if (est && !printMap[est]) printMap[est] = String(imp.printNodeId || '').trim();
    });

    var serieMap = {};
    seriesRows.forEach(function(r) {
      if (String(r.activo) === '0') return;
      var zId  = String(r.idZona        || '').trim();
      var tipo = String(r.tipoDocumento || '').toUpperCase().replace(/[\s_]/g, '');
      var ser  = String(r.serie         || '').trim();
      if (!zId || !ser) return;
      if (!serieMap[zId]) serieMap[zId] = { Serie_Nota: '', Serie_Boleta: '', Serie_Factura: '' };
      if (tipo === 'NOTAVENTA' || tipo === 'NV' || tipo === 'NOTADEVENTA') serieMap[zId].Serie_Nota   = ser;
      else if (tipo === 'BOLETA')  serieMap[zId].Serie_Boleta  = ser;
      else if (tipo === 'FACTURA') serieMap[zId].Serie_Factura = ser;
    });

    var zonasConfig = [];
    estRows.forEach(function(est) {
      if (String(est.activo) === '0') return;
      var origen = String(est.appOrigen || '').trim().toLowerCase();
      if (origen && origen !== 'mosexpress') return;
      var zId    = String(est.idZona     || '').trim();
      var estId  = String(est.idEstacion || '').trim();
      var nombre = String(est.nombre     || '').trim();
      if (!nombre) return;
      var series = serieMap[zId] || { Serie_Nota: '', Serie_Boleta: '', Serie_Factura: '' };
      zonasConfig.push({
        Zona_ID:         zId,
        Estacion_Nombre: nombre,
        idEstacion:      estId,
        PrintNode_ID:    printMap[estId] || '',
        Serie_Nota:      series.Serie_Nota,
        Serie_Boleta:    series.Serie_Boleta,
        Serie_Factura:   series.Serie_Factura,
        Admin_PIN:       String(est.adminPin || '').trim()
      });
    });

    catalogo['ZONAS_CONFIG'] = zonasConfig;
    Logger.log('MOS bridge ZONAS_CONFIG — ' + zonasConfig.length + ' estaciones');

  } catch(eZonas) {
    Logger.log('MOS bridge ERROR (zonas): ' + eZonas.message);
    catalogo['ZONAS_CONFIG'] = [];
  }

  // ── Datos operativos locales de ME ─────────────────────────
  ['PROMOCIONES', 'CLIENTES_FRECUENTES', 'STOCK_ZONAS'].forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    catalogo[name] = sheet ? obtenerDatosHojaComoJSON(sheet) : [];
  });

  catalogo['_meta'] = { fuente: 'MOS', timestamp: new Date().getTime() };

  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    data: catalogo
  })).setMimeType(ContentService.MimeType.JSON);
}

// Lee una hoja del Spreadsheet de MOS como array de objetos
function _obtenerHojaMOS(ss, nombreHoja) {
  var sheet = ss.getSheetByName(nombreHoja);
  if (!sheet) throw new Error('Hoja no encontrada en MOS: ' + nombreHoja);
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0].map(function(h) { return String(h).trim(); });
  return data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) {
      if (!h) return;
      var v = row[i];
      obj[h] = v instanceof Date
        ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : v;
    });
    return obj;
  }).filter(function(obj) {
    return Object.values(obj).some(function(v) { return v !== '' && v !== null && v !== undefined; });
  });
}

function _parsePrice(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return val;
  return parseFloat(String(val).replace(',', '.')) || 0;
}

// MOS almacena: "gravado"|"exonerado"|"inafecto" → ME: 1|2|3
function _convertirTipoIGV(tipoMos) {
  var t = String(tipoMos || '').toLowerCase();
  if (t === 'exonerado') return 9;   // catálogo 07 SUNAT
  if (t === 'inafecto')  return 11;  // catálogo 07 SUNAT
  if (t === 'ivap')      return 8;   // catálogo 07 SUNAT (arroz pilado 4%)
  return 1;
}

function verificarDispositivo(deviceId) {
  if (!deviceId) return generarRespuestaError("ID de dispositivo no proporcionado");

  var mosSsId  = PropertiesService.getScriptProperties().getProperty('MOS_SS_ID') || '';
  var datos    = [];
  var mosSS    = null;
  var usingMOS = false;

  try {
    if (mosSsId) {
      mosSS = SpreadsheetApp.openById(mosSsId);
      var mosSheet = mosSS.getSheetByName('DISPOSITIVOS');
      if (mosSheet) {
        datos    = obtenerDatosHojaComoJSON(mosSheet);
        usingMOS = true;
        Logger.log('verificarDispositivo — MOS DISPOSITIVOS (' + datos.length + ' registros)');
      }
    }
  } catch(e) {
    Logger.log('verificarDispositivo MOS ERROR: ' + e.message);
  }

  if (!usingMOS) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "success", autorizado: false, mensaje: "DISPOSITIVOS de MOS no disponible"
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // Buscar el registro del dispositivo (filtrar solo App=mosExpress)
  var registro = null;
  datos.forEach(function(d) {
    var idMatch  = (d.ID_Dispositivo === deviceId || d.idDispositivo === deviceId);
    var appMatch = !d.App || d.App === 'mosExpress';
    if (idMatch && appMatch) registro = d;
  });

  var autorizado = registro !== null &&
    (registro.Estado === 'ACTIVO' || registro.estado === 'ACTIVO' ||
     registro.activo === '1'      || registro.activo === 1);

  // Actualizar Ultima_Conexion en MOS
  if (autorizado && mosSS) {
    try {
      var dispSheet = mosSS.getSheetByName('DISPOSITIVOS');
      if (dispSheet) {
        var sheetData = dispSheet.getDataRange().getValues();
        var hdrs  = sheetData[0];
        var colUC = hdrs.indexOf('Ultima_Conexion');
        if (colUC >= 0) {
          var ahora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
          for (var i = 1; i < sheetData.length; i++) {
            if (String(sheetData[i][0]) === String(deviceId)) {
              dispSheet.getRange(i + 1, colUC + 1).setValue(ahora);
              break;
            }
          }
        }
      }
    } catch(eUC) {
      Logger.log('verificarDispositivo — no se pudo actualizar Ultima_Conexion: ' + eUC.message);
    }
  }

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

  // 1. Buscar en CLIENTES_FRECUENTES local
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
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
    var tipo     = doc.length === 11 ? 'ruc' : 'dni';
    var url      = 'https://dniruc.apisperu.com/api/v1/' + tipo + '/' + doc + '?token=' + token;
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var json     = JSON.parse(response.getContentText());

    var nombre    = '';
    var direccion = '';
    if (tipo === 'dni') {
      nombre = [json.nombres, json.apellidoPaterno, json.apellidoMaterno].filter(Boolean).join(' ').trim();
    } else {
      nombre    = (json.razonSocial || '').trim();
      direccion = (json.direccion   || '').trim();
    }

    if (!nombre) {
      return ContentService.createTextOutput(JSON.stringify({
        status: 'not_found', message: 'No se encontró información para ' + doc
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

  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error', message: 'Error consultando API: ' + e.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
