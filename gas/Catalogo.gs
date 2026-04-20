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
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var mosSsId  = PropertiesService.getScriptProperties().getProperty('MOS_SS_ID') || '';
  var catalogo = {};

  if (mosSsId) {
    // ── Phase 2: leer catálogo maestro desde MOS ──────────────
    try {
      var mosSS     = SpreadsheetApp.openById(mosSsId);
      var prodRows  = _obtenerHojaMOS(mosSS, 'PRODUCTOS_MASTER');
      var equivRows = _obtenerHojaMOS(mosSS, 'EQUIVALENCIAS');

      // Normalizar claves — trim + string para evitar espacios o tipos inesperados
      var _normKey = function(v) { return String(v === null || v === undefined ? '' : v).trim(); };
      // estado: excluir solo los explícitamente desactivados ('0' o 0)
      var _activo  = function(p) { return String(p.estado) !== '0'; };

      // ── Agrupar por skuBase (igual que MOS admin) ────────────────
      // Cada grupo = un SKU_Base. El miembro con menor factorConversion = base.
      // Productos sin skuBase usan su propio idProducto como grupo.
      var grupos = {};
      prodRows.forEach(function(p) {
        if (!_activo(p)) return;
        var sku = _normKey(p.skuBase) || _normKey(p.idProducto);
        if (!sku) return;
        if (!grupos[sku]) grupos[sku] = [];
        grupos[sku].push(p);
      });

      // Ordenar cada grupo por factor ascendente (factor 1 = base)
      var _pf = function(v) { return parseFloat(String(v === null || v === undefined ? '' : v).replace(',', '.')) || 1; };
      Object.keys(grupos).forEach(function(sku) {
        grupos[sku].sort(function(a, b) { return _pf(a.factorConversion) - _pf(b.factorConversion); });
      });

      Logger.log('MOS bridge — total filas: ' + prodRows.length + ' | grupos: ' + Object.keys(grupos).length);

      // PRODUCTO_BASE: un registro por grupo — representante = item con factor=1
      // (si no existe factor=1, usa el primero tras el sort ascendente)
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

      // PRESENTACIONES: todos los miembros del grupo con SKU_Base = sku del grupo
      catalogo['PRESENTACIONES'] = [];
      Object.keys(grupos).forEach(function(sku) {
        grupos[sku].forEach(function(p) {
          catalogo['PRESENTACIONES'].push({
            SKU_Base:     sku,
            SKU:          _normKey(p.idProducto),                            // SKU propio del producto (≠ barcode)
            Cod_Barras:   _normKey(p.codigoBarra) || _normKey(p.idProducto),
            Empaque:      p.descripcion || '',
            Precio_Venta: _parsePrice(p.precioVenta),
            Factor:       _pf(p.factorConversion)
          });
        });
      });

      // EQUIVALENCIAS: { Cod_Alias, Cod_Barras_Real }
      // Cod_Alias = codigo alternativo escaneado → Cod_Barras_Real = skuBase del producto base
      catalogo['EQUIVALENCIAS'] = equivRows
        .filter(function(e){ return String(e.activo) === '1'; })
        .map(function(e) {
          return {
            Cod_Alias:      e.codigoBarra,
            Cod_Barras_Real: e.skuBase
          };
        });

    } catch(e) {
      // Si falla la lectura de MOS, loguear y caer a tablas locales
      Logger.log('MOS bridge ERROR: ' + e.message + ' | stack: ' + e.stack);
      mosSsId = ''; // forzar fallback
    }
  }

  if (!mosSsId) {
    // ── Phase 1 fallback: tablas locales ──────────────────────
    ['PRODUCTO_BASE', 'PRESENTACIONES', 'EQUIVALENCIAS'].forEach(function(name) {
      var sheet = ss.getSheetByName(name);
      catalogo[name] = sheet ? obtenerDatosHojaComoJSON(sheet) : [];
    });
  }

  // ZONAS_CONFIG: si Phase 2 (MOS_SS_ID), construir desde MOS cruzando
  //   ESTACIONES + IMPRESORAS + SERIES_DOCUMENTALES
  // Si no → leer ZONAS_CONFIG local de ME (Phase 1 / fallback)
  if (mosSsId) {
    try {
      var mosSS2       = SpreadsheetApp.openById(mosSsId);
      var estRows      = _obtenerHojaMOS(mosSS2, 'ESTACIONES');
      var impRows      = _obtenerHojaMOS(mosSS2, 'IMPRESORAS');
      var seriesRows   = _obtenerHojaMOS(mosSS2, 'SERIES_DOCUMENTALES');

      // ── Mapa idEstacion → printNodeId (impresora TICKET activa de ME) ──
      var printMap = {};
      impRows.forEach(function(imp) {
        if (String(imp.activo) === '0') return;
        if (imp.appOrigen && imp.appOrigen !== 'mosExpress') return;
        var tipo = String(imp.tipo || '').toUpperCase();
        if (tipo && tipo !== 'TICKET') return; // solo impresoras de ticket
        var est = String(imp.idEstacion || '').trim();
        if (est && !printMap[est]) printMap[est] = String(imp.printNodeId || '').trim();
      });

      // ── Mapa idZona → { Serie_Nota, Serie_Boleta, Serie_Factura } ──
      var serieMap = {};
      seriesRows.forEach(function(r) {
        if (String(r.activo) === '0') return;
        var zId  = String(r.idZona        || '').trim();
        var tipo = String(r.tipoDocumento || '').toUpperCase().replace(/[\s_]/g, '');
        var ser  = String(r.serie         || '').trim();
        if (!zId || !ser) return;
        if (!serieMap[zId]) serieMap[zId] = { Serie_Nota: '', Serie_Boleta: '', Serie_Factura: '' };
        if (tipo === 'NOTAVENTA' || tipo === 'NV' || tipo === 'NOTADEVENTA') serieMap[zId].Serie_Nota    = ser;
        else if (tipo === 'BOLETA')  serieMap[zId].Serie_Boleta   = ser;
        else if (tipo === 'FACTURA') serieMap[zId].Serie_Factura  = ser;
      });

      // ── Una fila de ZONAS_CONFIG por cada estación activa de ME ──
      // appOrigen: incluir si es 'mosExpress' o si está vacío (no asignado aún)
      // Excluir solo los que explícitamente pertenecen a otra app (ej: 'warehouseMos')
      var zonasConfig = [];
      estRows.forEach(function(est) {
        if (String(est.activo) === '0') return;
        var origen = String(est.appOrigen || '').trim();
        if (origen && origen !== 'mosExpress') return;
        var zId     = String(est.idZona      || '').trim();
        var estId   = String(est.idEstacion  || '').trim();
        var nombre  = String(est.nombre      || '').trim();
        if (!nombre) return; // ignorar filas vacías
        var pin     = String(est.adminPin    || '').trim();
        var printer = printMap[estId] || '';
        var series  = serieMap[zId]   || { Serie_Nota: '', Serie_Boleta: '', Serie_Factura: '' };
        zonasConfig.push({
          Zona_ID:         zId,
          Estacion_Nombre: nombre,
          idEstacion:      estId,
          PrintNode_ID:    printer,
          Serie_Nota:      series.Serie_Nota,
          Serie_Boleta:    series.Serie_Boleta,
          Serie_Factura:   series.Serie_Factura,
          Admin_PIN:       pin
        });
      });

      // Si MOS devuelve vacío (ESTACIONES no migradas aún) → usar tabla local
      if (zonasConfig.length === 0) {
        Logger.log('MOS bridge ZONAS_CONFIG — ESTACIONES vacías, usando tabla local de ME');
        var fbSheet = ss.getSheetByName('ZONAS_CONFIG');
        catalogo['ZONAS_CONFIG'] = fbSheet ? obtenerDatosHojaComoJSON(fbSheet) : [];
      } else {
        catalogo['ZONAS_CONFIG'] = zonasConfig;
        Logger.log('MOS bridge ZONAS_CONFIG — ' + zonasConfig.length + ' estaciones desde MOS');
      }

    } catch(eZonas) {
      Logger.log('ZONAS_CONFIG bridge ERROR: ' + eZonas.message + ' — usando tabla local');
      var fallbackSheet = ss.getSheetByName('ZONAS_CONFIG');
      catalogo['ZONAS_CONFIG'] = fallbackSheet ? obtenerDatosHojaComoJSON(fallbackSheet) : [];
    }
  } else {
    var localSheet = ss.getSheetByName('ZONAS_CONFIG');
    catalogo['ZONAS_CONFIG'] = localSheet ? obtenerDatosHojaComoJSON(localSheet) : [];
  }

  // Estas siempre vienen de ME (datos operativos locales)
  ['PROMOCIONES', 'CLIENTES_FRECUENTES', 'STOCK_ZONAS'].forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    catalogo[name] = sheet ? obtenerDatosHojaComoJSON(sheet) : [];
  });

  // Indicar al frontend si estamos en modo MOS
  catalogo['_meta'] = {
    fuente:   mosSsId ? 'MOS' : 'LOCAL',
    timestamp: new Date().getTime()
  };

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
  // Trim headers para evitar problemas con espacios accidentales
  var headers = data[0].map(function(h) { return String(h).trim(); });
  return data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) {
      if (!h) return; // ignorar columnas sin cabecera
      var v = row[i];
      obj[h] = v instanceof Date
        ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : v;
    });
    return obj;
  }).filter(function(obj) {
    return Object.values(obj).some(function(v){ return v !== '' && v !== null && v !== undefined; });
  });
}

// Parsea precio aceptando tanto punto como coma como separador decimal
function _parsePrice(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return val;
  return parseFloat(String(val).replace(',', '.')) || 0;
}

// Convierte el Tipo_IGV de MOS (string) al código numérico que usa ME internamente
// MOS almacena: "gravado", "exonerado", "inafecto" (o vacío = gravado por defecto)
// ME PRODUCTO_BASE: 1=Gravado, 2=Exonerado, 3=Inafecto
function _convertirTipoIGV(tipoMos) {
  var t = String(tipoMos || '').toLowerCase();
  if (t === 'exonerado') return 2;
  if (t === 'inafecto')  return 3;
  return 1; // gravado por defecto
}

function verificarDispositivo(deviceId) {
  if (!deviceId) return generarRespuestaError("ID de dispositivo no proporcionado");

  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var mosSsId = PropertiesService.getScriptProperties().getProperty('MOS_SS_ID') || '';

  // Phase 2: verificar contra tabla DISPOSITIVOS de MOS (gestión centralizada)
  // Phase 1: tabla local DISPOSITIVOS de ME
  var datos   = [];
  var mosSS   = null;
  var usingMOS = false;
  try {
    if (mosSsId) {
      mosSS = SpreadsheetApp.openById(mosSsId);
      var mosSheet = mosSS.getSheetByName('DISPOSITIVOS');
      if (mosSheet) {
        datos = obtenerDatosHojaComoJSON(mosSheet);
        usingMOS = true;
        Logger.log('verificarDispositivo — usando DISPOSITIVOS de MOS (' + datos.length + ' registros)');
      }
    }
  } catch(e) {
    Logger.log('verificarDispositivo MOS ERROR: ' + e.message + ' — fallback a local');
  }

  // Fallback: tabla local si MOS no tiene la hoja o falló
  if (!usingMOS) {
    var localSheet = ss.getSheetByName('DISPOSITIVOS');
    if (!localSheet) {
      return ContentService.createTextOutput(JSON.stringify({
        status: "success", autorizado: false, mensaje: "Tabla DISPOSITIVOS no encontrada"
      })).setMimeType(ContentService.MimeType.JSON);
    }
    datos = obtenerDatosHojaComoJSON(localSheet);
  }

  // Buscar el registro del dispositivo
  var registro = null;
  datos.forEach(function(d) {
    var idMatch = (d.ID_Dispositivo === deviceId || d.idDispositivo === deviceId);
    // Si viene de MOS, filtrar solo dispositivos de mosExpress
    var appMatch = !usingMOS || !d.App || d.App === 'mosExpress';
    if (idMatch && appMatch) registro = d;
  });

  var autorizado = registro !== null &&
    (registro.Estado === 'ACTIVO' || registro.estado === 'ACTIVO' ||
     registro.activo === '1'      || registro.activo === 1);

  // Actualizar Ultima_Conexion en MOS si el dispositivo está autorizado
  if (autorizado && usingMOS && mosSS) {
    try {
      var dispSheet = mosSS.getSheetByName('DISPOSITIVOS');
      if (dispSheet) {
        var sheetData = dispSheet.getDataRange().getValues();
        var hdrs = sheetData[0];
        var colUC = hdrs.indexOf('Ultima_Conexion');
        if (colUC >= 0) {
          var tz = Session.getScriptTimeZone();
          var ahora = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
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
