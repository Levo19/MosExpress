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
    // esEnvasable=1 → producto a granel destinado a envasar (no se vende directo).
    // Solo se ocultan estas presentaciones; si dentro del mismo grupo SKU hay otra
    // presentación con esEnvasable=0/vacío, esa SÍ se muestra como "presentación
    // huérfana del grupo".
    var _esVendible = function(p) {
      return String(p.esEnvasable || '').trim() !== '1';
    };
    // Normalizar Unidad_Medida (código SUNAT). PRODUCTOS_MASTER tiene 2 columnas:
    //   - `unidad`         → libre (KGM, UND, KG, etc.)
    //   - `Unidad_Medida`  → código SUNAT autoritativo
    // Si `unidad` indica peso/granel, ese gana sobre `Unidad_Medida` (que puede
    // estar mal poblado con 'NIU' default por una migración previa).
    var _normUnidadMedida = function(p) {
      var u  = String(p.unidad || '').toUpperCase().trim();
      var um = String(p.Unidad_Medida || '').toUpperCase().trim();
      // Sinónimos de granel → KGM
      if (u === 'KGM' || u === 'KG' || u === 'KILO' || u === 'KILOS' ||
          u === 'GR'  || u === 'G'  || u === 'GRAMO' || u === 'GRAMOS') return 'KGM';
      // Sinónimos de litro → LTR
      if (u === 'LTR' || u === 'LT' || u === 'L' || u === 'LITRO' || u === 'LITROS') return 'LTR';
      // Sinónimos de metro → MTR
      if (u === 'MTR' || u === 'MT' || u === 'M' || u === 'METRO' || u === 'METROS') return 'MTR';
      // Si la col SUNAT tiene un código válido, usarlo
      if (um) return um;
      // Default
      return 'NIU';
    };

    // Agrupar por skuBase (todos los activos, sin filtrar por esEnvasable aún —
    // necesitamos el grupo completo para saber si HAY al menos un vendible)
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

    var skusOcultos = 0;
    Logger.log('MOS bridge — filas: ' + prodRows.length + ' | grupos: ' + Object.keys(grupos).length);

    // PRODUCTO_BASE: solo grupos que tengan AL MENOS un miembro vendible.
    // El representante se elige entre los vendibles (preferir factor=1).
    // Si el de factor=1 (producto base "nominal") es envasable y estamos
    // mostrando una presentación huérfana, concatenamos el nombre del granel
    // original + el nombre de la presentación para no perder contexto al buscar.
    // Ej: "Nakamito a granel" (factor=1 envasable) + "25kg saco" (vendible)
    //     → Nombre: "Nakamito a granel 25kg saco"
    catalogo['PRODUCTO_BASE'] = [];
    Object.keys(grupos).forEach(function(sku) {
      var members = grupos[sku];
      var vendibles = members.filter(_esVendible);
      if (vendibles.length === 0) {
        skusOcultos++;
        return; // todo el grupo es envasable → no mostrar
      }
      var miembroFactor1 = members.find(function(p) { return _pf(p.factorConversion) === 1; });
      var base = vendibles.find(function(p) { return _pf(p.factorConversion) === 1; }) || vendibles[0];
      var nombre;
      if (miembroFactor1 && !_esVendible(miembroFactor1) && base !== miembroFactor1) {
        // Producto "base nominal" envasable + presentación huérfana vendible
        var nomGranel = String(miembroFactor1.descripcion || '').trim();
        var nomPres   = String(base.descripcion || '').trim();
        nombre = nomGranel ? (nomGranel + ' ' + nomPres).trim() : nomPres;
      } else {
        nombre = String(base.descripcion || '').trim();
      }
      catalogo['PRODUCTO_BASE'].push({
        SKU_Base:      sku,
        Nombre:        nombre,
        Tipo_IGV:      _convertirTipoIGV(base.Tipo_IGV),
        Unidad_Medida: _normUnidadMedida(base),
        Cod_SUNAT:     base.Cod_SUNAT || ''
      });
    });

    // PRESENTACIONES: solo miembros vendibles (esEnvasable !== '1')
    catalogo['PRESENTACIONES'] = [];
    Object.keys(grupos).forEach(function(sku) {
      grupos[sku].forEach(function(p) {
        if (!_esVendible(p)) return; // saltar envasables individuales
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

    Logger.log('MOS bridge — grupos vendibles: ' + catalogo['PRODUCTO_BASE'].length +
               ' | grupos ocultos (todo envasable): ' + skusOcultos +
               ' | presentaciones vendibles: ' + catalogo['PRESENTACIONES'].length);

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
// [v2.5.59] consultarCliente — mejorado con:
//  · Respeta tipoDoc del frontend (FACTURA → solo RUC, NV/BOLETA → DNI o RUC)
//  · Reintentos automáticos (1 retry con 800ms backoff si falla red)
//  · Diagnóstico claro: distingue 'no_found' vs 'token_invalido' vs 'net_error'
//  · Logging del HTTP code y body para depurar
function consultarCliente(doc, tipoDocSolicitado) {
  if (!doc) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error', message: 'Documento requerido'
    })).setMimeType(ContentService.MimeType.JSON);
  }
  doc = String(doc).trim();
  tipoDocSolicitado = String(tipoDocSolicitado || '').toUpperCase();

  // [v2.5.59] Validación previa por tipoDoc — falla rápido sin gastar quota
  if (tipoDocSolicitado === 'FACTURA') {
    if (!/^\d{11}$/.test(doc)) {
      return ContentService.createTextOutput(JSON.stringify({
        status: 'validacion',
        message: 'FACTURA requiere RUC de 11 dígitos. Recibido: ' + doc.length + ' caracteres',
        codigo: 'FACTURA_NO_RUC'
      })).setMimeType(ContentService.MimeType.JSON);
    }
  } else {
    // NV / BOLETA: aceptar DNI (8) o RUC (11)
    if (!/^\d{8}$/.test(doc) && !/^\d{11}$/.test(doc)) {
      return ContentService.createTextOutput(JSON.stringify({
        status: 'validacion',
        message: 'Documento inválido. DNI son 8 dígitos, RUC son 11. Recibido: ' + doc.length,
        codigo: 'LONGITUD_INVALIDA'
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // 1. Buscar en CLIENTES_FRECUENTES local
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('CLIENTES_FRECUENTES');
  if (sheet) {
    var rows    = sheet.getDataRange().getValues();
    var headers = rows[0].map(function(h) { return String(h).trim(); });
    var docIdx  = headers.indexOf('Documento');
    // [v2.7.2] header real puede ser 'Nombre_RazonSocial' — aceptar ambos
    var nomIdx  = headers.indexOf('Nombre');
    if (nomIdx < 0) nomIdx = headers.indexOf('Nombre_RazonSocial');
    var dirIdx  = headers.indexOf('Direccion');
    // [v2.7.2] Forzar formato '@' en col Documento (idempotente, ahorra
    // pérdida de ceros en futuras escrituras desde aquí o desde fuera).
    if (docIdx >= 0) {
      try { sheet.getRange(1, docIdx + 1, sheet.getMaxRows(), 1).setNumberFormat('@'); } catch(_) {}
    }
    if (docIdx >= 0 && nomIdx >= 0) {
      var qDig = doc.replace(/\D/g, '');
      for (var i = 1; i < rows.length; i++) {
        // [v2.7.2] Comparación tolerante a DNI con cero perdido en row vieja
        var rawCell = String(rows[i][docIdx] == null ? '' : rows[i][docIdx]).trim();
        var aDig = rawCell.replace(/\D/g, '');
        var matchExacto   = (rawCell === doc) || (aDig === qDig);
        var matchPadStart = (qDig.length === 8 && aDig.length === 7 && ('0' + aDig) === qDig);
        if (matchExacto || matchPadStart) {
          return ContentService.createTextOutput(JSON.stringify({
            status:    'success',
            nombre:    String(rows[i][nomIdx]),
            documento: doc,  // siempre devolvemos el doc que el usuario pidió (con su cero)
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
      codigo: 'TOKEN_NO_CONFIGURADO',
      message: 'APISPERU_TOKEN no está en Script Properties. Contacta al admin del sistema.'
    })).setMimeType(ContentService.MimeType.JSON);
  }

  var tipo = doc.length === 11 ? 'ruc' : 'dni';
  var url  = 'https://dniruc.apisperu.com/api/v1/' + tipo + '/' + doc + '?token=' + token;

  // [v2.5.59] Reintentos: 1 retry con 800ms backoff si falla red o 5xx
  function _intentar(esRetry) {
    try {
      var response = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        // Timeout no es configurable en UrlFetchApp pero por defecto ~60s
      });
      var code = response.getResponseCode();
      var body = response.getContentText();
      // 401/403 = token rechazado · 402 = sin saldo · 404 = no encontrado · 5xx = error servidor
      if (code === 401 || code === 403) {
        Logger.log('[APISPeru] Token rechazado HTTP ' + code + ' body=' + body.substring(0, 200));
        return { _ko: true, status: 'error', codigo: 'TOKEN_RECHAZADO', message: 'Token APISPeru inválido o expirado (HTTP ' + code + '). Renovar.' };
      }
      if (code === 402 || /sin saldo|limit|excedid/i.test(body)) {
        Logger.log('[APISPeru] Sin saldo HTTP ' + code + ' body=' + body.substring(0, 200));
        return { _ko: true, status: 'error', codigo: 'SIN_SALDO', message: 'Token APISPeru sin saldo o cuota agotada. Renovar plan.' };
      }
      if (code === 404) {
        return { _ko: true, status: 'not_found', codigo: 'DOC_NO_ENCONTRADO', message: 'Documento ' + doc + ' no figura en ' + (tipo === 'ruc' ? 'SUNAT' : 'RENIEC') + '.' };
      }
      if (code >= 500 && code < 600) {
        Logger.log('[APISPeru] 5xx HTTP ' + code + ' body=' + body.substring(0, 200));
        if (!esRetry) return { _retry: true };
        return { _ko: true, status: 'error', codigo: 'API_5XX', message: 'APISPeru no responde (HTTP ' + code + '). Reintenta en unos segundos.' };
      }
      if (code !== 200) {
        Logger.log('[APISPeru] HTTP inesperado ' + code + ' body=' + body.substring(0, 200));
        return { _ko: true, status: 'error', codigo: 'HTTP_INESPERADO', message: 'APISPeru respondió HTTP ' + code };
      }
      var json;
      try { json = JSON.parse(body); }
      catch(eP) { return { _ko: true, status: 'error', codigo: 'PARSE_FAIL', message: 'APISPeru devolvió texto inválido: ' + body.substring(0, 100) }; }

      // [v2.6.4] Detectar formato nuevo: APISPeru a veces devuelve HTTP 200
      // con {success: false, message: "..."} en vez de HTTP 404 cuando el doc
      // no existe en su base. Sin esto, caía en NOMBRE_VACIO genérico.
      if (json && json.success === false) {
        return { _ko: true, status: 'not_found', codigo: 'DOC_NO_ENCONTRADO',
                 message: 'Documento ' + doc + ' no figura en ' + (tipo === 'ruc' ? 'SUNAT' : 'RENIEC') + (json.message ? ' (' + json.message + ')' : '') };
      }

      var nombre    = '';
      var direccion = '';
      if (tipo === 'dni') {
        // [v2.6.4] Aceptar variantes de nombre de campo (snake_case y camelCase)
        var nombres = json.nombres || json.nombre || json.first_name || '';
        var apePat  = json.apellidoPaterno || json.apellido_paterno || json.paterno || json.last_name || '';
        var apeMat  = json.apellidoMaterno || json.apellido_materno || json.materno || '';
        nombre = [nombres, apePat, apeMat].filter(Boolean).join(' ').trim();
      } else {
        nombre    = (json.razonSocial || json.razon_social || json.nombre || '').trim();
        direccion = (json.direccion   || json.domicilio || '').trim();
      }
      if (!nombre) {
        return { _ko: true, status: 'not_found', codigo: 'NOMBRE_VACIO',
                 message: 'APISPeru no devolvió nombre para ' + doc + ' (revisa /test_apisperu_dni?doc=' + doc + ' para ver respuesta cruda)' };
      }
      return {
        _ok: true,
        status:    'success',
        nombre:    nombre,
        documento: doc,
        tipo:      tipo === 'ruc' ? 'RUC' : 'DNI',
        fuente:    'api',
        direccion: direccion
      };
    } catch(eN) {
      Logger.log('[APISPeru] Network/Exception: ' + eN.message);
      if (!esRetry) return { _retry: true };
      return { _ko: true, status: 'error', codigo: 'NET_ERROR', message: 'Error de red: ' + eN.message };
    }
  }

  var resultado = _intentar(false);
  if (resultado._retry) {
    Utilities.sleep(800);
    resultado = _intentar(true);
  }
  // Limpiar flags internos antes de devolver
  delete resultado._ok; delete resultado._ko; delete resultado._retry;
  return ContentService.createTextOutput(JSON.stringify(resultado)).setMimeType(ContentService.MimeType.JSON);
}

// [v2.5.59] testApiSperu — endpoint diagnóstico para verificar que el token
// está vivo y la API responde. Útil cuando el cajero reporta que "no encuentra".
// Llamar: ?accion=test_apisperu
function testApiSperu() {
  var token = PropertiesService.getScriptProperties().getProperty('APISPERU_TOKEN');
  if (!token) {
    return ContentService.createTextOutput(JSON.stringify({
      ok: false, codigo: 'TOKEN_NO_CONFIGURADO',
      detalle: 'APISPERU_TOKEN no está en Script Properties'
    })).setMimeType(ContentService.MimeType.JSON);
  }
  var tokenTail = token.length > 8 ? '...' + token.substring(token.length - 6) : token;
  // Consulta de prueba con un RUC público conocido (SUNAT mismo)
  var rucPrueba = '20131312955'; // SUNAT
  var url = 'https://dniruc.apisperu.com/api/v1/ruc/' + rucPrueba + '?token=' + token;
  try {
    var t0 = Date.now();
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var ms = Date.now() - t0;
    var code = resp.getResponseCode();
    var body = resp.getContentText();
    var bodySnap = body.substring(0, 250);
    var diag = {
      ok: code === 200,
      tokenTail:   tokenTail,
      httpCode:    code,
      latenciaMs:  ms,
      bodyPreview: bodySnap
    };
    if (code === 200) {
      try {
        var j = JSON.parse(body);
        diag.razonSocial = j.razonSocial || j.nombre || '(sin campo)';
        diag.mensaje     = '✓ API funcionando. Token vivo. Consulta de prueba a SUNAT RUC ' + rucPrueba + ' en ' + ms + 'ms.';
      } catch(_){ diag.mensaje = 'HTTP 200 pero respuesta no es JSON válido.'; }
    } else if (code === 401 || code === 403) {
      diag.codigo  = 'TOKEN_INVALIDO';
      diag.mensaje = '✗ Token rechazado (HTTP ' + code + '). Renovar.';
    } else if (code === 402) {
      diag.codigo  = 'SIN_SALDO';
      diag.mensaje = '✗ Token sin saldo / cuota agotada (HTTP 402). Renovar plan.';
    } else {
      diag.codigo  = 'HTTP_' + code;
      diag.mensaje = '✗ HTTP inesperado ' + code + '. Ver bodyPreview.';
    }
    return ContentService.createTextOutput(JSON.stringify(diag)).setMimeType(ContentService.MimeType.JSON);
  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({
      ok: false, codigo: 'NET_ERROR', tokenTail: tokenTail,
      mensaje: 'Error de red al llamar APISPeru: ' + e.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// [v2.6.4] testApiSperuDoc — devuelve el body CRUDO de APISPeru para
// cualquier DNI/RUC. Útil para diagnosticar por qué un doc específico
// devuelve NOMBRE_VACIO (shape de respuesta inesperado).
// Llamar: ?accion=test_apisperu_doc&doc=46027897
function testApiSperuDoc(doc) {
  doc = String(doc || '').trim();
  if (!/^\d{8}$/.test(doc) && !/^\d{11}$/.test(doc)) {
    return ContentService.createTextOutput(JSON.stringify({
      ok: false, error: 'doc debe ser DNI(8) o RUC(11). Recibido: ' + doc.length + ' chars'
    })).setMimeType(ContentService.MimeType.JSON);
  }
  var token = PropertiesService.getScriptProperties().getProperty('APISPERU_TOKEN');
  if (!token) return ContentService.createTextOutput(JSON.stringify({
    ok: false, error: 'APISPERU_TOKEN no configurado'
  })).setMimeType(ContentService.MimeType.JSON);

  var tipo = doc.length === 11 ? 'ruc' : 'dni';
  var url = 'https://dniruc.apisperu.com/api/v1/' + tipo + '/' + doc + '?token=' + token;
  try {
    var t0 = Date.now();
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var ms = Date.now() - t0;
    var code = resp.getResponseCode();
    var body = resp.getContentText();
    var json = null;
    try { json = JSON.parse(body); } catch(_){}
    return ContentService.createTextOutput(JSON.stringify({
      ok: true, doc: doc, tipo: tipo, httpCode: code, latenciaMs: ms,
      bodyRaw: body.substring(0, 800),
      jsonKeys: json ? Object.keys(json) : null,
      json: json
    }, null, 2)).setMimeType(ContentService.MimeType.JSON);
  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({
      ok: false, error: 'NET_ERROR: ' + e.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
