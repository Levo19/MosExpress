// ============================================================
// MosExpress — Radio.gs
// Endpoints para la pantalla TV (radio.html):
//   - radioConfig()         → lee hoja RadioConfig (playlists, ticker, destacados)
//   - topProductosHoy()     → calcula los SKU más vendidos del día
//   - setupRadioSheet()     → crea la hoja RadioConfig con sample data
//                              (ejecutar UNA VEZ desde el editor de Apps Script)
// ============================================================

// ── radioConfig ────────────────────────────────────────────────
// Devuelve config para la pantalla TV: playlists por horario, mensajes
// del ticker, productos destacados manualmente y settings (intervalos).
// Si la hoja RadioConfig no existe, devuelve defaults razonables.
function radioConfig() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('RadioConfig');

  var defaults = {
    playlists: [
      { rango: [6, 12],  videoId: 'jfKfPfyJRdk', nombre: 'MAÑANA' },
      { rango: [12, 18], videoId: 'jfKfPfyJRdk', nombre: 'TARDE' },
      { rango: [18, 24], videoId: 'jfKfPfyJRdk', nombre: 'NOCHE' },
      { rango: [0, 6],   videoId: 'jfKfPfyJRdk', nombre: 'MADRUGADA' }
    ],
    ticker: [
      '🛒 MOSEXPRESS minimarket — abierto todos los días',
      '💳 PAGA CON YAPE · PLIN · TARJETA',
      '⭐ SÍGUENOS @MOSEXPRESS'
    ],
    destacados: [],
    imagenes: {},
    categorias: {},
    config: { rotarEstrellaSeg: 12, rotarCardsSeg: 7 }
  };

  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'ok', usando_defaults: true,
      playlists: defaults.playlists,
      ticker: defaults.ticker,
      destacados: defaults.destacados,
      imagenes: defaults.imagenes,
      categorias: defaults.categorias,
      config: defaults.config
    })).setMimeType(ContentService.MimeType.JSON);
  }

  var data = sheet.getDataRange().getValues();
  var playlists = [], ticker = [], destacados = [];
  var imagenes = {}, categorias = {};
  var config = { rotarEstrellaSeg: 12, rotarCardsSeg: 7 };

  // Estructura esperada: Tipo | Key | Valor   (header en fila 1)
  for (var i = 1; i < data.length; i++) {
    var tipo    = String(data[i][0] || '').trim().toLowerCase();
    var keyRaw  = data[i][1];
    var valor   = String(data[i][2] || '').trim();
    if (!tipo) continue;

    // Recuperar Key si Sheets lo convirtió a Date (ej. "6-12" → 6/dic en locale es-PE).
    // Extrae month+day del Date y reconstruye "X-Y" — funciona porque el rango
    // siempre es dos números separados por guión.
    var key;
    if (keyRaw instanceof Date) {
      var m = keyRaw.getMonth() + 1;
      var d = keyRaw.getDate();
      var lo = Math.min(m, d), hi = Math.max(m, d);
      key = lo + '-' + hi;
    } else {
      key = String(keyRaw || '').trim();
    }

    if (tipo === 'image') {
      // Key = SKU, Valor = URL (Drive share o cualquier URL directa)
      if (key && valor) imagenes[key] = valor;
      continue;
    }
    if (tipo === 'cat' || tipo === 'categoria') {
      // Key = SKU, Valor = nombre de categoría (bebidas, cerveza, snacks, etc)
      if (key && valor) categorias[key] = valor.toLowerCase();
      continue;
    }
    if (tipo === 'playlist') {
      // Key formato "6-12" o "18-24"
      var m = key.match(/^(\d+)\s*-\s*(\d+)$/);
      if (!m || !valor) continue;
      var inicio = parseInt(m[1], 10);
      var fin    = parseInt(m[2], 10);
      // Valor puede ser solo videoId, o "videoId|Nombre"
      var partes = valor.split('|');
      playlists.push({
        rango:   [inicio, fin],
        videoId: partes[0].trim(),
        nombre:  (partes[1] || '').trim() || _nombrePorHora(inicio)
      });
    } else if (tipo === 'ticker') {
      if (valor) ticker.push(valor);
    } else if (tipo === 'destacado') {
      if (key) destacados.push({ sku: key, prioridad: parseInt(valor, 10) || 99 });
    } else if (tipo === 'config') {
      if (key === 'rotar_estrella_seg') config.rotarEstrellaSeg = parseInt(valor, 10) || 12;
      if (key === 'rotar_cards_seg')    config.rotarCardsSeg    = parseInt(valor, 10) || 7;
    }
  }

  if (!playlists.length) playlists = defaults.playlists;
  if (!ticker.length)    ticker    = defaults.ticker;

  destacados.sort(function(a, b) { return a.prioridad - b.prioridad; });

  return ContentService.createTextOutput(JSON.stringify({
    status: 'ok',
    playlists: playlists,
    ticker: ticker,
    destacados: destacados,
    imagenes: imagenes,
    categorias: categorias,
    config: config
  })).setMimeType(ContentService.MimeType.JSON);
}

function _nombrePorHora(h) {
  if (h >= 5  && h < 12) return 'MAÑANA';
  if (h >= 12 && h < 18) return 'TARDE';
  if (h >= 18 && h < 24) return 'NOCHE';
  return 'MADRUGADA';
}

// ── topProductosHoy ────────────────────────────────────────────
// Devuelve los SKU más vendidos hoy desde VENTAS_CABECERA + VENTAS_DETALLE.
// Si hoy aún no hay ventas, hace fallback a los últimos 7 días.
// Frontend cruza el SKU con el catálogo descargado para obtener nombre/precio/etc.
//
// VENTAS_CABECERA cols: 0=ID_Venta 1=Fecha 6=Total 8=FormaPago 12=Estado_Envio
// VENTAS_DETALLE  cols: 0=ID_Venta 1=SKU 2=Nombre 3=Cantidad 4=Precio 5=Subtotal 6=Cod_Barras
function topProductosHoy() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shCab = ss.getSheetByName('VENTAS_CABECERA');
  var shDet = ss.getSheetByName('VENTAS_DETALLE');
  if (!shCab || !shDet) return generarRespuestaError('Hojas de ventas no encontradas');

  var tz   = Session.getScriptTimeZone();
  var hoy  = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var hace7 = new Date();
  hace7.setDate(hace7.getDate() - 7);
  var hace7Str = Utilities.formatDate(hace7, tz, 'yyyy-MM-dd');

  var cabData = shCab.getDataRange().getValues();
  var ventasHoy = {}, ventas7d = {}, ventasNoAnuladas = {};
  for (var v = 1; v < cabData.length; v++) {
    var fechaRaw = cabData[v][1];
    if (!fechaRaw) continue;
    var fechaStr = (fechaRaw instanceof Date)
      ? Utilities.formatDate(fechaRaw, tz, 'yyyy-MM-dd')
      : String(fechaRaw).substr(0, 10);
    var estado = String(cabData[v][12] || 'COMPLETADO');
    if (estado === 'ANULADO') continue;
    var idVenta = String(cabData[v][0] || '');
    if (!idVenta) continue;
    ventasNoAnuladas[idVenta] = true;
    if (fechaStr === hoy)     ventasHoy[idVenta] = true;
    if (fechaStr >= hace7Str) ventas7d[idVenta]  = true;
  }

  // Agregar cantidades por SKU + recolectar todos los SKUs que esta tienda
  // ha vendido alguna vez (filtro vital para el radio: NO mostrar productos
  // del MOS master que esta tienda no sale en su catálogo activo).
  var detData = shDet.getDataRange().getValues();
  var sumHoy = {}, sumNombreHoy = {};
  var sum7d  = {}, sumNombre7d  = {};
  var skusDeLaTienda = {};
  for (var d = 1; d < detData.length; d++) {
    var idV    = String(detData[d][0] || '');
    var sku    = String(detData[d][1] || '').trim();
    var nombre = String(detData[d][2] || '');
    var qty    = parseFloat(detData[d][3]) || 0;
    if (!sku || qty <= 0) continue;
    if (!ventasNoAnuladas[idV]) continue;
    skusDeLaTienda[sku] = true;
    if (ventasHoy[idV]) {
      sumHoy[sku] = (sumHoy[sku] || 0) + qty;
      sumNombreHoy[sku] = nombre;
    }
    if (ventas7d[idV]) {
      sum7d[sku] = (sum7d[sku] || 0) + qty;
      sumNombre7d[sku] = nombre;
    }
  }

  var hoyArr = Object.keys(sumHoy).map(function(sku) {
    return { sku: sku, nombre: sumNombreHoy[sku], vendidos: sumHoy[sku] };
  }).sort(function(a, b) { return b.vendidos - a.vendidos; });

  var fallback = false;
  if (!hoyArr.length) {
    fallback = true;
    hoyArr = Object.keys(sum7d).map(function(sku) {
      return { sku: sku, nombre: sumNombre7d[sku], vendidos: sum7d[sku] };
    }).sort(function(a, b) { return b.vendidos - a.vendidos; });
  }

  return ContentService.createTextOutput(JSON.stringify({
    status: 'ok',
    fecha:  hoy,
    es_fallback_7d: fallback,
    productos: hoyArr.slice(0, 20),
    skus_de_la_tienda: Object.keys(skusDeLaTienda)
  })).setMimeType(ContentService.MimeType.JSON);
}

// ── setupRadioSheet ────────────────────────────────────────────
// EJECUTAR UNA VEZ desde el editor de Apps Script:
//   1) abre script.google.com
//   2) selecciona la función setupRadioSheet en el dropdown
//   3) click ▶ Ejecutar
// Crea la hoja RadioConfig con headers + ejemplos editables.
function setupRadioSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('RadioConfig');
  if (sh) {
    Logger.log('RadioConfig ya existe — no se sobreescribe.');
    return;
  }
  sh = ss.insertSheet('RadioConfig');
  sh.getRange('A1:C1')
    .setValues([['Tipo', 'Key', 'Valor']])
    .setFontWeight('bold')
    .setBackground('#10b981')
    .setFontColor('#ffffff');

  // Sample rows — el usuario edita estos valores
  var samples = [
    ['playlist', '6-12',  'jfKfPfyJRdk|MAÑANA CHILL'],
    ['playlist', '12-18', 'jfKfPfyJRdk|TARDE POP'],
    ['playlist', '18-24', 'jfKfPfyJRdk|NOCHE FIESTA'],
    ['playlist', '0-6',   'jfKfPfyJRdk|MADRUGADA LO-FI'],
    ['ticker',   '1', '🛒 MOSEXPRESS — abierto todos los días'],
    ['ticker',   '2', '💳 PAGA CON YAPE · PLIN · TARJETA'],
    ['ticker',   '3', '🍺 CERVEZAS HELADAS siempre'],
    ['ticker',   '4', '🥖 PAN RECIÉN HORNEADO en la mañana'],
    ['ticker',   '5', '⭐ SÍGUENOS @MOSEXPRESS'],
    ['destacado', 'SKU_AQUI',  '1'],
    ['image',    'SKU_AQUI',  'https://drive.google.com/file/d/PEGAR_FILE_ID_AQUI/view'],
    ['cat',      'SKU_AQUI',  'bebidas'],
    ['config',   'rotar_estrella_seg', '12'],
    ['config',   'rotar_cards_seg',    '7']
  ];
  // Formato 'texto plano' en col Key — evita que Sheets convierta "6-12" a fecha
  sh.getRange(2, 2, samples.length, 1).setNumberFormat('@');
  sh.getRange(2, 1, samples.length, 3).setValues(samples);

  // Notas explicativas en una columna lateral
  sh.getRange('E1').setValue('GUÍA').setFontWeight('bold');
  sh.getRange('E2').setValue('Tipo "playlist": Key = "6-12" (rango horas), Valor = "videoId|Nombre"');
  sh.getRange('E3').setValue('Tipo "ticker": Key = orden numérico, Valor = mensaje');
  sh.getRange('E4').setValue('Tipo "destacado": Key = SKU del producto, Valor = prioridad (menor = más arriba)');
  sh.getRange('E5').setValue('Tipo "config": rotar_estrella_seg / rotar_cards_seg (segundos)');
  sh.getRange('E6').setValue('videoId de YouTube: lo que va después de v= en la URL');
  sh.getRange('E7').setValue('Tipo "image": Key = SKU del producto, Valor = URL (Drive público o cualquier URL)');
  sh.getRange('E8').setValue('Tipo "cat": Key = SKU, Valor = categoría (bebidas, cerveza, snacks, chocolate, lacteos, panaderia, abarrotes, limpieza, higiene, agua, golosinas, cigarros)');

  sh.setColumnWidth(1, 110);
  sh.setColumnWidth(2, 100);
  sh.setColumnWidth(3, 280);
  sh.setColumnWidth(5, 90);

  Logger.log('✓ Hoja RadioConfig creada con sample data. Edita los valores y la TV los toma.');
}
