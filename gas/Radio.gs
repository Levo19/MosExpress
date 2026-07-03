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
// [CERO-GAS · Migración Sheet→Supabase] Los SKU más vendidos ahora vienen 100%
// de la sombra Supabase vía la RPC me.radio_ventas (SQL 325) — YA NO se lee la
// Hoja VENTAS_CABECERA/DETALLE. Misma forma de salida que antes; el frontend
// (radio.html/radioProductos) sólo usa sku + vendidos + skus_de_la_tienda.
// Sin fallback a la Hoja (directriz cero-GAS): si la RPC falla, devuelve vacío
// y el radio muestra el catálogo sin conteos (degradación limpia, no error).
//
// Reglas espejadas server-side (325): excluye ANULADO/HUERFANA_LIMPIADA, día de
// negocio en TZ Lima, top 20 de hoy (o 7d si hoy vacío), skus 30d (o alguna vez).
function topProductosHoy() {
  var vacio = { status: 'ok', fecha: '', es_fallback_7d: false,
                productos: [], skus_de_la_tienda: [], rango_filtro: '30d' };
  var out = vacio;
  try {
    var r = _sbRpc('me', 'radio_ventas', {});
    if (r && r.ok && r.data && r.data.status === 'ok') out = r.data;
  } catch (e) { out = vacio; }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
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

// ── radioProductos ─────────────────────────────────────────────
// Endpoint TODO-EN-UNO para radio.html. Hace el trabajo pesado server-side
// y devuelve un payload liviano (~40KB vs ~1MB del catálogo completo):
//   { status, productos:[{sku,nombre,precio,vendidos,categoria,img}],
//     playlists, ticker, destacados, config }
// Cachea el resultado 5 min (CacheService) — la cuenta pesada corre 1 sola
// vez cada 5 min aunque haya varias TVs conectadas.
function radioProductos() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('radio_productos_v1');
  if (cached) {
    return ContentService.createTextOutput(cached)
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 1. Catálogo — reusa descargarCatalogo() parseando su salida
  //    (cero riesgo: no toca Catalogo.gs)
  var catalogo = {};
  try {
    var catJson = JSON.parse(descargarCatalogo().getContent());
    catalogo = (catJson && catJson.data) ? catJson.data : {};
  } catch(e) { catalogo = {}; }
  var presentaciones = catalogo.PRESENTACIONES || [];
  var productoBase   = catalogo.PRODUCTO_BASE  || [];

  // 2. Ventas — reusa topProductosHoy()
  var topData = { productos: [], skus_de_la_tienda: [] };
  try { topData = JSON.parse(topProductosHoy().getContent()); } catch(e) {}

  // 3. Config — reusa radioConfig()
  var cfg = { playlists: [], ticker: [], destacados: [], imagenes: {}, categorias: {}, config: {} };
  try { cfg = JSON.parse(radioConfig().getContent()); } catch(e) {}

  // ── Merge ──
  var baseMap = {};
  for (var i = 0; i < productoBase.length; i++) {
    baseMap[productoBase[i].SKU_Base] = productoBase[i];
  }

  // SKUs que vende la tienda (VENTAS_DETALLE guarda el SKU_Base)
  var skusTienda = {};
  var sdt = topData.skus_de_la_tienda || [];
  for (var s = 0; s < sdt.length; s++) skusTienda[sdt[s]] = true;

  // Rollup de vendidos por SKU_Base
  var vendidosPorBase = {};
  var tprods = topData.productos || [];
  for (var t = 0; t < tprods.length; t++) {
    var tk = tprods[t].sku;
    vendidosPorBase[tk] = (vendidosPorBase[tk] || 0) + (tprods[t].vendidos || 0);
  }

  // Canónico por grupo (PRESENTACION con Factor==1, o la primera del grupo)
  var canonicoPorGrupo = {};
  for (var p = 0; p < presentaciones.length; p++) {
    var pr = presentaciones[p];
    var factor = parseFloat(pr.Factor) || 1;
    var ex = canonicoPorGrupo[pr.SKU_Base];
    if (!ex || (factor === 1 && (parseFloat(ex.Factor) || 1) !== 1)) {
      canonicoPorGrupo[pr.SKU_Base] = pr;
    }
  }

  var overrideImg = cfg.imagenes || {};
  var overrideCat = cfg.categorias || {};
  var hayFiltro = sdt.length > 0;

  var productos = [];
  var grupos = Object.keys(canonicoPorGrupo);
  for (var g = 0; g < grupos.length; g++) {
    var skuBase = grupos[g];
    if (hayFiltro && !skusTienda[skuBase]) continue; // solo lo que vende la tienda
    var pres = canonicoPorGrupo[skuBase];
    var base = baseMap[skuBase] || {};
    var nombre = String(base.Nombre || pres.Empaque || pres.SKU || '').toUpperCase();
    var catKey = overrideCat[pres.SKU] || overrideCat[skuBase] || _categorizarRadio(nombre);
    var imgRaw = overrideImg[pres.SKU] || overrideImg[skuBase] || '';
    productos.push({
      sku:       pres.SKU,
      nombre:    nombre,
      precio:    parseFloat(pres.Precio_Venta) || 0,
      vendidos:  vendidosPorBase[skuBase] || 0,
      categoria: catKey,
      img:       _normalizarImgRadio(imgRaw)
    });
  }

  // Ordenar por más vendidos y topear a 300 (lean, cabe en CacheService 100KB)
  productos.sort(function(a, b) { return b.vendidos - a.vendidos; });
  if (productos.length > 300) productos = productos.slice(0, 300);

  var payload = JSON.stringify({
    status: 'ok',
    productos: productos,
    playlists: cfg.playlists || [],
    ticker: cfg.ticker || [],
    destacados: cfg.destacados || [],
    config: cfg.config || { rotarEstrellaSeg: 12, rotarCardsSeg: 7 }
  });

  // Cachear 5 min — si pasara de 100KB, put falla silencioso (sin problema)
  try { cache.put('radio_productos_v1', payload, 300); } catch(e) {}

  return ContentService.createTextOutput(payload)
    .setMimeType(ContentService.MimeType.JSON);
}

// Drive share-link → URL directa (w=600, liviana para TV)
function _normalizarImgRadio(url) {
  url = String(url || '');
  if (!url) return '';
  var m = url.match(/drive\.google\.com\/file\/d\/([^\/]+)/);
  if (!m) m = url.match(/[?&]id=([^&]+)/);
  if (m) return 'https://lh3.googleusercontent.com/d/' + m[1] + '=w600';
  return url;
}

// Categorización por keyword (corre server-side, el cliente solo recibe el string)
function _categorizarRadio(nombre) {
  var n = String(nombre || '');
  var K = [
    [/cerveza|pilsen|cristal|cusque|corona|heineken|backus/i, 'cerveza'],
    [/coca\s?cola|inca\s?kola|sprite|fanta|gaseosa|kola|pepsi|7\s?up/i, 'bebidas'],
    [/agua|san luis|cielo|san mateo/i, 'agua'],
    [/jugo|nectar|n[eé]ctar|frugos|tampico/i, 'bebidas'],
    [/\bsal\b|pimienta|comino|paprika|achiote|oregano|or[eé]gano|canela|sazonador|condimento|sazon|saz[oó]n/i, 'especerias'],
    [/aceite|oliva|vinagre|aderezo/i, 'aderezos'],
    [/salsa|ketchup|mayonesa|mostaza|\baji\b|aj[ií]|sillao|soya|huancaina/i, 'salsas'],
    [/atun|at[uú]n|conserva|menestra|frijol|lenteja|garbanzo|enlatado/i, 'conservas'],
    [/chocolate|sublime|princesa|cua\s?cua|hershey|kit\s?kat|cocoa/i, 'chocolate'],
    [/galleta|casino|soda field|\bfield\b|oreo|morochas|margarita|wafer/i, 'galletas'],
    [/papita|chizito|cheese|\blay|pringles|piqueo|doritos|cheetos|cancha/i, 'snacks'],
    [/caramelo|chupetin|chupet[ií]n|chicle|halls|mentos|gomas/i, 'golosinas'],
    [/leche|gloria|laive|pura vida|yogur|queso|mantequilla/i, 'lacteos'],
    [/\bpan\b|bagueta|tostada|paneton|panet[oó]n/i, 'panaderia'],
    [/arroz|fideo|azucar|az[uú]car|harina|avena|quinua|qu[ií]nua|kiwicha/i, 'abarrotes'],
    [/detergente|jabon liquid|lejia|lej[ií]a|sapolio|ariel|bolivar|bol[ií]var|lavavajilla/i, 'limpieza'],
    [/papel higi|higienico|higi[eé]nico|toalla|servilleta|kotex|pa[nñ]al|pampers|\bjabon\b|shampoo|crema dental/i, 'higiene'],
    [/cigarro|tabaco|hamilton|lucky|marlboro|winston|caribe/i, 'cigarros'],
    [/helado|donofrio|cassata/i, 'helados'],
    [/huevo/i, 'abarrotes'],
    [/ajinomoto|glutamato/i, 'especerias']
  ];
  for (var i = 0; i < K.length; i++) {
    if (K[i][0].test(n)) return K[i][1];
  }
  return 'default';
}
