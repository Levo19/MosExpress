// ============================================================
// MosExpress — ReporteCierre.gs
// Sirve el HTML de cierre de caja vía doGet accion=ver_cierre
// No almacena nada — genera el reporte en vivo al abrirlo.
// ============================================================

function getCierreHtml(idCaja) {
  if (!idCaja) return HtmlService.createHtmlOutput('<h2 style="font-family:sans-serif;padding:40px">ID de caja requerido</h2>');

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── 1. Leer CAJAS ────────────────────────────────────────────
  var caja = null;
  var cajasSheet = ss.getSheetByName('CAJAS');
  if (cajasSheet) {
    var cd = cajasSheet.getDataRange().getValues();
    for (var i = 1; i < cd.length; i++) {
      if (String(cd[i][0]) === String(idCaja)) {
        caja = {
          id:            String(cd[i][0]),
          vendedor:      String(cd[i][1] || ''),
          estacion:      String(cd[i][2] || ''),
          fechaApertura: cd[i][3] instanceof Date ? Utilities.formatDate(cd[i][3], Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : String(cd[i][3] || ''),
          montoInicial:  parseFloat(cd[i][4]) || 0,
          estado:        String(cd[i][5] || ''),
          montoFinal:    parseFloat(cd[i][6]) || 0,
          fechaCierre:   cd[i][7] instanceof Date ? Utilities.formatDate(cd[i][7], Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : String(cd[i][7] || ''),
          zona:          String(cd[i][8] || '')
        };
        break;
      }
    }
  }
  if (!caja) return HtmlService.createHtmlOutput('<h2 style="font-family:sans-serif;padding:40px">Caja "' + idCaja + '" no encontrada.</h2>');

  // ── 2. Leer VENTAS_CABECERA ──────────────────────────────────
  var ventas = [];
  var ventasSheet = ss.getSheetByName('VENTAS_CABECERA');
  if (ventasSheet) {
    var vd = ventasSheet.getDataRange().getValues();
    for (var v = 1; v < vd.length; v++) {
      if (String(vd[v][10]) === String(idCaja)) {
        ventas.push({
          idVenta:     String(vd[v][0]),
          fechaISO:    vd[v][1] instanceof Date ? Utilities.formatDate(vd[v][1], Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : String(vd[v][1] || ''),
          hora:        vd[v][1] instanceof Date ? vd[v][1].getHours() : -1,
          clienteDoc:  String(vd[v][4] || ''),
          clienteNom:  String(vd[v][5] || ''),
          total:       parseFloat(vd[v][6]) || 0,
          tipoDoc:     String(vd[v][7] || 'NOTA_DE_VENTA'),
          metodo:      String(vd[v][8] || 'EFECTIVO'),
          correlativo: String(vd[v][9] || ''),
          estado:      String(vd[v][12] || 'COMPLETADO'),
          obs:         String(vd[v][14] || '')
        });
      }
    }
  }

  // ── 3. Leer VENTAS_DETALLE ───────────────────────────────────
  var detMap = {};
  var prodTotales = {};
  var detSheet = ss.getSheetByName('VENTAS_DETALLE');
  if (detSheet && ventas.length > 0) {
    var vtIds = {};
    ventas.forEach(function(x){ vtIds[x.idVenta] = true; });
    var dd = detSheet.getDataRange().getValues();
    for (var d = 1; d < dd.length; d++) {
      var dvId = String(dd[d][0]);
      if (!vtIds[dvId]) continue;
      if (!detMap[dvId]) detMap[dvId] = [];
      var item = {
        nombre:   String(dd[d][2] || ''),
        cantidad: parseFloat(dd[d][3]) || 0,
        precio:   parseFloat(dd[d][4]) || 0,
        subtotal: parseFloat(dd[d][5]) || 0
      };
      detMap[dvId].push(item);
      var venta = ventas.filter(function(x){ return x.idVenta === dvId; })[0];
      if (venta && venta.estado !== 'ANULADO') {
        var pn = item.nombre || 'Sin nombre';
        prodTotales[pn] = (prodTotales[pn] || 0) + item.subtotal;
      }
    }
  }

  // ── 4. Leer MOVIMIENTOS_EXTRA ────────────────────────────────
  var extras = [];
  var extSheet = ss.getSheetByName('MOVIMIENTOS_EXTRA');
  if (extSheet) {
    var ed = extSheet.getDataRange().getValues();
    for (var e2 = 1; e2 < ed.length; e2++) {
      if (String(ed[e2][1]) === String(idCaja)) {
        extras.push({
          tipo:     String(ed[e2][3] || 'EGRESO'),
          monto:    parseFloat(ed[e2][4]) || 0,
          concepto: String(ed[e2][5] || ''),
          hora:     ed[e2][2] instanceof Date ? Utilities.formatDate(ed[e2][2], Session.getScriptTimeZone(), 'HH:mm') : ''
        });
      }
    }
  }

  // ── 5. Calcular analítica ────────────────────────────────────
  var activas    = ventas.filter(function(x){ return x.estado !== 'ANULADO'; });
  var anuladas   = ventas.filter(function(x){ return x.estado === 'ANULADO'; });
  var sinCobrar  = activas.filter(function(x){ return x.metodo === 'POR_COBRAR'; });
  var creditos   = activas.filter(function(x){ return x.metodo === 'CREDITO'; });
  var cobradas   = activas.filter(function(x){ return x.metodo !== 'POR_COBRAR'; });

  var totalVentas = cobradas.reduce(function(a,x){ return a+x.total; }, 0);

  // Totales por grupo: MIXTO se divide en su parte efectivo (EFE) y parte virtual (VIR)
  var totalEfectivo = 0;
  var totalVirtual  = 0;
  cobradas.filter(function(x){ return x.metodo !== 'CREDITO'; }).forEach(function(x) {
    if (x.metodo === 'EFECTIVO') {
      totalEfectivo += x.total;
    } else if (x.metodo.indexOf('MIXTO') === 0) {
      var efeM = x.metodo.match(/EFE:([\d.]+)/i);
      var virM = x.metodo.match(/VIR:([\d.]+)/i);
      var efe  = efeM ? parseFloat(efeM[1]) : 0;
      var vir  = virM ? parseFloat(virM[1]) : x.total - efe;
      totalEfectivo += efe;
      totalVirtual  += vir;
    } else {
      totalVirtual += x.total;
    }
  });
  var totalCredito = creditos.reduce(function(s,x){ return s+x.total; }, 0);
  var totalOtros   = totalVirtual + totalCredito;

  var totalEntradas = extras.filter(function(x){return x.tipo==='INGRESO';}).reduce(function(a,x){return a+x.monto;},0);
  var totalSalidas  = extras.filter(function(x){return x.tipo==='EGRESO'; }).reduce(function(a,x){return a+x.monto;},0);

  // Efectivo esperado en caja = calculado directo desde Sheets (fuente autoritativa)
  var efectivoEsperado = caja.montoInicial + totalEfectivo + totalEntradas - totalSalidas;

  // Por método pago (solo cobradas)
  var byMetodo = {};
  cobradas.forEach(function(x){
    if (!byMetodo[x.metodo]) byMetodo[x.metodo] = { count:0, total:0 };
    byMetodo[x.metodo].count++;
    byMetodo[x.metodo].total += x.total;
  });

  // Por tipo doc (excluir CREDITO para no duplicar con sección dedicada)
  var byDoc = {};
  var docLabels = { 'NOTA_DE_VENTA':'Nota de Venta','BOLETA':'Boleta','FACTURA':'Factura' };
  cobradas.filter(function(x){ return x.metodo !== 'CREDITO'; }).forEach(function(x){
    var td = x.tipoDoc;
    if (!byDoc[td]) byDoc[td] = { count:0, total:0 };
    byDoc[td].count++;
    byDoc[td].total += x.total;
  });

  // Por hora (ventas activas)
  var byHora = {};
  activas.forEach(function(x){
    if (x.hora >= 0) { byHora[x.hora] = (byHora[x.hora]||0) + x.total; }
  });
  var horasArr = [], totalHoraArr = [];
  for (var h = 0; h <= 23; h++) {
    if (byHora[h] !== undefined) { horasArr.push(h+':00'); totalHoraArr.push(byHora[h].toFixed(2)); }
  }

  // Top 8 productos
  var prodPares = Object.keys(prodTotales).map(function(k){ return [k, prodTotales[k]]; });
  prodPares.sort(function(a,b){ return b[1]-a[1]; });
  var topProdNom = prodPares.slice(0,8).map(function(p){ return p[0].length>22?p[0].substring(0,22)+'…':p[0]; });
  var topProdVal = prodPares.slice(0,8).map(function(p){ return p[1].toFixed(2); });

  // Chart data
  var metodoKeys = Object.keys(byMetodo);
  var metodoVals = metodoKeys.map(function(k){ return byMetodo[k].total.toFixed(2); });
  var docKeys    = Object.keys(byDoc);
  var docVals    = docKeys.map(function(k){ return byDoc[k].total.toFixed(2); });
  var docLabArr  = docKeys.map(function(k){ return docLabels[k]||k; });

  // virtualMets: métodos no-efectivo no-crédito (para el chart y conteo de tickets)
  var METODOS_NO_VIRTUAL = { 'EFECTIVO': true, 'CREDITO': true };
  var virtualMets  = Object.keys(byMetodo).filter(function(k){ return !METODOS_NO_VIRTUAL[k]; });
  // totalVirtual ya calculado arriba con MIXTO correctamente dividido (no se usa byMetodo aquí)

  var fm  = function(n){ return 'S/ ' + parseFloat(n||0).toFixed(2); };
  var pct = function(n,t){ return t===0?'0':((n/t)*100).toFixed(1); };

  // ── 6. Generar HTML ──────────────────────────────────────────
  var H = [];

  H.push('<!DOCTYPE html><html lang="es"><head>');
  H.push('<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">');
  H.push('<title>Cierre — ' + caja.vendedor + ' ' + caja.fechaCierre + '</title>');
  H.push('<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>');
  H.push('<style>');
  H.push('*{box-sizing:border-box;margin:0;padding:0}');
  H.push('body{font-family:system-ui,-apple-system,sans-serif;background:#f1f5f9;color:#1e293b;font-size:14px}');
  H.push('.wrap{max-width:1100px;margin:0 auto;padding:24px 16px}');

  // Header
  H.push('.header{background:linear-gradient(135deg,#4f46e5,#7c3aed);color:#fff;border-radius:16px;padding:24px 28px;margin-bottom:20px}');
  H.push('.header h1{font-size:22px;font-weight:700;margin-bottom:4px}');
  H.push('.header .sub{opacity:.85;font-size:13px}');
  H.push('.header .meta{display:flex;flex-wrap:wrap;gap:16px;margin-top:14px;font-size:13px}');
  H.push('.header .meta span{background:rgba(255,255,255,.15);border-radius:8px;padding:4px 12px}');

  // Arqueo (nuevo — reemplaza kpi-grid)
  H.push('.arqueo-wrap{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:16px}');
  H.push('.arqueo-panel{background:#fff;border-radius:14px;border:1px solid #e2e8f0;padding:20px}');
  H.push('.arqueo-title{font-size:11px;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:.08em;margin-bottom:14px;display:flex;align-items:center;gap:6px}');
  H.push('.arqueo-row{display:flex;justify-content:space-between;align-items:center;padding:5px 0;font-size:13px;border-bottom:1px dashed #f1f5f9}');
  H.push('.arqueo-row:last-child{border:none}');
  H.push('.arqueo-row .lbl{color:#64748b}');
  H.push('.arqueo-row .op{font-weight:700;font-size:12px;width:16px;color:#94a3b8;flex-shrink:0}');
  H.push('.arqueo-row .amt{font-weight:600;color:#1e293b}');
  H.push('.arqueo-row .amt.pos{color:#16a34a}.arqueo-row .amt.neg{color:#dc2626}');
  H.push('.arqueo-divider{border:none;border-top:2px solid #e2e8f0;margin:12px 0 0}');
  H.push('.arqueo-total{text-align:center;padding:14px 0 4px}');
  H.push('.arqueo-total .lbl{font-size:11px;font-weight:600;color:#64748b;text-transform:uppercase;letter-spacing:.07em;margin-bottom:6px}');
  H.push('.arqueo-total .amt{font-size:28px;font-weight:800;color:#16a34a;letter-spacing:-.5px}');
  // Panel otros: lista limpia de 3 ítems
  H.push('.otros-item{display:flex;align-items:center;justify-content:space-between;padding:11px 0;border-bottom:1px solid #f1f5f9}');
  H.push('.otros-item:last-child{border:none}');
  H.push('.otros-item .oi-lbl{font-size:13px;color:#475569;font-weight:500}');
  H.push('.otros-item .oi-cnt{font-size:11px;color:#94a3b8;margin-left:6px}');
  H.push('.otros-item .oi-amt{font-size:15px;font-weight:700;color:#475569}');
  H.push('.otros-item.virt .oi-lbl{color:#1d4ed8;font-weight:700}');
  H.push('.otros-item.virt .oi-amt{font-size:20px;font-weight:800;color:#1d4ed8}');
  H.push('.stats-row{display:flex;flex-wrap:wrap;gap:8px;margin-top:14px;padding-top:14px;border-top:1px solid #f1f5f9}');
  H.push('.stat-chip{background:#f8fafc;border:1px solid #e2e8f0;border-radius:20px;padding:4px 12px;font-size:12px;font-weight:600;color:#475569;display:flex;align-items:center;gap:5px}');
  H.push('.stat-chip .n{font-size:16px;font-weight:800;color:#1e293b}');

  // Diff box
  H.push('.diff-box{border-radius:12px;padding:14px 20px;font-size:16px;font-weight:700;margin-bottom:20px;display:flex;align-items:center;gap:10px}');
  H.push('.diff-ok{background:#f0fdf4;color:#15803d;border:1px solid #bbf7d0}');
  H.push('.diff-neg{background:#fef2f2;color:#dc2626;border:1px solid #fecaca}');
  H.push('.diff-eq{background:#f8fafc;color:#475569;border:1px solid #e2e8f0}');

  // Charts — grilla fija 2×2
  H.push('.charts-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:18px;margin-bottom:22px}');
  H.push('.card{background:#fff;border-radius:14px;border:1px solid #e2e8f0;padding:20px}');
  H.push('.card h3{font-size:12px;font-weight:700;color:#475569;text-transform:uppercase;letter-spacing:.06em;margin-bottom:16px}');
  H.push('.chart-wrap{position:relative;height:240px}');
  H.push('.chart-legend{margin-top:14px;display:flex;flex-direction:column;gap:4px}');
  H.push('.chart-legend-row{display:flex;justify-content:space-between;font-size:12px;padding:3px 0;border-bottom:1px solid #f8fafc;color:#475569}');
  H.push('.chart-legend-row:last-child{border:none}');
  H.push('.chart-legend-row span:last-child{font-weight:600;color:#1e293b}');

  // Secciones
  H.push('.section{background:#fff;border-radius:12px;border:1px solid #e2e8f0;margin-bottom:14px;overflow:hidden}');
  H.push('.sec-head{display:flex;align-items:center;justify-content:space-between;padding:14px 18px;cursor:pointer;user-select:none;border-bottom:1px solid #f1f5f9}');
  H.push('.sec-head:hover{background:#f8fafc}');
  H.push('.sec-title{font-weight:600;font-size:14px;display:flex;align-items:center;gap:8px}');
  H.push('.sec-badge{background:#ede9fe;color:#6d28d9;border-radius:20px;padding:2px 10px;font-size:12px;font-weight:600}');
  H.push('.sec-total{font-weight:700;font-size:15px;color:#1e293b}');
  H.push('.sec-body{display:none;padding:0}');
  H.push('.sec-body.open{display:block}');

  // Tabla
  H.push('.tbl{width:100%;border-collapse:collapse;font-size:13px}');
  H.push('.tbl th{padding:10px 14px;text-align:left;font-size:11px;font-weight:600;color:#64748b;text-transform:uppercase;letter-spacing:.05em;border-bottom:2px solid #f1f5f9;background:#f8fafc;white-space:nowrap}');
  H.push('.tbl td{padding:9px 14px;border-bottom:1px solid #f8fafc;vertical-align:middle}');
  H.push('.tbl tr.det-row{cursor:pointer}.tbl tr.det-row:hover td{background:#f8fafc}');
  H.push('.tbl tr.items-row td{background:#f8fafc;padding:0}.items-inner{padding:10px 14px 10px 40px}');
  H.push('.tbl tr.items-row{display:none}.tbl tr.items-row.open{display:table-row}');
  H.push('.item-line{display:flex;justify-content:space-between;font-size:12px;color:#475569;padding:3px 0;border-bottom:1px dashed #e2e8f0}');
  H.push('.item-line:last-child{border:none}');

  // Totales por cobro al pie de sección
  H.push('.cobro-totals{display:flex;flex-wrap:wrap;gap:8px;padding:12px 16px;background:#f8fafc;border-top:1px solid #f1f5f9}');
  H.push('.cobro-chip{display:flex;align-items:center;gap:6px;background:#fff;border:1px solid #e2e8f0;border-radius:8px;padding:6px 12px;font-size:12px}');
  H.push('.cobro-chip .cc-lbl{color:#64748b;font-weight:500}');
  H.push('.cobro-chip .cc-amt{font-weight:700;color:#1e293b}');
  H.push('.cobro-chip .cc-cnt{color:#94a3b8;font-size:11px}');

  // Badges
  H.push('.badge{display:inline-block;padding:2px 9px;border-radius:20px;font-size:11px;font-weight:700;letter-spacing:.02em}');
  H.push('.b-nv{background:#eff6ff;color:#2563eb}.b-bo{background:#f0fdf4;color:#15803d}');
  H.push('.b-fa{background:#fff7ed;color:#c2410c}.b-an{background:#fef2f2;color:#dc2626}');
  H.push('.b-cr{background:#fdf4ff;color:#7e22ce}.b-pc{background:#fffbeb;color:#b45309}');
  H.push('.b-ef{background:#dcfce7;color:#15803d}.b-tr{background:#eff6ff;color:#1d4ed8}');
  H.push('.b-yp{background:#faf5ff;color:#7e22ce}.b-vi{background:#ecfdf5;color:#065f46}');

  H.push('.arrow{transition:transform .2s;display:inline-block}');
  H.push('.empty{text-align:center;color:#94a3b8;padding:28px;font-size:13px}');
  H.push('.footer{text-align:center;color:#94a3b8;font-size:12px;padding:20px 0 8px}');
  H.push('@media print{.sec-body{display:block!important}.no-print{display:none!important}}');
  H.push('@media(max-width:600px){.charts-grid{grid-template-columns:1fr}.arqueo-wrap{grid-template-columns:1fr}}');
  H.push('</style></head><body>');

  H.push('<div class="wrap">');

  // ── Header ──────────────────────────────────────────────────
  H.push('<div class="header">');
  H.push('<div style="display:flex;align-items:flex-start;justify-content:space-between;flex-wrap:wrap;gap:8px">');
  H.push('<div><h1>Cierre de Turno</h1><div class="sub">MOSexpress · Reporte generado el ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') + '</div></div>');
  H.push('<button onclick="window.print()" class="no-print" style="background:rgba(255,255,255,.2);border:none;color:#fff;border-radius:8px;padding:8px 14px;cursor:pointer;font-size:13px">🖨 Imprimir</button>');
  H.push('</div>');
  H.push('<div class="meta">');
  H.push('<span>👤 ' + caja.vendedor + '</span>');
  H.push('<span>🏪 ' + (caja.zona || caja.estacion || '—') + '</span>');
  H.push('<span>🟢 Apertura: ' + caja.fechaApertura + '</span>');
  H.push('<span>🔴 Cierre: ' + (caja.fechaCierre || '—') + '</span>');
  H.push('</div></div>');

  // ── Arqueo de caja (nuevo diseño) ────────────────────────────
  H.push('<div class="arqueo-wrap">');

  // Panel izquierdo: flujo de efectivo
  H.push('<div class="arqueo-panel">');
  H.push('<div class="arqueo-title">💵 Arqueo físico</div>');
  H.push('<div class="arqueo-row"><span class="op"> </span><span class="lbl">Monto inicial</span><span class="amt">' + fm(caja.montoInicial) + '</span></div>');
  H.push('<div class="arqueo-row"><span class="op pos">+</span><span class="lbl">Ventas efectivo</span><span class="amt pos">' + fm(totalEfectivo) + '</span></div>');
  if (totalEntradas > 0) {
    H.push('<div class="arqueo-row"><span class="op pos">+</span><span class="lbl">Ingresos extra</span><span class="amt pos">' + fm(totalEntradas) + '</span></div>');
  }
  if (totalSalidas > 0) {
    H.push('<div class="arqueo-row"><span class="op neg">−</span><span class="lbl">Salidas extra</span><span class="amt neg">' + fm(totalSalidas) + '</span></div>');
  }
  H.push('<hr class="arqueo-divider">');
  H.push('<div class="arqueo-total">');
  H.push('<div class="lbl">= Esperado en caja</div>');
  H.push('<div class="amt">' + fm(efectivoEsperado) + '</div>');
  H.push('</div>');
  H.push('</div>'); // panel izquierdo

  // Panel derecho: OTROS — listado limpio de 3 ítems
  H.push('<div class="arqueo-panel">');
  H.push('<div class="arqueo-title">Otros</div>');

  // 1. Total virtual — resaltado en azul (contar tickets que tienen parte virtual)
  var cntVirtual = cobradas.filter(function(x){
    return x.metodo !== 'EFECTIVO' && x.metodo !== 'CREDITO';
  }).length;
  H.push('<div class="otros-item virt">');
  H.push('<div><span class="oi-lbl">Total virtual</span><span class="oi-cnt">(' + cntVirtual + ' tickets)</span></div>');
  H.push('<span class="oi-amt">' + fm(totalVirtual) + '</span>');
  H.push('</div>');

  // 2. Total anulados
  var totalAnulados = anuladas.reduce(function(s,x){ return s + x.total; }, 0);
  H.push('<div class="otros-item">');
  H.push('<div><span class="oi-lbl">Total anulados</span><span class="oi-cnt">(' + anuladas.length + ' tickets)</span></div>');
  H.push('<span class="oi-amt">' + fm(totalAnulados) + '</span>');
  H.push('</div>');

  // 3. Total crédito
  var cntCredito   = creditos.length;
  var totalCredito = creditos.reduce(function(s,x){ return s + x.total; }, 0);
  H.push('<div class="otros-item">');
  H.push('<div><span class="oi-lbl">Total crédito</span><span class="oi-cnt">(' + cntCredito + ' tickets)</span></div>');
  H.push('<span class="oi-amt">' + fm(totalCredito) + '</span>');
  H.push('</div>');

  // Chips de estadísticas generales
  H.push('<div class="stats-row">');
  H.push('<div class="stat-chip"><span class="n">' + cobradas.length + '</span> tickets cobrados</div>');
  if (sinCobrar.length > 0) {
    H.push('<div class="stat-chip" style="border-color:#fde68a;color:#b45309"><span class="n" style="color:#b45309">' + sinCobrar.length + '</span> por cobrar</div>');
  }
  H.push('</div>');
  H.push('</div>'); // panel derecho

  H.push('</div>'); // arqueo-wrap

  // ── Resumen de caja ──────────────────────────────────────────
  // El efectivo esperado se calcula siempre desde Sheets (fuente autoritativa).
  // No existe "declarado" ni "diferencia" — el cajero nunca ingresa un conteo físico.
  H.push('<div class="diff-box diff-eq">✅ <span>Efectivo esperado en caja: ' + fm(efectivoEsperado) + '</span>');
  H.push('<span style="font-size:13px;font-weight:400;margin-left:auto;opacity:.8">Virtual: ' + fm(totalVirtual) + ' · Crédito: ' + fm(totalCredito) + '</span>');
  H.push('</div>');

  // ── Charts 2×2 ───────────────────────────────────────────────
  H.push('<div class="charts-grid">');

  // Donut: métodos de pago
  H.push('<div class="card"><h3>Métodos de Pago</h3><div class="chart-wrap"><canvas id="chartMetodo"></canvas></div>');
  H.push('<div class="chart-legend">');
  metodoKeys.forEach(function(k){
    var m = byMetodo[k];
    H.push('<div class="chart-legend-row"><span>' + k + ' (' + m.count + ')</span><span>' + fm(m.total) + '</span></div>');
  });
  H.push('</div></div>');

  // Donut: tipo documento
  H.push('<div class="card"><h3>Por Tipo de Documento</h3><div class="chart-wrap"><canvas id="chartDoc"></canvas></div>');
  H.push('<div class="chart-legend">');
  docKeys.forEach(function(k){
    var bd = byDoc[k];
    H.push('<div class="chart-legend-row"><span>' + (docLabels[k]||k) + ' (' + bd.count + ')</span><span>' + fm(bd.total) + '</span></div>');
  });
  H.push('</div></div>');

  // Bar: ventas por hora
  if (horasArr.length > 0) {
    H.push('<div class="card"><h3>Ventas por Hora</h3><div class="chart-wrap"><canvas id="chartHoras"></canvas></div></div>');
  } else {
    H.push('<div class="card" style="display:flex;align-items:center;justify-content:center"><div class="empty">Sin datos por hora</div></div>');
  }

  // Bar horizontal: top productos
  if (topProdNom.length > 0) {
    H.push('<div class="card"><h3>Top Productos</h3><div class="chart-wrap"><canvas id="chartProds"></canvas></div></div>');
  } else {
    H.push('<div class="card" style="display:flex;align-items:center;justify-content:center"><div class="empty">Sin productos registrados</div></div>');
  }

  H.push('</div>'); // charts-grid

  // ── Secciones de tickets por tipo ─────────────────────────────
  var tiposOrden = ['NOTA_DE_VENTA','BOLETA','FACTURA'];
  tiposOrden.forEach(function(td) {
    var grupo = cobradas.filter(function(x){ return x.tipoDoc === td && x.metodo !== 'CREDITO'; });
    if (grupo.length === 0) return;
    var totalGrupo = grupo.reduce(function(a,x){return a+x.total;},0);
    var label = docLabels[td] || td;
    var badgeCls = td==='BOLETA'?'b-bo':td==='FACTURA'?'b-fa':'b-nv';
    H.push(_section('sec-' + td, '📄 ' + label, grupo.length, fm(totalGrupo), badgeCls,
      _tablaVentas(grupo, detMap, badgeCls)));
  });

  // ── Anulados ──────────────────────────────────────────────────
  if (anuladas.length > 0) {
    var totalAn = anuladas.reduce(function(a,x){return a+x.total;},0);
    H.push(_section('sec-anulados', '❌ Anulados', anuladas.length, fm(totalAn), 'b-an',
      _tablaVentas(anuladas, detMap, 'b-an')));
  }

  // ── Sin cobrar ────────────────────────────────────────────────
  if (sinCobrar.length > 0) {
    var totalSC = sinCobrar.reduce(function(a,x){return a+x.total;},0);
    H.push(_section('sec-sincobrar', '⏳ Sin Cobrar', sinCobrar.length, fm(totalSC), 'b-pc',
      _tablaVentas(sinCobrar, detMap, 'b-pc')));
  }

  // ── Créditos ──────────────────────────────────────────────────
  if (creditos.length > 0) {
    var totalCr = creditos.reduce(function(a,x){return a+x.total;},0);
    H.push(_section('sec-creditos', '🔁 Créditos', creditos.length, fm(totalCr), 'b-cr',
      _tablaVentas(creditos, detMap, 'b-cr')));
  }

  // ── Movimientos extra ─────────────────────────────────────────
  if (extras.length > 0) {
    var extHtml = ['<table class="tbl"><thead><tr><th>Hora</th><th>Tipo</th><th>Concepto</th><th>Monto</th></tr></thead><tbody>'];
    extras.forEach(function(e3){
      var cls = e3.tipo==='INGRESO' ? 'b-ef' : 'b-an';
      extHtml.push('<tr><td style="color:#94a3b8">' + (e3.hora||'—') + '</td><td><span class="badge ' + cls + '">' + e3.tipo + '</span></td><td>' + e3.concepto + '</td><td style="font-weight:600">' + fm(e3.monto) + '</td></tr>');
    });
    extHtml.push('</tbody></table>');
    H.push(_section('sec-extras', '💰 Movimientos Extra', extras.length,
      fm(totalEntradas - totalSalidas), '', extHtml.join('')));
  }

  // ── Footer ────────────────────────────────────────────────────
  H.push('<div class="footer">MOSexpress · Cierre ID: ' + idCaja + ' · ' + caja.vendedor + ' · ' + caja.fechaCierre + '</div>');
  H.push('</div>'); // wrap

  // ── Chart.js scripts ──────────────────────────────────────────
  var metColors  = ['#4f46e5','#0ea5e9','#f59e0b','#ef4444','#10b981','#8b5cf6','#ec4899'];
  H.push('<script>');
  H.push('var metKeys=' + JSON.stringify(metodoKeys) + ';');
  H.push('var metVals=' + JSON.stringify(metodoVals) + ';');
  H.push('var docKeys=' + JSON.stringify(docLabArr) + ';');
  H.push('var docVals=' + JSON.stringify(docVals) + ';');
  H.push('var horasArr=' + JSON.stringify(horasArr) + ';');
  H.push('var horasVals=' + JSON.stringify(totalHoraArr) + ';');
  H.push('var topNom=' + JSON.stringify(topProdNom) + ';');
  H.push('var topVal=' + JSON.stringify(topProdVal) + ';');
  H.push('var COLS=' + JSON.stringify(metColors) + ';');
  H.push('function mkDoughnut(id,labels,vals){var c=document.getElementById(id);if(!c)return;new Chart(c,{type:"doughnut",data:{labels:labels,datasets:[{data:vals,backgroundColor:COLS.slice(0,labels.length),borderWidth:2,borderColor:"#fff"}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:"bottom",labels:{font:{size:11},padding:8,boxWidth:12}}}}})}');
  H.push('function mkBar(id,labels,vals,horiz){var c=document.getElementById(id);if(!c)return;new Chart(c,{type:"bar",data:{labels:labels,datasets:[{data:vals,backgroundColor:horiz?"#6366f1":"#4f46e5",borderRadius:4}]},options:{indexAxis:horiz?"y":"x",responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{ticks:{font:{size:11},color:"#64748b"},grid:{color:"#f1f5f9"}},y:{ticks:{font:{size:11},color:"#64748b"},grid:{color:"#f1f5f9"}}}}})}');
  H.push('mkDoughnut("chartMetodo",metKeys,metVals);');
  H.push('mkDoughnut("chartDoc",docKeys,docVals);');
  H.push('if(document.getElementById("chartHoras"))mkBar("chartHoras",horasArr,horasVals,false);');
  H.push('if(document.getElementById("chartProds"))mkBar("chartProds",topNom,topVal,true);');
  H.push('document.querySelectorAll(".sec-head").forEach(function(h){h.addEventListener("click",function(){var b=h.nextElementSibling;var open=b.classList.toggle("open");h.querySelector(".arrow").style.transform=open?"rotate(90deg)":""})});');
  H.push('document.querySelectorAll(".det-row").forEach(function(r){r.addEventListener("click",function(){var nr=r.nextElementSibling;if(nr&&nr.classList.contains("items-row"))nr.classList.toggle("open")})});');
  H.push('<\/script>');
  H.push('</body></html>');

  return HtmlService.createHtmlOutput(H.join(''))
    .setTitle('Cierre — ' + caja.vendedor)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Helpers HTML ─────────────────────────────────────────────

function _kpi(label, val, note, cls) {
  return '<div class="kpi"><div class="label">' + label + '</div><div class="val ' + (cls||'') + '">' + val + '</div><div class="note">' + (note||'') + '</div></div>';
}

function _section(id, title, count, total, badgeCls, body) {
  return '<div class="section">' +
    '<div class="sec-head" id="h-' + id + '">' +
      '<div class="sec-title"><span class="arrow" style="color:#94a3b8">▶</span>' + title +
        (count > 0 ? ' <span class="sec-badge">' + count + '</span>' : '') +
      '</div>' +
      '<div class="sec-total">' + total + '</div>' +
    '</div>' +
    '<div class="sec-body" id="b-' + id + '">' + (body || '<div class="empty">Sin registros</div>') + '</div>' +
  '</div>';
}

// Badge de método de cobro con colores diferenciados
function _metodoBadge(metodo) {
  var m = (metodo || '').toUpperCase();
  var cls;
  if (m === 'EFECTIVO')   cls = 'b-ef';
  else if (m === 'CREDITO')    cls = 'b-cr';
  else if (m === 'POR_COBRAR') cls = 'b-pc';
  else if (m === 'ANULADO')    cls = 'b-an';
  else if (m.indexOf('YAPE') >= 0 || m.indexOf('PLIN') >= 0) cls = 'b-yp';
  else if (m.indexOf('VISA') >= 0 || m.indexOf('TARJETA') >= 0 || m.indexOf('MASTERCARD') >= 0) cls = 'b-vi';
  else cls = 'b-tr';
  return '<span class="badge ' + cls + '">' + metodo + '</span>';
}

function _tablaVentas(lista, detMap, badgeCls) {
  if (!lista || lista.length === 0) return '<div class="empty">Sin tickets</div>';

  // Calcular sub-totales por método
  var byM = {};
  lista.forEach(function(v) {
    var m = v.metodo || 'EFECTIVO';
    if (!byM[m]) byM[m] = { count: 0, total: 0 };
    byM[m].count++;
    byM[m].total += v.total;
  });

  var fm = function(n){ return 'S/ ' + parseFloat(n||0).toFixed(2); };

  // Filas de la tabla
  var rows = ['<table class="tbl"><thead><tr>'];
  rows.push('<th>#</th><th>Correlativo</th><th>Cliente</th><th>Cobro</th><th style="text-align:right">Total</th>');
  rows.push('</tr></thead><tbody>');

  lista.forEach(function(v, idx) {
    var det = detMap[v.idVenta] || [];
    var clienteStr = v.clienteNom || (v.clienteDoc ? 'Doc: '+v.clienteDoc : '—');
    rows.push('<tr class="det-row">');
    rows.push('<td style="color:#94a3b8;font-size:12px">' + (idx+1) + '</td>');
    rows.push('<td><span class="badge ' + badgeCls + '">' + (v.correlativo||'—') + '</span>' +
      (v.hora >= 0 ? '<div style="font-size:11px;color:#94a3b8;margin-top:2px">' + (v.fechaISO||'').substring(11,16) + '</div>' : '') + '</td>');
    rows.push('<td style="max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:12px">' + clienteStr + '</td>');
    rows.push('<td>' + _metodoBadge(v.metodo) + '</td>');
    rows.push('<td style="font-weight:700;text-align:right">S/ ' + v.total.toFixed(2) + '</td>');
    rows.push('</tr>');
    if (det.length > 0) {
      rows.push('<tr class="items-row"><td colspan="5"><div class="items-inner">');
      det.forEach(function(it){
        rows.push('<div class="item-line"><span>' + it.cantidad + ' × ' + it.nombre + '</span><span>S/ ' + it.subtotal.toFixed(2) + '</span></div>');
      });
      rows.push('</div></td></tr>');
    }
  });
  rows.push('</tbody></table>');

  // Barra de totales por cobro al pie
  rows.push('<div class="cobro-totals">');
  rows.push('<span style="font-size:11px;font-weight:600;color:#94a3b8;align-self:center;margin-right:4px">TOTALES:</span>');
  Object.keys(byM).forEach(function(m) {
    rows.push('<div class="cobro-chip">' +
      _metodoBadge(m) +
      '<span class="cc-amt">' + fm(byM[m].total) + '</span>' +
      '<span class="cc-cnt">(' + byM[m].count + ')</span>' +
    '</div>');
  });
  rows.push('</div>');

  return rows.join('');
}
