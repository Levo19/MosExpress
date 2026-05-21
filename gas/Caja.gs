// ============================================================
// MosExpress — Caja.gs
// Apertura/cierre de turno, cobros, anulaciones, créditos,
// movimientos extra de caja y query de cajero activo.
// ============================================================

// ── Helper: chequea si ya existe guía SALIDA_VENTAS para una caja ────────────
// Mismo criterio que generarGuiaSalidaVentas (Tipo='SALIDA_VENTAS' + Obs
// contiene cajaId). Retorna true/false sin generar nada.
function _existeGuiaSalidaVentasParaCaja(ss, cajaId) {
  var sh = ss.getSheetByName('GUIAS_CABECERA');
  if (!sh) return false;
  var d = sh.getDataRange().getValues();
  var idStr = String(cajaId);
  for (var i = 1; i < d.length; i++) {
    if (String(d[i][4]) === 'SALIDA_VENTAS' &&
        String(d[i][5] || '').indexOf(idStr) >= 0) {
      return true;
    }
  }
  return false;
}

// ── Helper: cierra automáticamente cajas ABIERTA de días anteriores ──────────
// [v2.5.10] Refactor: ahora usa _cerrarCajaAtomicoCore para cada caja zombi,
// igual que el cierre forzado de admin. Eso significa que cada auto-cierre:
//   ✓ Anula POR_COBRAR de esa caja
//   ✓ Calcula montoFinal (apertura + efectivo + ingresos - egresos)
//   ✓ Escribe CERRADA_AUTO + fechaCierre
//   ✓ Genera guía SALIDA_VENTAS
//   ✓ Audita la operación con source 'AUTO_CIERRE_DIA'
//   ✓ Notifica al cajero original (push)
//   ✓ Notifica a MOS
//   ✗ NO imprime ticket Z automático (no hay sesión browser; tendrías que
//     ir a 'Ver cierre' desde MOS para reimprimir)
// Antes solo marcaba el estado y nada más → POR_COBRAR quedaban colgados,
// sin guía, sin descuento de stock, sin notificación.
//
// Toma un solo LockService para procesar todas las cajas en batch (más
// eficiente que pedir lock N veces).
function _autoCerrarCajasViejas(sheetCajas) {
  var tz  = Session.getScriptTimeZone();
  var hoy = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var filas = sheetCajas.getDataRange().getValues();

  // Detectar cajas a auto-cerrar primero (lectura)
  var idsZombi = [];
  for (var c = 1; c < filas.length; c++) {
    if (String(filas[c][5]) !== 'ABIERTA') continue;
    var fApert = filas[c][3];
    if (!fApert) continue;
    var diaApert = Utilities.formatDate(
      fApert instanceof Date ? fApert : new Date(fApert), tz, 'yyyy-MM-dd'
    );
    if (diaApert < hoy) {
      idsZombi.push(String(filas[c][0]));
    }
  }
  if (!idsZombi.length) return 0;

  // Tomar el lock global UNA VEZ para todo el batch
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(60000);
  } catch(e) {
    Logger.log('_autoCerrarCajasViejas: no se pudo tomar lock, skipping');
    return 0;
  }

  var cerradas = 0;
  try {
    idsZombi.forEach(function(idCaja) {
      try {
        var resp = _cerrarCajaAtomicoCore({
          idCaja:      idCaja,
          montoFinal:  null,                // auto: calcular
          idsAnular:   null,                // auto: detectar POR_COBRAR
          esForzado:   true,                // forzado por sistema
          estadoFinal: 'CERRADA_AUTO',      // distingue del cierre manual
          adminAuth: {
            nombre:     'auto-sistema',
            rol:        'SISTEMA',
            via:        'AUTO_CIERRE_DIA',
            idPersonal: ''
          },
          motivo: 'Auto-cierre de caja del día anterior (jornada vencida)'
        });
        // Verificar respuesta success
        var parsed = null;
        try { parsed = JSON.parse(resp.getContent()); } catch(_){}
        if (parsed && parsed.status === 'success') cerradas++;
      } catch(eC) {
        Logger.log('_autoCerrarCajasViejas error en ' + idCaja + ': ' + eC.message);
      }
    });
  } finally {
    try { lock.releaseLock(); } catch(_){}
  }

  return cerradas;
}

function procesarAperturaCaja(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetCajas = ss.getSheetByName("CAJAS");
  if (!sheetCajas) return generarRespuestaError("Pestaña CAJAS no encontrada.");

  // Auto-cerrar cajas de días anteriores y forzar escritura antes de re-leer
  var _cajasAutoCerradas = _autoCerrarCajasViejas(sheetCajas);

  // Asegurar que la columna 'PrintNode_ID' existe (col 10) — auto-creación idempotente
  // para que warehouseMos pueda leer a qué impresora mandar avisos de preingreso.
  try {
    var lastCol = sheetCajas.getLastColumn();
    var headers = sheetCajas.getRange(1, 1, 1, Math.max(lastCol, 1)).getValues()[0];
    var hasPrintNode = headers.some(function(h) { return String(h).trim() === 'PrintNode_ID'; });
    if (!hasPrintNode) sheetCajas.getRange(1, lastCol + 1).setValue('PrintNode_ID');
  } catch(e) { /* no-fatal */ }

  // Un solo cajero activo por zona a la vez
  if (data.zona) {
    var filasActualizadas = sheetCajas.getDataRange().getValues();
    for (var i = 1; i < filasActualizadas.length; i++) {
      if (String(filasActualizadas[i][5]) === 'ABIERTA' &&
          String(filasActualizadas[i][8] || '') === String(data.zona)) {
        return generarRespuestaError(
          "Ya hay un turno activo en " + data.zona + " (cajero: " + filasActualizadas[i][1] + "). Cierra ese turno primero."
        );
      }
    }
  }

  var idCaja = "CAJA-" + new Date().getTime();
  var _tz    = Session.getScriptTimeZone();
  var _ahora = Utilities.formatDate(new Date(), _tz, 'yyyy-MM-dd HH:mm:ss');
  // Columnas: ID_Caja | Vendedor | Estacion | Fecha_Apertura | Monto_Inicial | Estado | Monto_Final | Fecha_Cierre | Zona_ID | PrintNode_ID
  sheetCajas.appendRow([
    idCaja, data.vendedor, data.estacion, _ahora,
    data.montoInicial || 0, "ABIERTA", "", "", data.zona || '',
    data.printNodeId || ''   // NUEVA: ID PrintNode de la impresora asignada a esta caja
  ]);
  SpreadsheetApp.flush(); // garantiza que appendRow llegue a Sheets antes de retornar el ID al frontend

  // Notificar a admins/master en MOS — solo a ellos, no al cajero mismo
  try {
    var horaStr = Utilities.formatDate(new Date(), _tz, 'HH:mm');
    _notificarMOS(
      '🛒 ' + (data.vendedor || 'Cajero') + ' aperturó caja',
      (data.estacion || data.zona || '') + ' · ' + horaStr,
      data.vendedor || null,
      'ME_CAJA_APERTURA'
    );
  } catch(eP) { Logger.log('Push apertura caja: ' + eP.message); }

  return ContentService.createTextOutput(JSON.stringify({
    status: "success", idCaja: idCaja,
    mensaje: "Caja aperturada exitosamente",
    cajasAutoCerradas: _cajasAutoCerradas
  })).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// CIERRE DE CAJA — refactor atómico unificado (v2.5.7)
// ============================================================
// Antes había dos endpoints separados sin atomicidad:
//   - ANULACION_MASIVA (POST 1)
//   - CIERRE_CAJA      (POST 2)
// Si el POST 2 fallaba, quedaba estado inconsistente (POR_COBRAR anulados
// pero caja ABIERTA). Caso real: CAJA-1778848407996 el 2026-05-15.
//
// Ahora un único helper _cerrarCajaAtomico hace TODO con LockService:
//   1. Validar caja existe + estado actual (idempotente si ya CERRADA)
//   2. Anular POR_COBRAR (idsAnular dados, o detectados automáticamente)
//   3. Calcular montoFinal (si no viene del cajero, calcular auto)
//   4. Escribir CERRADA + montoFinal + fechaCierre + flush
//   5. Auditoría con historialCambios (lo que faltaba antes)
//   6. Generar guía SALIDA_VENTAS (no bloquea respuesta)
//   7. Push a MOS y opcionalmente al cajero original
//
// procesarCierreCaja (cajero) y cerrarCajaForzado (admin) delegan acá.
function _cerrarCajaAtomico(opts) {
  opts = opts || {};
  var idCaja = String(opts.idCaja || '');
  if (!idCaja) return generarRespuestaError('idCaja requerido');

  // [v2.5.10] Si ya estamos dentro de un lock externo (ej. auto-cierre que
  // procesa múltiples cajas en batch), skipear LockService propio.
  if (opts._skipLock) {
    return _cerrarCajaAtomicoCore(opts);
  }

  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch(e) { return generarRespuestaError('LOCK_TIMEOUT: otra operación en curso, reintentá en unos segundos'); }

  try {
    return _cerrarCajaAtomicoCore(opts);
  } finally {
    try { lock.releaseLock(); } catch(_){}
  }
}

// Lógica del cierre — sin LockService (el caller decide si envuelve o no).
function _cerrarCajaAtomicoCore(opts) {
  var idCaja = String(opts.idCaja || '');
  // [v2.5.10] estadoFinal: 'CERRADA' default, 'CERRADA_AUTO' para auto-cierre
  var estadoFinal = String(opts.estadoFinal || 'CERRADA');

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetCajas = ss.getSheetByName('CAJAS');
    var sheetVentas = ss.getSheetByName('VENTAS_CABECERA');
    var sheetExtra = ss.getSheetByName('MOVIMIENTOS_EXTRA');
    if (!sheetCajas)  return generarRespuestaError('CAJAS no encontrada');
    if (!sheetVentas) return generarRespuestaError('VENTAS_CABECERA no encontrada');

    // ── 1. Localizar la caja ──
    var filasCajas = sheetCajas.getDataRange().getValues();
    var filaCaja = -1;
    var cajaRow = null;
    for (var i = 1; i < filasCajas.length; i++) {
      if (String(filasCajas[i][0]) === idCaja) {
        filaCaja = i;
        cajaRow = filasCajas[i];
        break;
      }
    }
    if (filaCaja < 0) return generarRespuestaError('Caja ' + idCaja + ' no encontrada');

    // Leer info de la caja ANTES del check yaCerrada porque puede que
    // necesitemos regenerar la guía SALIDA_VENTAS si falta (modo reparación).
    var cajaVendedor = String(cajaRow[1] || '');
    var cajaEstacion = String(cajaRow[2] || '');
    var cajaZona     = String(cajaRow[8] || '');
    var montoInicial = parseFloat(cajaRow[4]) || 0;
    // [v2.5.8] PrintNode_ID guardado al abrir caja
    var printNodeId  = String(cajaRow[9] || '');

    // Idempotencia + REPARACIÓN: si ya está cerrada (CERRADA o CERRADA_AUTO)
    // pero la guía SALIDA_VENTAS nunca se generó, regenerar ahora.
    // generarGuiaSalidaVentas tiene defensa anti-duplicado internamente.
    var estadoActual = String(cajaRow[5] || '');
    if (estadoActual === 'CERRADA' || estadoActual === 'CERRADA_AUTO') {
      var guiaRegenerada = false;
      var existeGuia = false;
      if (cajaZona) {
        try { existeGuia = _existeGuiaSalidaVentasParaCaja(ss, idCaja); } catch(_){}
        if (!existeGuia) {
          try {
            generarGuiaSalidaVentas(ss, idCaja, cajaVendedor, cajaZona);
            guiaRegenerada = true;
            // Auditoría de la reparación
            try {
              if (typeof auditarLog === 'function') {
                auditarLog('CAJAS', idCaja, {
                  usuario: String((opts.adminAuth && opts.adminAuth.nombre) || cajaVendedor),
                  rol:     String((opts.adminAuth && opts.adminAuth.rol) || 'CAJERO'),
                  source:  'ME_REGEN_GUIA_VENTAS',
                  accion:  'regenerar_guia_salida_ventas',
                  autorizadoPor: opts.adminAuth || null,
                  ref: { idCaja: idCaja, vendedor: cajaVendedor, zona: cajaZona },
                  motivo: 'Caja ya CERRADA pero sin guía SALIDA_VENTAS — regeneración automática',
                  ts: new Date().toISOString()
                });
              }
            } catch(_){}
          } catch(eR) { Logger.log('Regen guia: ' + eR.message); }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        yaCerrada: true,
        guiaRegenerada: guiaRegenerada,
        guiaExistia: existeGuia,
        mensaje: guiaRegenerada
          ? 'Caja ya cerrada — guía SALIDA_VENTAS regenerada'
          : 'Caja ya estaba cerrada (guía OK)',
        idCaja: idCaja,
        vendedor: cajaVendedor,
        zona: cajaZona,
        montoFinal: parseFloat(cajaRow[6]) || 0,
        printNodeId: printNodeId
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // ── 2. Anular POR_COBRAR ──
    // Si el frontend mandó idsAnular, usar esa lista. Si no, detectar
    // automáticamente todos los POR_COBRAR de esta caja.
    var idsTarget = Array.isArray(opts.idsAnular) ? opts.idsAnular.map(String) : null;
    var filasV = sheetVentas.getDataRange().getValues();
    var idsAnulados = [];
    var efectivoVentas = 0;
    for (var v = 1; v < filasV.length; v++) {
      var idCajaV = String(filasV[v][10] || '');
      var idV     = String(filasV[v][0] || '');
      var formaPago = String(filasV[v][8] || '').toUpperCase();
      var total = parseFloat(filasV[v][6]) || 0;

      // Anular POR_COBRAR: usar lista explícita si vino, si no auto-detectar
      // por caja actual. Importante: nunca anular CREDITO, ya está formal.
      var debeAnular = false;
      if (idsTarget) {
        if (idsTarget.indexOf(idV) !== -1) debeAnular = true;
      } else if (idCajaV === idCaja && formaPago === 'POR_COBRAR') {
        debeAnular = true;
      }
      if (debeAnular && formaPago === 'POR_COBRAR') {
        sheetVentas.getRange(v + 1, 9).setValue('ANULADO');
        idsAnulados.push(idV);
        continue;
      }

      // Si esta venta pertenece a la caja, sumar efectivo (para auto-cálculo)
      if (idCajaV === idCaja) {
        if (formaPago === 'EFECTIVO') {
          efectivoVentas += total;
        } else if (formaPago.indexOf('MIXTO') === 0) {
          var m = formaPago.match(/EFE:([\d.]+)/);
          if (m) efectivoVentas += parseFloat(m[1]) || 0;
        }
      }
    }

    // ── 3. Sumar ingresos/egresos extra de la caja ──
    var ingresosEfe = 0, egresosEfe = 0;
    if (sheetExtra) {
      var filasE = sheetExtra.getDataRange().getValues();
      for (var x = 1; x < filasE.length; x++) {
        if (String(filasE[x][1] || '') !== idCaja) continue;
        var tipoE = String(filasE[x][3] || '');
        var mtoE  = parseFloat(filasE[x][4]) || 0;
        if      (tipoE === 'INGRESO') ingresosEfe += mtoE;
        else if (tipoE === 'EGRESO')  egresosEfe  += mtoE;
      }
    }

    // ── 4. Determinar montoFinal ──
    // Si el cajero lo declaró explícitamente, respetar (puede haber descuadre).
    // Si no viene, calcular automático.
    var montoFinalAuto = Math.round((montoInicial + efectivoVentas + ingresosEfe - egresosEfe) * 100) / 100;
    var montoFinal;
    if (opts.montoFinal !== null && opts.montoFinal !== undefined && opts.montoFinal !== '') {
      montoFinal = parseFloat(opts.montoFinal);
      if (isNaN(montoFinal)) montoFinal = montoFinalAuto;
    } else {
      montoFinal = montoFinalAuto;
    }

    // ── 5. Escribir CERRADA + montoFinal + fechaCierre ──
    var tz = Session.getScriptTimeZone();
    var fechaCierre = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
    sheetCajas.getRange(filaCaja + 1, 6).setValue(estadoFinal);
    sheetCajas.getRange(filaCaja + 1, 7).setValue(montoFinal);
    sheetCajas.getRange(filaCaja + 1, 8).setValue(fechaCierre);
    SpreadsheetApp.flush();

    // ── 6. Auditoría ──
    try {
      if (typeof auditarLog === 'function') {
        // [v2.5.10] Distinguir source: ME_CIERRE_CAJA (cajero), MOS_CIERRE_FORZADO
        // (admin desde MOS), AUTO_CIERRE_DIA (sistema, jornada vencida).
        var auditSource, auditAccion;
        if (opts.adminAuth && opts.adminAuth.via === 'AUTO_CIERRE_DIA') {
          auditSource = 'AUTO_CIERRE_DIA';
          auditAccion = 'auto_cerrar_caja_jornada_vencida';
        } else if (opts.esForzado) {
          auditSource = 'MOS_CIERRE_FORZADO';
          auditAccion = 'cerrar_caja_forzado';
        } else {
          auditSource = 'ME_CIERRE_CAJA';
          auditAccion = 'cerrar_caja';
        }
        auditarLog('CAJAS', idCaja, {
          usuario: String((opts.adminAuth && opts.adminAuth.nombre) || cajaVendedor),
          rol: String((opts.adminAuth && opts.adminAuth.rol) || 'CAJERO'),
          source: auditSource,
          accion: auditAccion,
          autorizadoPor: opts.adminAuth || null,
          cambios: [
            { campo: 'Estado',      antes: 'ABIERTA', despues: 'CERRADA' },
            { campo: 'Monto_Final', antes: '',         despues: montoFinal }
          ],
          ref: {
            idCaja: idCaja, vendedor: cajaVendedor, zona: cajaZona,
            idsAnulados: idsAnulados.length, montoFinal: montoFinal,
            montoFinalAuto: montoFinalAuto, descuadre: montoFinal - montoFinalAuto
          },
          motivo: String(opts.motivo || ''),
          ts: new Date().toISOString()
        });
      }
    } catch(eA) { Logger.log('Audit cierre: ' + eA.message); }

    // ── 7. Generar guía SALIDA_VENTAS (no bloquea respuesta) ──
    if (cajaZona) {
      try { generarGuiaSalidaVentas(ss, idCaja, cajaVendedor, cajaZona); }
      catch(eG) { Logger.log('Error guia ventas: ' + eG.toString()); }
    }

    // ── 8. Push al cajero (solo si fue forzado) ──
    if (opts.esForzado && cajaVendedor) {
      try {
        if (typeof enviarPushUsuario === 'function') {
          var admin = String((opts.adminAuth && opts.adminAuth.nombre) || 'admin');
          enviarPushUsuario(cajaVendedor,
            '🔐 Tu caja fue cerrada por admin',
            admin + ' cerró tu turno · Monto final S/ ' + montoFinal.toFixed(2),
            { idNotif: 'ME_CAJA_CERRADA_POR_ADMIN', idCaja: idCaja });
        }
      } catch(eU) { Logger.log('Push cajero: ' + eU.message); }
    }

    // ── 9. Push a MOS confirmando ──
    try {
      var hora = Utilities.formatDate(new Date(), tz, 'HH:mm');
      var titulo = opts.esForzado
        ? ('🔐 Cierre forzado · ' + cajaVendedor)
        : ('🔐 Caja cerrada · ' + hora);
      var detalle = opts.esForzado
        ? (String((opts.adminAuth && opts.adminAuth.nombre) || 'admin') + ' · S/ ' + montoFinal.toFixed(2) + (idsAnulados.length ? ' · ' + idsAnulados.length + ' anulados' : ''))
        : (cajaVendedor + (cajaZona ? ' · ' + cajaZona : '') + ' · S/ ' + montoFinal.toFixed(2));
      _notificarMOS(titulo, detalle, cajaVendedor, opts.esForzado ? 'ME_CAJA_CIERRE_FORZADO' : 'ME_CAJA_CIERRE');
    } catch(eM) { Logger.log('Push MOS cierre: ' + eM.message); }

    return ContentService.createTextOutput(JSON.stringify({
      status:         'success',
      idCaja:         idCaja,
      vendedor:       cajaVendedor,
      estacion:       cajaEstacion,
      zona:           cajaZona,
      printNodeId:    printNodeId,        // [v2.5.8] para imprimir Z desde MOS
      montoInicial:   montoInicial,
      efectivoVentas: Math.round(efectivoVentas * 100) / 100,
      ingresos:       Math.round(ingresosEfe * 100) / 100,
      egresos:        Math.round(egresosEfe * 100) / 100,
      montoFinal:     montoFinal,
      montoFinalAuto: montoFinalAuto,
      descuadre:      Math.round((montoFinal - montoFinalAuto) * 100) / 100,
      anulados:       idsAnulados.length,
      idsAnulados:    idsAnulados,
      fechaCierre:    fechaCierre,
      cerradoPor:     String((opts.adminAuth && opts.adminAuth.nombre) || cajaVendedor),
      esForzado:      !!opts.esForzado,
      mensaje:        opts.esForzado ? 'Caja cerrada forzadamente por admin' : 'Caja cerrada correctamente'
    })).setMimeType(ContentService.MimeType.JSON);

  } catch(eC) {
    Logger.log('_cerrarCajaAtomicoCore error: ' + (eC && eC.message || eC));
    return generarRespuestaError('Error interno cierre: ' + (eC && eC.message || eC));
  }
}

// Endpoint del cajero — flow normal, delega al helper atómico.
// Backward compat: acepta data.cajaId (legacy) o data.idCaja.
function procesarCierreCaja(data) {
  return _cerrarCajaAtomico({
    idCaja:     data.idCaja || data.cajaId,
    montoFinal: data.montoFinal,         // si viene, respetar (cajero lo declara)
    idsAnular:  data.idsAnular || null,   // opcional, si no viene se auto-detecta
    esForzado:  false,
    adminAuth:  null,
    motivo:     ''
  });
}

// Endpoint admin/master — cierre forzado desde MOS. Delega al helper atómico
// pasando esForzado=true para que la auditoría/push lleven la marca de admin.
function cerrarCajaForzado(data) {
  if (!data || !data.idCaja) return generarRespuestaError('idCaja requerido');
  return _cerrarCajaAtomico({
    idCaja:    data.idCaja,
    montoFinal: null,             // admin no declara: calculamos auto
    idsAnular:  null,             // auto-detectar POR_COBRAR de la caja
    esForzado:  true,
    adminAuth:  data.adminAuth || {},
    motivo:     String(data.motivo || 'Cierre forzado por admin desde MOS')
  });
}

// [v2.5.33] Devuelve TODOS los cajeros activos por zona en una sola llamada.
// El wizard moderno lo usa para poblar las cards de zona en el paso 2 con la
// info de "🔴 cajero ya activo: X" sin tener que disparar N requests.
function cajerosActivosTodos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("CAJAS");
  if (!sheet) return generarRespuestaError("CAJAS no encontrada");
  var _cerradas = _autoCerrarCajasViejas(sheet);
  var data = sheet.getDataRange().getValues();
  var porZona = {};
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][5]) === 'ABIERTA') {
      var z = String(data[i][8] || '');
      if (!porZona[z]) porZona[z] = { vendedor: String(data[i][1]), idCaja: String(data[i][0]), desde: data[i][3] };
    }
  }
  return ContentService.createTextOutput(JSON.stringify({
    status: 'success', porZona: porZona, cajasAutoCerradas: _cerradas
  })).setMimeType(ContentService.MimeType.JSON);
}

// [v2.5.40] Cambiar la impresora ASIGNADA de una caja ABIERTA.
// Usado cuando el cajero/vendedor cambia de impresora desde el modal admin.
// Actualiza CAJAS col 3 (Estacion) y col 10 (PrintNode_ID) para que cualquier
// flujo que use estos datos (ticket cobro, ingreso proveedor desde almacén,
// reimpresión, cierre Z) imprima en la nueva impresora.
function cambiarImpresoraCaja(data) {
  if (!data.idCaja)        return generarRespuestaError('idCaja requerido');
  if (!data.estacionNombre) return generarRespuestaError('estacionNombre requerido');
  if (!data.printNodeId)   return generarRespuestaError('printNodeId requerido');
  if (!data.adminAuth || !data.adminAuth.nombre) {
    return generarRespuestaError('adminAuth requerido (requiere admin/master)');
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('CAJAS');
  if (!sheet) return generarRespuestaError('CAJAS no encontrada');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.idCaja)) {
      if (String(rows[i][5]) !== 'ABIERTA') {
        return generarRespuestaError('La caja no está ABIERTA — no se puede cambiar impresora');
      }
      var estacionPrev = String(rows[i][2] || '');
      var printNodePrev = String(rows[i][9] || '');
      // Update Estacion (col 3 = index 2) y PrintNode_ID (col 10 = index 9)
      sheet.getRange(i + 1, 3).setValue(String(data.estacionNombre));
      sheet.getRange(i + 1, 10).setValue(String(data.printNodeId));
      SpreadsheetApp.flush();
      // Auditar
      try {
        auditarLog('CAJAS_CAMBIO_IMPRESORA', data.idCaja, {
          usuario: 'MOS-Admin', rol: 'ADMIN',
          source: 'ME_CAMBIO_IMPRESORA',
          accion: 'update_PrintNode_ID',
          autorizadoPor: { nombre: data.adminAuth.nombre, rol: data.adminAuth.rol || 'ADMIN', via: 'PIN_8DIG' },
          ref: {
            idCaja: data.idCaja,
            estacionPrev: estacionPrev, printNodePrev: printNodePrev,
            estacionNueva: data.estacionNombre, printNodeNuevo: data.printNodeId
          },
          motivo: String(data.motivo || '')
        });
      } catch(_) {}
      return ContentService.createTextOutput(JSON.stringify({
        status: 'success', idCaja: data.idCaja,
        estacionPrev: estacionPrev, printNodePrev: printNodePrev,
        estacionNueva: data.estacionNombre, printNodeNuevo: data.printNodeId
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  return generarRespuestaError('Caja ' + data.idCaja + ' no encontrada');
}

function cajeroActivo(zona) {
  if (!zona) return generarRespuestaError("zona requerida");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("CAJAS");
  if (!sheet) return generarRespuestaError("CAJAS no encontrada");
  // Auto-cerrar cajas viejas antes de consultar (evita falso positivo de "hay cajero activo")
  var _cerradas = _autoCerrarCajasViejas(sheet);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][5]) === 'ABIERTA' && String(data[i][8] || '') === zona) {
      return ContentService.createTextOutput(JSON.stringify({
        status: 'success', activo: true,
        vendedor: String(data[i][1]), idCaja: String(data[i][0]), desde: data[i][3],
        cajasAutoCerradas: _cerradas
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  return ContentService.createTextOutput(JSON.stringify({
    status: 'success', activo: false, cajasAutoCerradas: _cerradas
  })).setMimeType(ContentService.MimeType.JSON);
}

function cobrarVentaExistente(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("VENTAS_CABECERA");
  if (!sheet) return generarRespuestaError("VENTAS_CABECERA no encontrada");

  var filas = sheet.getDataRange().getValues();
  for (var i = filas.length - 1; i > 0; i--) {  // buscar desde el final (más probable)
    if (String(filas[i][0]) === String(data.idVenta)) {
      var formaAnt = String(filas[i][8] || '');
      var cajaAnt  = String(filas[i][10] || '');
      sheet.getRange(i + 1, 9).setValue(data.metodo);
      if (data.cajaId) sheet.getRange(i + 1, 11).setValue(String(data.cajaId));

      // Log de auditoría
      try {
        var actor = _audExtraerActor(data);
        var cambios = [{ campo:'FormaPago', antes: formaAnt, despues: String(data.metodo) }];
        if (data.cajaId && data.cajaId !== cajaAnt) {
          cambios.push({ campo:'ID_Caja', antes: cajaAnt, despues: String(data.cajaId) });
        }
        auditarLog('VENTAS_CABECERA', data.idVenta, {
          usuario: actor.usuario, rol: actor.rol,
          source: 'ME_COBRAR_VENTA',
          accion: 'cobrar_venta',
          cambios: cambios,
          autorizadoPor: actor.autorizadoPor || null,
          motivo: data.motivo || ''
        });
      } catch(_){}

      return ContentService.createTextOutput(JSON.stringify({
        status: "success", mensaje: "Venta cobrada correctamente"
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  return generarRespuestaError("Venta con ID " + data.idVenta + " no encontrada.");
}

function creditarVenta(data) {
  if (!data.idVenta) return generarRespuestaError("idVenta requerido");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("VENTAS_CABECERA");
  if (!sheet) return generarRespuestaError("VENTAS_CABECERA no encontrada");
  var filas = sheet.getDataRange().getValues();
  for (var i = filas.length - 1; i > 0; i--) {
    if (String(filas[i][0]) === String(data.idVenta)) {
      var formaAnt = String(filas[i][8] || '');
      var obsAnt   = String(filas[i][14] || '');
      sheet.getRange(i + 1, 9).setValue('CREDITO');
      sheet.getRange(i + 1, 15).setValue(String(data.obs || ''));

      try {
        var actor = _audExtraerActor(data);
        auditarLog('VENTAS_CABECERA', data.idVenta, {
          usuario: actor.usuario, rol: actor.rol,
          source: 'ME_CREDITAR_VENTA',
          accion: 'convertir_a_credito',
          cambios: [
            { campo:'FormaPago', antes: formaAnt, despues:'CREDITO' },
            { campo:'Obs',       antes: obsAnt,   despues: String(data.obs || '') }
          ],
          autorizadoPor: actor.autorizadoPor || null,
          motivo: data.motivo || data.obs || ''
        });
      } catch(_){}

      return ContentService.createTextOutput(JSON.stringify({
        status: "success", mensaje: "Crédito registrado"
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  return generarRespuestaError("Venta " + data.idVenta + " no encontrada.");
}

function anularVentaIndividual(data) {
  if (!data.ventaId) return generarRespuestaError("No se proporcionó ventaId.");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("VENTAS_CABECERA");
  if (!sheet) return generarRespuestaError("VENTAS_CABECERA no encontrada.");
  var filas = sheet.getDataRange().getValues();
  for (var i = filas.length - 1; i > 0; i--) {
    if (String(filas[i][0]) === String(data.ventaId)) {
      var formaAnt = String(filas[i][8] || '');
      sheet.getRange(i + 1, 9).setValue('ANULADO');

      try {
        var actor = _audExtraerActor(data);
        auditarLog('VENTAS_CABECERA', data.ventaId, {
          usuario: actor.usuario, rol: actor.rol,
          source: 'ME_ANULAR_VENTA',
          accion: 'anular_venta_interna',
          cambios: [{ campo:'FormaPago', antes: formaAnt, despues:'ANULADO' }],
          autorizadoPor: actor.autorizadoPor || null,
          motivo: data.motivo || ''
        });
      } catch(_){}

      // Avisar a WH que descuente del pickup origen (si existe y no está cerrado).
      // No bloquea ni rompe la anulación si falla.
      try { notificarAnulacionPickupAWH(data.ventaId); } catch(_){}

      return ContentService.createTextOutput(JSON.stringify({
        status: "success", mensaje: "Venta anulada correctamente"
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  return generarRespuestaError("Venta con ID " + data.ventaId + " no encontrada.");
}

// Anula en masa todos los tickets POR_COBRAR no cobrados al cierre del turno
function anulacionMasiva(data) {
  if (!data.ids || !data.ids.length) return generarRespuestaError("No se enviaron IDs a anular.");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("VENTAS_CABECERA");
  if (!sheet) return generarRespuestaError("VENTAS_CABECERA no encontrada.");

  var filas = sheet.getDataRange().getValues();
  var anulados = 0;
  var idsAnulados = [];
  for (var i = 1; i < filas.length; i++) {
    if (data.ids.indexOf(String(filas[i][0])) !== -1) {
      sheet.getRange(i + 1, 9).setValue('ANULADO');
      anulados++;
      idsAnulados.push(String(filas[i][0]));
    }
  }
  // Notificar WH para descontar de pickups origen (no bloquea)
  try { idsAnulados.forEach(function(id){ notificarAnulacionPickupAWH(id); }); } catch(_){}
  return ContentService.createTextOutput(JSON.stringify({
    status: "success", anulados: anulados
  })).setMimeType(ContentService.MimeType.JSON);
}

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
    String(data.cajaId      || ''),
    new Date(),
    String(data.tipo        || 'EGRESO'),
    parseFloat(data.monto)  || 0,
    String(data.concepto    || ''),
    String(data.obs         || ''),
    String(data.registradoPor || '')
  ]);

  // Alerta de recojo de efectivo: tras INGRESO sube monto → posible cruce de
  // threshold (alerta). Tras EGRESO baja monto → bandera se reajusta sola
  // sin enviar push. Virtuales no cuentan (no tocan caja física).
  try {
    if (data.cajaId) _chequearAlertaEfectivo(data.cajaId);
  } catch(eA) { Logger.log('Alerta efectivo (extra): ' + eA.message); }

  return ContentService.createTextOutput(JSON.stringify({
    status: "success", idExtra: id
  })).setMimeType(ContentService.MimeType.JSON);
}

function getExtrasCaja(cajaId) {
  if (!cajaId) return generarRespuestaError("cajaId requerido");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("MOVIMIENTOS_EXTRA");
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({ status: "success", extras: [] }))
      .setMimeType(ContentService.MimeType.JSON);
  }
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
  return ContentService.createTextOutput(JSON.stringify({ status: "success", extras: result }))
    .setMimeType(ContentService.MimeType.JSON);
}
