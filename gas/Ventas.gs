// ============================================================
// MosExpress — Ventas.gs
// Registro de ventas, correlativo O(1), consultas de ventas.
// ============================================================

function procesarVenta(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetCabecera = ss.getSheetByName("VENTAS_CABECERA");
  var sheetDetalle  = ss.getSheetByName("VENTAS_DETALLE");

  if (!sheetCabecera) throw new Error("Pestaña VENTAS_CABECERA no encontrada.");
  if (!sheetDetalle)  throw new Error("Pestaña VENTAS_DETALLE no encontrada.");

  var auth   = data.auth        || {};
  var pos    = data.pos_config  || {};
  var header = data.header      || {};
  var items  = data.items       || [];

  // [v2.5.45] GUARD ANTI-HUÉRFANAS — Rechazar payloads malformados ANTES de
  // generar un correlativo "undefined-XXX" que ensucia VENTAS_CABECERA.
  // Causas históricas: pendingSales con raw_data corrupto · venta disparada
  // con config.estacion=null · POST vacío de algún probe externo · SW
  // reintentando jobs incompletos.
  var _missing = [];
  if (!auth.vendedor || String(auth.vendedor).trim() === '') _missing.push('auth.vendedor');
  if (!pos.serieActual || String(pos.serieActual).trim() === '') _missing.push('pos_config.serieActual');
  if (!Array.isArray(items) || items.length === 0)              _missing.push('items (lista vacía)');
  var _totalNum = parseFloat(header.total);
  if (isNaN(_totalNum) || _totalNum <= 0)                       _missing.push('header.total (debe ser > 0)');
  if (_missing.length > 0) {
    try {
      Logger.log('[procesarVenta] RECHAZADA — payload inválido. Faltan: ' + _missing.join(', ') +
                 ' · refLocal=' + (data.data_sync && data.data_sync.last_sync) +
                 ' · deviceId=' + (auth.deviceId || '(sin)') +
                 ' · raw=' + JSON.stringify(data).substring(0, 500));
    } catch(_) {}
    return {
      idVenta: null, correlativo: null, printDispatched: false, dedupVenta: false,
      error: 'PAYLOAD_INVALIDO',
      campos_faltantes: _missing,
      mensaje: 'Venta rechazada: faltan ' + _missing.join(', ')
    };
  }

  // Seguridad de rol: dispositivo cajero sin caja abierta → POR_COBRAR
  if (auth.esCajero && !pos.cajaId) {
    header.metodo = 'POR_COBRAR';
  }

  // [v2.6.0] DEFENSA: vendedor (no cajero) creando venta debe tener caja
  // activa en su zona. Si vino con cajaId pero esa caja ya está CERRADA,
  // buscar la caja activa actual de la zona y reasignar. Si no hay ninguna
  // caja abierta en la zona → rechazar con código NO_CAJA_ACTIVA_EN_ZONA
  // (el frontend bloquea el POS con overlay y reintenta cuando reabra caja).
  if (!auth.esCajero) {
    var zonaV = String(auth.zona || pos.zona || '').trim();
    var cajaIdEnviada = String(pos.cajaId || '').trim();
    var cajaValida = false;
    if (zonaV) {
      try {
        var ssV = SpreadsheetApp.getActiveSpreadsheet();
        var shCV = ssV.getSheetByName('CAJAS');
        if (shCV) {
          var fcv = shCV.getDataRange().getValues();
          // Cols CAJAS: 0 ID · 1 Vendedor · 2 Estacion · 3 Fec_Ap · 4 Monto_Ini
          //             · 5 Estado · 6 Mon_Fin · 7 Fec_Cie · 8 Zona · 9 PNode
          // 1) Si vino cajaId, verificar que esa caja siga ABIERTA en la zona
          if (cajaIdEnviada) {
            for (var iv = 1; iv < fcv.length; iv++) {
              if (String(fcv[iv][0]) === cajaIdEnviada
                  && String(fcv[iv][5] || '').toUpperCase() === 'ABIERTA'
                  && String(fcv[iv][8] || '').trim() === zonaV) {
                cajaValida = true;
                break;
              }
            }
          }
          // 2) Si la caja enviada NO está abierta (o no vino), buscar la activa de la zona
          if (!cajaValida) {
            for (var jv = fcv.length - 1; jv >= 1; jv--) {
              if (String(fcv[jv][5] || '').toUpperCase() === 'ABIERTA'
                  && String(fcv[jv][8] || '').trim() === zonaV) {
                pos.cajaId = String(fcv[jv][0]);  // reasignar a la caja activa actual
                cajaValida = true;
                break;
              }
            }
          }
        }
      } catch(eCV) { Logger.log('[procesarVenta] check caja activa zona falló: ' + eCV.message); }
    }
    if (!cajaValida) {
      Logger.log('[procesarVenta] RECHAZADA NO_CAJA_ACTIVA_EN_ZONA · vendedor=' + auth.vendedor + ' zona=' + zonaV);
      return {
        idVenta: null, correlativo: null, printDispatched: false, dedupVenta: false,
        error: 'NO_CAJA_ACTIVA_EN_ZONA',
        zona: zonaV,
        mensaje: 'No hay caja abierta en tu zona. Pide al cajero que abra caja antes de vender.'
      };
    }
  }

  var fechaActual = new Date();
  var idVenta = "V-" + fechaActual.getTime();

  // ── Idempotencia: evita duplicados cuando el browser reintenta ──────────
  // [v40 fix ticket doble] Retornar `dedupVenta:true` para que el frontend
  // NO dispare el fallback de impresión. Antes retornaba sólo
  // `printDispatched:false` → frontend creía que GAS no había impreso y
  // mandaba un 2do print job a PrintNode → 2 tickets físicos.
  var refLocal = (data.data_sync && data.data_sync.last_sync) ? String(data.data_sync.last_sync) : '';
  if (refLocal) {
    var totalFilas  = sheetCabecera.getLastRow();
    var buscarDesde = Math.max(2, totalFilas - 199);
    var filasBuscar = sheetCabecera.getRange(buscarDesde, 1, totalFilas - buscarDesde + 1, 16).getValues();
    for (var fi = filasBuscar.length - 1; fi >= 0; fi--) {
      if (String(filasBuscar[fi][13]) === refLocal) {
        return {
          idVenta:         String(filasBuscar[fi][0]),
          correlativo:     String(filasBuscar[fi][9]),
          printDispatched: false,
          dedupVenta:      true   // ← frontend ve este flag y NO reimprime
        };
      }
    }
  }

  // ── Correlativo: pre-reserva (NV) o atómico (CPE) ───────────────────────
  // [v2.5.58] Si el cliente pre-reservó al abrir el modal de pago (caso NV),
  // viene `idReserva` en el header y usamos ese número. Eso garantiza
  // impresión instantánea sin esperar GAS para numerar. Para CPE
  // (boleta/factura) NO pre-reservamos (gaps requieren reporte SUNAT) →
  // sigue el flujo legacy con obtenerSiguienteCorrelativoRapido.
  var correlativoNumero, correlativoFinal;
  if (header.idReserva && String(header.tipoDoc || '').toUpperCase() === 'NOTA_DE_VENTA') {
    var consumido = _consumirReserva(header.idReserva, idVenta, auth.deviceId);
    if (!consumido.ok) {
      // Reserva inválida — fallback al método atómico
      Logger.log('[Venta] Reserva inválida (' + consumido.error + ') — cayendo a método atómico');
      correlativoNumero = obtenerSiguienteCorrelativoRapido(ss, pos.serieActual);
    } else {
      correlativoNumero = consumido.numero;
    }
  } else {
    correlativoNumero = obtenerSiguienteCorrelativoRapido(ss, pos.serieActual);
  }
  correlativoFinal = pos.serieActual + "-" + ("000000" + correlativoNumero).slice(-6);

  // ── VENTAS_CABECERA (19 columnas) ────────────────────────────────────────
  // ID_Venta | Fecha | Vendedor | Estacion | Cliente_Doc | Cliente_Nombre | Total
  // | Tipo_Doc | FormaPago | Correlativo | ID_Caja | ID_Dispositivo | Estado_Envio
  // | Ref_Local | Obs | Tipo_Doc_Cliente | NF_Estado | NF_Hash | NF_Enlace
  var tipoDocCliente = parseInt((header.cliente && header.cliente.tipo) || 0, 10);
  sheetCabecera.appendRow([
    idVenta, fechaActual, auth.vendedor, auth.estacion,
    (header.cliente && header.cliente.doc)    || '',
    (header.cliente && header.cliente.nombre) || '',
    header.total,
    header.tipoDoc,
    header.metodo || 'EFECTIVO',
    correlativoFinal, pos.cajaId, auth.deviceId, "COMPLETADO",
    refLocal,
    String(header.obs || ''),
    tipoDocCliente,
    '', '', ''   // NF_Estado, NF_Hash, NF_Enlace — se llenan después si aplica
  ]);

  // ── VENTAS_DETALLE (10 columnas) — escritura por lote ────────────────────
  // ID_Venta | SKU | Nombre | Cantidad | Precio | Subtotal | Cod_Barras
  // | Valor_Unitario | Tipo_IGV | Unidad_Medida
  if (items.length > 0) {
    var detalleRows = items.map(function(item) {
      var valorUnitario = parseFloat(item.valor_unitario) ||
                          Math.round(parseFloat(item.precio || 0) / 1.18 * 100) / 100;
      return [
        idVenta,
        item.sku,
        item.nombre,
        item.cantidad,
        item.precio,
        item.subtotal,
        String(item.codBarras || ''),
        Math.round(valorUnitario * 100) / 100,
        parseInt(item.tipo_igv || 1, 10),
        String(item.unidad_de_medida || 'NIU')
      ];
    });
    var lastRow = sheetDetalle.getLastRow();
    var rangeDetalle = sheetDetalle.getRange(lastRow + 1, 1, detalleRows.length, detalleRows[0].length);
    // Forzar texto en col 7 (Cod_Barras) y col 2 (SKU) antes de escribir
    // para que Sheets no elimine ceros a la izquierda de códigos numéricos
    sheetDetalle.getRange(lastRow + 1, 7, detalleRows.length, 1).setNumberFormat('@STRING@');
    sheetDetalle.getRange(lastRow + 1, 2, detalleRows.length, 1).setNumberFormat('@STRING@');
    rangeDetalle.setValues(detalleRows);
  }

  // ── Registrar cliente frecuente ──────────────────────────────────────────
  // [v40] Guardar SIEMPRE que haya doc válido y no sea VARIOS (66666). Antes
  // sólo guardaba en BOLETA/FACTURA → las NV con DNI/RUC real perdían el
  // contacto. Ahora cualquier venta con cliente identificado lo persiste.
  var _cliDoc    = (header.cliente && header.cliente.doc)    || '';
  var _cliNombre = (header.cliente && header.cliente.nombre) || '';
  if (_cliDoc && _cliDoc !== '66666' && _cliNombre) {
    verificarYAgregaCliente(
      _cliDoc, _cliNombre, header.tipoDoc,
      (header.cliente && header.cliente.direccion) || ''
    );
  }

  // ── Emitir CPE en NubeFact (solo BOLETA y FACTURA) ───────────────────────
  var nfEstado = 'NA';
  var nfHash   = '';
  var nfEnlace = '';
  var nfResult = null;

  if (header.tipoDoc === 'BOLETA' || header.tipoDoc === 'FACTURA') {
    // [v40.1] Validaciones SUNAT defensivas — bloquear ANTES de pegarle a NubeFact
    // para evitar emisiones rechazadas y consumir cuota del API.
    // Reglas reales SUNAT:
    //   - Boleta < S/700: VARIOS permitido (tipo=0 "consumidor anónimo")
    //   - Boleta ≥ S/700: DNI (8) o RUC (11) obligatorio
    //   - Factura: SIEMPRE RUC + dirección fiscal
    var _docCli    = String((header.cliente && header.cliente.doc) || '');
    var _nomCli    = String((header.cliente && header.cliente.nombre) || '').trim();
    var _dirCli    = String((header.cliente && header.cliente.direccion) || '').trim();
    var _tipoCli   = parseInt((header.cliente && header.cliente.tipo) || 0, 10); // 0/1/4/6/7
    var _totalVta  = parseFloat(header.total || 0);
    var _bloqueoCPE = '';

    // BOLETA ≥ S/700 requiere doc identificado: DNI(8) | RUC(11) | CE(tipo=4) | PAS(tipo=7)
    if (header.tipoDoc === 'BOLETA' && _totalVta >= 700) {
      var _esCEoPas = (_tipoCli === 4 || _tipoCli === 7) && _docCli && _docCli !== '66666';
      var _esDniRuc = (_docCli.length === 8 || _docCli.length === 11) && _docCli !== '66666' && _docCli !== '0';
      if (!_esCEoPas && !_esDniRuc) {
        _bloqueoCPE = 'BOLETA >=S/700 requiere DNI/RUC/CE/Pasaporte (SUNAT)';
      }
    }
    // FACTURA: RUC 11 dig + dirección obligatoria (NubeFact rechaza sin)
    if (!_bloqueoCPE && header.tipoDoc === 'FACTURA') {
      if (_docCli === '66666' || _docCli.length !== 11) {
        _bloqueoCPE = 'FACTURA requiere RUC de 11 digitos (no VARIOS)';
      }
      else if (!_dirCli) _bloqueoCPE = 'FACTURA requiere direccion fiscal';
    }
    // Denominación no puede estar vacía (válido para boleta y factura)
    if (!_bloqueoCPE && !_nomCli) {
      _bloqueoCPE = header.tipoDoc + ' requiere denominacion del cliente';
    }

    if (_bloqueoCPE) {
      nfEstado = 'ERROR';
      nfResult = { ok: false, error: _bloqueoCPE };
      Logger.log('NubeFact bloqueado venta ' + idVenta + ': ' + _bloqueoCPE);
    } else {
      nfResult = emitirNubeFact(data, correlativoFinal);
      nfEstado = nfResult.ok ? 'EMITIDO' : 'ERROR';
      nfHash   = nfResult.hash   || '';
      nfEnlace = nfResult.enlace || '';
    }
    var nfRow = sheetCabecera.getLastRow();
    sheetCabecera.getRange(nfRow, 17, 1, 3).setValues([[nfEstado, nfHash, nfEnlace]]);
    if (!nfResult.ok) Logger.log('NubeFact error venta ' + idVenta + ': ' + (nfResult.error || ''));
  }
  var nfQrString = (nfResult && nfResult.qrString) ? nfResult.qrString : '';

  // ── Imprimir si el browser lo pidió explícitamente ───────────────────────
  // pos.print_request=true: solo index.html v39+. Sin este flag → browser imprime por su cuenta.
  var printDispatched = false;
  if (pos.print_request === true && pos.printerId) {
    printDispatched = imprimirTicketInternamente(data, correlativoFinal, pos.printerId, nfResult);
  }

  // ── Log de auditoría: creación de venta ────────────────────────────────────
  try {
    var actor = _audExtraerActor(data);
    auditarLog('VENTAS_CABECERA', idVenta, {
      usuario: actor.usuario, rol: actor.rol,
      source: 'ME_CREAR_VENTA',
      accion: 'crear',
      autorizadoPor: actor.autorizadoPor || null,
      ref: {
        correlativo: correlativoFinal,
        tipoDoc:     header.tipoDoc,
        formaPago:   header.metodo || 'EFECTIVO',
        total:       header.total,
        nfEstado:    nfEstado,
        idCaja:      pos.cajaId || ''
      },
      motivo: ''
    });
  } catch(_){}

  // ── Auto-registro de jornada en MOS (idempotente por nombre + fecha) ───────
  try { _registrarJornadaEnMOS(String(auth.vendedor || '')); } catch(eJ) {
    Logger.log('Auto-jornada MOS: ' + eJ.message);
  }

  // ── [v40] Push admin/master cuando se autoriza venta a CRÉDITO ─────────────
  // Cualquier admin con clave 8-dígitos que autoriza una venta a crédito
  // dispara una notificación push a TODOS los admin+master en MOS. Esto da
  // trazabilidad inmediata sin necesidad de imprimir copia interna.
  // idNotif=ME_CREDITO_AUTORIZADO → editable desde MOS/configuraciones.
  try {
    var _adminAuth = data.adminAuth || (header.adminAuth) || null;
    var _formaPago = header.metodo || '';
    if (_adminAuth && _adminAuth.via === 'PIN_8DIG' && _formaPago === 'CREDITO') {
      var _vendedorTxt = String(auth.vendedor || 'desconocido');
      var _clienteTxt  = _cliNombre || 'VARIOS';
      var _adminTxt    = String(_adminAuth.nombre || 'admin').replace(/^admin:/i, '');
      var _montoTxt    = parseFloat(header.total || 0).toFixed(2);
      var _obsTxt      = String(header.obs || '').trim();
      var titulo = '💳 Crédito autorizado · ' + correlativoFinal;
      var cuerpo = 'Cajero: ' + _vendedorTxt
                 + '\nCliente: ' + _clienteTxt
                 + '\nMonto: S/ ' + _montoTxt
                 + '\nAutoriza: ' + _adminTxt
                 + (_obsTxt ? '\nNota: ' + _obsTxt : '');
      _notificarMOS(titulo, cuerpo, null, 'ME_CREDITO_AUTORIZADO');
    }
  } catch(ePushC) { Logger.log('Push crédito autorizado: ' + ePushC.message); }

  // ── Alerta de recojo de efectivo (preventivo robos): cruzar S/500 inicial
  //    o cada S/250 después dispara push a admins+master. Solo evalúa cajas
  //    ABIERTAS y solo cuenta efectivo físico (EFECTIVO o parte EFE de MIXTO).
  try {
    if (pos.cajaId) _chequearAlertaEfectivo(pos.cajaId);
  } catch(eA) { Logger.log('Alerta efectivo: ' + eA.message); }

  return { idVenta: idVenta, correlativo: correlativoFinal, printDispatched: printDispatched,
           dedupVenta: false,
           nfEstado: nfEstado, nfHash: nfHash, nfEnlace: nfEnlace, nfQrString: nfQrString };
}

// Registra la jornada del vendedor en ProyectoMOS al procesar su primera venta del día.
// Idempotente: si ya existe una jornada con el mismo nombre y fecha no inserta duplicados.
function _registrarJornadaEnMOS(nombreVendedor) {
  if (!nombreVendedor) return;
  var mosSsId = PropertiesService.getScriptProperties().getProperty('MOS_SS_ID');
  if (!mosSsId) return;

  var tz    = Session.getScriptTimeZone();
  var fecha = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var ss    = SpreadsheetApp.openById(mosSsId);
  var sheet = ss.getSheetByName('JORNADAS');
  if (!sheet) return;

  // Idempotencia: verificar si ya existe la jornada hoy
  var tz2  = Session.getScriptTimeZone();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var fechaFila = data[i][1] instanceof Date
      ? Utilities.formatDate(data[i][1], tz2, 'yyyy-MM-dd')
      : String(data[i][1] || '').substring(0, 10);
    if (String(data[i][3]).toLowerCase() === nombreVendedor.toLowerCase() && fechaFila === fecha) return;
  }

  // Buscar montoBase en PERSONAL_MASTER de MOS
  var monto = 0;
  try {
    var pm    = ss.getSheetByName('PERSONAL_MASTER');
    if (pm) {
      var pmData = pm.getDataRange().getValues();
      var pmHdrs = pmData[0].map(function(h){ return String(h).trim(); });
      var idxNom = pmHdrs.indexOf('nombre');
      var idxMon = pmHdrs.indexOf('montoBase');
      for (var j = 1; j < pmData.length; j++) {
        if (String(pmData[j][idxNom]).toLowerCase() === nombreVendedor.toLowerCase()) {
          monto = parseFloat(pmData[j][idxMon]) || 0;
          break;
        }
      }
    }
  } catch(e2) {}

  sheet.appendRow([
    'JOR' + new Date().getTime(), fecha, '', nombreVendedor,
    'VENDEDOR', 'mosExpress', '', monto, '', 'AUTO', 'AUTO_VENTA'
  ]);
}

// Devuelve todas las ventas de hoy de la zona del cajero (filtradas por prefijos de serie)
function ventasHoyZona(prefijosStr, desdeStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("VENTAS_CABECERA");
  if (!sheet) return generarRespuestaError("VENTAS_CABECERA no encontrada");

  var prefijos = prefijosStr ? prefijosStr.split(',').map(function(p) { return p.trim(); }) : [];
  var data = sheet.getDataRange().getValues();
  var hoy   = new Date().toDateString();
  // Si se envía "desde" (ISO datetime de apertura de caja), filtrar por turno.
  // Si no, usar el filtro legacy de "hoy".
  var desde = (desdeStr && desdeStr.trim()) ? new Date(desdeStr.trim()) : null;
  var result = [];

  for (var i = 1; i < data.length; i++) {
    var fechaTk = data[i][1] instanceof Date ? data[i][1] : new Date(data[i][1]);
    if (desde) {
      // Solo tickets emitidos en o después de la apertura del turno
      if (fechaTk < desde) continue;
    } else {
      if (fechaTk.toDateString() !== hoy) continue;
    }
    var correlativo = String(data[i][9]);
    if (prefijos.length > 0) {
      var enZona = prefijos.some(function(p) { return correlativo.indexOf(p) === 0; });
      if (!enZona) continue;
    }
    result.push({
      id_venta:       data[i][0],
      fecha:          data[i][1],
      vendedor:       String(data[i][2] || ''),
      estacion:       String(data[i][3] || ''),
      cliente_doc:    String(data[i][4] || ''),
      cliente_nombre: String(data[i][5] || ''),
      total:          parseFloat(data[i][6]) || 0,
      tipo_doc:       String(data[i][7] || ''),
      forma_pago:     String(data[i][8] || ''),
      correlativo:    correlativo,
      id_caja:        String(data[i][10] || ''),
      id_dispositivo: String(data[i][11] || ''),
      status:         String(data[i][12] || ''),
      ref_local:      String(data[i][13] || ''),
      obs:            String(data[i][14] || '')
    });
  }

  return ContentService.createTextOutput(JSON.stringify({
    status: "success", ventas: result
  })).setMimeType(ContentService.MimeType.JSON);
}

function detalleVenta(idVenta) {
  if (!idVenta) return generarRespuestaError("id_venta requerido");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("VENTAS_DETALLE");
  if (!sheet) return generarRespuestaError("VENTAS_DETALLE no encontrada");

  var data = sheet.getDataRange().getValues();
  var items = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(idVenta)) {
      items.push({
        sku:      String(data[i][1] || ''),
        nombre:   String(data[i][2] || ''),
        cantidad: parseFloat(data[i][3]) || 0,
        precio:   parseFloat(data[i][4]) || 0,
        subtotal: parseFloat(data[i][5]) || 0
      });
    }
  }
  return ContentService.createTextOutput(JSON.stringify({
    status: "success", items: items
  })).setMimeType(ContentService.MimeType.JSON);
}

function verificarYAgregaCliente(doc, nombre, tipoDoc, direccion) {
  if (!doc) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("CLIENTES_FRECUENTES");
  if (!sheet) return;
  // [v40] Antes: si el cliente ya existía, salía sin hacer nada → si un cliente
  // viejo no tenía dirección guardada y ahora APISPeru la trajo, no se actualizaba.
  // Ahora: si existe Y la dirección está vacía Y llega una nueva no vacía, refrescar.
  var rangeAll = sheet.getDataRange().getValues();
  if (rangeAll.length < 2) {
    sheet.appendRow([doc, nombre, tipoDoc, new Date(), String(direccion || '')]);
    return;
  }
  var hdrs   = rangeAll[0].map(function(h){ return String(h).trim(); });
  var idxDoc = hdrs.indexOf('Documento');
  var idxNom = hdrs.indexOf('Nombre');
  var idxDir = hdrs.indexOf('Direccion');
  if (idxDoc < 0) return;
  for (var i = 1; i < rangeAll.length; i++) {
    if (String(rangeAll[i][idxDoc]) === String(doc)) {
      // Ya existe: refrescar dirección y nombre si llegan datos nuevos
      var dirVieja = idxDir >= 0 ? String(rangeAll[i][idxDir] || '').trim() : '';
      var dirNueva = String(direccion || '').trim();
      if (idxDir >= 0 && !dirVieja && dirNueva) {
        sheet.getRange(i + 1, idxDir + 1).setValue(dirNueva);
      }
      var nomViejo = idxNom >= 0 ? String(rangeAll[i][idxNom] || '').trim() : '';
      var nomNuevo = String(nombre || '').trim();
      // Solo actualizar nombre si el viejo era vacío (no pisar correcciones manuales)
      if (idxNom >= 0 && !nomViejo && nomNuevo) {
        sheet.getRange(i + 1, idxNom + 1).setValue(nomNuevo);
      }
      return;
    }
  }
  // No existe: agregar nuevo
  sheet.appendRow([doc, nombre, tipoDoc, new Date(), String(direccion || '')]);
}

// ── Correlativo O(1) con LockService ─────────────────────────────────────────
// Hoja CORRELATIVOS: encabezados Serie | Siguiente
// Crea la hoja automáticamente si no existe.
//
// IMPORTANTE — Anti race condition:
// Timeout subido a 30s y SIN fallback Math.max() (peligroso: bajo carga simultánea
// dos requests podían retornar el mismo número). Si timeout, se LANZA EXCEPCIÓN
// con código RESERVA_OCUPADA para que el frontend reintente con backoff exponencial.
// [v2.5.45] Limpia las filas huérfanas históricas de VENTAS_CABECERA generadas
// antes del guard. Las marca con Estado_Envio='HUERFANA_LIMPIADA' para que NO
// aparezcan en KPIs/reportes pero queden trazables en historialCambios para
// auditoría. Es admin-only — invocar desde editor GAS o desde MOS via PIN.
function limpiarVentasHuerfanas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('VENTAS_CABECERA');
  if (!sheet) return { ok: false, error: 'VENTAS_CABECERA no encontrada' };
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idxCorr = headers.indexOf('Correlativo');
  var idxEstado = headers.indexOf('Estado_Envio');
  if (idxCorr < 0 || idxEstado < 0) return { ok: false, error: 'Columnas Correlativo/Estado_Envio no encontradas' };
  var limpiadas = 0;
  var detalles = [];
  for (var i = 1; i < data.length; i++) {
    var corr = String(data[i][idxCorr] || '');
    var estadoActual = String(data[i][idxEstado] || '');
    if (corr.indexOf('undefined-') === 0 && estadoActual !== 'HUERFANA_LIMPIADA') {
      sheet.getRange(i + 1, idxEstado + 1).setValue('HUERFANA_LIMPIADA');
      limpiadas++;
      detalles.push({ fila: i + 1, correlativo: corr, idVenta: String(data[i][0] || '') });
    }
  }
  SpreadsheetApp.flush();
  Logger.log('[limpiarVentasHuerfanas] ' + limpiadas + ' filas marcadas como HUERFANA_LIMPIADA');
  return { ok: true, limpiadas: limpiadas, detalles: detalles };
}

// ============================================================
// [v2.5.47] Endpoints para Centro Tributario (consumidos por MOS)
// ============================================================
function tributarioVentasMes(mes, anio) {
  mes = parseInt(mes, 10); anio = parseInt(anio, 10);
  if (!mes || !anio) { var h = new Date(); mes = h.getMonth()+1; anio = h.getFullYear(); }
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('VENTAS_CABECERA');
  if (!sheet) return generarRespuestaError('VENTAS_CABECERA no encontrada');
  var data = sheet.getDataRange().getValues();
  var hdrs = data[0].map(function(h){ return String(h).trim(); });
  var idxFecha = hdrs.indexOf('Fecha');
  var idxTotal = hdrs.indexOf('Total');
  var idxTipo  = hdrs.indexOf('Tipo_Doc');
  var idxFP    = hdrs.indexOf('FormaPago');
  var idxNFE   = hdrs.indexOf('NF_Estado');
  var idxEstE  = hdrs.indexOf('Estado_Envio');

  var totalVentas = 0;
  var totalIGVEmitido = 0;
  var cpeTotal = 0, cpeEmitidos = 0, cpePendientes = 0, cpeErrores = 0, cpeAnulados = 0;

  for (var i = 1; i < data.length; i++) {
    var f = data[i][idxFecha];
    var d = f instanceof Date ? f : new Date(f);
    if (isNaN(d.getTime())) continue;
    if (d.getFullYear() !== anio || (d.getMonth() + 1) !== mes) continue;

    var tipo = String(data[i][idxTipo] || '');
    var fp   = String(data[i][idxFP] || '');
    var nfe  = String(data[i][idxNFE] || '');
    var ee   = String(data[i][idxEstE] || '');
    if (ee === 'HUERFANA_LIMPIADA') continue;
    if (fp === 'ANULADO') continue;

    var total = parseFloat(data[i][idxTotal]) || 0;
    if (tipo === 'BOLETA' || tipo === 'FACTURA') {
      cpeTotal++;
      totalVentas += total;
      // IGV emitido: total / 1.18 * 0.18 (asumiendo todo gravado por simplicidad)
      var igvAprox = Math.round((total - (total / 1.18)) * 100) / 100;
      totalIGVEmitido += igvAprox;
      if (nfe === 'EMITIDO')              cpeEmitidos++;
      else if (nfe === 'RECHAZADO_SUNAT') cpeErrores++;
      else if (nfe === 'ERROR')           cpeErrores++;
      else if (nfe === 'PENDIENTE' || nfe === '' || nfe === 'NA') cpePendientes++;
    }
    // NOTA_DE_VENTA suma a ventas pero NO a IGV ni a CPE
    if (tipo === 'NOTA_DE_VENTA' && fp !== 'POR_COBRAR' && fp !== 'CREDITO') {
      totalVentas += total;
    }
  }
  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    mes: mes, anio: anio,
    totalVentas:    Math.round(totalVentas * 100) / 100,
    totalIGVEmitido: Math.round(totalIGVEmitido * 100) / 100,
    cpeTotal: cpeTotal,
    cpeEmitidos: cpeEmitidos,
    cpePendientes: cpePendientes,
    cpeErrores: cpeErrores,
    cpeAnulados: cpeAnulados
  })).setMimeType(ContentService.MimeType.JSON);
}

function tributarioCPEMes(mes, anio) {
  mes = parseInt(mes, 10); anio = parseInt(anio, 10);
  if (!mes || !anio) { var h = new Date(); mes = h.getMonth()+1; anio = h.getFullYear(); }
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('VENTAS_CABECERA');
  if (!sheet) return generarRespuestaError('VENTAS_CABECERA no encontrada');
  var data = sheet.getDataRange().getValues();
  var hdrs = data[0].map(function(h){ return String(h).trim(); });
  var H = {};
  hdrs.forEach(function(h, i) { H[h] = i; });

  var lista = [];
  for (var i = 1; i < data.length; i++) {
    var f = data[i][H.Fecha];
    var d = f instanceof Date ? f : new Date(f);
    if (isNaN(d.getTime())) continue;
    if (d.getFullYear() !== anio || (d.getMonth() + 1) !== mes) continue;
    var tipo = String(data[i][H.Tipo_Doc] || '');
    if (tipo !== 'BOLETA' && tipo !== 'FACTURA') continue;
    var ee = String(data[i][H.Estado_Envio] || '');
    if (ee === 'HUERFANA_LIMPIADA') continue;

    lista.push({
      idVenta:     String(data[i][H.ID_Venta] || ''),
      fecha:       d.toISOString(),
      correlativo: String(data[i][H.Correlativo] || ''),
      tipo:        tipo,
      cliente:     String(data[i][H.Cliente_Nombre] || ''),
      clienteDoc:  String(data[i][H.Cliente_Doc] || ''),
      total:       parseFloat(data[i][H.Total]) || 0,
      formaPago:   String(data[i][H.FormaPago] || ''),
      nfEstado:    String(data[i][H.NF_Estado] || ''),
      nfHash:      String(data[i][H.NF_Hash] || ''),
      nfEnlace:    String(data[i][H.NF_Enlace] || '')
    });
  }
  lista.sort(function(a, b) { return new Date(b.fecha) - new Date(a.fecha); });
  return ContentService.createTextOutput(JSON.stringify({ status: 'success', cpe: lista })).setMimeType(ContentService.MimeType.JSON);
}

function tributarioReintentarCPE(idVenta) {
  if (!idVenta) return generarRespuestaError('idVenta requerido');
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('VENTAS_CABECERA');
  if (!sheet) return generarRespuestaError('VENTAS_CABECERA no encontrada');
  var data = sheet.getDataRange().getValues();
  var hdrs = data[0].map(function(h){ return String(h).trim(); });
  var H = {};
  hdrs.forEach(function(h, i) { H[h] = i; });

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][H.ID_Venta]) === String(idVenta)) {
      var corr = String(data[i][H.Correlativo] || '');
      var tipo = String(data[i][H.Tipo_Doc] || '');
      if (tipo !== 'BOLETA' && tipo !== 'FACTURA') return generarRespuestaError('Solo se reintentan BOLETA/FACTURA');
      var partes = corr.split('-');
      var serie = partes[0];
      var numero = parseInt(partes[partes.length - 1], 10);
      // Consultar estado actual
      var cons = consultarCPENubeFact(serie, numero, tipo);
      if (cons.ok) {
        // Actualizar la fila
        var nuevoEstado = cons.aceptada ? 'EMITIDO' : 'RECHAZADO_SUNAT';
        sheet.getRange(i + 1, H.NF_Estado + 1).setValue(nuevoEstado);
        if (cons.hash)   sheet.getRange(i + 1, H.NF_Hash + 1).setValue(cons.hash);
        if (cons.enlace) sheet.getRange(i + 1, H.NF_Enlace + 1).setValue(cons.enlace);
        SpreadsheetApp.flush();
        return ContentService.createTextOutput(JSON.stringify({
          status: 'success', idVenta: idVenta, nuevoEstado: nuevoEstado,
          aceptada: cons.aceptada, sunatDescription: cons.sunatDescription
        })).setMimeType(ContentService.MimeType.JSON);
      }
      return ContentService.createTextOutput(JSON.stringify({
        status: 'error', error: cons.error, noExiste: cons.noExiste || false
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  return generarRespuestaError('Venta ' + idVenta + ' no encontrada');
}

function obtenerSiguienteCorrelativoRapido(ss, serie) {
  var sheet = ss.getSheetByName('CORRELATIVOS');
  if (!sheet) {
    var sheetCab = ss.getSheetByName('VENTAS_CABECERA');
    var initial  = sheetCab ? obtenerSiguienteCorrelativo(sheetCab, serie) : 1;
    sheet = ss.insertSheet('CORRELATIVOS');
    sheet.appendRow(['Serie', 'Siguiente']);
    sheet.appendRow([serie, initial + 1]);
    return initial;
  }

  var lock = LockService.getScriptLock();
  var lockOK = false;
  try { lock.waitLock(30000); lockOK = true; }  // 30s timeout
  catch (e) {
    // NO fallback Math.max() — propaga error explícito para que frontend reintente.
    throw new Error('RESERVA_OCUPADA: no se pudo obtener correlativo tras 30s. Reintenta.');
  }

  try {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === serie) {
        var siguiente = parseInt(data[i][1], 10) || 1;
        // Anti-duplicado defensivo: si CORRELATIVOS quedó detrás (ej. por edición
        // manual de la hoja), avanzar hasta encontrar uno libre.
        while (correlativoYaExiste(ss, serie, siguiente)) { siguiente++; }
        sheet.getRange(i + 1, 2).setValue(siguiente + 1);
        SpreadsheetApp.flush();  // garantiza commit antes de soltar el lock
        return siguiente;
      }
    }
    // Serie nueva
    var sheetCab2 = ss.getSheetByName('VENTAS_CABECERA');
    var initial2  = sheetCab2 ? obtenerSiguienteCorrelativo(sheetCab2, serie) : 1;
    sheet.appendRow([serie, initial2 + 1]);
    SpreadsheetApp.flush();
    return initial2;
  } finally {
    if (lockOK) lock.releaseLock();
  }
}

// ============================================================
// [v2.5.58] PRE-RESERVA DE CORRELATIVOS (NV) — anti-LOCAL
// ============================================================
// El cliente reserva el correlativo en el momento que abre el modal
// de pago (con tipoDoc default NV). Cuando el cajero presiona Cobrar,
// la venta se procesa CON el correlativo ya en mano — impresión
// instantánea sin esperar GAS para numerar.
//
// Filosofía:
//   - Pre-reserva solo se usa para NV (gaps internos son OK)
//   - BOLETA/FACTURA NO usa pre-reserva (gaps requieren reporte SUNAT)
//   - LockService garantiza no-colisiones entre cajeros simultáneos
//   - Si la reserva se cancela o expira (>5min), queda como GAP en NV
//   - Trigger de limpieza cada hora marca ACTIVAS viejas como EXPIRADA
// ============================================================
var RES_CORR_HEADERS = [
  'idReserva', 'serie', 'numero', 'vendedor', 'deviceId',
  'reservadoAt', 'estado', 'usadoAt', 'idVenta'
];

function _getHojaReservasCorrelativos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('RESERVAS_CORRELATIVOS');
  if (!sh) {
    sh = ss.insertSheet('RESERVAS_CORRELATIVOS');
    sh.getRange(1, 1, 1, RES_CORR_HEADERS.length).setValues([RES_CORR_HEADERS]);
    sh.setFrozenRows(1);
  }
  return sh;
}

// ─────────────────────────────────────────────────────────────
// reservarCorrelativo({serie, vendedor, deviceId})
// LockService asegura no-colisión. Devuelve {numero, idReserva}.
// ─────────────────────────────────────────────────────────────
function reservarCorrelativo(data) {
  if (!data || !data.serie) return generarRespuestaError('serie requerida');
  var serie = String(data.serie).trim();
  var vendedor = String(data.vendedor || 'desconocido');
  var deviceId = String(data.deviceId || '');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // 1. Obtener número atómicamente
  var numero;
  try {
    numero = obtenerSiguienteCorrelativoRapido(ss, serie);
  } catch(e) {
    return generarRespuestaError('No se pudo reservar: ' + e.message);
  }
  // 2. Registrar la reserva
  var idReserva = 'RES-' + Date.now() + '-' + Math.floor(Math.random() * 1000);
  try {
    var sh = _getHojaReservasCorrelativos();
    sh.appendRow([
      idReserva, serie, numero, vendedor, deviceId,
      new Date(), 'ACTIVA', '', ''
    ]);
    SpreadsheetApp.flush();
  } catch(eR) {
    // Si falla el log, igual devolvemos el número (ya está reservado en CORRELATIVOS)
    Logger.log('Reserva log fallo: ' + eR.message);
  }
  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    numero: numero,
    idReserva: idReserva,
    correlativoFmt: serie + '-' + ('000000' + numero).slice(-6)
  })).setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────────────────────────
// cancelarReservaCorrelativo({idReserva})
// Marca la reserva como CANCELADA. El número NO se reusa (gap).
// Para NV el gap es OK. Para CPE no aplica porque no pre-reservamos.
// ─────────────────────────────────────────────────────────────
function cancelarReservaCorrelativo(data) {
  if (!data || !data.idReserva) return generarRespuestaError('idReserva requerida');
  var idReserva = String(data.idReserva);
  var sh = _getHojaReservasCorrelativos();
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === idReserva) {
      if (String(rows[i][6]) !== 'ACTIVA') {
        return ContentService.createTextOutput(JSON.stringify({
          status: 'success', yaProcesada: true, estado: String(rows[i][6])
        })).setMimeType(ContentService.MimeType.JSON);
      }
      sh.getRange(i + 1, 7).setValue('CANCELADA');
      sh.getRange(i + 1, 8).setValue(new Date());
      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({
        status: 'success', cancelada: true
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  return generarRespuestaError('idReserva no encontrada');
}

// ─────────────────────────────────────────────────────────────
// _consumirReserva(idReserva, idVentaFinal) — interno
// Llamado desde procesarVenta cuando viene idReserva en payload.
// Valida estado ACTIVA + marca USADA. Retorna {ok, numero} o {ok:false}.
// ─────────────────────────────────────────────────────────────
function _consumirReserva(idReserva, idVentaFinal, deviceIdActual) {
  var sh = _getHojaReservasCorrelativos();
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) !== String(idReserva)) continue;
    if (String(rows[i][6]) !== 'ACTIVA') {
      return { ok: false, error: 'Reserva no ACTIVA (estado: ' + rows[i][6] + ')' };
    }
    // Defensa: deviceId debe coincidir (anti-fraude entre cajeros)
    if (deviceIdActual && String(rows[i][4]) && String(rows[i][4]) !== String(deviceIdActual)) {
      return { ok: false, error: 'deviceId no coincide con reserva' };
    }
    sh.getRange(i + 1, 7).setValue('USADA');
    sh.getRange(i + 1, 8).setValue(new Date());
    sh.getRange(i + 1, 9).setValue(String(idVentaFinal || ''));
    SpreadsheetApp.flush();
    return {
      ok: true,
      serie: String(rows[i][1]),
      numero: parseInt(rows[i][2], 10)
    };
  }
  return { ok: false, error: 'idReserva no encontrada' };
}

// ─────────────────────────────────────────────────────────────
// _limpiarReservasViejas — trigger horario, marca ACTIVAS>5min
// como EXPIRADA. Auto-instala el trigger al primer uso.
// ─────────────────────────────────────────────────────────────
function _limpiarReservasViejas() {
  try {
    var sh = _getHojaReservasCorrelativos();
    var rows = sh.getDataRange().getValues();
    var ahora = Date.now();
    var LIMITE_MS = 5 * 60 * 1000; // 5 minutos
    var n = 0;
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][6]) !== 'ACTIVA') continue;
      var t = rows[i][5] instanceof Date ? rows[i][5].getTime() : new Date(rows[i][5]).getTime();
      if (isNaN(t)) continue;
      if (ahora - t > LIMITE_MS) {
        sh.getRange(i + 1, 7).setValue('EXPIRADA');
        sh.getRange(i + 1, 8).setValue(new Date());
        n++;
      }
    }
    if (n > 0) {
      SpreadsheetApp.flush();
      Logger.log('[ReservasCorrelativos] ' + n + ' reservas marcadas EXPIRADA');
    }
    return { ok: true, expiradas: n };
  } catch(e) { return { ok: false, error: e.message }; }
}

function _ensureTriggerLimpiarReservas() {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === '_limpiarReservasViejas') return;
    }
    ScriptApp.newTrigger('_limpiarReservasViejas')
      .timeBased().everyHours(1).create();
    Logger.log('Trigger _limpiarReservasViejas instalado');
  } catch(eT) { Logger.log('Trigger reservas: ' + eT.message); }
}

// Fallback O(n): scan de VENTAS_CABECERA. Solo se usa cuando CORRELATIVOS no existe todavía.
function obtenerSiguienteCorrelativo(sheet, serie) {
  var data = sheet.getDataRange().getValues();
  var prefijo = serie + "-";
  var maxCorrelativo = 0;
  for (var i = 1; i < data.length; i++) {
    var valorSerie = String(data[i][9]);
    if (valorSerie.indexOf(prefijo) === 0) {
      var num = parseInt(valorSerie.substring(prefijo.length), 10);
      if (!isNaN(num) && num > maxCorrelativo) maxCorrelativo = num;
    }
  }
  return maxCorrelativo + 1;
}

// Verifica si un correlativo ya existe en las últimas 100 filas (bounded O(100))
function correlativoYaExiste(ss, serie, numero) {
  var sheetCab = ss.getSheetByName('VENTAS_CABECERA');
  if (!sheetCab) return false;
  var objetivo  = serie + '-' + ('000000' + numero).slice(-6);
  var totalRows = sheetCab.getLastRow();
  if (totalRows < 2) return false;
  var desde = Math.max(2, totalRows - 99);
  var filas = sheetCab.getRange(desde, 10, totalRows - desde + 1, 1).getValues();
  for (var i = 0; i < filas.length; i++) {
    if (String(filas[i][0]) === objetivo) return true;
  }
  return false;
}
