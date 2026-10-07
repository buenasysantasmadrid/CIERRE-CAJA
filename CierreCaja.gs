function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    // Inicio de sesión con Google (ver "Acceso" al final de este archivo).
    // Todo lo demás necesita una sesión válida.
    if (data.accionIniciarSesion) {
      return ContentService
        .createTextOutput(JSON.stringify(iniciarSesion_(data.idToken)))
        .setMimeType(ContentService.MimeType.JSON);
    }
    if (!sesionValida_(data.sesion)) return respuestaSinSesion_();

    if (data.test) {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, msg: 'Conexión OK' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Leer una factura desde una foto (botón "Escanear" de la app). No
    // guarda nada en el Sheet: solo devuelve los datos leídos.
    if (data.accionEscanearFactura) {
      return ContentService
        .createTextOutput(JSON.stringify(escanearFactura_(data)))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Entender un gasto dictado por voz (botón "Dictar gasto"). Tampoco
    // guarda nada: solo devuelve los campos.
    if (data.accionDictarGasto) {
      return ContentService
        .createTextOutput(JSON.stringify(dictarGasto_(data)))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Leer uno o varios tickets de Glovo desde una foto (paso 3 del cierre).
    // Tampoco guarda nada: devuelve nº de pedido e importe de cada uno.
    if (data.accionEscanearGlovo) {
      return ContentService
        .createTextOutput(JSON.stringify(escanearTicketGlovo_(data)))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Los archivos del menú de tres rayas (Producción, Cambio de tinta...):
    // ver Menu.gs. No tocan los días ni los albaranes.
    var delMenu = accionDelMenu_(data);
    if (delMenu) return delMenu;

    // Lo mismo para un ingreso, un retiro o unas empanadas dictados.
    if (data.accionDictarMovimiento) {
      return ContentService
        .createTextOutput(JSON.stringify(dictarMovimiento_(data)))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Aviso del TPV 1: "Cambiado" o "Después lo cambio", para todos los
    // dispositivos. No toca los días ni los albaranes.
    if (data.accionAvisoTpv1) {
      return ContentService
        .createTextOutput(JSON.stringify(guardarAvisoTpv1_(data.mes, data.estado)))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Cualquier guardado cambia los días o los albaranes: que la próxima
    // consulta de listados los vuelva a armar.
    invalidarCacheListados_();

    if (data.accionCambiarFormaPago) {
      return ContentService
        .createTextOutput(JSON.stringify(cambiarFormaPagoAlbaran_(data)))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Planilla de Cierre de Caja DEL MES al que pertenece esta fecha (ver
    // Indice.gs) — se crea sola la primera vez que hace falta.
    var ss = planillaCierreCaja_(data.fecha);

    var TURNOS = [
      { key: 'mediodia', label: 'Mediodía' },
      { key: 'noche', label: 'Noche' }
    ];

    var registro = getOrCrearRegistro_(ss);
    var mov = getOrCrearMovimientos_(ss);

    // Una copia vieja de la caja (de un dispositivo que no vio el último
    // cambio que hizo otro) no pisa lo que ya está guardado.
    var desactualizada = cajaDesactualizada_(registro, data);
    if (desactualizada) {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, conflicto: true, error: desactualizada }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var ahora = new Date();
    TURNOS.forEach(function (turnoInfo) {
      var t = data[turnoInfo.key];
      if (!t) return;
      escribirRegistroYMovimientos_(registro, mov, data, t, turnoInfo.label, ahora);
    });

    // Cada uno de estos dos pasos va en su propio try/catch: si uno falla
    // (ej. falta la hoja plantilla "1", o hay un problema con la planilla
    // de Contabilidad) no debe impedir que se intente el otro, ni tapar que
    // Registro/Movimientos (arriba) ya se guardaron bien.
    var hojaDia = { ok: true };
    try {
      escribirHojaDelDiaExacta_(ss, data);
    } catch (errHoja) {
      hojaDia = { ok: false, error: String(errHoja) };
      Logger.log('escribirHojaDelDiaExacta_ error: ' + errHoja);
    }

    // Los demás dispositivos preguntan cada pocos segundos si esta caja
    // cambió (GET ?verCaja=fecha): se avisa ya, antes de Contabilidad y
    // Albaranes, que tardan.
    marcarCajaCambiada_(data.fecha, ahora);

    var contabilidad = escribirContabilidad_(data);
    var albaranes = escribirAlbaranes_(data); // ver Albaranes.gs
    var glovo = escribirGlovo_(data); // ver Glovo.gs

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, actualizadoEn: ahora.toISOString(), hojaDia: hojaDia, contabilidad: contabilidad, albaranes: albaranes, glovo: glovo }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getOrCrearRegistro_(ss) {
  var registro = ss.getSheetByName('Registro');
  if (!registro) {
    registro = ss.insertSheet('Registro');
    registro.appendRow([
      'ID', 'Estado', 'Fecha', 'Turno', 'Negocio', 'Fondo fijo', 'Total facturado', 'Empanadas',
      'Facturado neto', 'TPV 1', 'TPV 2', 'Ventas en efectivo', 'Ingresos efectivo',
      'Gastos efectivo', 'Retiros', 'Efectivo esperado', 'Efectivo contado', 'Diferencia',
      'Gastos no efectivo (info)', 'Cant. movimientos', 'Última actualización', 'Datos JSON (uso interno)', 'TPV 3'
    ]);
    registro.setFrozenRows(1);
  }
  return registro;
}

function getOrCrearMovimientos_(ss) {
  var mov = ss.getSheetByName('Movimientos');
  if (!mov) {
    mov = ss.insertSheet('Movimientos');
    mov.appendRow([
      'ID', 'ID Movimiento', 'Fecha', 'Turno', 'Tipo', 'Subtipo', 'Proveedor / Motivo', 'Responsable',
      'Nº Factura', 'Info', 'IVA (€)', 'Importe', 'Última actualización', 'Fecha factura'
    ]);
    mov.setFrozenRows(1);
  }
  return mov;
}

// Columna (0-based) con ese encabezado en la fila 1; si la pestaña es de
// antes y no la tiene, se agrega al final.
function asegurarColumna_(hoja, nombre) {
  var ultima = hoja.getLastColumn();
  var header = hoja.getRange(1, 1, 1, ultima).getValues()[0];
  var idx = header.indexOf(nombre);
  if (idx > -1) return idx;
  hoja.getRange(1, ultima + 1).setValue(nombre);
  return ultima;
}

// Columna (0-based) de "Fecha factura" en "Movimientos": la fecha de la
// factura de cada gasto, que puede no ser la de la caja en la que se cargó.
function columnaFechaFactura_(mov) {
  return asegurarColumna_(mov, 'Fecha factura');
}

// Fila de "Movimientos" con la fecha de factura en su columna.
function filaMovimiento_(valores, idxFechaFactura, fechaFactura) {
  var fila = valores.slice();
  while (fila.length <= idxFechaFactura) fila.push('');
  fila[idxFechaFactura] = fechaFactura || '';
  return fila;
}

// ============================================================================
// MIGRACIÓN MANUAL — correr una sola vez desde el editor de Apps Script
// (menú Ejecutar, elegir esta función) si la pestaña "Movimientos" es de
// una versión vieja del script (le faltan las columnas "ID Movimiento" e
// "Info", como pasa si su fila 1 no tiene esos textos). Sin esas columnas,
// funciones como "buscarFacturas" o "listarFacturas" (Albaranes) no pueden
// funcionar.
//
// No borra nada: renombra la pestaña vieja como respaldo y genera una
// "Movimientos" nueva, completa y con las columnas correctas, reconstruida
// desde "Registro" (la columna "Datos JSON (uso interno)" ahí sí tiene
// siempre el detalle completo y sin corrimientos de cada movimiento, así
// que es la fuente confiable para reconstruir esto). Reutiliza el mismo
// mapeo tipo/subtipo -> etiqueta que ya usa la reconciliación manual más
// abajo (TIPO_LABEL_MAP_ / SUBTIPO_LABEL_MAP_).
//
// Desde que Cierre de Caja es una planilla por mes (ver Indice.gs), esta
// función ya no tiene una "planilla activa" implícita: pasale el ID de la
// planilla del mes que haga falta reconstruir (Extensiones > Apps Script >
// elegir esta función > Ejecutar te va a pedir que edites este llamado, o
// se puede correr manualmente desde el editor con
// reconstruirMovimientosDesdeRegistro('<ID de esa planilla mensual>')).
// ============================================================================
function reconstruirMovimientosDesdeRegistro(idPlanillaCierreCaja) {
  if (!idPlanillaCierreCaja) throw new Error('Pasale el ID de la planilla mensual de Cierre de Caja a reconstruir, ej.: reconstruirMovimientosDesdeRegistro("1AbC...")');
  var ss = SpreadsheetApp.openById(idPlanillaCierreCaja);
  var registro = ss.getSheetByName('Registro');
  if (!registro) throw new Error('No existe la pestaña "Registro".');

  var viejo = ss.getSheetByName('Movimientos');
  if (viejo) {
    var nombreBackup = 'Movimientos (vieja, backup)';
    var backupPrevio = ss.getSheetByName(nombreBackup);
    if (backupPrevio) ss.deleteSheet(backupPrevio); // por si se corre esto más de una vez
    viejo.setName(nombreBackup);
  }
  var mov = getOrCrearMovimientos_(ss);
  var idxFechaFactura = columnaFechaFactura_(mov);

  var valores = registro.getDataRange().getValues();
  var header = valores[0];
  var idxFecha = header.indexOf('Fecha');
  var idxTurno = header.indexOf('Turno');
  var idxUltima = header.indexOf('Última actualización');
  var idxJSON = -1;
  for (var h = 0; h < header.length; h++) {
    if (String(header[h] || '').indexOf('Datos JSON') === 0) { idxJSON = h; break; }
  }
  if (idxJSON === -1) throw new Error('La pestaña "Registro" no tiene la columna "Datos JSON (uso interno)".');

  var filas = [];
  for (var i = 1; i < valores.length; i++) {
    var fila = valores[i];
    var jsonTexto = fila[idxJSON];
    if (!jsonTexto) continue;

    var turnoData;
    try { turnoData = JSON.parse(jsonTexto); } catch (errParse) { continue; }

    var idTurno = turnoData.id || fila[0];
    var fechaRaw = fila[idxFecha];
    var fechaStr = (fechaRaw instanceof Date)
      ? Utilities.formatDate(fechaRaw, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(fechaRaw || '');
    var turnoLabel = fila[idxTurno];
    var actualizado = idxUltima > -1 ? fila[idxUltima] : new Date();

    (turnoData.movimientos || []).forEach(function (m) {
      var proveedorMotivo = m.proveedor || m.motivo || '';
      if (PROVEEDORES_CON_DETALLE_LIBRE_.indexOf(m.proveedor) > -1 && m.proveedorDetalle) proveedorMotivo += ' — ' + m.proveedorDetalle;
      filas.push(filaMovimiento_([
        idTurno, m.id || '', fechaStr, turnoLabel,
        TIPO_LABEL_MAP_[m.tipo] || m.tipo || '', SUBTIPO_LABEL_MAP_[m.subtipo] || '',
        proveedorMotivo, m.responsable || '', m.factura || '', (m.info || ''),
        (m.iva != null ? m.iva : ''), m.importe, actualizado
      ], idxFechaFactura, m.tipo === 'gasto' ? (m.fecha || '') : ''));
    });
  }

  if (filas.length) {
    mov.getRange(2, 1, filas.length, filas[0].length).setValues(filas);
  }

  Logger.log('Reconstruidas ' + filas.length + ' filas en "Movimientos" a partir de "Registro".');
  return filas.length;
}

// Cada dispositivo manda, por turno, la "Última actualización" que vio la
// última vez que trajo o guardó ese turno (`base`) y su identificador
// (`dispositivo`). Si desde entonces otro dispositivo (o una edición a mano
// en el Sheet) guardó ese turno, lo que manda es una copia vieja: no se
// guarda, y la app trae lo último. Las filas de antes de esto (sin
// "escritoPor") y la app vieja (sin `dispositivo`) no se revisan.
function cajaDesactualizada_(registro, data) {
  if (!data.dispositivo) return null;
  var header = registro.getRange(1, 1, 1, registro.getLastColumn()).getValues()[0];
  var idxUltima = header.indexOf('Última actualización');
  var idxJSON = -1;
  for (var h = 0; h < header.length; h++) {
    if (String(header[h] || '').indexOf('Datos JSON') === 0) { idxJSON = h; break; }
  }
  if (idxUltima === -1 || idxJSON === -1) return null;

  var claves = ['mediodia', 'noche'];
  for (var i = 0; i < claves.length; i++) {
    var t = data[claves[i]];
    if (!t || !t.id) continue;
    var fila = buscarFilaPorId_(registro, t.id);
    if (fila === -1) continue;
    var valores = registro.getRange(fila, 1, 1, header.length).getValues()[0];
    var escritoPor = '';
    try { escritoPor = (JSON.parse(valores[idxJSON] || '{}') || {}).escritoPor || ''; } catch (errJson) {}
    if (!escritoPor || escritoPor === data.dispositivo) continue;
    var ultima = valores[idxUltima] instanceof Date ? valores[idxUltima].getTime() : 0;
    var base = t.base ? new Date(t.base).getTime() : 0;
    if (ultima > (isFinite(base) ? base : 0) + 1000) {
      return 'La caja del ' + data.fecha + ' (' + (claves[i] === 'noche' ? 'Noche' : 'Mediodía') + ') se cambió desde otro dispositivo después de la última vez que este la trajo.';
    }
  }
  return null;
}

function escribirRegistroYMovimientos_(registro, mov, data, t, turnoLabel, ahora) {
  var c = t.calc || {};
  var id = t.id || '';
  var estado = t.estado || 'Sincronizado';

  var datosJSON = JSON.stringify({
    id: t.id,
    fondoFijo: t.fondoFijo,
    totalFacturado: t.totalFacturado,
    tpv1: t.tpv1,
    tpv2: t.tpv2,
    tpv3: t.tpv3 || 0,
    fondoConteo: t.fondoConteo || null,
    escritoPor: data.dispositivo || '',
    denom: t.denom,
    movimientos: t.movimientosRaw || t.movimientos,
    glovo: t.glovo || [],
    calc: t.calc
  });

  var filaRegistro = [
    id, estado, data.fecha, turnoLabel, data.negocio,
    t.fondoFijo, t.totalFacturado, c.empanadas, c.facturadoNeto,
    t.tpv1, t.tpv2, c.ventasEfectivo, c.ingresoEfectivo,
    c.gastoEfectivo, c.egreso, c.esperado, c.totalContado, c.diferencia,
    c.gastoNoEfectivo, (t.movimientos || []).length, ahora || new Date(), datosJSON
  ];
  // "TPV 3" va al final (se agregó después): en las pestañas de antes se
  // crea la columna sola.
  var idxTpv3 = asegurarColumna_(registro, 'TPV 3');
  while (filaRegistro.length <= idxTpv3) filaRegistro.push('');
  filaRegistro[idxTpv3] = t.tpv3 || 0;

  var filaExistente = id ? buscarFilaPorId_(registro, id) : -1;
  if (filaExistente > -1) {
    registro.getRange(filaExistente, 1, 1, filaRegistro.length).setValues([filaRegistro]);
  } else {
    registro.appendRow(filaRegistro);
  }

  if (id) borrarFilasPorId_(mov, id);
  var idxFechaFactura = columnaFechaFactura_(mov);
  (t.movimientos || []).forEach(function (m) {
    var proveedorMotivo = m.proveedor || m.motivo || '';
    if (PROVEEDORES_CON_DETALLE_LIBRE_.indexOf(m.proveedor) > -1 && m.proveedorDetalle) proveedorMotivo += ' — ' + m.proveedorDetalle;
    mov.appendRow(filaMovimiento_([
      id, m.id || '', data.fecha, turnoLabel, m.tipo, m.subtipo,
      proveedorMotivo, m.responsable || '',
      m.factura || '', (m.info || ''), (m.iva != null ? m.iva : ''), m.importe, new Date()
    ], idxFechaFactura, m.fechaFactura));
  });
}

function nombreHojaDia_(fechaISO) {
  var partes = String(fechaISO || '').split('-');
  if (partes.length !== 3) return 'Sin fecha';
  return partes[2] + '-' + partes[1] + '-' + partes[0];
}

// "DD-MM-YYYY" (nombre de pestaña) -> "YYYY-MM-DD" (fecha interna) —
// inverso de nombreHojaDia_.
function fechaDesdeNombreHoja_(nombre) {
  var p = String(nombre || '').split('-');
  if (p.length !== 3) return null;
  return p[2] + '-' + p[1] + '-' + p[0];
}

// Copias de la plantilla "1" que quedaron sin renombrar ("Copia de 1",
// "Copia de 1 2", ...), de cuando dos guardados del mismo día crearon la
// pestaña a la vez.
function borrarCopiasSueltasPlantilla_(ss) {
  ss.getSheets().forEach(function (h) {
    if (/^(Copia de|Copy of) 1( \d+)?$/.test(h.getName())) ss.deleteSheet(h);
  });
}

// Dónde va cada cosa en la pestaña de un día. La plantilla "1" con TPV 3
// (desde el 28/09/2026) tiene dos columnas más (L:M) para TPV 3, así que
// EFECTIVO ANTIGUO pasó de M:R a O:T. Las pestañas hechas con la plantilla
// de antes siguen funcionando: ahí TPV 3 va en H7:I8 / H64:I66.
function celdasHojaDia_(hoja) {
  var conTpv3 = normalizarClave_(hoja.getRange('O3').getValue()) === 'EFECTIVO ANTIGUO';
  return conTpv3
    ? { conTpv3: true, colsEfectivoAntiguo: ['O', 'P', 'Q', 'R', 'S', 'T'], tpv3Md: 'M5', tpv3Nc: 'M61', tpv3NochePropio: 'M62' }
    : { conTpv3: false, colsEfectivoAntiguo: ['M', 'N', 'O', 'P', 'Q', 'R'], tpv3Md: 'I8', tpv3Nc: 'I65', tpv3NochePropio: 'I66' };
}

function escribirHojaDelDiaExacta_(ss, data) {
  var nombre = nombreHojaDia_(data.fecha);
  var hoja = ss.getSheetByName(nombre);

  if (!hoja) {
    // De a un guardado por vez: si llegan dos del mismo día juntos (ej. dos
    // dispositivos, o uno que vuelve a mandar una caja recién borrada), los
    // dos veían que no estaba la pestaña, los dos copiaban la plantilla y
    // la copia del segundo quedaba suelta como "Copia de 1".
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      hoja = ss.getSheetByName(nombre); // la pudo crear otro guardado mientras se esperaba
      if (!hoja) {
        var plantilla = ss.getSheetByName('1');
        if (!plantilla) {
          throw new Error('No se encontró la hoja "1" (la plantilla) en esta planilla, así que no se pudo crear la pestaña del día.');
        }
        borrarCopiasSueltasPlantilla_(ss);
        var copia = plantilla.copyTo(ss);
        try {
          copia.setName(nombre);
          hoja = copia;
        } catch (errNombre) {
          ss.deleteSheet(copia);
          hoja = ss.getSheetByName(nombre);
          if (!hoja) throw errNombre;
        }
        ss.setActiveSheet(hoja);
        ss.moveActiveSheet(ss.getNumSheets());
        SpreadsheetApp.flush();
      }
    } finally {
      lock.releaseLock();
    }
  }

  var md = data.mediodia || {};
  var nc = data.noche || {};
  var movsMd = md.movimientos || [];
  var movsNc = nc.movimientos || [];

  hoja.getRange('B2').setValue(textoFechaLarga_(data.fecha));
  var celdas = celdasHojaDia_(hoja);

  // ---- MEDIODÍA (posiciones de la plantilla "1" nueva) ----
  hoja.getRange('C45').setValue(md.totalFacturado || 0); // "TOTAL CIERRE SISTEMA"
  hoja.getRange('I5').setValue(md.tpv1 || 0);
  hoja.getRange('K5').setValue(md.tpv2 || 0);
  hoja.getRange(celdas.tpv3Md).setValue(md.tpv3 || 0);
  if (!celdas.conTpv3) {
    // Pestaña con la plantilla de antes (sin lugar para TPV 3): va en
    // H7:I8, que ahí está libre, y en I6 la suma de los 3 (la usa "SUMA
    // INGRESOS").
    hoja.getRange('H6:I7').setValues([
      ['TPV 1/2/3 DIA', (md.tpv1 || 0) + (md.tpv2 || 0) + (md.tpv3 || 0)],
      ['TPV 3', '']
    ]);
    hoja.getRange('H8').setValue('TOTAL');
  }
  hoja.getRange('C43').setValue(md.fondoFijo || 0);

  escribirFilasFijas_(hoja, 6, 16, ['A', 'B', 'C', 'D', 'E', 'F'], filtrarPorSubtipo_(movsMd, 'Efectivo').map(filaGasto_));
  escribirFilasFijas_(hoja, 5, 6, celdas.colsEfectivoAntiguo, filtrarPorSubtipo_(movsMd, 'Efectivo antiguo').map(filaGasto_));
  escribirFilasFijas_(hoja, 26, 15, ['A', 'B', 'C', 'D', 'E', 'F'], filtrarPorSubtiposNoEfectivo_(movsMd).map(filaGasto_));
  escribirFilasFijas_(hoja, 29, 5, ['H', 'I', 'J'], filtrarPorTipo_(movsMd, 'Egreso').map(filaMotivoImporte_));
  escribirFilasFijas_(hoja, 37, 4, ['H', 'I', 'J'], filtrarPorTipo_(movsMd, 'Ingreso').map(filaMotivoImporte_));
  escribirFilasFijas_(hoja, 50, 3, ['H', 'I', 'J'], filtrarPorTipo_(movsMd, 'Empanadas').map(filaEmpanada_));
  escribirDenomBilletes_(hoja, md.denom, 13);
  escribirDenomMonedas_(hoja, md.denom, 18);

  // ---- NOCHE (posiciones de la plantilla "1" nueva) ----
  // En I61/K61 va la lectura de TODO el día (lo que se carga en la app como
  // "TPV 1/2 — día completo"). En I62/K62 va la parte que le corresponde
  // solo a Noche (ese total menos lo que ya se cargó en Mediodía — se
  // calcula acá mismo, no se toma de "calc", para que quede bien aunque
  // esta pestaña se regenere desde una sincronización normal de la app,
  // no solo desde una edición manual). En I63 va la suma de esos dos.
  hoja.getRange('C87').setValue(nc.totalFacturado || 0); // "TOTAL CIERRE SISTEMA" (noche)
  hoja.getRange('I61').setValue(nc.tpv1 || 0);
  hoja.getRange('K61').setValue(nc.tpv2 || 0);
  hoja.getRange('C86').setValue(nc.fondoFijo || 0);

  var tpv1NochePropio = Math.max(0, (nc.tpv1 || 0) - (md.tpv1 || 0));
  var tpv2NochePropio = Math.max(0, (nc.tpv2 || 0) - (md.tpv2 || 0));
  hoja.getRange('I62').setValue(tpv1NochePropio);
  hoja.getRange('K62').setValue(tpv2NochePropio);
  // TPV 3 de Noche: el total del día completo y la parte de Noche, igual
  // que TPV 1 y 2.
  var tpv3NochePropio = Math.max(0, (nc.tpv3 || 0) - (md.tpv3 || 0));
  hoja.getRange(celdas.tpv3Nc).setValue(nc.tpv3 || 0);
  hoja.getRange(celdas.tpv3NochePropio).setValue(tpv3NochePropio);
  if (!celdas.conTpv3) {
    hoja.getRange('H64:H66').setValues([['TPV 3'], ['TOTAL'], ['TPV 3 NOCHE']]);
  }
  hoja.getRange('I63').setValue(tpv1NochePropio + tpv2NochePropio + tpv3NochePropio);
  hoja.getRange('I67').setValue( // TOT TARJETA del día (si Noche no cargó un TPV, vale el de Mediodía)
    Math.max(md.tpv1 || 0, nc.tpv1 || 0) + Math.max(md.tpv2 || 0, nc.tpv2 || 0) + Math.max(md.tpv3 || 0, nc.tpv3 || 0));

  escribirFilasFijas_(hoja, 63, 10, ['A', 'B', 'C', 'D', 'E', 'F'], filtrarPorSubtipo_(movsNc, 'Efectivo').map(filaGasto_));
  escribirFilasFijas_(hoja, 63, 6, celdas.colsEfectivoAntiguo, filtrarPorSubtipo_(movsNc, 'Efectivo antiguo').map(filaGasto_));
  escribirFilasFijas_(hoja, 75, 8, ['A', 'B', 'C', 'D', 'E', 'F'], filtrarPorSubtiposNoEfectivo_(movsNc).map(filaGasto_));
  escribirFilasFijas_(hoja, 88, 5, ['H', 'I', 'J'], filtrarPorTipo_(movsNc, 'Egreso').map(filaMotivoImporte_));
  escribirFilasFijas_(hoja, 97, 5, ['H', 'I', 'J'], filtrarPorTipo_(movsNc, 'Ingreso').map(filaMotivoImporte_));
  escribirFilasFijas_(hoja, 106, 5, ['H', 'I', 'J'], filtrarPorTipo_(movsNc, 'Empanadas').map(filaEmpanada_));
  escribirDenomBilletes_(hoja, nc.denom, 71);
  escribirDenomMonedas_(hoja, nc.denom, 76);
}

function escribirFilasFijas_(hoja, startRow, maxFilas, colLetras, filas) {
  var numCols = colLetras.length;
  var primeraCol = colLetras[0];
  var ultimaCol = colLetras[colLetras.length - 1];
  var matriz = [];
  for (var i = 0; i < maxFilas; i++) {
    if (i < filas.length) {
      matriz.push(filas[i]);
    } else {
      matriz.push(new Array(numCols).fill(''));
    }
  }
  hoja.getRange(primeraCol + startRow + ':' + ultimaCol + (startRow + maxFilas - 1)).setValues(matriz);
}

function filaGasto_(m) {
  var colB = (PROVEEDORES_CON_DETALLE_LIBRE_.indexOf(m.proveedor) > -1) ? (m.proveedorDetalle || '') : '';
  return [m.proveedor || m.motivo || '', colB, m.factura || '', (m.info || ''), (m.iva != null ? m.iva : ''), m.importe || 0];
}
function filaMotivoImporte_(m) {
  return [m.proveedor || m.motivo || '', m.responsable || '', m.importe || 0];
}
function filaEmpanada_(m) {
  return [m.proveedor || m.motivo || '', '', m.importe || 0];
}

function filtrarPorTipo_(movs, tipoLabel) {
  return (movs || []).filter(function (m) { return m.tipo === tipoLabel; });
}
function filtrarPorSubtipo_(movs, subtipoLabel) {
  return (movs || []).filter(function (m) { return m.subtipo === subtipoLabel; });
}
function filtrarPorSubtiposNoEfectivo_(movs) {
  var etiquetas = ['No efectivo', 'Tarjeta', 'No pagado', 'Transferencia'];
  return (movs || []).filter(function (m) { return etiquetas.indexOf(m.subtipo) > -1; });
}

function escribirDenomBilletes_(hoja, denom, startRow) {
  var valores = [100, 50, 20, 10, 5];
  var billetes = (denom && denom.billetes) || {};
  var matriz = valores.map(function (v) { return [Number(billetes[v]) || 0]; });
  hoja.getRange('J' + startRow + ':J' + (startRow + valores.length - 1)).setValues(matriz);
}

function escribirDenomMonedas_(hoja, denom, startRow) {
  var valores = [2, 1, 0.5, 0.2, 0.1, 0.05];
  var blister = (denom && denom.monedasBlister) || {};
  var sueltas = (denom && denom.monedasSueltas) || {};
  var matrizBlister = valores.map(function (v) { return [Number(blister[v]) || 0]; });
  var matrizSueltas = valores.map(function (v) { return [Number(sueltas[v]) || 0]; });
  hoja.getRange('I' + startRow + ':I' + (startRow + valores.length - 1)).setValues(matrizBlister);
  hoja.getRange('J' + startRow + ':J' + (startRow + valores.length - 1)).setValues(matrizSueltas);
}

function textoFechaLarga_(fechaISO) {
  var partes = String(fechaISO || '').split('-');
  if (partes.length !== 3) return '';
  var anio = parseInt(partes[0], 10);
  var mes = parseInt(partes[1], 10) - 1;
  var dia = parseInt(partes[2], 10);
  var fecha = new Date(anio, mes, dia);
  var diasSemana = ["Domingo", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"];
  var meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
  return diasSemana[fecha.getDay()] + ' ' + dia + ' de ' + meses[mes] + ' de ' + anio;
}

function buscarFilaPorId_(hoja, id) {
  var ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return -1;
  var valores = hoja.getRange(2, 1, ultimaFila - 1, 1).getValues();
  for (var i = 0; i < valores.length; i++) {
    if (valores[i][0] === id) return i + 2;
  }
  return -1;
}

function borrarFilasPorId_(hoja, id) {
  var ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return;
  var valores = hoja.getRange(2, 1, ultimaFila - 1, 1).getValues();
  for (var i = valores.length - 1; i >= 0; i--) {
    if (valores[i][0] === id) hoja.deleteRow(i + 2);
  }
}

function puntajeTurno_(t) {
  var movs = (t.movimientos || []).length;
  var c = t.calc || {};
  var tieneActividad = (c.totalContado || 0) !== 0 || (t.fondoFijo || 0) !== 0 ||
    (t.totalFacturado || 0) !== 0 || (t.tpv1 || 0) !== 0 || (t.tpv2 || 0) !== 0 || (t.tpv3 || 0) !== 0;
  return movs * 1000 + (tieneActividad ? 1 : 0);
}

function doGet(e) {
  var parametros = (e && e.parameter) || {};
  // La app pregunta si este Apps Script ya pide inicio de sesión.
  if (parametros.loginInfo) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, login: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // Todo lo demás necesita una sesión válida (ver "Acceso" al final).
  if (!sesionValida_(parametros.sesion)) return respuestaSinSesion_();

  // Diagnóstico rápido desde el navegador (GET ?diag=1, opcionalmente
  // &fecha=YYYY-MM-DD): lista el nombre exacto de cada pestaña de la
  // planilla de Cierre de Caja del mes de esa fecha (hoy por defecto), para
  // poder comprobar sin entrar al editor de Apps Script si existe la hoja
  // plantilla "1" (o si tiene un espacio/caracter invisible en el nombre).
  var diag = e && e.parameter && e.parameter.diag;
  if (diag) {
    var salidaDiag;
    try {
      var ssDiag = planillaCierreCaja_((e.parameter && e.parameter.fecha) || hoyISO_());
      salidaDiag = {
        ok: true,
        planilla: ssDiag.getName(),
        pestanas: ssDiag.getSheets().map(function (h) { return h.getName(); })
      };
    } catch (errDiagCierre) {
      salidaDiag = { ok: false, error: String(errDiagCierre) };
    }
    return ContentService
      .createTextOutput(JSON.stringify(salidaDiag))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // Diagnóstico de Contabilidad (GET ?diagContabilidad=1): intenta abrir la
  // planilla de Contabilidad del año (la que figura en el Índice) y listar
  // sus pestañas, sin escribir nada ni crear la planilla si no existe. Sirve
  // para ver el error real (permiso, ID incorrecto, etc.) sin tener que
  // volver a sincronizar un día desde la app.
  var diagContabilidad = e && e.parameter && e.parameter.diagContabilidad;
  if (diagContabilidad) {
    try {
      var anioDiag = parseInt((e.parameter && e.parameter.anio) || new Date().getFullYear(), 10);
      var idDiag = buscarEnIndice_('CONTABILIDAD', String(anioDiag));
      if (!idDiag) throw new Error('La pestaña "Archivos" del Índice no tiene una fila CONTABILIDAD para ' + anioDiag + ' (se crea sola al guardar el primer día de ese año).');
      var ssC = SpreadsheetApp.openById(idDiag);
      var mesActual = MESES_MAYUS_[new Date().getMonth()];
      var hojaMes = ssC.getSheetByName(mesActual);
      return ContentService
        .createTextOutput(JSON.stringify({
          ok: true,
          anio: anioDiag,
          idUsado: idDiag,
          planilla: ssC.getName(),
          pestanas: ssC.getSheets().map(function (h) { return h.getName(); }),
          existePestanaMesActual: !!hojaMes,
          filaHoy: hojaMes ? hojaMes.getRange('A' + (3 + (new Date().getDate() - 1)) + ':O' + (3 + (new Date().getDate() - 1))).getValues()[0] : null
        }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (errDiag) {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, error: String(errDiag) }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }
  // Prueba manual: abrir esta URL con ?testContabilidad=1 (opcionalmente
  // &fecha=YYYY-MM-DD) para repetir, con los datos reales ya guardados de
  // ese día, el mismo escribirContabilidad_ que corre en cada guardado
  // normal — útil para diagnosticar sin depender del registro de
  // Ejecuciones (que no muestra estos errores porque quedan atrapados
  // adentro de esa función) ni de las herramientas de desarrollador del
  // navegador.
  var testContabilidad = e && e.parameter && e.parameter.testContabilidad;
  if (testContabilidad) {
    var fechaTest = (e.parameter && e.parameter.fecha) || hoyISO_();
    var resultadoTest;
    try {
      var ssCierreTest = planillaCierreCaja_(fechaTest);
      var diaTest = obtenerDiaJSON_(ssCierreTest, fechaTest);
      resultadoTest = escribirContabilidad_({ fecha: fechaTest, mediodia: diaTest.mediodia, noche: diaTest.noche });
      resultadoTest.diaEncontradoEnCierreDeCaja = !!diaTest.encontrado;
    } catch (errTest) {
      resultadoTest = { ok: false, error: String(errTest) };
    }
    return ContentService
      .createTextOutput(JSON.stringify(resultadoTest))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Pregunta rápida (sin abrir ninguna planilla): ¿cuándo se guardó por
  // última vez la caja de esta fecha? La app la hace cada pocos segundos y
  // solo trae la caja entera si cambió.
  if (parametros.verCaja) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, v: versionCaja_(parametros.verCaja) }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var avisoTpv1Mes = e && e.parameter && e.parameter.avisoTpv1;
  if (avisoTpv1Mes) {
    return ContentService
      .createTextOutput(JSON.stringify(leerAvisoTpv1_(avisoTpv1Mes)))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var listarDias = e && e.parameter && e.parameter.listarDias;
  if (listarDias) {
    return ContentService
      .createTextOutput(respuestaConCache_('dias', listarDiasConDatos_))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var listarFacturas = e && e.parameter && e.parameter.listarFacturas;
  if (listarFacturas) {
    var proveedorFiltro = (e.parameter && e.parameter.proveedor) || '';
    var detalleFiltro = (e.parameter && e.parameter.detalle) || '';
    return ContentService
      .createTextOutput(respuestaConCache_('facturas_' + proveedorFiltro + '_' + detalleFiltro, function () {
        return listarFacturasProveedores_(proveedorFiltro, detalleFiltro);
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var fecha = e && e.parameter && e.parameter.fecha;
  if (fecha) {
    var resultadoDia;
    try {
      resultadoDia = obtenerDiaJSON_(planillaCierreCaja_(fecha), fecha);
    } catch (errSs) {
      resultadoDia = { ok: false, error: String(errSs) };
    }
    return ContentService
      .createTextOutput(JSON.stringify(resultadoDia))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, msg: 'API de Cierre de Caja activa. Usá POST para enviar datos, GET ?fecha=YYYY-MM-DD para leer un día ya guardado, o GET ?listarFacturas=1 (con ?proveedor=<nombre> opcional) para ver todas las facturas y albaranes de proveedores.' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- versión de cada caja (para que los dispositivos se enteren solos) ----------
// Cada guardado deja en la memoria rápida de Google la hora del último
// cambio de esa fecha. Si no está (pasaron 6 horas, o se editó a mano en
// "Registro"), la app igual consulta la caja entera cada tanto.
function marcarCajaCambiada_(fecha, cuando) {
  if (!fecha) return;
  try { CacheService.getScriptCache().put('caja_ver_' + fecha, (cuando || new Date()).toISOString(), 21600); } catch (e) {}
}

function versionCaja_(fecha) {
  try { return CacheService.getScriptCache().get('caja_ver_' + fecha) || null; } catch (e) { return null; }
}

// ---------- memoria rápida de los listados (días y albaranes) ----------
// Armar la lista de días o la de albaranes obliga a abrir varias planillas
// mensuales (varios segundos). Se guarda la respuesta 10 minutos en la
// memoria rápida de Google y se tira en cuanto algo cambia: cualquier
// guardado desde la app (doPost), una edición a mano (onEdit) o las
// herramientas del menú (borrar / mover días). Para tirar todo de una vez
// sin tener que conocer cada clave, las claves llevan un número de versión:
// invalidar = cambiar la versión.
var CACHE_LISTADOS_SEGUNDOS_ = 600;

function versionCacheListados_() {
  var cache = CacheService.getScriptCache();
  var v = cache.get('listados_version');
  if (!v) {
    v = String(Date.now());
    cache.put('listados_version', v, 21600);
  }
  return v;
}

function invalidarCacheListados_() {
  try { CacheService.getScriptCache().put('listados_version', String(Date.now()), 21600); } catch (e) {}
}

// Devuelve el JSON de fn(), usando la copia guardada si hay una. Solo se
// guardan las respuestas correctas (ok: true) y que entren en la memoria
// rápida (máximo 100 KB por valor).
function respuestaConCache_(nombre, fn) {
  var cache = CacheService.getScriptCache();
  var clave = 'listado_' + versionCacheListados_() + '_' + nombre;
  if (clave.length > 250) clave = clave.slice(0, 250);
  var guardado = cache.get(clave);
  if (guardado) return guardado;
  var resultado = fn();
  var texto = JSON.stringify(resultado);
  if (resultado && resultado.ok) {
    try { cache.put(clave, texto, CACHE_LISTADOS_SEGUNDOS_); } catch (e) {}
  }
  return texto;
}

function coincideProveedor_(textoProv, proveedor, detalleNorm) {
  if (!proveedor) return true;
  if (proveedor === 'Servicios') return esProveedorServicio_(textoProv);
  if (proveedor === 'Varios') {
    return textoProv.indexOf('Varios') === 0 && (!detalleNorm || textoProv.toLowerCase().indexOf(detalleNorm) > -1);
  }
  return (textoProv === proveedor) || (textoProv.indexOf(proveedor + ' —') === 0);
}

// Trae TODO lo cargado como "Gasto" — pagado o no — en UNA planilla mensual
// de Cierre de Caja, para la pantalla de Albaranes, donde se puede
// consultar y volver a abrir cualquier proveedor/movimiento para editarlo,
// o marcarlo como pagado.
//
// NOTA: esto sigue siendo la búsqueda de Albaranes "vieja" (todavía no está
// construida la pantalla nueva pensada para esto — ver Indice.gs, tipo
// ALBARANES). Por ahora, para que no deje de funcionar al partir Cierre de
// Caja en un archivo por mes, listarFacturasProveedores_ (más abajo) junta
// esto de los últimos MESES_HISTORIAL_ meses.
function listarFacturasProveedoresEnPlanilla_(ss, proveedor, detalleNorm) {
    var mov = ss.getSheetByName('Movimientos');
    if (!mov) return { facturas: [] };

    var valores = mov.getDataRange().getValues();
    var header = valores[0];
    var idxIdTurno = header.indexOf('ID');
    var idxIdMov = header.indexOf('ID Movimiento');
    var idxFecha = header.indexOf('Fecha');
    var idxTurno = header.indexOf('Turno');
    var idxTipo = header.indexOf('Tipo');
    var idxSubtipo = header.indexOf('Subtipo');
    var idxProvMotivo = header.indexOf('Proveedor / Motivo');
    var idxFactura = header.indexOf('Nº Factura');
    var idxInfo = header.indexOf('Info');
    var idxIva = header.indexOf('IVA (€)');
    var idxImporte = header.indexOf('Importe');
    var idxFechaFactura = header.indexOf('Fecha factura');

    if (idxIdMov === -1) {
      return { facturas: [], error: 'La pestaña "Movimientos" es de una versión anterior y no tiene la columna "ID Movimiento". Volvé a sincronizar un cambio desde la app para que se agregue.' };
    }

    var resultado = [];

    for (var i = 1; i < valores.length; i++) {
      var fila = valores[i];
      if (fila[idxTipo] !== 'Gasto') continue;

      // La pata "Efectivo antiguo" que sale de la caja de HOY (ver
      // aplicarPagoEfectivoAntiguoAlbaran en el front-end) es un movimiento
      // aparte, solo para descontar esa caja — no un albarán en sí. El
      // albarán original ya queda marcado como "Efectivo antiguo" con su
      // "PAGADO <fecha>" en Info, así que esta pata se oculta acá para no
      // mostrar el mismo pago dos veces.
      var infoFila = String(fila[idxInfo] || '');
      if (infoFila.indexOf('Pago de factura ') === 0) continue;

      var textoProv = String(fila[idxProvMotivo] || '');
      if (!coincideProveedor_(textoProv, proveedor, detalleNorm)) continue;

      var fechaRaw = fila[idxFecha];
      var fechaCaja = (fechaRaw instanceof Date) ? Utilities.formatDate(fechaRaw, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(fechaRaw || '');
      var fechaFacturaRaw = idxFechaFactura > -1 ? fila[idxFechaFactura] : '';
      var fechaFactura = (fechaFacturaRaw instanceof Date) ? Utilities.formatDate(fechaFacturaRaw, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(fechaFacturaRaw || '');
      resultado.push({
        idMovimiento: fila[idxIdMov],
        idTurno: fila[idxIdTurno],
        fecha: fechaCaja, // la caja en la que se cargó (para abrirlo o pagarlo)
        fechaFactura: fechaFactura || fechaCaja,
        turno: fila[idxTurno],
        subtipo: fila[idxSubtipo],
        proveedorTexto: textoProv,
        factura: fila[idxFactura],
        info: String(fila[idxInfo] || ''),
        iva: fila[idxIva],
        importe: fila[idxImporte]
      });
    }

    resultado.sort(function (a, b) { return a.fechaFactura < b.fechaFactura ? 1 : (a.fechaFactura > b.fechaFactura ? -1 : 0); });

    return { facturas: resultado };
}

// Junta listarFacturasProveedoresEnPlanilla_ de los últimos MESES_HISTORIAL_
// meses (los que ya existen — ver planillasCierreCajaRecientes_ en
// Indice.gs) para que Albaranes no pierda historial al partir Cierre de
// Caja en un archivo por mes.
function listarFacturasProveedores_(proveedor, detalle) {
  try {
    var detalleNorm = (detalle || '').toLowerCase().trim();
    var planillas = planillasCierreCajaRecientes_(MESES_HISTORIAL_);
    var resultado = [];
    var error = null;

    // Un mismo movimiento puede aparecer en dos planillas de mes (un día
    // que quedó guardado en la del mes de al lado) o repetido en la misma:
    // se muestra una sola vez, preferentemente el de la planilla de su mes.
    var vistos = {};
    planillas.forEach(function (p) {
      var r = listarFacturasProveedoresEnPlanilla_(p.ss, proveedor, detalleNorm);
      if (r.error) error = r.error;
      r.facturas.forEach(function (f) {
        var claves = [
          'id:' + f.idMovimiento,
          'datos:' + [f.fecha, f.turno, normalizarClave_(f.proveedorTexto), normalizarClave_(f.factura), Number(f.importe) || 0].join('|')
        ];
        var deSuMes = String(f.fecha).slice(0, 7) === p.periodo;
        var previo = null;
        claves.forEach(function (k) { if (vistos[k]) previo = vistos[k]; });
        if (previo) {
          if (deSuMes && !previo.deSuMes) {
            resultado[previo.pos] = f;
            previo.deSuMes = true;
          }
          return;
        }
        var reg = { pos: resultado.length, deSuMes: deSuMes };
        claves.forEach(function (k) { vistos[k] = reg; });
        resultado.push(f);
      });
    });

    // Solo los últimos 2 meses (este y el anterior, por fecha de factura).
    // Los albaranes sin pagar más viejos se siguen mostrando, para poder
    // pagarlos.
    var hoy = new Date();
    var desde = Utilities.formatDate(new Date(hoy.getFullYear(), hoy.getMonth() - 1, 1, 12), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    resultado = resultado.filter(function (f) {
      return f.fechaFactura >= desde || f.subtipo === 'No pagado';
    });

    resultado.sort(function (a, b) { return a.fechaFactura < b.fechaFactura ? 1 : (a.fechaFactura > b.fechaFactura ? -1 : 0); });

    var salida = { ok: true, facturas: resultado };
    if (error) salida.error = error;
    return salida;
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

function obtenerDiaJSON_(ss, fecha) {
  try {
    var registro = ss.getSheetByName('Registro');
    if (!registro) {
      return { ok: true, encontrado: false, fecha: fecha, negocio: '', mediodia: null, noche: null };
    }

    var valores = registro.getDataRange().getValues();
    var header = valores[0];
    var idxFecha = header.indexOf('Fecha');
    var idxTurno = header.indexOf('Turno');
    var idxNegocio = header.indexOf('Negocio');
    var idxUltima = header.indexOf('Última actualización');
    var idxJSON = -1;
    for (var h = 0; h < header.length; h++) {
      if (String(header[h] || '').indexOf('Datos JSON') === 0) { idxJSON = h; break; }
    }

    if (idxJSON === -1) {
      return {
        ok: true, encontrado: false, fecha: fecha, negocio: '', mediodia: null, noche: null,
        error: 'La pestaña "Registro" es de una versión anterior del script y no tiene la columna "Datos JSON (uso interno)". Volvé a sincronizar un cambio desde la app para que se agregue.'
      };
    }

    var resultado = { ok: true, encontrado: false, fecha: fecha, negocio: '', mediodia: null, noche: null };
    var mejorPorTurno = {};

    for (var i = 1; i < valores.length; i++) {
      var fila = valores[i];
      var filaFechaRaw = fila[idxFecha];
      var filaFecha = (filaFechaRaw instanceof Date)
        ? Utilities.formatDate(filaFechaRaw, Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : String(filaFechaRaw || '');
      if (filaFecha !== fecha) continue;

      var jsonTexto = fila[idxJSON];
      if (!jsonTexto) continue;

      try {
        var turnoData = JSON.parse(jsonTexto);
        var turnoKey = (String(fila[idxTurno]).toLowerCase() === 'noche') ? 'noche' : 'mediodia';
        var puntaje = puntajeTurno_(turnoData);

        if (!mejorPorTurno[turnoKey] || puntaje >= mejorPorTurno[turnoKey].puntaje) {
          var actualizadoRaw = idxUltima > -1 ? fila[idxUltima] : null;
          turnoData.actualizadoEn = (actualizadoRaw instanceof Date) ? actualizadoRaw.toISOString() : null;
          mejorPorTurno[turnoKey] = { puntaje: puntaje, datos: turnoData, fila: i + 1 };
          resultado.negocio = fila[idxNegocio] || resultado.negocio;
        }
      } catch (errParse) {
      }
    }

    if (mejorPorTurno.mediodia) { resultado.mediodia = mejorPorTurno.mediodia.datos; resultado.encontrado = true; }
    if (mejorPorTurno.noche) { resultado.noche = mejorPorTurno.noche.datos; resultado.encontrado = true; }

    return resultado;

  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

// ============================================================================
// PARTE 3 — Ida y vuelta con el Sheet (edición manual → app)
// ============================================================================
// Dos superficies de edición manual:
//
// 1) Pestaña "Registro": corregir a mano "Fondo fijo", "Total facturado",
//    "TPV 1" o "TPV 2" de una fila ya existente.
// 2) La pestaña bonita de cada día (la que tiene el nombre de la fecha, ej.
//    "09-09-2026"): corregir el conteo de billetes/monedas, el "Fondo
//    fijo", el "Total Cierre Sistema" (Total facturado) o el TPV 1 / TPV 2
//    de Mediodía o Noche.
//
// En los dos casos, el disparador recalcula todo lo que depende de esos
// números (igual que hace la app), actualiza "Datos JSON" de la fila en
// Registro y vuelve a generar la pestaña bonita del día — así la próxima
// vez que la app consulte ese día (cada 20 segundos, o al abrirlo), ve el
// cambio.
//
// OJO — alcance de esto: los MOVIMIENTOS (proveedores, ingresos, egresos,
// empanadas — tanto en la pestaña "Movimientos" como en las tablas de la
// pestaña bonita) NO son una superficie de edición. Cada movimiento tiene
// un ID interno que la app usa
// para poder editarlo/borrarlo y para el flujo de "Efectivo antiguo" (pagar
// una factura vieja); esas tablas no muestran ese ID, así que reconstruir
// los movimientos desde ahí obligaría a inventarles un ID nuevo cada vez, y
// eso rompería en silencio esos dos flujos. Para cargar o corregir
// movimientos, seguís usando la app.
//
// Esto son simples triggers (onEdit) — se activan solo con la edición de
// una persona real en la hoja; los cambios que hace el propio script (como
// los que estos mismos disparadores escriben) no los vuelven a activar, así
// que no hay riesgo de bucle infinito.
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var hoja = e.range.getSheet();
    var nombreHoja = hoja.getName();
    if (nombreHoja === 'Registro') {
      invalidarCacheListados_();
      manejarEdicionRegistro_(e, hoja);
    } else if (/^\d{2}-\d{2}-\d{4}$/.test(nombreHoja)) {
      invalidarCacheListados_();
      manejarEdicionHojaDelDia_(e, hoja, nombreHoja);
      var p = nombreHoja.split('-');
      marcarCajaCambiada_(p[2] + '-' + p[1] + '-' + p[0]);
    }
  } catch (err) {
    Logger.log('onEdit error: ' + err);
  }
}

function manejarEdicionRegistro_(e, hoja) {
  var fila = e.range.getRow();
  if (fila === 1) return; // fila de encabezados

  var header = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  var idxFondoFijo = header.indexOf('Fondo fijo');
  var idxTotalFacturado = header.indexOf('Total facturado');
  var idxTpv1 = header.indexOf('TPV 1');
  var idxTpv2 = header.indexOf('TPV 2');
  var idxTpv3 = header.indexOf('TPV 3');
  var idxJSON = -1;
  for (var h = 0; h < header.length; h++) {
    if (String(header[h] || '').indexOf('Datos JSON') === 0) { idxJSON = h; break; }
  }
  if (idxJSON === -1) return; // pestaña vieja, sin dónde guardar el resultado

  var colsEditables = [idxFondoFijo, idxTotalFacturado, idxTpv1, idxTpv2, idxTpv3]
    .filter(function (i) { return i > -1; })
    .map(function (i) { return i + 1; });
  var colInicio = e.range.getColumn();
  var colFin = colInicio + e.range.getNumColumns() - 1;
  var tocaAlgunaEditable = colsEditables.some(function (c) { return c >= colInicio && c <= colFin; });
  if (!tocaAlgunaEditable) return; // se editó otra columna (por ej. la propia "Datos JSON"): no hacer nada

  var idxId = header.indexOf('ID');
  var idxFecha = header.indexOf('Fecha');
  var idxTurno = header.indexOf('Turno');
  var idxNegocio = header.indexOf('Negocio');

  var filaValores = hoja.getRange(fila, 1, 1, header.length).getValues()[0];
  var idEditado = filaValores[idxId];
  var fechaRaw = filaValores[idxFecha];
  var fechaStr = (fechaRaw instanceof Date)
    ? Utilities.formatDate(fechaRaw, Session.getScriptTimeZone(), 'yyyy-MM-dd')
    : String(fechaRaw || '');
  if (!idEditado || !fechaStr) return;

  var jsonTexto = filaValores[idxJSON];
  if (!jsonTexto) return; // fila que todavía no tiene "Datos JSON" (nunca se sincronizó desde la app)
  var turnoEditado;
  try { turnoEditado = JSON.parse(jsonTexto); } catch (errParse) { return; }

  // Se trabaja sobre el JSON de la fila que efectivamente se editó (no
  // sobre "la mejor fila" que arma obtenerDiaJSON_ para esa fecha/turno) —
  // así funciona aunque haya filas duplicadas para el mismo día/turno.
  // Antes, si la fila editada no era la que ganaba el desempate, la edición
  // se perdía en silencio, sin ningún aviso.
  turnoEditado.id = idEditado;
  turnoEditado.fondoFijo = Number(filaValores[idxFondoFijo]) || 0;
  turnoEditado.totalFacturado = Number(filaValores[idxTotalFacturado]) || 0;
  turnoEditado.tpv1 = Number(filaValores[idxTpv1]) || 0;
  turnoEditado.tpv2 = Number(filaValores[idxTpv2]) || 0;
  if (idxTpv3 > -1) turnoEditado.tpv3 = Number(filaValores[idxTpv3]) || 0;

  var turnoKey = (String(filaValores[idxTurno]).toLowerCase() === 'noche') ? 'noche' : 'mediodia';
  var diaExistente = obtenerDiaJSON_(hoja.getParent(), fechaStr) || {};

  var dia = {
    mediodia: turnoKey === 'mediodia' ? turnoEditado : (diaExistente.mediodia || null),
    noche: turnoKey === 'noche' ? turnoEditado : (diaExistente.noche || null)
  };
  recalcularDia_(dia);
  guardarDiaCompleto_(hoja.getParent(), fechaStr, filaValores[idxNegocio] || diaExistente.negocio, dia);
}

// "J" -> 10, "AA" -> 27, etc. (Apps Script no trae un helper corto para esto)
function columnaLetraANumero_(letra) {
  var n = 0;
  for (var i = 0; i < letra.length; i++) {
    n = n * 26 + (letra.charCodeAt(i) - 64);
  }
  return n;
}

// Filas donde vive cada tabla de conteo (mismos números que usan
// escribirDenomBilletes_/escribirDenomMonedas_ al escribir).
var BILLETES_VALORES_ = [100, 50, 20, 10, 5];
var MONEDAS_VALORES_ = [2, 1, 0.5, 0.2, 0.1, 0.05];

function manejarEdicionHojaDelDia_(e, hoja, nombreHoja) {
  var fechaStr = fechaDesdeNombreHoja_(nombreHoja);
  if (!fechaStr) return;

  var col = e.range.getColumn();
  var colFin = col + e.range.getNumColumns() - 1;
  var fila = e.range.getRow();
  var filaFin = fila + e.range.getNumRows() - 1;

  function tocaCelda_(colLetra, filaObjetivo) {
    var colNum = columnaLetraANumero_(colLetra);
    return colNum >= col && colNum <= colFin && filaObjetivo >= fila && filaObjetivo <= filaFin;
  }
  function tocaRango_(colLetra, filaDesde, filaHasta) {
    var colNum = columnaLetraANumero_(colLetra);
    return colNum >= col && colNum <= colFin && filaHasta >= fila && filaDesde <= filaFin;
  }

  var celdas = celdasHojaDia_(hoja);
  function tocaA1_(a1) { return tocaCelda_(a1.replace(/\d+/, ''), Number(a1.replace(/\D+/, ''))); }

  var tocaMd = tocaCelda_('C', 45) || tocaCelda_('I', 5) || tocaCelda_('K', 5) || tocaA1_(celdas.tpv3Md) || tocaCelda_('C', 43) ||
    tocaRango_('J', 13, 17) || tocaRango_('I', 18, 23) || tocaRango_('J', 18, 23);
  var tocaNc = tocaCelda_('C', 87) || tocaCelda_('I', 61) || tocaCelda_('K', 61) || tocaA1_(celdas.tpv3Nc) || tocaCelda_('C', 86) ||
    tocaRango_('J', 71, 75) || tocaRango_('I', 76, 81) || tocaRango_('J', 76, 81);

  if (!tocaMd && !tocaNc) return; // se editó otra celda de esta pestaña (movimientos, etc.)

  var diaExistente = obtenerDiaJSON_(hoja.getParent(), fechaStr);
  if (!diaExistente || !diaExistente.encontrado) return; // este día nunca se sincronizó desde la app

  var huboCambio = false;
  if (tocaMd && diaExistente.mediodia && diaExistente.mediodia.id) {
    aplicarEdicionTurnoDesdeHoja_(hoja, diaExistente.mediodia, 'C45', 'I5', 'K5', celdas.tpv3Md, 'C43', 13, 18);
    huboCambio = true;
  }
  if (tocaNc && diaExistente.noche && diaExistente.noche.id) {
    aplicarEdicionTurnoDesdeHoja_(hoja, diaExistente.noche, 'C87', 'I61', 'K61', celdas.tpv3Nc, 'C86', 71, 76);
    huboCambio = true;
  }
  if (!huboCambio) return;

  recalcularDia_(diaExistente);
  guardarDiaCompleto_(hoja.getParent(), fechaStr, diaExistente.negocio, diaExistente);
}

// Lee de la pestaña bonita el Total facturado / TPV 1 / 2 / 3 / Fondo fijo
// y el conteo completo de billetes y monedas de un turno, y los vuelca
// sobre el objeto `turno` (que ya viene con el resto de sus datos —
// movimientos, id, etc.— intactos, para no perder nada de lo que la app
// cargó).
function aplicarEdicionTurnoDesdeHoja_(hoja, turno, celdaFacturado, celdaTpv1, celdaTpv2, celdaTpv3, celdaFondoFijo, filaBilletesDesde, filaMonedasDesde) {
  turno.totalFacturado = Number(hoja.getRange(celdaFacturado).getValue()) || 0;
  turno.tpv1 = Number(hoja.getRange(celdaTpv1).getValue()) || 0;
  turno.tpv2 = Number(hoja.getRange(celdaTpv2).getValue()) || 0;
  turno.tpv3 = Number(hoja.getRange(celdaTpv3).getValue()) || 0;
  turno.fondoFijo = Number(hoja.getRange(celdaFondoFijo).getValue()) || 0;

  var denom = { billetes: {}, monedasBlister: {}, monedasSueltas: {} };
  BILLETES_VALORES_.forEach(function (v, i) {
    denom.billetes[v] = Number(hoja.getRange('J' + (filaBilletesDesde + i)).getValue()) || 0;
  });
  MONEDAS_VALORES_.forEach(function (v, i) {
    var filaCelda = filaMonedasDesde + i;
    denom.monedasBlister[v] = Number(hoja.getRange('I' + filaCelda).getValue()) || 0;
    denom.monedasSueltas[v] = Number(hoja.getRange('J' + filaCelda).getValue()) || 0;
  });
  turno.denom = denom;
}

function recalcularDia_(dia) {
  var ROLL_VALUE = { 2: 50, 1: 25, 0.5: 20, 0.2: 8, 0.1: 4, 0.05: 2.5 };
  var BILLETES = [100, 50, 20, 10, 5];
  var MONEDAS = [2, 1, 0.5, 0.2, 0.1, 0.05];
  var NO_EFECTIVO = ['no_efectivo', 'tarjeta', 'no_pagado', 'transferencia'];

  function totalContadoDesdeDenom_(denom) {
    denom = denom || {};
    var total = 0;
    BILLETES.forEach(function (v) {
      total += (Number((denom.billetes || {})[v]) || 0) * v;
    });
    MONEDAS.forEach(function (v) {
      var sueltas = Number((denom.monedasSueltas || {})[v]) || 0;
      var blister = Number((denom.monedasBlister || {})[v]) || 0;
      total += sueltas * v + blister * (ROLL_VALUE[v] || 0);
    });
    return total;
  }

  var md = dia.mediodia, nc = dia.noche;
  [md, nc].forEach(function (turno, idx) {
    if (!turno) return;
    var movs = turno.movimientos || [];
    var empanadas = 0, gastoEfectivo = 0, gastoNoEfectivo = 0, ingresoEfectivo = 0, egreso = 0;
    movs.forEach(function (m) {
      var importe = Number(m.importe) || 0;
      if (m.tipo === 'gasto') {
        if (NO_EFECTIVO.indexOf(m.subtipo) > -1) gastoNoEfectivo += importe;
        else gastoEfectivo += importe;
      } else if (m.tipo === 'ingreso') {
        ingresoEfectivo += importe;
      } else if (m.tipo === 'egreso') {
        egreso += importe;
      } else if (m.tipo === 'empanadas') {
        empanadas += importe;
      }
    });

    var esNoche = idx === 1;
    var totalFacturadoPropio = turno.totalFacturado || 0;
    var tpv1Propio = turno.tpv1 || 0;
    var tpv2Propio = turno.tpv2 || 0;
    var tpv3Propio = turno.tpv3 || 0;
    if (esNoche && md) {
      totalFacturadoPropio = Math.max(0, (turno.totalFacturado || 0) - (md.totalFacturado || 0));
      tpv1Propio = Math.max(0, (turno.tpv1 || 0) - (md.tpv1 || 0));
      tpv2Propio = Math.max(0, (turno.tpv2 || 0) - (md.tpv2 || 0));
      tpv3Propio = Math.max(0, (turno.tpv3 || 0) - (md.tpv3 || 0));
    }

    var facturadoNeto = totalFacturadoPropio - empanadas;
    var ventasEfectivo = facturadoNeto - (tpv1Propio + tpv2Propio + tpv3Propio);
    var totalContado = totalContadoDesdeDenom_(turno.denom);
    var esperado = (turno.fondoFijo || 0) + ventasEfectivo + ingresoEfectivo - gastoEfectivo - egreso;
    var diferencia = totalContado - esperado;

    turno.calc = {
      empanadas: empanadas, facturadoNeto: facturadoNeto, ventasEfectivo: ventasEfectivo,
      ingresoEfectivo: ingresoEfectivo, gastoEfectivo: gastoEfectivo, egreso: egreso,
      esperado: esperado, totalContado: totalContado, diferencia: diferencia,
      gastoNoEfectivo: gastoNoEfectivo,
      totalFacturadoPropio: totalFacturadoPropio, tpv1Propio: tpv1Propio, tpv2Propio: tpv2Propio, tpv3Propio: tpv3Propio
    };
  });

  return dia;
}

function guardarDiaCompleto_(ss, fecha, negocio, dia) {
  var registro = ss.getSheetByName('Registro');
  if (!registro) return;
  var header = registro.getRange(1, 1, 1, registro.getLastColumn()).getValues()[0];
  var idxJSON = -1;
  for (var h = 0; h < header.length; h++) {
    if (String(header[h] || '').indexOf('Datos JSON') === 0) { idxJSON = h; break; }
  }
  if (idxJSON === -1) return;

  var idxFondoFijo = header.indexOf('Fondo fijo');
  var idxTotalFacturado = header.indexOf('Total facturado');
  var idxTpv1 = header.indexOf('TPV 1');
  var idxTpv2 = header.indexOf('TPV 2');
  var idxTpv3 = header.indexOf('TPV 3');
  var idxEsperado = header.indexOf('Efectivo esperado');
  var idxContado = header.indexOf('Efectivo contado');
  var idxDiferencia = header.indexOf('Diferencia');
  var idxFacturadoNeto = header.indexOf('Facturado neto');
  var idxUltima = header.indexOf('Última actualización');

  ['mediodia', 'noche'].forEach(function (t) {
    var turno = dia[t];
    if (!turno || !turno.id) return;
    var filaN = buscarFilaPorId_(registro, turno.id);
    if (filaN === -1) return;
    var c = turno.calc || {};
    if (idxFondoFijo > -1) registro.getRange(filaN, idxFondoFijo + 1).setValue(turno.fondoFijo || 0);
    if (idxTotalFacturado > -1) registro.getRange(filaN, idxTotalFacturado + 1).setValue(turno.totalFacturado || 0);
    if (idxTpv1 > -1) registro.getRange(filaN, idxTpv1 + 1).setValue(turno.tpv1 || 0);
    if (idxTpv2 > -1) registro.getRange(filaN, idxTpv2 + 1).setValue(turno.tpv2 || 0);
    if (idxTpv3 > -1) registro.getRange(filaN, idxTpv3 + 1).setValue(turno.tpv3 || 0);
    if (idxEsperado > -1) registro.getRange(filaN, idxEsperado + 1).setValue(c.esperado || 0);
    if (idxContado > -1) registro.getRange(filaN, idxContado + 1).setValue(c.totalContado || 0);
    if (idxDiferencia > -1) registro.getRange(filaN, idxDiferencia + 1).setValue(c.diferencia || 0);
    if (idxFacturadoNeto > -1) registro.getRange(filaN, idxFacturadoNeto + 1).setValue(c.facturadoNeto || 0);
    turno.escritoPor = 'sheet'; // editado a mano: una copia vieja de la app no lo pisa
    registro.getRange(filaN, idxJSON + 1).setValue(JSON.stringify(turno));
    if (idxUltima > -1) registro.getRange(filaN, idxUltima + 1).setValue(new Date());
  });

  var dataParaHoja = {
    fecha: fecha, negocio: negocio,
    mediodia: turnoParaHojaExacta_(dia.mediodia),
    noche: turnoParaHojaExacta_(dia.noche)
  };
  escribirHojaDelDiaExacta_(ss, dataParaHoja);
}

// Proveedores genéricos (ver PROVEEDORES_CON_DETALLE_LIBRE en el
// front-end): al elegirlos, la app pide escribir a mano el nombre real, que
// se guarda acá como "<genérico> — <detalle>".
var PROVEEDORES_CON_DETALLE_LIBRE_ = ['Varios', 'Super'];
var TIPO_LABEL_MAP_ = { gasto: 'Gasto', ingreso: 'Ingreso', egreso: 'Egreso', empanadas: 'Empanadas' };
var SUBTIPO_LABEL_MAP_ = {
  efectivo: 'Efectivo', no_efectivo: 'No efectivo', efectivo_antiguo: 'Efectivo antiguo',
  tarjeta: 'Tarjeta', no_pagado: 'No pagado', transferencia: 'Transferencia',
  cambio: 'Cambio', ingreso_arroba: 'Ingreso @', empanadas_ing: 'Empanadas', empleados: 'Empleados',
  varios: 'Varios', retiro: 'Retiro de dinero', empanadas: 'Empanadas'
};
function turnoParaHojaExacta_(turno) {
  if (!turno) return null;
  var copia = {};
  for (var k in turno) copia[k] = turno[k];
  copia.estado = 'Sincronizado';
  copia.movimientos = (turno.movimientos || []).map(function (m) {
    var mCopia = {};
    for (var k2 in m) mCopia[k2] = m[k2];
    mCopia.tipo = TIPO_LABEL_MAP_[m.tipo] || m.tipo;
    mCopia.subtipo = SUBTIPO_LABEL_MAP_[m.subtipo] || m.subtipo;
    return mCopia;
  });
  return copia;
}

// Etiquetas que la app deja marcadas sola en el campo Info según la forma
// de pago (ver GASTO_INFO_TAG en el front-end) — para poder sacarlas antes
// de poner la etiqueta nueva al cambiar la forma de pago de un albarán ya
// cargado.
var FORMA_PAGO_LABEL_ = { tarjeta: 'Tarjeta', transferencia: 'Transferencia', efectivo_antiguo: 'Efectivo antiguo' };
// Palabra de modalidad que se pone en Info junto con "PAGADO" — para
// Efectivo antiguo se pone "EFECTIVO" (es la forma de pago real; "antiguo"
// solo aclara que salió de una caja de un día posterior).
var FORMA_PAGO_MODALIDAD_ = { tarjeta: 'TARJETA', transferencia: 'TRANSFERENCIA', efectivo_antiguo: 'EFECTIVO' };
var INFO_TAGS_CONOCIDAS_ = ['TARJETA', 'NO PAGADO', 'TRANSFERENCIA'];

function quitarTagInfo_(info) {
  info = String(info || '');
  // Si ya tenía un "PAGADO <MODALIDAD> <fecha>" puesto de una vez anterior,
  // lo saca primero (para no duplicarlo si se cambia la forma de pago de
  // nuevo). Hasta 2 palabras después de PAGADO: modalidad y fecha.
  var sinPagado = info.replace(/^PAGADO(\s+\S+){0,2}\s*(?:·\s*)?/, '');
  if (sinPagado !== info) info = sinPagado;
  for (var i = 0; i < INFO_TAGS_CONOCIDAS_.length; i++) {
    var tag = INFO_TAGS_CONOCIDAS_[i];
    if (info === tag) return '';
    var prefijo = tag + ' · ';
    if (info.indexOf(prefijo) === 0) return info.slice(prefijo.length);
  }
  return info;
}

// Desde la pantalla de Albaranes: al tocar un albarán "No pagado" y elegir
// cómo se pagó. Tarjeta/Transferencia solo cambian la etiqueta de este
// movimiento viejo (no afectan ninguna caja). Efectivo antiguo se usa
// cuando además se cargó, en el día de hoy, un gasto "Efectivo antiguo" que
// sale de esa caja — acá solo se deja marcado el albarán viejo como pagado
// de esa forma.
function cambiarFormaPagoAlbaran_(data) {
  try {
    var nuevaFormaPago = data.nuevaFormaPago;
    if (!FORMA_PAGO_LABEL_[nuevaFormaPago]) {
      return { ok: false, error: 'Forma de pago no válida.' };
    }

    if (!data.fecha) {
      return { ok: false, error: 'Falta la fecha del albarán — hace falta para saber en qué planilla mensual buscarlo. Volvé a sincronizar un cambio desde la app para que se mande.' };
    }
    var ss;
    try { ss = planillaCierreCaja_(data.fecha); }
    catch (errSs) { return { ok: false, error: String(errSs) }; }

    var mov = ss.getSheetByName('Movimientos');
    if (!mov) return { ok: false, error: 'No existe la pestaña "Movimientos".' };

    var header = mov.getRange(1, 1, 1, mov.getLastColumn()).getValues()[0];
    var idxIdMov = header.indexOf('ID Movimiento');
    var idxIdTurno = header.indexOf('ID');
    var idxSubtipo = header.indexOf('Subtipo');
    var idxInfo = header.indexOf('Info');
    var idxUltima = header.indexOf('Última actualización');
    if (idxIdMov === -1 || idxSubtipo === -1 || idxInfo === -1) {
      return { ok: false, error: 'La pestaña "Movimientos" no tiene las columnas necesarias (¿versión vieja del script?).' };
    }

    var valores = mov.getDataRange().getValues();
    var filaEncontrada = -1;
    for (var i = 1; i < valores.length; i++) {
      if (valores[i][idxIdMov] === data.idMovimiento) { filaEncontrada = i + 1; break; }
    }
    if (filaEncontrada === -1) {
      return { ok: false, error: 'No se encontró ese movimiento en "Movimientos" (puede que ya se haya reescrito).' };
    }

    var infoActual = String(valores[filaEncontrada - 1][idxInfo] || '');
    var infoSinTag = quitarTagInfo_(infoActual);
    var fechaPago = data.fechaPago ? String(data.fechaPago) : '';
    var modalidad = FORMA_PAGO_MODALIDAD_[nuevaFormaPago] || '';
    var nuevoInfo = 'PAGADO' + (modalidad ? ' ' + modalidad : '') + (fechaPago ? ' ' + fechaPago : '') + (infoSinTag ? ' · ' + infoSinTag : '');
    var nuevoSubtipoLabel = FORMA_PAGO_LABEL_[nuevaFormaPago];

    mov.getRange(filaEncontrada, idxSubtipo + 1).setValue(nuevoSubtipoLabel);
    mov.getRange(filaEncontrada, idxInfo + 1).setValue(nuevoInfo);
    if (idxUltima > -1) mov.getRange(filaEncontrada, idxUltima + 1).setValue(new Date());

    var idTurno = valores[filaEncontrada - 1][idxIdTurno];
    actualizarMovimientoEnRegistro_(ss, idTurno, data.idMovimiento, { subtipo: nuevaFormaPago, info: nuevoInfo });

    // Que el archivo de Albaranes muestre la forma de pago nueva.
    var albaranes;
    try { albaranes = sincronizarAlbaranesDelDia_(data.fecha, gastosDelDia_(obtenerDiaJSON_(ss, data.fecha)), true); }
    catch (errAlb) { albaranes = { ok: false, error: String(errAlb) }; }

    return { ok: true, albaranes: albaranes };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

function actualizarMovimientoEnRegistro_(ss, idTurno, idMovimiento, cambios) {
  var registro = ss.getSheetByName('Registro');
  if (!registro) return;
  var filaN = buscarFilaPorId_(registro, idTurno);
  if (filaN === -1) return;

  var header = registro.getRange(1, 1, 1, registro.getLastColumn()).getValues()[0];
  var idxJSON = -1;
  for (var h = 0; h < header.length; h++) {
    if (String(header[h] || '').indexOf('Datos JSON') === 0) { idxJSON = h; break; }
  }
  if (idxJSON === -1) return;

  var jsonTexto = registro.getRange(filaN, idxJSON + 1).getValue();
  if (!jsonTexto) return;

  var turnoData;
  try { turnoData = JSON.parse(jsonTexto); } catch (e) { return; }

  var seEncontro = false;
  (turnoData.movimientos || []).forEach(function (m) {
    if (m.id === idMovimiento) {
      for (var k in cambios) m[k] = cambios[k];
      seEncontro = true;
    }
  });
  if (!seEncontro) return;

  registro.getRange(filaN, idxJSON + 1).setValue(JSON.stringify(turnoData));
  var idxUltima = header.indexOf('Última actualización');
  if (idxUltima > -1) registro.getRange(filaN, idxUltima + 1).setValue(new Date());

  var idxFecha = header.indexOf('Fecha');
  var idxNegocio = header.indexOf('Negocio');
  var filaCompleta = registro.getRange(filaN, 1, 1, header.length).getValues()[0];
  var fechaRaw = filaCompleta[idxFecha];
  var fechaStr = (fechaRaw instanceof Date)
    ? Utilities.formatDate(fechaRaw, Session.getScriptTimeZone(), 'yyyy-MM-dd')
    : String(fechaRaw || '');
  if (!fechaStr) return;

  var dia = obtenerDiaJSON_(ss, fechaStr);
  if (!dia || !dia.encontrado) return;
  var dataParaHoja = {
    fecha: fechaStr,
    negocio: filaCompleta[idxNegocio] || dia.negocio,
    mediodia: turnoParaHojaExacta_(dia.mediodia),
    noche: turnoParaHojaExacta_(dia.noche)
  };
  escribirHojaDelDiaExacta_(ss, dataParaHoja);
}

// Junta los días con datos de los últimos MESES_HISTORIAL_ meses (los que
// ya existen) para el calendario y para el fondo fijo sugerido del turno
// anterior — ver planillasCierreCajaRecientes_ en Indice.gs.
//
// Además devuelve `cierres`: por fecha, el efectivo contado de cada turno y
// si ese día tuvo actividad — la app lo usa para el fondo fijo sugerido
// (efectivo contado de la última caja trabajada), aunque esa caja se haya
// cargado desde otro dispositivo. También el TPV 1 de cada turno, para
// el aviso de la app cuando el TPV 1 llega al límite del mes.
function listarDiasConDatos_() {
  try {
    var fechasSet = {};
    var cierres = {};

    planillasCierreCajaRecientes_(MESES_HISTORIAL_).forEach(function (p) {
      var registro = p.ss.getSheetByName('Registro');
      if (!registro) return;

      var valores = registro.getDataRange().getValues();
      var header = valores[0];
      var idxFecha = header.indexOf('Fecha');
      if (idxFecha === -1) return;
      var idxTurno = header.indexOf('Turno');
      var idxContado = header.indexOf('Efectivo contado');
      var idxCantMov = header.indexOf('Cant. movimientos');
      var idxFacturado = header.indexOf('Total facturado');
      var idxTpv1 = header.indexOf('TPV 1');
      var idxTpv2 = header.indexOf('TPV 2');
      var idxTpv3 = header.indexOf('TPV 3');
      function num(fila, idx) { return idx > -1 ? (Number(fila[idx]) || 0) : 0; }

      for (var i = 1; i < valores.length; i++) {
        var raw = valores[i][idxFecha];
        var fechaStr = (raw instanceof Date)
          ? Utilities.formatDate(raw, Session.getScriptTimeZone(), 'yyyy-MM-dd')
          : String(raw || '');
        if (!fechaStr) continue;
        fechasSet[fechaStr] = true;

        var fila = valores[i];
        var contado = num(fila, idxContado);
        var activo = num(fila, idxCantMov) > 0 || contado !== 0 || num(fila, idxFacturado) !== 0 ||
          num(fila, idxTpv1) !== 0 || num(fila, idxTpv2) !== 0 || num(fila, idxTpv3) !== 0;
        var c = cierres[fechaStr] || (cierres[fechaStr] = { mediodia: 0, noche: 0, actividad: false, tpv1Mediodia: 0, tpv1Noche: 0 });
        var turno = String(idxTurno > -1 ? fila[idxTurno] : '').toLowerCase() === 'noche' ? 'noche' : 'mediodia';
        if (activo) {
          c.actividad = true;
          c[turno] = contado;
          c[turno === 'noche' ? 'tpv1Noche' : 'tpv1Mediodia'] = num(fila, idxTpv1);
        }
      }
    });

    return { ok: true, fechas: Object.keys(fechasSet), cierres: cierres };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

// ============================================================================
// PARTE 4 — Envío a la planilla de Contabilidad (otra planilla, una pestaña
// por mes: ENERO, FEBRERO, ... DICIEMBRE)
// ============================================================================
// Fila fija por día del mes: fila 3 = día 1, fila 4 = día 2, ... así un
// mismo día sincronizado varias veces siempre pisa la misma fila (nunca
// duplica), sin importar qué haya quedado antes ahí.
//
// Columnas de cada día (fila 2 = encabezados). El resto (P-U, V, X-AG a mano o vinculadas a otro Sheet) no se toca:
//   A  días trabajados de ese día (0 / 0,5 / 1 — ver turnoTieneActividad_)
//   B  DIA — fecha
//   C  TOTAL SIST — total facturado del día sin empanadas (D + E)
//   D  MEDIO DIA — total facturado de Mediodía, sin sus empanadas
//   E  NOCHE — la parte propia de Noche (total del día completo - Mediodía),
//      sin sus empanadas
//   (Las empanadas van aparte, en G. Hasta el 5/10/2026 C, D y E las
//   incluían.)
//   F  GLOVO — suma de los tickets de Glovo del día (paso 3 del cierre). Si
//      el día no tiene tickets cargados en la app, no se toca (puede estar
//      puesto a mano).
//   G  EMPANADAS — Mediodía + Noche
//   H  TOTAL — fórmula GLOVO + EMPANADAS
//   I  TPV 1 — acumulado del día completo (el que ya carga Noche)
//   J  TPV 2 — ídem
//   K  TPV 3 — ídem
//   L  TARJETAS — fórmula TPV 1 + TPV 2 + TPV 3
//   M  EFEVO — efectivo contado al cerrar el último turno trabajado (es lo
//      que queda como fondo fijo para el día siguiente)
//   N  DIFERENCIA — Mediodía + Noche
//   O  RETIRA — egresos de Mediodía + Noche
// A34 queda con la fórmula =SUM(A3:A33), así el total de días trabajados
// del mes se actualiza solo cada vez que se escribe una fila nueva.
//
// Hasta septiembre de 2026 las columnas eran otras (F UBER EAT, G WEB,
// H EMPANADAS, I TOTAL, J TPV 1, K TPV 2, L TARJETAS): cada pestaña con el
// formato viejo se pasa sola al nuevo (adaptarPestanaContabilidadTresTpv_)
// la primera vez que se escribe en ella.
var MESES_MAYUS_ = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];
// Columnas que escribe la app (se vacían al copiar un mes o vaciar un día).
var COLS_CONTABILIDAD_APP_ = ['A', 'B', 'C', 'D', 'E', 'G', 'I', 'J', 'K', 'M', 'N', 'O'];

// Pasa una pestaña de mes del formato viejo (con WEB y 2 TPV) al nuevo (sin
// WEB, GLOVO en vez de UBER EAT, 3 TPV). Los datos de cada día se mueven a
// su columna nueva; lo que hubiera en WEB se pierde (se avisa en el
// resultado). Si la pestaña ya tiene el formato nuevo, no hace nada.
function adaptarPestanaContabilidadTresTpv_(hoja) {
  var enc = hoja.getRange('A2:U2').getValues()[0];
  function h(col) { return normalizarClave_(enc[columnaLetraANumero_(col) - 1]); }
  if (h('G') !== 'WEB' || h('K') !== 'TPV 2') return null;

  var viejos = hoja.getRange('F3:L33').getValues(); // UBER, WEB, EMPANADAS, TOTAL, TPV 1, TPV 2, TARJETAS
  var diasConWeb = [];
  var nuevos = viejos.map(function (v, i) {
    var r = i + 3;
    if (v[1] !== '' && v[1] != null && Number(v[1]) !== 0) diasConWeb.push((i + 1) + ': ' + v[1]);
    return [v[0], v[2], '=F' + r + '+G' + r, v[4], v[5], '', '=I' + r + '+J' + r + '+K' + r];
  });
  hoja.getRange('F3:L33').setValues(nuevos);
  hoja.getRange('F2:L2').setValues([['GLOVO', 'EMPANADAS', 'TOTAL', 'TPV 1', 'TPV 2', 'TPV 3', 'TARJETAS']]);

  // Promedios y totales de abajo que dependían de las columnas movidas.
  hoja.getRange('G35').setFormula('=G34/$A$34');
  hoja.getRange('H35').setFormula('=H34/$A$34');
  hoja.getRange('K35').setFormula('=K34/$A$34');
  if (normalizarClave_(hoja.getRange('G36').getValue()) === 'TOTAL EXTRAS') hoja.getRange('H36').setFormula('=F34+G34');

  // Comparación con los TPV del banco (P/Q): TPV 1 y TPV 2 ahora están en I y J.
  if (h('P') === 'TPV 1' && h('Q') === 'TPV 2') {
    var comparacion = [];
    for (var r = 3; r <= 33; r++) {
      comparacion.push(['=I' + r + '-P' + r, '=J' + r + '-Q' + r, '=IFERROR(R' + r + '/I' + r + ', "")', '=IFERROR(S' + r + '/J' + r + ', "")']);
    }
    hoja.getRange('R3:U33').setValues(comparacion);
  }

  return { pestana: hoja.getName(), diasConWeb: diasConWeb };
}

// Desde el menú: pasa al formato nuevo todas las pestañas de mes (y MASTER)
// de todas las planillas de Contabilidad del Índice.
function adaptarContabilidadTresTpv() {
  var resumen = [];
  planillasDelTipo_('CONTABILIDAD').forEach(function (p) {
    var ss = SpreadsheetApp.openById(p.id);
    MESES_MAYUS_.concat(['MASTER']).forEach(function (nombre) {
      var hoja = ss.getSheetByName(nombre);
      if (!hoja) return;
      var r = adaptarPestanaContabilidadTresTpv_(hoja);
      if (!r) return;
      resumen.push(p.nombre + ' / ' + nombre + ': adaptada' + (r.diasConWeb.length ? ' — tenía WEB (se borró) en los días ' + r.diasConWeb.join(', ') : ''));
    });
  });
  return resumen.length ? resumen : ['No había ninguna pestaña con el formato viejo.'];
}

function adaptarContabilidadTresTpvDesdeMenu() {
  var ui = SpreadsheetApp.getUi();
  var ok = ui.alert(
    'Contabilidad: pasar a 3 TPV',
    'Esto cambia las columnas de todas las pestañas de mes de Contabilidad: saca WEB, UBER EAT pasa a llamarse GLOVO, agrega TPV 3 y TARJETAS suma los 3 TPV. Los datos de cada día se mueven solos a su columna nueva. ¿Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (ok !== ui.Button.YES) return;
  ui.alert('Listo', adaptarContabilidadTresTpv().join('\n'), ui.ButtonSet.OK);
}

// El fondo fijo solo NO cuenta como actividad: la app lo pone sola al abrir
// un día (viene de la caja anterior), así que un día abierto sin cargar
// nada no es un día trabajado (antes quedaba con diferencia = -fondo fijo).
function turnoTieneActividad_(t) {
  var c = t.calc || {};
  var movs = t.movimientosRaw || t.movimientos || [];
  return movs.length > 0 || (t.glovo || []).length > 0 || (c.totalContado || 0) !== 0 ||
    (t.totalFacturado || 0) !== 0 || (t.tpv1 || 0) !== 0 || (t.tpv2 || 0) !== 0 || (t.tpv3 || 0) !== 0;
}

// Trae la pestaña del mes en la planilla de Contabilidad DEL AÑO que
// corresponda (ver Indice.gs — esa planilla anual se crea sola la primera
// vez que hace falta), creando la pestaña del mes si todavía no existe.
//
// Los meses nuevos se copian SIEMPRE de la pestaña "MASTER": un mes en
// blanco con los títulos, el formato y las fórmulas (se puede cambiar a mano
// y los meses siguientes salen así). Cada planilla anual tiene su MASTER: si
// la de un año no la tiene (por ejemplo, el año recién creado), se copia la
// del año anterior. Solo si no hay ninguna MASTER se usa, como antes, el mes
// anterior, vaciando las columnas que escribe la app.
var NOMBRE_MASTER_CONTABILIDAD_ = 'MASTER';

function getOrCrearPestanaContabilidad_(ss, mesIndex /* 0-11 */) {
  var nombre = MESES_MAYUS_[mesIndex];
  var hoja = ss.getSheetByName(nombre);
  if (hoja) return hoja;

  var master = masterContabilidad_(ss);
  var anterior = ss.getSheetByName(MESES_MAYUS_[(mesIndex + 11) % 12]);
  var origen = master || anterior;
  if (!origen) throw new Error('No se encontró la pestaña MASTER ni ningún mes para duplicar en la planilla de Contabilidad.');

  hoja = origen.copyTo(ss);
  hoja.setName(nombre);
  // Va detrás del mes anterior si existe; si no, delante de la MASTER.
  ss.setActiveSheet(hoja);
  if (anterior) ss.moveActiveSheet(anterior.getIndex() + 1);
  else if (master) ss.moveActiveSheet(master.getIndex());

  adaptarPestanaContabilidadTresTpv_(hoja);
  if (!master) {
    // Copia de un mes con datos: se vacían los días que escribe la app.
    COLS_CONTABILIDAD_APP_.concat(['F']).forEach(function (col) {
      hoja.getRange(col + '3:' + col + '33').clearContent();
    });
  }
  hoja.getRange('A34').setFormula('=SUM(A3:A33)');
  protegerPestanaContabilidad_(hoja);
  // La MASTER de una planilla recién creada desde la plantilla aún no tiene
  // la protección.
  if (master && !master.getProtections(SpreadsheetApp.ProtectionType.SHEET).some(function (p) {
    return p.getDescription() === DESCRIPCION_PROTECCION_CONTABILIDAD_;
  })) protegerPestanaContabilidad_(master);

  return hoja;
}

// La MASTER de esta planilla; si no tiene, copia la de otro año (el anterior
// más cercano que la tenga) y la deja al final, visible. null si no hay
// ninguna.
function masterContabilidad_(ss) {
  var propia = ss.getSheetByName(NOMBRE_MASTER_CONTABILIDAD_);
  if (propia) return propia;
  var otras = planillasDelTipo_('CONTABILIDAD')
    .filter(function (p) { return p.id !== ss.getId(); })
    .sort(function (a, b) { return a.periodo < b.periodo ? 1 : -1; });
  for (var i = 0; i < otras.length; i++) {
    try {
      var m = SpreadsheetApp.openById(otras[i].id).getSheetByName(NOMBRE_MASTER_CONTABILIDAD_);
      if (!m) continue;
      var copia = m.copyTo(ss);
      copia.setName(NOMBRE_MASTER_CONTABILIDAD_);
      ss.setActiveSheet(copia);
      ss.moveActiveSheet(ss.getNumSheets());
      protegerPestanaContabilidad_(copia);
      return copia;
    } catch (err) {
      Logger.log('No se pudo copiar la MASTER de ' + otras[i].nombre + ': ' + err);
    }
  }
  return null;
}

// Contabilidad solo la edita el dueño de la planilla; el resto de la gente con
// acceso puede editar únicamente las columnas P, Q y W de los meses y la
// MASTER (las demás pestañas quedan protegidas enteras). La app escribe como
// el dueño (executeAs USER_DEPLOYING), así que la protección no la frena.
// Se aplica sola a cada mes y MASTER nuevos; para las que ya existen, desde
// el menú (protegerContabilidad).
var DESCRIPCION_PROTECCION_CONTABILIDAD_ = 'Contabilidad: solo el dueño (P, Q y W libres)';
var COLS_CONTABILIDAD_LIBRES_ = ['P', 'Q', 'W'];
// La única cuenta que puede editar Contabilidad y Albaranes enteras. Google
// no deja que quien pone la protección se quite a sí mismo, así que hay que
// ponerla con esta cuenta (si no, no se protege nada y se avisa).
var EDITOR_UNICO_ = 'lfisbein@gmail.com';

// Protege la pestaña entera para que solo la edite EDITOR_UNICO_ (y el dueño
// del archivo, que Google no deja quitar), salvo las columnas libres. Si ya
// tenía esta protección, la rehace.
function protegerPestanaSoloEditor_(hoja, descripcion, columnasLibres) {
  hoja.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) {
    if (p.getDescription() === descripcion) p.remove();
  });
  var proteccion = hoja.protect().setDescription(descripcion);
  if (columnasLibres && columnasLibres.length) {
    proteccion.setUnprotectedRanges(columnasLibres.map(function (col) {
      return hoja.getRange(col + ':' + col);
    }));
  }
  proteccion.addEditor(EDITOR_UNICO_);
  proteccion.removeEditors(proteccion.getEditors().filter(function (u) {
    return u.getEmail().toLowerCase() !== EDITOR_UNICO_;
  }));
  if (proteccion.canDomainEdit()) proteccion.setDomainEdit(false);
}

function protegerPestanaContabilidad_(hoja) {
  var nombre = hoja.getName();
  var esMes = MESES_MAYUS_.indexOf(nombre) !== -1 || nombre === NOMBRE_MASTER_CONTABILIDAD_;
  protegerPestanaSoloEditor_(hoja, DESCRIPCION_PROTECCION_CONTABILIDAD_, esMes ? COLS_CONTABILIDAD_LIBRES_ : []);
}

// Protege todas las pestañas de todas las planillas de un tipo del Índice.
// Solo si lo corre EDITOR_UNICO_; avisa si el dueño de un archivo es otro.
function protegerPlanillasDelTipo_(tipo, protegerPestana) {
  var quien = String(Session.getEffectiveUser().getEmail() || '').toLowerCase();
  if (quien !== EDITOR_UNICO_) {
    return ['No se protegió nada: esto lo está haciendo ' + (quien || 'una cuenta desconocida') +
      ', y esa cuenta quedaría pudiendo editar. Entra con ' + EDITOR_UNICO_ + ' y vuelve a usar esta opción.'];
  }
  var resumen = [];
  planillasDelTipo_(tipo).forEach(function (p) {
    try {
      var ss = SpreadsheetApp.openById(p.id);
      var hojas = ss.getSheets();
      hojas.forEach(protegerPestana);
      var duenio = String((ss.getOwner() && ss.getOwner().getEmail()) || '').toLowerCase();
      resumen.push(p.nombre + ': ' + hojas.length + ' pestañas protegidas' +
        (duenio && duenio !== EDITOR_UNICO_
          ? ' — OJO: la dueña del archivo es ' + duenio + ' y el dueño siempre puede editar todo. Hay que pasar la propiedad a ' + EDITOR_UNICO_ + '.'
          : ''));
    } catch (err) {
      resumen.push(p.nombre + ': no se pudo proteger — ' + err);
    }
  });
  return resumen.length ? resumen : ['No hay planillas de ' + tipo + ' en el Índice.'];
}

function protegerContabilidad() {
  return protegerPlanillasDelTipo_('CONTABILIDAD', protegerPestanaContabilidad_);
}

function protegerContabilidadDesdeMenu() {
  SpreadsheetApp.getUi().alert('Contabilidad protegida', protegerContabilidad().join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

// ----------------------------------------------------------------------------
// Tablas de abajo de cada mes de Contabilidad
// ----------------------------------------------------------------------------
// 1) Semanas (C40:E44, títulos en D39/E39): semanas de lunes a domingo. Una
//    semana que cruza dos meses es del mes que tiene 4 o más días de ella (el
//    del jueves): un mes que empieza de lunes a jueves suma en su semana 1 los
//    días del final del mes anterior; uno que empieza viernes, sábado o
//    domingo deja esos días en la última semana del mes anterior.
//      C  "semana N (dd/mm – dd/mm)"
//      D  SEMANA — promedio por día trabajado: TOTAL SIST / días trabajados
//         (A: 0,5 si solo Mediodía, 1 si Mediodía y Noche). Sin Glovo ni
//         empanadas.
//      E  FIN DE SEMANA — TOTAL SIST de viernes, sábado y domingo de esa
//         semana
//    Como mira otros meses (y en enero/diciembre la planilla de otro año), no
//    son fórmulas: se calculan y escriben al guardar cada día (ese mes y los
//    de al lado), y desde el menú para todo el año.
// 2) Días de la semana (C48:K55): una columna por día (E lunes … K domingo).
//    Filas 49-51 totales de TOTAL SIST, MEDIO DIA y NOCHE; 52 vacía; 53-55
//    promedios: cada total entre los días (o turnos) que facturaron algo.
//    Aquí no se usa el 0,5 de la columna A: un domingo solo de mediodía con
//    1.882 daba 3.764 de promedio. Valores, no fórmulas (las fórmulas daban #ERROR! en
//    la planilla); se recalcula igual que las semanas. En la MASTER queda
//    en blanco.
// Solo desde octubre de 2026: los meses anteriores quedan como estaban.
var DESDE_TABLAS_CONTABILIDAD_ = { anio: 2026, mes: 9 };
var FILA_SEMANAS_CONTABILIDAD_ = 40; // semana 1; hasta 5 semanas
var FILA_DIAS_SEMANA_CONTABILIDAD_ = 48;

// Lunes de cada semana que pertenece al mes (4 o 5).
function semanasDelMesContabilidad_(anio, mesIndex) {
  var primero = new Date(anio, mesIndex, 1, 12);
  var lunes = new Date(anio, mesIndex, 1 - ((primero.getDay() + 6) % 7), 12);
  var semanas = [];
  for (var d = lunes; ; d = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 7, 12)) {
    var jueves = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 3, 12);
    if (jueves.getFullYear() * 12 + jueves.getMonth() > anio * 12 + mesIndex) break;
    if (jueves.getMonth() === mesIndex) semanas.push(d);
  }
  return semanas;
}

// Lee días de las pestañas de mes (sin crear nada) con caché por pestaña.
// Devuelve { dias: A, total: C } del día (TOTAL SIST: sin Glovo ni
// empanadas), o null si no hay pestaña.
function lectorDiasContabilidad_() {
  var planillas = {}, pestanas = {};
  return function (fecha) {
    var anio = fecha.getFullYear(), mes = fecha.getMonth();
    var clave = anio + '-' + mes;
    if (!(clave in pestanas)) {
      if (!(anio in planillas)) {
        var id = buscarEnIndice_('CONTABILIDAD', String(anio));
        planillas[anio] = id ? SpreadsheetApp.openById(id) : null;
      }
      var hoja = planillas[anio] && planillas[anio].getSheetByName(MESES_MAYUS_[mes]);
      pestanas[clave] = hoja ? hoja.getRange('A3:H33').getValues() : null;
    }
    var filas = pestanas[clave];
    if (!filas) return null;
    var f = filas[fecha.getDate() - 1];
    return { dias: Number(f[0]) || 0, total: Number(f[2]) || 0 };
  };
}

function llevaTablasContabilidad_(anio, mesIndex) {
  return anio * 12 + mesIndex >= DESDE_TABLAS_CONTABILIDAD_.anio * 12 + DESDE_TABLAS_CONTABILIDAD_.mes;
}

function escribirSemanasContabilidad_(hoja, anio, mesIndex, leerDia) {
  if (!llevaTablasContabilidad_(anio, mesIndex)) return;
  var dosDig = function (n) { return (n < 10 ? '0' : '') + n; };
  var corto = function (d) { return dosDig(d.getDate()) + '/' + dosDig(d.getMonth() + 1); };
  var filas = semanasDelMesContabilidad_(anio, mesIndex).map(function (lunes, i) {
    var total = 0, dias = 0, finde = 0, domingo;
    for (var k = 0; k < 7; k++) {
      var d = new Date(lunes.getFullYear(), lunes.getMonth(), lunes.getDate() + k, 12);
      var dia = leerDia(d);
      if (dia) {
        total += dia.total;
        dias += dia.dias;
        if (k >= 4) finde += dia.total;
      }
      domingo = d;
    }
    return [
      'semana ' + (i + 1) + ' (' + corto(lunes) + ' – ' + corto(domingo) + ')',
      dias > 0 ? Math.round(total / dias * 100) / 100 : '',
      finde ? Math.round(finde * 100) / 100 : ''
    ];
  });
  while (filas.length < 5) filas.push(['semana ' + (filas.length + 1), '', '']);
  hoja.getRange('D' + (FILA_SEMANAS_CONTABILIDAD_ - 1) + ':E' + (FILA_SEMANAS_CONTABILIDAD_ - 1)).setValues([['SEMANA', 'FIN DE SEMANA']]);
  hoja.getRange('C' + FILA_SEMANAS_CONTABILIDAD_ + ':E' + (FILA_SEMANAS_CONTABILIDAD_ + 4)).setValues(filas);
}

function escribirDiasSemanaContabilidad_(hoja, anio, mesIndex) {
  var f0 = FILA_DIAS_SEMANA_CONTABILIDAD_;
  var redondo = function (n) { return Math.round(n * 100) / 100; };
  // Por día de la semana (0 = lunes): totales de C, D, E y cuántos días,
  // mediodías y noches facturaron algo (lo cerrado no cuenta).
  var t = [0, 1, 2, 3, 4, 5, 6].map(function () { return { c: 0, d: 0, e: 0, dias: 0, md: 0, nc: 0 }; });
  var hayDatos = anio != null;
  if (hayDatos) {
    hoja.getRange('A3:E33').getValues().forEach(function (f, i) {
      var fecha = new Date(anio, mesIndex, i + 1, 12);
      if (fecha.getMonth() !== mesIndex) return; // 29-31 en meses más cortos
      var x = t[(fecha.getDay() + 6) % 7];
      x.c += Number(f[2]) || 0;
      x.d += Number(f[3]) || 0;
      x.e += Number(f[4]) || 0;
      if ((Number(f[2]) || 0) > 0) x.dias++;
      if ((Number(f[3]) || 0) > 0) x.md++;
      if ((Number(f[4]) || 0) > 0) x.nc++;
    });
  }
  var fila = function (fn) { return t.map(function (x) { return hayDatos ? fn(x) : ''; }); };
  var prom = function (total, n) { return n > 0 ? redondo(total / n) : ''; };
  var valores = [
    fila(function (x) { return redondo(x.c); }),
    fila(function (x) { return redondo(x.d); }),
    fila(function (x) { return redondo(x.e); }),
    ['', '', '', '', '', '', ''], // fila 52: sin uso
    fila(function (x) { return prom(x.c, x.dias); }),
    fila(function (x) { return prom(x.d, x.md); }),
    fila(function (x) { return prom(x.e, x.nc); })
  ];
  hoja.getRange('E' + f0 + ':K' + f0).setValues([['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo']]);
  hoja.getRange('D' + (f0 + 1) + ':D' + (f0 + 7)).setValues([['TOTAL SIST'], ['MEDIO DIA'], ['NOCHE'], [''], ['TOTAL SIST'], ['MEDIO DIA'], ['NOCHE']]);
  hoja.getRange('E' + (f0 + 1) + ':K' + (f0 + 7)).setValues(valores);
}

// 3) Objetivos del día (filas 3-33): se pinta la celda cuando se supera el
//    objetivo (todo sin empanadas, que ya van aparte):
//      MEDIO DIA (D) de lunes a viernes      > 1.100 €
//      NOCHE (E) del jueves                  >   870 €
//      NOCHE (E) del viernes y del sábado    > 1.570 €
//      MEDIO DIA (D) del sábado y del domingo > 1.570 €
//    Solo por turnos: el total del día (C) no se pinta nunca.
//    Son formatos condicionales de "mayor que" sobre cada celda (sin
//    fórmulas), así que se actualizan solos si se cambia un número a mano.
//    Antes se borran TODOS los formatos condicionales de la pestaña. En la
//    MASTER solo se borran (no tiene fechas).
var COLOR_OBJETIVO_CONTABILIDAD_ = '#b7e1cd';
var OBJETIVOS_CONTABILIDAD_ = [
  { col: 'D', dias: [1, 2, 3, 4, 5], minimo: 1100 }, // getDay(): 0 domingo … 6 sábado
  { col: 'E', dias: [4], minimo: 870 },
  { col: 'E', dias: [5, 6], minimo: 1570 },
  { col: 'D', dias: [6, 0], minimo: 1570 }
];

function colorearObjetivosContabilidad_(hoja, anio, mesIndex) {
  if (anio == null) { hoja.setConditionalFormatRules([]); return; } // MASTER
  if (!llevaTablasContabilidad_(anio, mesIndex)) return;
  var diasMes = new Date(anio, mesIndex + 1, 0).getDate();
  var reglas = OBJETIVOS_CONTABILIDAD_.map(function (obj) {
    var rangos = [];
    for (var d = 1; d <= diasMes; d++) {
      if (obj.dias.indexOf(new Date(anio, mesIndex, d, 12).getDay()) !== -1) rangos.push(hoja.getRange(obj.col + (2 + d)));
    }
    return SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(obj.minimo)
      .setBackground(COLOR_OBJETIVO_CONTABILIDAD_)
      .setRanges(rangos)
      .build();
  });
  hoja.setConditionalFormatRules(reglas);
}

// Al guardar un día: semanas de ese mes y de los de al lado (una semana que
// cruza meses puede ser del otro). Solo pestañas que ya existen.
function actualizarSemanasAlrededor_(anio, mesIndex) {
  var leerDia = lectorDiasContabilidad_();
  [-1, 0, 1].forEach(function (delta) {
    var d = new Date(anio, mesIndex + delta, 1, 12);
    var id = buscarEnIndice_('CONTABILIDAD', String(d.getFullYear()));
    var hoja = id && SpreadsheetApp.openById(id).getSheetByName(MESES_MAYUS_[d.getMonth()]);
    if (!hoja) return;
    escribirSemanasContabilidad_(hoja, d.getFullYear(), d.getMonth(), leerDia);
    // La tabla de días de la semana (fórmulas) en el mes que se guardó.
    if (delta === 0 && llevaTablasContabilidad_(d.getFullYear(), d.getMonth())) {
      escribirDiasSemanaContabilidad_(hoja, d.getFullYear(), d.getMonth());
      colorearObjetivosContabilidad_(hoja, d.getFullYear(), d.getMonth());
    }
  });
}

// Desde el menú: las dos tablas en todos los meses (y la de días de la
// semana también en la MASTER) de todas las planillas de Contabilidad.
function tablasContabilidad() {
  var leerDia = lectorDiasContabilidad_();
  var resumen = [];
  planillasDelTipo_('CONTABILIDAD').forEach(function (p) {
    var anio = parseInt(String(p.periodo).slice(0, 4), 10);
    try {
      var ss = SpreadsheetApp.openById(p.id);
      var master = ss.getSheetByName(NOMBRE_MASTER_CONTABILIDAD_);
      if (master) {
        escribirDiasSemanaContabilidad_(master, null);
        colorearObjetivosContabilidad_(master, null);
      }
      var hechas = 0;
      MESES_MAYUS_.forEach(function (nombre, mes) {
        var hoja = ss.getSheetByName(nombre);
        if (!hoja || !llevaTablasContabilidad_(anio, mes)) return;
        escribirDiasSemanaContabilidad_(hoja, anio, mes);
        colorearObjetivosContabilidad_(hoja, anio, mes);
        escribirSemanasContabilidad_(hoja, anio, mes, leerDia);
        hechas++;
      });
      resumen.push(p.nombre + ': ' + hechas + ' meses' + (master ? ' y la MASTER' : ''));
    } catch (err) {
      resumen.push(p.nombre + ': error — ' + err);
    }
  });
  return resumen.length ? resumen : ['No hay planillas de Contabilidad en el Índice.'];
}

function tablasContabilidadDesdeMenu() {
  SpreadsheetApp.getUi().alert('Contabilidad: semanas y días de la semana', tablasContabilidad().join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

// Lo facturado por Mediodía y la parte propia de Noche (Noche carga el
// total del día completo), cada uno sin sus empanadas. Un turno sin total
// cargado queda en 0 aunque tenga empanadas.
function totalesSinEmpanadas_(md, nc) {
  md = md || {}; nc = nc || {};
  var empMd = (md.calc || {}).empanadas || 0, empNc = (nc.calc || {}).empanadas || 0;
  var totalMd = md.totalFacturado || 0;
  var propioNc = Math.max(0, (nc.totalFacturado || 0) - totalMd);
  return {
    mediodia: totalMd > 0 ? Math.round((totalMd - empMd) * 100) / 100 : 0,
    noche: propioNc > 0 ? Math.round((propioNc - empNc) * 100) / 100 : 0
  };
}

// Las semanas no deben cortar el guardado del día si fallan.
function semanasTrasGuardar_(anio, mesIndex) {
  try {
    SpreadsheetApp.flush();
    actualizarSemanasAlrededor_(anio, mesIndex);
  } catch (err) {
    Logger.log('Semanas de Contabilidad: ' + err);
  }
  try {
    actualizarPromediosTotalesAlbaranes_(anio, mesIndex);
  } catch (err2) {
    Logger.log('Promedios de TOTALES (Albaranes): ' + err2);
  }
}

function escribirContabilidad_(data) {
  try {
    var partes = String(data.fecha || '').split('-');
    if (partes.length !== 3) return { ok: false, error: 'Fecha inválida' };
    var anio = parseInt(partes[0], 10), mes = parseInt(partes[1], 10), diaDelMes = parseInt(partes[2], 10);
    if (!anio || !mes || !diaDelMes) return { ok: false, error: 'Fecha inválida' };

    var md = data.mediodia || {};
    var nc = data.noche || {};
    var cMd = md.calc || {};
    var cNc = nc.calc || {};
    var mdActivo = turnoTieneActividad_(md);
    var ncActivo = turnoTieneActividad_(nc);

    var ssContabilidad = planillaContabilidad_(data.fecha);
    var hoja = getOrCrearPestanaContabilidad_(ssContabilidad, mes - 1);
    adaptarPestanaContabilidadTresTpv_(hoja);
    var fila = 3 + (diaDelMes - 1);

    // Día sin nada cargado (o que se vació desde la app): la fila queda en
    // blanco en vez de con ceros y diferencia negativa.
    if (!mdActivo && !ncActivo) {
      COLS_CONTABILIDAD_APP_.forEach(function (col) {
        hoja.getRange(col + fila).clearContent();
      });
      semanasTrasGuardar_(anio, mes - 1);
      return { ok: true, pestana: hoja.getName(), fila: fila, vacio: true };
    }

    var diasHoy = (mdActivo && ncActivo) ? 1 : ((mdActivo || ncActivo) ? 0.5 : 0);
    var mediodiaTotal = totalesSinEmpanadas_(md, nc).mediodia;
    var nochePropio = totalesSinEmpanadas_(md, nc).noche;
    var empanadasDia = (cMd.empanadas || 0) + (cNc.empanadas || 0);
    // Noche carga el acumulado del día completo, que nunca puede ser menor
    // que lo que ya marcó Mediodía: si en Noche quedó en 0 (no se cargó),
    // vale el de Mediodía.
    var tpv1Dia = Math.max(md.tpv1 || 0, nc.tpv1 || 0);
    var tpv2Dia = Math.max(md.tpv2 || 0, nc.tpv2 || 0);
    var tpv3Dia = Math.max(md.tpv3 || 0, nc.tpv3 || 0);
    var efevoDia = ncActivo ? (cNc.totalContado || 0) : (cMd.totalContado || 0);
    // Solo los turnos que se trabajaron: un turno sin nada cargado tiene
    // igual un fondo fijo (el de Noche sale de lo contado en Mediodía), y su
    // "diferencia" sería ese fondo entero en negativo.
    var diferenciaDia = (mdActivo ? (cMd.diferencia || 0) : 0) + (ncActivo ? (cNc.diferencia || 0) : 0);
    var retiraDia = (cMd.egreso || 0) + (cNc.egreso || 0);

    hoja.getRange('A' + fila).setValue(diasHoy);
    // Mediodía (no medianoche): si el huso horario de esta planilla de
    // Contabilidad no coincide exactamente con el del proyecto de Apps
    // Script, medianoche puede caer del lado del día anterior al mostrarse,
    // corriendo la fecha visible un día para atrás. Al mediodía queda lejos
    // de cualquier límite de huso horario real.
    hoja.getRange('B' + fila).setValue(new Date(anio, mes - 1, diaDelMes, 12));
    hoja.getRange('C' + fila + ':E' + fila).setValues([[mediodiaTotal + nochePropio, mediodiaTotal, nochePropio]]);
    if (ticketsGlovoDelDia_(data).length) hoja.getRange('F' + fila).setValue(totalGlovoDelDia_(data));
    hoja.getRange('G' + fila + ':H' + fila).setValues([[empanadasDia, '=F' + fila + '+G' + fila]]);
    hoja.getRange('I' + fila + ':O' + fila).setValues([[
      tpv1Dia, tpv2Dia, tpv3Dia, '=I' + fila + '+J' + fila + '+K' + fila,
      efevoDia, diferenciaDia, retiraDia
    ]]);
    semanasTrasGuardar_(anio, mes - 1);

    return { ok: true, pestana: hoja.getName(), fila: fila };
  } catch (err) {
    // No corta el guardado normal del Cierre de Caja si esto falla.
    Logger.log('escribirContabilidad_ error: ' + err);
    return { ok: false, error: String(err) };
  }
}

// ============================================================================
// PARTE 5 — Herramienta de pruebas: borrar un día completo desde un menú en
// el propio Google Sheet (para no tener que pedir un script nuevo cada vez
// que se quiere repetir una prueba).
// ============================================================================
// Simple trigger: corre solo al abrir la planilla a la que está pegado este
// proyecto (la planilla Índice), agrega el menú "Cierre de Caja — Pruebas"
// en la barra de arriba.
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('Cierre de Caja — Pruebas')
    .addItem('Borrar un día completo…', 'borrarDiaCompleto')
    .addItem('Conectar los meses cargados en el Índice', 'conectarMesesCierreCajaDesdeMenu')
    .addItem('Mover días que quedaron en otro mes', 'moverDiasAlMesCorrectoDesdeMenu')
    .addSeparator()
    .addItem('Albaranes: restaurar datos viejos 2026 (una sola vez)', 'restaurarAlbaranesViejos2026DesdeMenu')
    .addItem('Albaranes: reenviar un mes desde Cierre de Caja…', 'reenviarMesAAlbaranesDesdeMenu')
    .addItem('Albaranes: pasar al formato con forma y día de pago', 'albaranesPasarAFormatoNuevoDesdeMenu')
    .addSeparator()
    .addItem('Contabilidad: pasar a 3 TPV (sin WEB, con GLOVO)', 'adaptarContabilidadTresTpvDesdeMenu')
    .addItem('Contabilidad: proteger (solo el dueño; P, Q y W libres)', 'protegerContabilidadDesdeMenu')
    .addItem('Contabilidad: recalcular semanas, días de la semana y colores', 'tablasContabilidadDesdeMenu')
    .addItem('Albaranes: proteger (solo el dueño; columna VARIOS libre)', 'protegerAlbaranesDesdeMenu')
    .addItem('Albaranes: pasar NOO a FRUTAPRO y sacarla del TOTAL (una vez)', 'completarFrutaproDesdeNooDesdeMenu')
    .addItem('Albaranes: completar promedios por día y semana en TOTALES', 'actualizarPromediosTotalesAlbaranesDesdeMenu')
    .addItem('Cierre de Caja: pasar la plantilla y los días a 3 TPV', 'pasarCierresATresTpvDesdeMenu')
    .addToUi();
}

// Días que quedaron guardados en la planilla de otro mes (por ejemplo, el
// 01-10 cargado en septiembre antes de partir Cierre de Caja en un archivo
// por mes): los pasa a la planilla de su mes — filas de "Registro" y
// "Movimientos" y la pestaña del día. Sin esto la app no los encuentra,
// porque busca cada día en la planilla de su mes.
function moverDiasAlMesCorrecto() {
  invalidarCacheListados_();
  var resumen = [];
  planillasDelTipo_('CIERRE_CAJA').forEach(function (p) {
    if (!/^\d{4}-\d{2}$/.test(p.periodo)) return;
    var ss = SpreadsheetApp.openById(p.id);
    var registro = ss.getSheetByName('Registro');
    if (!registro) return;
    var valores = registro.getDataRange().getValues();
    var idxFecha = valores[0].indexOf('Fecha');
    if (idxFecha === -1) return;

    var fechas = {};
    for (var i = 1; i < valores.length; i++) {
      var f = textoFecha_(valores[i][idxFecha]);
      if (/^\d{4}-\d{2}-\d{2}$/.test(f) && f.substring(0, 7) !== p.periodo) fechas[f] = true;
    }
    Object.keys(fechas).sort().forEach(function (fecha) {
      try {
        var destino = planillaCierreCaja_(fecha);
        if (destino.getId() === ss.getId()) return;
        resumen.push(fecha + ': de "' + p.nombre + '" a "' + destino.getName() + '" — ' + moverDia_(ss, destino, fecha));
      } catch (err) {
        resumen.push(fecha + ': ERROR — ' + err);
      }
    });
  });
  Logger.log(resumen.join('\n'));
  return resumen.length ? resumen : ['No hay días en una planilla de otro mes.'];
}

function textoFecha_(raw) {
  return (raw instanceof Date)
    ? Utilities.formatDate(raw, Session.getScriptTimeZone(), 'yyyy-MM-dd')
    : String(raw || '');
}

function moverDia_(origen, destino, fecha) {
  var partes = [];
  partes.push(moverFilasPorFecha_(origen.getSheetByName('Registro'), getOrCrearRegistro_(destino), fecha) + ' filas de Registro');
  partes.push(moverFilasPorFecha_(origen.getSheetByName('Movimientos'), getOrCrearMovimientos_(destino), fecha) + ' de Movimientos');

  var nombre = nombreHojaDia_(fecha);
  var hojaDia = origen.getSheetByName(nombre);
  if (hojaDia) {
    if (destino.getSheetByName(nombre)) {
      partes.push('la pestaña "' + nombre + '" ya existía en el destino: se dejó la del destino y se borró la de origen');
    } else {
      hojaDia.copyTo(destino).setName(nombre);
      partes.push('pestaña "' + nombre + '" movida');
    }
    origen.deleteSheet(hojaDia);
  }
  return partes.join(', ');
}

// Pasa las filas de `fecha` de una pestaña a otra (ubicando cada columna por
// su encabezado) y las borra del origen. Devuelve cuántas movió.
function moverFilasPorFecha_(hOrigen, hDestino, fecha) {
  if (!hOrigen || hOrigen.getLastRow() < 2) return 0;
  var valores = hOrigen.getDataRange().getValues();
  var header = valores[0];
  var idxFecha = header.indexOf('Fecha');
  if (idxFecha === -1) return 0;
  var headerDestino = hDestino.getRange(1, 1, 1, hDestino.getLastColumn()).getValues()[0];
  var mapa = headerDestino.map(function (h) { return header.indexOf(h); });

  var filasOrigen = [], filasDestino = [];
  for (var i = 1; i < valores.length; i++) {
    if (textoFecha_(valores[i][idxFecha]) !== fecha) continue;
    filasOrigen.push(i + 1);
    filasDestino.push(mapa.map(function (j) { return j > -1 ? valores[i][j] : ''; }));
  }
  if (!filasDestino.length) return 0;

  hDestino.getRange(hDestino.getLastRow() + 1, 1, filasDestino.length, headerDestino.length).setValues(filasDestino);
  for (var k = filasOrigen.length - 1; k >= 0; k--) hOrigen.deleteRow(filasOrigen[k]);
  return filasDestino.length;
}

function moverDiasAlMesCorrectoDesdeMenu() {
  var ui = SpreadsheetApp.getUi();
  var ok = ui.alert(
    'Mover días al mes correcto',
    'Busca en todas las planillas de Cierre de Caja los días que quedaron guardados en otro mes y los pasa a la planilla de su mes. Antes, revisá que en "Archivos" esté la planilla de cada mes (si falta, se crea una nueva desde la plantilla). ¿Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (ok !== ui.Button.YES) return;
  ui.alert('Listo', moverDiasAlMesCorrecto().join('\n'), ui.ButtonSet.OK);
}

function conectarMesesCierreCajaDesdeMenu() {
  var resumen = conectarMesesCierreCaja();
  SpreadsheetApp.getUi().alert('Meses de Cierre de Caja', resumen.join('\n') || 'No hay filas CIERRE_CAJA en "Archivos".', SpreadsheetApp.getUi().ButtonSet.OK);
}

// Pide una fecha, confirma, y borra TODO lo de ese día: filas de "Registro"
// y "Movimientos" con esa fecha, la pestaña de ese día (si existe), y la
// fila correspondiente en la planilla de Contabilidad de ese año/mes (solo
// las columnas que escribe la app: A,B,D,E,H,J,K,M,N,O — el resto de esa
// fila no se toca). Busca el día en todas las planillas de Cierre de Caja
// del Índice y la Contabilidad de ese año; si esta no existe, no la crea solo para borrar algo
// ahí — no hay nada que borrar en un archivo que no existe.
//
// El menú solo aparece en la planilla a la que está pegado este proyecto de
// Apps Script, pero sirve para borrar un día de CUALQUIER mes.
function borrarDiaCompleto() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt('Borrar un día completo', 'Fecha a borrar (formato AAAA-MM-DD, ej. 2026-10-05):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  var fecha = String(resp.getResponseText() || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha)) {
    ui.alert('Fecha inválida. Tiene que ser AAAA-MM-DD, ej. 2026-10-05.');
    return;
  }

  var confirmacion = ui.alert(
    'Confirmar borrado',
    'Esto borra TODO lo del ' + fecha + ': filas de Registro y Movimientos, la pestaña del día (si existe), y la fila correspondiente en Contabilidad. No se puede deshacer. ¿Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (confirmacion !== ui.Button.YES) return;

  var resumen = borrarDiaEnTodosLados_(fecha);
  ui.alert('Listo', JSON.stringify(resumen, null, 2), ui.ButtonSet.OK);
}

// Las planillas (Cierre de Caja del mes, Contabilidad del año) se buscan en
// el Índice con buscarEnIndice_, que — a diferencia de obtenerOCrearPlanilla_
// — NO las crea si no existen: para borrar no tiene sentido crear un
// archivo nuevo solo para no encontrar nada que tocar en él.
function borrarDiaEnTodosLados_(fecha) {
  invalidarCacheListados_();
  var resumen = { fecha: fecha, registro: 0, movimientos: 0, pestanaDia: null, contabilidad: null };

  function fechaDeFila_(fila, idxFecha) {
    var raw = fila[idxFecha];
    return (raw instanceof Date)
      ? Utilities.formatDate(raw, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(raw || '');
  }
  function borrarFilasPorFecha_(hoja) {
    if (!hoja) return 0;
    var ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) return 0;
    var header = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
    var idxFecha = header.indexOf('Fecha');
    if (idxFecha === -1) return 0;
    var valores = hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn()).getValues();
    var borradas = 0;
    for (var i = valores.length - 1; i >= 0; i--) {
      if (fechaDeFila_(valores[i], idxFecha) === fecha) {
        hoja.deleteRow(i + 2);
        borradas++;
      }
    }
    return borradas;
  }

  // Se busca en TODAS las planillas de Cierre de Caja del Índice, no solo
  // en la del mes de la fecha: un día puede haber quedado guardado en otro
  // mes (por ejemplo, días cargados antes de partir Cierre de Caja en un
  // archivo por mes).
  var nombreDia = nombreHojaDia_(fecha);
  resumen.planillas = [];
  planillasDelTipo_('CIERRE_CAJA').forEach(function (p) {
    try {
      var ss = SpreadsheetApp.openById(p.id);
      var reg = borrarFilasPorFecha_(ss.getSheetByName('Registro'));
      var mov = borrarFilasPorFecha_(ss.getSheetByName('Movimientos'));
      var hojaDia = ss.getSheetByName(nombreDia);
      if (hojaDia) { ss.deleteSheet(hojaDia); resumen.pestanaDia = nombreDia + ' (en ' + p.nombre + ')'; }
      resumen.registro += reg;
      resumen.movimientos += mov;
      if (reg || mov || hojaDia) resumen.planillas.push(p.nombre);
    } catch (errP) {
      resumen.planillas.push(p.nombre + ': error — ' + errP);
    }
  });

  var partes = fecha.split('-');
  var anio = parseInt(partes[0], 10), mes = parseInt(partes[1], 10), diaDelMes = parseInt(partes[2], 10);
  if (anio && mes && diaDelMes) {
    try {
      var idPlanilla = buscarEnIndice_('CONTABILIDAD', String(anio));
      if (idPlanilla) {
        var ssC = SpreadsheetApp.openById(idPlanilla);
        var hojaMes = ssC.getSheetByName(MESES_MAYUS_[mes - 1]);
        if (hojaMes) {
          adaptarPestanaContabilidadTresTpv_(hojaMes);
          var filaC = 3 + (diaDelMes - 1);
          COLS_CONTABILIDAD_APP_.forEach(function (col) {
            hojaMes.getRange(col + filaC).clearContent();
          });
          resumen.contabilidad = MESES_MAYUS_[mes - 1] + ' fila ' + filaC;
        }
      }
    } catch (errC) {
      resumen.contabilidad = 'error: ' + errC;
    }
  }

  // Albaranes: sacar los gastos de ese día (solo los que cargó la app).
  try {
    var rAlb = sincronizarAlbaranesDelDia_(fecha, [], true);
    resumen.albaranes = rAlb.ok ? 'listo' : 'error: ' + rAlb.error;
  } catch (errAlb) {
    resumen.albaranes = 'error: ' + errAlb;
  }

  // Glovo: sacar los tickets de ese día (solo los que cargó la app).
  try {
    var rGlovo = sincronizarGlovoDelDia_(fecha, []);
    resumen.glovo = rGlovo.ok ? 'listo' : 'error: ' + rGlovo.error;
  } catch (errGlovo) {
    resumen.glovo = 'error: ' + errGlovo;
  }

  return resumen;
}

// ============================================================================
// Una vez: pasar Cierre de Caja a la plantilla "1" con TPV 3 (menú).
// ============================================================================
// 1) Busca, entre las planillas de Cierre de Caja del Índice (y la planilla
//    plantilla de "Tipos"), una cuya pestaña "1" ya tenga TPV 3, y copia esa
//    "1" a las que todavía tienen la de antes — así los meses nuevos y los
//    días nuevos salen con TPV 3.
// 2) Vuelve a armar, con esa plantilla, las pestañas de los días que
//    todavía tienen el formato viejo, con los datos guardados de la app
//    (Registro). Lo que se haya escrito a mano en esas pestañas fuera de lo
//    que carga la app se pierde.
function pasarCierresATresTpv() {
  var resumen = [];
  var archivos = planillasDelTipo_('CIERRE_CAJA').map(function (p) { return { id: p.id, nombre: p.nombre, esPlantilla: false }; });
  var plantillaId = '';
  try { plantillaId = normalizarTexto_(configTipo_('CIERRE_CAJA').plantillaId); } catch (errCfg) {}
  if (plantillaId) archivos.push({ id: plantillaId, nombre: 'Plantilla de Cierre de Caja', esPlantilla: true });

  archivos.forEach(function (a) { a.ss = SpreadsheetApp.openById(a.id); });
  var origen = null;
  archivos.forEach(function (a) {
    var uno = a.ss.getSheetByName('1');
    if (!origen && uno && celdasHojaDia_(uno).conTpv3) origen = uno;
  });
  if (!origen) throw new Error('Ninguna planilla de Cierre de Caja tiene todavía la pestaña "1" con TPV 3 (EFECTIVO ANTIGUO en la columna O).');
  // En la plantilla, los títulos de TPV 3 decían "TPV 2".
  origen.getRange('L4').setValue('TPV 3');
  origen.getRange('L60').setValue('TPV 3');

  archivos.forEach(function (a) {
    var ss = a.ss;
    var uno = ss.getSheetByName('1');
    if (!uno || !celdasHojaDia_(uno).conTpv3) {
      var posicion = uno ? uno.getIndex() : 1;
      var copia = origen.copyTo(ss);
      if (uno) ss.deleteSheet(uno);
      copia.setName('1');
      ss.setActiveSheet(copia);
      ss.moveActiveSheet(posicion);
      resumen.push(a.nombre + ': plantilla "1" actualizada');
    }
    if (a.esPlantilla) return;

    var rehechos = [], sinDatos = [];
    ss.getSheets().forEach(function (hoja) {
      var nombre = hoja.getName();
      if (!/^\d{2}-\d{2}-\d{4}$/.test(nombre) || celdasHojaDia_(hoja).conTpv3) return;
      var fecha = fechaDesdeNombreHoja_(nombre);
      var dia = obtenerDiaJSON_(ss, fecha);
      if (!dia || !dia.encontrado) { sinDatos.push(nombre); return; }
      var posicionDia = hoja.getIndex();
      ss.deleteSheet(hoja);
      escribirHojaDelDiaExacta_(ss, {
        fecha: fecha,
        negocio: dia.negocio,
        mediodia: turnoParaHojaExacta_(dia.mediodia),
        noche: turnoParaHojaExacta_(dia.noche)
      });
      var nueva = ss.getSheetByName(nombre);
      if (nueva) { ss.setActiveSheet(nueva); ss.moveActiveSheet(posicionDia); }
      rehechos.push(nombre);
    });
    if (rehechos.length) resumen.push(a.nombre + ': ' + rehechos.length + ' días rehechos con TPV 3');
    if (sinDatos.length) resumen.push(a.nombre + ': sin datos de la app, no se tocaron: ' + sinDatos.join(', '));
  });
  archivos.forEach(function (a) { borrarCopiasSueltasPlantilla_(a.ss); });
  return resumen.length ? resumen : ['Todo ya estaba con TPV 3.'];
}

function pasarCierresATresTpvDesdeMenu() {
  var ui = SpreadsheetApp.getUi();
  var ok = ui.alert(
    'Cierre de Caja: pasar a 3 TPV',
    'Esto copia la pestaña "1" con TPV 3 a todas las planillas de Cierre de Caja (y a la plantilla), y vuelve a armar las pestañas de los días que tienen el formato viejo con los datos de la app. Lo que se haya escrito a mano en esas pestañas (fuera de lo que carga la app) se pierde. ¿Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (ok !== ui.Button.YES) return;
  ui.alert('Listo', pasarCierresATresTpv().join('\n'), ui.ButtonSet.OK);
}

// ---------- Escanear facturas con Claude ----------
// La clave de Anthropic NO va en el código (el repositorio es público): se
// guarda en Configuración del proyecto → Propiedades del script, con el
// nombre ANTHROPIC_API_KEY.
var MODELO_ESCANEO = 'claude-sonnet-5';

var ESQUEMA_FACTURA = {
  type: 'object',
  properties: {
    proveedor: { type: 'string' },
    numero_factura: { type: 'string' },
    fecha: { type: 'string' },
    importe_total: { type: 'string' },
    importe_iva: { type: 'string' },
    base_imponible: { type: 'string' },
    tipo_iva: { type: 'string' },
    proveedor_lista: { type: 'string' }
  },
  required: ['proveedor', 'numero_factura', 'fecha', 'importe_total', 'importe_iva', 'base_imponible', 'tipo_iva', 'proveedor_lista'],
  additionalProperties: false
};

var INSTRUCCIONES_FACTURA =
  'Esta es la foto de una factura, albarán o ticket de compra de un restaurante ' +
  '(el restaurante es el CLIENTE: "Buenas y Santas"). Extrae estos datos:\n' +
  '- proveedor: nombre de la empresa que EMITE la factura (el vendedor), tal como aparece. Nunca el cliente.\n' +
  '- numero_factura: número de factura, albarán o ticket.\n' +
  '- fecha: fecha de la factura en formato YYYY-MM-DD.\n' +
  '- importe_total: total a pagar, IVA incluido.\n' +
  '- importe_iva: suma de todas las cuotas de IVA (si hay varios tipos, súmalas).\n' +
  '- base_imponible: suma de las bases imponibles.\n' +
  '- tipo_iva: porcentaje o porcentajes de IVA, por ejemplo "10" o "4, 10, 21".\n' +
  '- proveedor_lista: a cuál de los proveedores habituales (lista de abajo) corresponde el emisor, escrito EXACTAMENTE como el nombre de la lista. ' +
  'Vale aunque en la factura salga con otra razón social, abreviado, con el nombre del grupo o de la marca, o mal escrito. ' +
  'Si no es ninguno de la lista, o no estás seguro, déjalo vacío.\n' +
  'Los importes con punto como separador decimal y sin símbolo de euro (ejemplo: 1234.56). ' +
  'Si un dato no aparece o no se lee con seguridad, deja el campo vacío ("") en vez de inventarlo.';

// La app manda sus proveedores habituales, cada uno con los nombres con que
// puede salir en la factura: "Disbesa (en la factura: INICIATIVAS SODEXO
// S.L., SODEXO)". Así la IA dice directamente cuál es.
function escanearFactura_(data) {
  if (!data.imageBase64) return { ok: false, error: 'No llegó la foto' };
  var lista = (data.proveedores || []).join('\n');
  return pedirJsonAClaude_([
    { type: 'image', source: { type: 'base64', media_type: data.mediaType || 'image/jpeg', data: data.imageBase64 } },
    { type: 'text', text: INSTRUCCIONES_FACTURA + '\n\nProveedores habituales:\n' + (lista || '(no hay lista)') }
  ], ESQUEMA_FACTURA);
}

var ESQUEMA_DICTADO = {
  type: 'object',
  properties: {
    proveedor: { type: 'string' },
    numero_factura: { type: 'string' },
    fecha: { type: 'string' },
    importe_total: { type: 'string' },
    importe_iva: { type: 'string' },
    forma_pago: { type: 'string', enum: ['efectivo', 'tarjeta', 'no_pagado', 'transferencia', ''] },
    responsable: { type: 'string' },
    info: { type: 'string' }
  },
  required: ['proveedor', 'numero_factura', 'fecha', 'importe_total', 'importe_iva', 'forma_pago', 'responsable', 'info'],
  additionalProperties: false
};

// Un gasto de proveedor dicho en voz alta ("factura de Monbake de hoy,
// número 113, importe 99, IVA 1, pagada en efectivo, responsable Luciana").
function dictarGasto_(data) {
  var texto = String(data.texto || '').trim();
  if (!texto) return { ok: false, error: 'No llegó el texto' };
  var instrucciones =
    'Un empleado de un restaurante dictó por voz un gasto a un proveedor. Hoy es ' + (data.fechaHoy || '') + ' (YYYY-MM-DD).\n' +
    'Proveedores de la lista:\n' + (data.proveedores || []).join('\n') + '\n' +
    'Empleados: ' + (data.responsables || []).join(', ') + '.\n' +
    'Extrae:\n' +
    '- proveedor: si coincide con uno de la lista (aunque el dictado lo diga un poco distinto, o diga el nombre que sale en la factura), escribe exactamente el nombre de la lista, sin lo que va entre paréntesis; si no, el nombre tal como se dijo.\n' +
    '- numero_factura: número de factura o albarán.\n' +
    '- fecha: en formato YYYY-MM-DD. "hoy", "ayer", "el lunes"... se calculan a partir de hoy.\n' +
    '- importe_total: importe total (con punto decimal, sin símbolo, ej. 99.50).\n' +
    '- importe_iva: importe del IVA en euros (no el porcentaje).\n' +
    '- forma_pago: efectivo, tarjeta, no_pagado (si dice que no se pagó o queda pendiente) o transferencia.\n' +
    '- responsable: el empleado que lo carga, escrito como en la lista si coincide.\n' +
    '- info: cualquier otro detalle u observación que se haya dicho; si no hay, vacío.\n' +
    'El texto viene de un reconocimiento de voz y puede tener errores (ej. "faceta" por "factura"). ' +
    'Si un dato no se dijo, deja el campo vacío ("") en vez de inventarlo.\n\n' +
    'Dictado: """' + texto + '"""';
  return pedirJsonAClaude_([{ type: 'text', text: instrucciones }], ESQUEMA_DICTADO);
}

// ============================================================================
// Aviso de la app cuando el TPV 1 llega al límite del mes: lo que se eligió
// ("Cambiado" o "Después lo cambio" hasta cierta hora) queda guardado acá,
// en las propiedades del script, para que lo vean todos los dispositivos.
// Una propiedad por mes: avisoTpv1_YYYY-MM.
// ============================================================================
function leerAvisoTpv1_(mes) {
  if (!/^\d{4}-\d{2}$/.test(String(mes || ''))) return { ok: false, error: 'Mes no válido' };
  var texto = PropertiesService.getScriptProperties().getProperty('avisoTpv1_' + mes);
  var estado = {};
  try { estado = texto ? JSON.parse(texto) : {}; } catch (err) { estado = {}; }
  return { ok: true, mes: mes, estado: estado };
}

function guardarAvisoTpv1_(mes, estado) {
  if (!/^\d{4}-\d{2}$/.test(String(mes || ''))) return { ok: false, error: 'Mes no válido' };
  var limpio = {};
  if (estado && estado.cambiado) limpio.cambiado = true;
  if (estado && Number(estado.posponerHasta)) limpio.posponerHasta = Number(estado.posponerHasta);
  PropertiesService.getScriptProperties().setProperty('avisoTpv1_' + mes, JSON.stringify(limpio));
  return { ok: true, mes: mes, estado: limpio };
}

var ESQUEMA_DICTADO_MOVIMIENTO = {
  type: 'object',
  properties: {
    subtipo: { type: 'string', enum: ['cambio', 'varios', 'empanadas_ing', 'empleados', ''] },
    pagado: { type: 'string', enum: ['si', 'no', ''] },
    cliente: { type: 'string' },
    importe: { type: 'string' },
    responsable: { type: 'string' },
    info: { type: 'string' }
  },
  required: ['subtipo', 'pagado', 'cliente', 'importe', 'responsable', 'info'],
  additionalProperties: false
};

var DICTADO_TIPO_TEXTO_ = {
  ingreso: 'un INGRESO de dinero a la caja',
  egreso: 'un RETIRO de dinero de la caja',
  empanadas: 'una venta de EMPANADAS a un cliente'
};

// Un ingreso, un retiro o unas empanadas dichos en voz alta. El tipo lo
// pone la app (la hoja que está abierta), acá solo se sacan los campos.
function dictarMovimiento_(data) {
  var texto = String(data.texto || '').trim();
  if (!texto) return { ok: false, error: 'No llegó el texto' };
  var tipoTexto = DICTADO_TIPO_TEXTO_[data.tipo];
  if (!tipoTexto) return { ok: false, error: 'Tipo de movimiento no válido' };
  var instrucciones =
    'Un empleado de un restaurante dictó por voz ' + tipoTexto + '.\n' +
    'Empleados: ' + (data.responsables || []).join(', ') + '.\n' +
    'Extrae:\n' +
    '- subtipo: solo si es un ingreso: cambio (entra cambio/sencillo para la caja), empanadas_ing (pago de empanadas), empleados (dinero que pone un empleado) o varios (cualquier otro). Si no es un ingreso, vacío.\n' +
    '- pagado: solo si son empanadas: si (pagaron) o no (no pagaron, quedan debiendo). Si no se dijo o no son empanadas, vacío.\n' +
    '- cliente: solo si son empanadas: uno de ' + (data.clientesEmpanadas || []).join(', ') + ' si coincide (escríbelo igual), si no el nombre tal como se dijo.\n' +
    '- importe: importe en euros (con punto decimal, sin símbolo, ej. 45.50).\n' +
    '- responsable: el empleado que lo carga, escrito como en la lista si coincide.\n' +
    '- info: motivo, detalle u observación que se haya dicho; si no hay, vacío.\n' +
    'El texto viene de un reconocimiento de voz y puede tener errores. ' +
    'Si un dato no se dijo, deja el campo vacío ("") en vez de inventarlo.\n\n' +
    'Dictado: """' + texto + '"""';
  return pedirJsonAClaude_([{ type: 'text', text: instrucciones }], ESQUEMA_DICTADO_MOVIMIENTO);
}

// Llama a Claude con el contenido dado y devuelve el JSON que pide el
// esquema, con ok: true (o { ok: false, error } si algo falla).
function pedirJsonAClaude_(contenido, esquema) {
  var clave = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!clave) return { ok: false, error: 'Falta la clave ANTHROPIC_API_KEY en las Propiedades del script' };

  var cuerpo = {
    model: MODELO_ESCANEO,
    max_tokens: 16000,
    output_config: {
      effort: 'low',
      format: { type: 'json_schema', schema: esquema }
    },
    messages: [{ role: 'user', content: contenido }]
  };

  // Si Anthropic está saturado (429 / 529 / 5xx) se reintenta un par de veces.
  var resp, codigo, json;
  for (var intento = 0; intento < 3; intento++) {
    resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-api-key': clave, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify(cuerpo),
      muteHttpExceptions: true
    });
    codigo = resp.getResponseCode();
    if (codigo !== 429 && codigo < 500) break;
    Utilities.sleep(2000 * (intento + 1));
  }

  try { json = JSON.parse(resp.getContentText()); } catch (err) { json = null; }
  if (codigo !== 200 || !json) {
    var msg = (json && json.error && json.error.message) || ('Error de la IA (' + codigo + ')');
    return { ok: false, error: msg };
  }
  if (json.stop_reason === 'refusal') return { ok: false, error: 'La IA no quiso responder a esto' };

  var texto = (json.content || []).filter(function (b) { return b.type === 'text'; })
    .map(function (b) { return b.text; }).join('');
  try {
    var datos = JSON.parse(texto);
    datos.ok = true;
    return datos;
  } catch (err) {
    return { ok: false, error: 'No se pudo entender la respuesta de la IA' };
  }
}

// Ejecutar UNA VEZ a mano desde el editor (menú de funciones → autorizarEscaneo
// → Ejecutar) para que Google pida el permiso de conectarse con Claude.
// También sirve para comprobar que la clave está bien puesta.
function autorizarEscaneo() {
  var clave = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!clave) { Logger.log('Falta la propiedad ANTHROPIC_API_KEY'); return; }
  var resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/models', {
    headers: { 'x-api-key': clave, 'anthropic-version': '2023-06-01' },
    muteHttpExceptions: true
  });
  Logger.log(resp.getResponseCode() === 200 ? 'OK: permiso dado y clave correcta' : 'Error ' + resp.getResponseCode() + ': ' + resp.getContentText());
}

// ============================================================================
// Acceso: solo estas cuentas de Google pueden usar la app.
// ============================================================================
// Para dar o quitar acceso, cambiar esta lista, pegar el archivo y crear una
// Nueva versión. Quitar una cuenta de acá la deja afuera enseguida, aunque
// tenga una sesión abierta.
var CUENTAS_PERMITIDAS_ = [
  'lfisbein@gmail.com',
  'buenasysantasmadrid@gmail.com',
  'buenasysantas9@gmail.com'
];
// ID de cliente de OAuth (Google Cloud, proyecto "Buenas y Santas"). No es
// secreto: es el mismo que usa la página.
var GOOGLE_CLIENT_ID_ = '831497268247-b3hhi4pqlvnd18vv7dph2ihodmih2jmp.apps.googleusercontent.com';
var DIAS_SESION_ = 90;

// La app manda el "credential" que le dio Google al iniciar sesión. Se
// comprueba con Google que es válido y para esta app, y que la cuenta está
// en la lista. Si todo va bien se crea una sesión (un código al azar que
// vale DIAS_SESION_ días) guardada en las propiedades del script.
function iniciarSesion_(idToken) {
  if (!idToken) return { ok: false, error: 'No llegó el inicio de sesión de Google.' };
  var resp = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken), { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) return { ok: false, error: 'Google no dio por válido el inicio de sesión. Prueba de nuevo.' };
  var info = JSON.parse(resp.getContentText());
  if (info.aud !== GOOGLE_CLIENT_ID_) return { ok: false, error: 'El inicio de sesión no es de esta app.' };
  var email = String(info.email || '').toLowerCase();
  if (String(info.email_verified) !== 'true') return { ok: false, error: 'La cuenta ' + email + ' no está verificada por Google.' };
  if (CUENTAS_PERMITIDAS_.indexOf(email) === -1) return { ok: false, error: 'La cuenta ' + email + ' no tiene acceso a la caja.' };

  var props = PropertiesService.getScriptProperties();
  borrarSesionesVencidas_(props);
  var sesion = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  var expira = Date.now() + DIAS_SESION_ * 24 * 60 * 60 * 1000;
  props.setProperty('sesion_' + sesion, JSON.stringify({ email: email, expira: expira }));
  return { ok: true, sesion: sesion, email: email, expira: expira };
}

// Devuelve la cuenta de la sesión, o null si no es válida (no existe,
// venció o la cuenta ya no está en la lista).
function sesionValida_(sesion) {
  if (!sesion || !/^[a-f0-9]{64}$/.test(String(sesion))) return null;
  var texto = PropertiesService.getScriptProperties().getProperty('sesion_' + sesion);
  if (!texto) return null;
  var s;
  try { s = JSON.parse(texto); } catch (err) { return null; }
  if (!s || !(s.expira > Date.now()) || CUENTAS_PERMITIDAS_.indexOf(s.email) === -1) return null;
  return s.email;
}

function respuestaSinSesion_() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, sinSesion: true, error: 'Hay que iniciar sesión con una cuenta autorizada.' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function borrarSesionesVencidas_(props) {
  var todas = props.getProperties();
  Object.keys(todas).forEach(function (clave) {
    if (clave.indexOf('sesion_') !== 0) return;
    var s = null;
    try { s = JSON.parse(todas[clave]); } catch (err) {}
    if (!s || !(s.expira > Date.now())) props.deleteProperty(clave);
  });
}

// Para correr a mano desde el editor (por ejemplo si se pierde un móvil):
// cierra la sesión en todos los dispositivos. Cada uno tendrá que volver
// a iniciar sesión con Google.
function cerrarTodasLasSesiones() {
  var props = PropertiesService.getScriptProperties();
  var borradas = 0;
  Object.keys(props.getProperties()).forEach(function (clave) {
    if (clave.indexOf('sesion_') === 0) { props.deleteProperty(clave); borradas++; }
  });
  Logger.log('Sesiones cerradas: ' + borradas);
}
