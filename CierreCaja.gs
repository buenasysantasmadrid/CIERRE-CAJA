function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    if (data.test) {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, msg: 'Conexión OK' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

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

    TURNOS.forEach(function (turnoInfo) {
      var t = data[turnoInfo.key];
      if (!t) return;
      escribirRegistroYMovimientos_(registro, mov, data, t, turnoInfo.label);
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

    var contabilidad = escribirContabilidad_(data);
    var albaranes = escribirAlbaranes_(data); // ver Albaranes.gs

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, hojaDia: hojaDia, contabilidad: contabilidad, albaranes: albaranes }))
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
      'Gastos no efectivo (info)', 'Cant. movimientos', 'Última actualización', 'Datos JSON (uso interno)'
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
      'Nº Factura', 'Info', 'IVA (€)', 'Importe', 'Última actualización'
    ]);
    mov.setFrozenRows(1);
  }
  return mov;
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
      filas.push([
        idTurno, m.id || '', fechaStr, turnoLabel,
        TIPO_LABEL_MAP_[m.tipo] || m.tipo || '', SUBTIPO_LABEL_MAP_[m.subtipo] || '',
        proveedorMotivo, m.responsable || '', m.factura || '', (m.info || ''),
        (m.iva != null ? m.iva : ''), m.importe, actualizado
      ]);
    });
  }

  if (filas.length) {
    mov.getRange(2, 1, filas.length, filas[0].length).setValues(filas);
  }

  Logger.log('Reconstruidas ' + filas.length + ' filas en "Movimientos" a partir de "Registro".');
  return filas.length;
}

function escribirRegistroYMovimientos_(registro, mov, data, t, turnoLabel) {
  var c = t.calc || {};
  var id = t.id || '';
  var estado = t.estado || 'Sincronizado';

  var datosJSON = JSON.stringify({
    id: t.id,
    fondoFijo: t.fondoFijo,
    totalFacturado: t.totalFacturado,
    tpv1: t.tpv1,
    tpv2: t.tpv2,
    denom: t.denom,
    movimientos: t.movimientosRaw || t.movimientos,
    calc: t.calc
  });

  var filaRegistro = [
    id, estado, data.fecha, turnoLabel, data.negocio,
    t.fondoFijo, t.totalFacturado, c.empanadas, c.facturadoNeto,
    t.tpv1, t.tpv2, c.ventasEfectivo, c.ingresoEfectivo,
    c.gastoEfectivo, c.egreso, c.esperado, c.totalContado, c.diferencia,
    c.gastoNoEfectivo, (t.movimientos || []).length, new Date(), datosJSON
  ];

  var filaExistente = id ? buscarFilaPorId_(registro, id) : -1;
  if (filaExistente > -1) {
    registro.getRange(filaExistente, 1, 1, filaRegistro.length).setValues([filaRegistro]);
  } else {
    registro.appendRow(filaRegistro);
  }

  if (id) borrarFilasPorId_(mov, id);
  (t.movimientos || []).forEach(function (m) {
    var proveedorMotivo = m.proveedor || m.motivo || '';
    if (PROVEEDORES_CON_DETALLE_LIBRE_.indexOf(m.proveedor) > -1 && m.proveedorDetalle) proveedorMotivo += ' — ' + m.proveedorDetalle;
    mov.appendRow([
      id, m.id || '', data.fecha, turnoLabel, m.tipo, m.subtipo,
      proveedorMotivo, m.responsable || '',
      m.factura || '', (m.info || ''), (m.iva != null ? m.iva : ''), m.importe, new Date()
    ]);
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

function escribirHojaDelDiaExacta_(ss, data) {
  var nombre = nombreHojaDia_(data.fecha);
  var hoja = ss.getSheetByName(nombre);

  if (!hoja) {
    var plantilla = ss.getSheetByName('1');
    if (!plantilla) {
      throw new Error('No se encontró la hoja "1" (la plantilla) en esta planilla, así que no se pudo crear la pestaña del día.');
    }
    hoja = plantilla.copyTo(ss);
    hoja.setName(nombre);
    ss.setActiveSheet(hoja);
    ss.moveActiveSheet(ss.getNumSheets());
  }

  var md = data.mediodia || {};
  var nc = data.noche || {};
  var movsMd = md.movimientos || [];
  var movsNc = nc.movimientos || [];

  hoja.getRange('B2').setValue(textoFechaLarga_(data.fecha));

  // ---- MEDIODÍA (posiciones de la plantilla "1" nueva) ----
  hoja.getRange('C45').setValue(md.totalFacturado || 0); // "TOTAL CIERRE SISTEMA"
  hoja.getRange('I5').setValue(md.tpv1 || 0);
  hoja.getRange('K5').setValue(md.tpv2 || 0);
  hoja.getRange('C43').setValue(md.fondoFijo || 0);

  escribirFilasFijas_(hoja, 6, 16, ['A', 'B', 'C', 'D', 'E', 'F'], filtrarPorSubtipo_(movsMd, 'Efectivo').map(filaGasto_));
  escribirFilasFijas_(hoja, 5, 6, ['M', 'N', 'O', 'P', 'Q', 'R'], filtrarPorSubtipo_(movsMd, 'Efectivo antiguo').map(filaGasto_));
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
  hoja.getRange('I63').setValue(tpv1NochePropio + tpv2NochePropio);

  escribirFilasFijas_(hoja, 63, 10, ['A', 'B', 'C', 'D', 'E', 'F'], filtrarPorSubtipo_(movsNc, 'Efectivo').map(filaGasto_));
  escribirFilasFijas_(hoja, 63, 6, ['M', 'N', 'O', 'P', 'Q', 'R'], filtrarPorSubtipo_(movsNc, 'Efectivo antiguo').map(filaGasto_));
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
    (t.totalFacturado || 0) !== 0 || (t.tpv1 || 0) !== 0 || (t.tpv2 || 0) !== 0;
  return movs * 1000 + (tieneActividad ? 1 : 0);
}

function doGet(e) {
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

  var listarDias = e && e.parameter && e.parameter.listarDias;
  if (listarDias) {
    return ContentService
      .createTextOutput(JSON.stringify(listarDiasConDatos_()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var listarFacturas = e && e.parameter && e.parameter.listarFacturas;
  if (listarFacturas) {
    var proveedorFiltro = (e.parameter && e.parameter.proveedor) || '';
    var detalleFiltro = (e.parameter && e.parameter.detalle) || '';
    return ContentService
      .createTextOutput(JSON.stringify(listarFacturasProveedores_(proveedorFiltro, detalleFiltro)))
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

function coincideProveedor_(textoProv, proveedor, detalleNorm) {
  if (!proveedor) return true;
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
      resultado.push({
        idMovimiento: fila[idxIdMov],
        idTurno: fila[idxIdTurno],
        fecha: (fechaRaw instanceof Date) ? Utilities.formatDate(fechaRaw, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(fechaRaw || ''),
        turno: fila[idxTurno],
        subtipo: fila[idxSubtipo],
        proveedorTexto: textoProv,
        factura: fila[idxFactura],
        info: String(fila[idxInfo] || ''),
        iva: fila[idxIva],
        importe: fila[idxImporte]
      });
    }

    resultado.sort(function (a, b) { return a.fecha < b.fecha ? 1 : (a.fecha > b.fecha ? -1 : 0); });

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

    planillas.forEach(function (p) {
      var r = listarFacturasProveedoresEnPlanilla_(p.ss, proveedor, detalleNorm);
      if (r.error) error = r.error;
      resultado = resultado.concat(r.facturas);
    });

    resultado.sort(function (a, b) { return a.fecha < b.fecha ? 1 : (a.fecha > b.fecha ? -1 : 0); });

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
      manejarEdicionRegistro_(e, hoja);
    } else if (/^\d{2}-\d{2}-\d{4}$/.test(nombreHoja)) {
      manejarEdicionHojaDelDia_(e, hoja, nombreHoja);
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
  var idxJSON = -1;
  for (var h = 0; h < header.length; h++) {
    if (String(header[h] || '').indexOf('Datos JSON') === 0) { idxJSON = h; break; }
  }
  if (idxJSON === -1) return; // pestaña vieja, sin dónde guardar el resultado

  var colsEditables = [idxFondoFijo, idxTotalFacturado, idxTpv1, idxTpv2]
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

  var tocaMd = tocaCelda_('C', 45) || tocaCelda_('I', 5) || tocaCelda_('K', 5) || tocaCelda_('C', 43) ||
    tocaRango_('J', 13, 17) || tocaRango_('I', 18, 23) || tocaRango_('J', 18, 23);
  var tocaNc = tocaCelda_('C', 87) || tocaCelda_('I', 61) || tocaCelda_('K', 61) || tocaCelda_('C', 86) ||
    tocaRango_('J', 71, 75) || tocaRango_('I', 76, 81) || tocaRango_('J', 76, 81);

  if (!tocaMd && !tocaNc) return; // se editó otra celda de esta pestaña (movimientos, etc.)

  var diaExistente = obtenerDiaJSON_(hoja.getParent(), fechaStr);
  if (!diaExistente || !diaExistente.encontrado) return; // este día nunca se sincronizó desde la app

  var huboCambio = false;
  if (tocaMd && diaExistente.mediodia && diaExistente.mediodia.id) {
    aplicarEdicionTurnoDesdeHoja_(hoja, diaExistente.mediodia, 'C45', 'I5', 'K5', 'C43', 13, 18);
    huboCambio = true;
  }
  if (tocaNc && diaExistente.noche && diaExistente.noche.id) {
    aplicarEdicionTurnoDesdeHoja_(hoja, diaExistente.noche, 'C87', 'I61', 'K61', 'C86', 71, 76);
    huboCambio = true;
  }
  if (!huboCambio) return;

  recalcularDia_(diaExistente);
  guardarDiaCompleto_(hoja.getParent(), fechaStr, diaExistente.negocio, diaExistente);
}

// Lee de la pestaña bonita el Total facturado / TPV 1 / TPV 2 / Fondo fijo
// y el conteo completo de billetes y monedas de un turno, y los vuelca
// sobre el objeto `turno` (que ya viene con el resto de sus datos —
// movimientos, id, etc.— intactos, para no perder nada de lo que la app
// cargó).
function aplicarEdicionTurnoDesdeHoja_(hoja, turno, celdaFacturado, celdaTpv1, celdaTpv2, celdaFondoFijo, filaBilletesDesde, filaMonedasDesde) {
  turno.totalFacturado = Number(hoja.getRange(celdaFacturado).getValue()) || 0;
  turno.tpv1 = Number(hoja.getRange(celdaTpv1).getValue()) || 0;
  turno.tpv2 = Number(hoja.getRange(celdaTpv2).getValue()) || 0;
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
    if (esNoche && md) {
      totalFacturadoPropio = Math.max(0, (turno.totalFacturado || 0) - (md.totalFacturado || 0));
      tpv1Propio = Math.max(0, (turno.tpv1 || 0) - (md.tpv1 || 0));
      tpv2Propio = Math.max(0, (turno.tpv2 || 0) - (md.tpv2 || 0));
    }

    var facturadoNeto = totalFacturadoPropio - empanadas;
    var ventasEfectivo = facturadoNeto - (tpv1Propio + tpv2Propio);
    var totalContado = totalContadoDesdeDenom_(turno.denom);
    var esperado = (turno.fondoFijo || 0) + ventasEfectivo + ingresoEfectivo - gastoEfectivo - egreso;
    var diferencia = totalContado - esperado;

    turno.calc = {
      empanadas: empanadas, facturadoNeto: facturadoNeto, ventasEfectivo: ventasEfectivo,
      ingresoEfectivo: ingresoEfectivo, gastoEfectivo: gastoEfectivo, egreso: egreso,
      esperado: esperado, totalContado: totalContado, diferencia: diferencia,
      gastoNoEfectivo: gastoNoEfectivo,
      totalFacturadoPropio: totalFacturadoPropio, tpv1Propio: tpv1Propio, tpv2Propio: tpv2Propio
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
    if (idxEsperado > -1) registro.getRange(filaN, idxEsperado + 1).setValue(c.esperado || 0);
    if (idxContado > -1) registro.getRange(filaN, idxContado + 1).setValue(c.totalContado || 0);
    if (idxDiferencia > -1) registro.getRange(filaN, idxDiferencia + 1).setValue(c.diferencia || 0);
    if (idxFacturadoNeto > -1) registro.getRange(filaN, idxFacturadoNeto + 1).setValue(c.facturadoNeto || 0);
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
function listarDiasConDatos_() {
  try {
    var fechasSet = {};

    planillasCierreCajaRecientes_(MESES_HISTORIAL_).forEach(function (p) {
      var registro = p.ss.getSheetByName('Registro');
      if (!registro) return;

      var valores = registro.getDataRange().getValues();
      var header = valores[0];
      var idxFecha = header.indexOf('Fecha');
      if (idxFecha === -1) return;

      for (var i = 1; i < valores.length; i++) {
        var raw = valores[i][idxFecha];
        var fechaStr = (raw instanceof Date)
          ? Utilities.formatDate(raw, Session.getScriptTimeZone(), 'yyyy-MM-dd')
          : String(raw || '');
        if (fechaStr) fechasSet[fechaStr] = true;
      }
    });

    return { ok: true, fechas: Object.keys(fechasSet) };
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
// Columnas que esta función escribe — el resto (C/F/G/I/L fórmulas propias
// de la hoja, P-U, V, Y-AG a mano o vinculadas a otro Sheet) no se toca:
//   A  días trabajados de ese día (0 / 0,5 / 1 — ver turnoTieneActividad_)
//   B  fecha
//   D  MEDIO DIA — total facturado de Mediodía
//   E  NOCHE — la parte propia de Noche (total del día completo - Mediodía)
//   H  EMPANADAS — Mediodía + Noche
//   J  TPV 1 — acumulado del día completo (el que ya carga Noche)
//   K  TPV 2 — ídem
//   M  EFEVO — efectivo contado al cerrar el último turno trabajado (es lo
//      que queda como fondo fijo para el día siguiente)
//   N  DIFERENCIA — Mediodía + Noche
//   O  RETIRA — egresos de Mediodía + Noche
// A34 queda con la fórmula =SUM(A3:A33), así el total de días trabajados
// del mes se actualiza solo cada vez que se escribe una fila nueva.
var MESES_MAYUS_ = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];

function turnoTieneActividad_(t) {
  var c = t.calc || {};
  return (c.totalContado || 0) !== 0 || (t.fondoFijo || 0) !== 0 ||
    (t.totalFacturado || 0) !== 0 || (t.tpv1 || 0) !== 0 || (t.tpv2 || 0) !== 0;
}

// Trae la pestaña del mes en la planilla de Contabilidad DEL AÑO que
// corresponda (ver Indice.gs — esa planilla anual se crea sola la primera
// vez que hace falta), creando la pestaña del mes si todavía no existe:
// duplica la pestaña del mes anterior o, si tampoco existe (por ejemplo, el
// primer mes del año), "MASTER", la renombra y limpia las filas de días
// 3-33 en las columnas de arriba (para no arrastrar datos viejos del mes
// que se copió).
function getOrCrearPestanaContabilidad_(ss, mesIndex /* 0-11 */) {
  var nombre = MESES_MAYUS_[mesIndex];
  var hoja = ss.getSheetByName(nombre);
  if (hoja) return hoja;

  var origenNombre = MESES_MAYUS_[(mesIndex + 11) % 12];
  var origen = ss.getSheetByName(origenNombre) || ss.getSheetByName('MASTER');
  if (!origen) throw new Error('No se encontró ninguna pestaña de mes para duplicar en la planilla de Contabilidad.');

  hoja = origen.copyTo(ss);
  hoja.setName(nombre);
  var idxOrigen = origen.getIndex();
  ss.setActiveSheet(hoja);
  ss.moveActiveSheet(idxOrigen + 1);

  ['A', 'B', 'D', 'E', 'H', 'J', 'K', 'M', 'N', 'O'].forEach(function (col) {
    hoja.getRange(col + '3:' + col + '33').clearContent();
  });
  hoja.getRange('A34').setFormula('=SUM(A3:A33)');

  return hoja;
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
    var fila = 3 + (diaDelMes - 1);

    var diasHoy = (mdActivo && ncActivo) ? 1 : ((mdActivo || ncActivo) ? 0.5 : 0);
    var mediodiaTotal = md.totalFacturado || 0;
    var nochePropio = Math.max(0, (nc.totalFacturado || 0) - mediodiaTotal);
    var empanadasDia = (cMd.empanadas || 0) + (cNc.empanadas || 0);
    var tpv1Dia = ncActivo ? (nc.tpv1 || 0) : (md.tpv1 || 0);
    var tpv2Dia = ncActivo ? (nc.tpv2 || 0) : (md.tpv2 || 0);
    var efevoDia = ncActivo ? (cNc.totalContado || 0) : (cMd.totalContado || 0);
    var diferenciaDia = (cMd.diferencia || 0) + (cNc.diferencia || 0);
    var retiraDia = (cMd.egreso || 0) + (cNc.egreso || 0);

    hoja.getRange('A' + fila).setValue(diasHoy);
    // Mediodía (no medianoche): si el huso horario de esta planilla de
    // Contabilidad no coincide exactamente con el del proyecto de Apps
    // Script, medianoche puede caer del lado del día anterior al mostrarse,
    // corriendo la fecha visible un día para atrás. Al mediodía queda lejos
    // de cualquier límite de huso horario real.
    hoja.getRange('B' + fila).setValue(new Date(anio, mes - 1, diaDelMes, 12));
    hoja.getRange('D' + fila).setValue(mediodiaTotal);
    hoja.getRange('E' + fila).setValue(nochePropio);
    hoja.getRange('H' + fila).setValue(empanadasDia);
    hoja.getRange('J' + fila).setValue(tpv1Dia);
    hoja.getRange('K' + fila).setValue(tpv2Dia);
    hoja.getRange('M' + fila).setValue(efevoDia);
    hoja.getRange('N' + fila).setValue(diferenciaDia);
    hoja.getRange('O' + fila).setValue(retiraDia);

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
    .addSeparator()
    .addItem('Albaranes: congelar datos viejos (una sola vez)', 'congelarAlbaranesViejosDesdeMenu')
    .addItem('Albaranes: reenviar un mes desde Cierre de Caja…', 'reenviarMesAAlbaranesDesdeMenu')
    .addToUi();
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
          var filaC = 3 + (diaDelMes - 1);
          ['A', 'B', 'D', 'E', 'H', 'J', 'K', 'M', 'N', 'O'].forEach(function (col) {
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

  return resumen;
}

