// ============================================================================
// ALBARANES — cada gasto de proveedor que se carga en la app se escribe
// también en la planilla anual de Albaranes (tipo ALBARANES del Índice, ver
// Indice.gs): en la pestaña de ese proveedor, en el bloque del mes de la
// fecha.
//
// Forma de cada pestaña de proveedor (la de siempre): 12 bloques uno debajo
// del otro, uno por mes (enero arriba). Cada bloque tiene una fila de
// encabezados que empieza con "FECHA" en la columna A, 42 filas de datos y
// una fila de total (ej. enero: encabezados fila 3, datos 4–45, total 46).
// Las columnas se ubican leyendo los encabezados de cada bloque (no por
// posición), porque no todas las pestañas son iguales. Formato actual
// (desde el 28/09/2026, ver albaranesPasarAFormatoNuevo más abajo):
//   FECHA · Nº FACTURA · IVA · IMPORTE · OBERVACIONES · FORMA DE PAGO · DIA DE PAGO · VARIOS   (la mayoría)
//   FECHA · PROVEEDOR · Nº FACTURA · IVA · IMPORTE · OBERVACIONES · FORMA DE PAGO · DIA DE PAGO · VARIOS   (VARIOS y SUPER)
//   FECHA · NOMBRE · IMPORTE   (EXTRAS)
// Formato de antes (se sigue entendiendo, por si queda algún bloque así):
//   FECHA · FACTURA Nº · IMPORTE · IVA · OBERVACIONES   (la mayoría)
//   FECHA · INFO · FACTURA Nº · IMPORTE · IVA · OBERVACIONES   (VINO)
//   FECHA · SUPER · IMPORTE · IVA · OBERVACIONES   (SUPER)
//   FECHA · PROVEEDOR · OBERVACIONES · IMPORTE · IVA   (VARIOS)
//
// La columna Z (oculta) guarda el ID del movimiento de la app, para poder
// actualizarlo o borrarlo sin duplicar: la app manda el día entero en cada
// sincronización. La pestaña oculta "IDs app" lleva la cuenta de qué ID
// está en qué pestaña y de qué fecha es. Las filas cargadas a mano (sin ID
// en la columna Z) no se tocan, salvo para ordenarlas por fecha dentro del
// bloque.
//
// Los datos de 2026 que venían del sistema viejo (fórmulas IMPORTRANGE) se
// cargan como valores fijos con "Albaranes: restaurar datos viejos 2026"
// (menú del Índice, ver restaurarAlbaranesViejos2026 más abajo). Mientras
// un bloque tenga fórmulas, el script no escribe en él, para no pisarlas.
// ============================================================================

var ALBARANES_COL_ID_ = 26; // Z
var ALBARANES_ANCHO_ = 9;   // A–I: la parte visible de cada bloque
var ALBARANES_FILAS_DATOS_ = 42;
var ALBARANES_HOJA_IDS_ = 'IDs app';
// Pestañas del archivo de Albaranes que no son de un proveedor con la forma
// de siempre: no se restauran ni se escribe en ellas.
var ALBARANES_PESTANAS_EXCLUIDAS_ = ['TOTALES', 'VERDE REBELDE', ALBARANES_HOJA_IDS_];

// Tildes y demás marcas que quedan sueltas después de normalize('NFD').
var DIACRITICOS_ = new RegExp('[' + String.fromCharCode(0x300) + '-' + String.fromCharCode(0x36f) + ']', 'g');

function normalizarClave_(v) {
  return String(v == null ? '' : v)
    .normalize('NFD').replace(DIACRITICOS_, '')
    .toUpperCase().replace(/\s+/g, ' ').trim();
}

// Filas (1-based) de los encabezados de cada bloque de mes: las celdas de la
// columna A que dicen "FECHA", en orden (la primera es enero).
function filasEncabezadoAlbaranes_(hoja) {
  var ultima = hoja.getLastRow();
  if (ultima < 1) return [];
  var colA = hoja.getRange(1, 1, ultima, 1).getValues();
  var filas = [];
  for (var i = 0; i < colA.length; i++) {
    if (normalizarClave_(colA[i][0]) === 'FECHA') filas.push(i + 1);
  }
  return filas;
}

// Qué columna (0-based dentro de A–I) lleva cada dato en este bloque.
function mapaColumnasAlbaranes_(encabezados) {
  var mapa = {};
  encabezados.forEach(function (h, i) {
    var k = normalizarClave_(h).replace(/[^A-Z]/g, '');
    if (k === 'FECHA') mapa.fecha = i;
    else if (k === 'FORMADEPAGO') mapa.formaPago = i;
    else if (k === 'DIADEPAGO') mapa.diaPago = i;
    else if (k.indexOf('FACTURA') > -1 || k === 'ALBARAN') mapa.factura = i; // FACTURA Nº / Nº FACTURA
    else if (k === 'IMPORTE' || k === 'TOTAL') mapa.importe = i;
    else if (k === 'IVA') mapa.iva = i;
    else if (k.indexOf('OB') === 0) mapa.obs = i; // OBERVACIONES / OBSERVACIONES
    else if (k === 'INFO') mapa.info = i;
    else if (k === 'SUPER' || k === 'PROVEEDOR' || k === 'NOMBRE') mapa.detalle = i;
  });
  return mapa;
}

function pestanaAlbaranesParaProveedor_(ss, proveedor) {
  var buscado = normalizarClave_(proveedor);
  if (!buscado) return null;
  var hojas = ss.getSheets();
  for (var i = 0; i < hojas.length; i++) {
    var nombre = hojas[i].getName();
    if (ALBARANES_PESTANAS_EXCLUIDAS_.indexOf(nombre) > -1) continue;
    if (normalizarClave_(nombre) === buscado) return hojas[i];
  }
  return null;
}

function getOrCrearHojaIdsAlbaranes_(ss) {
  var hoja = ss.getSheetByName(ALBARANES_HOJA_IDS_);
  if (!hoja) {
    hoja = ss.insertSheet(ALBARANES_HOJA_IDS_);
    hoja.appendRow(['ID Movimiento', 'Fecha', 'Pestaña', 'Fecha factura']);
    hoja.setFrozenRows(1);
    hoja.hideSheet();
  }
  return hoja;
}

// Gastos de proveedor del día, con las claves internas de la app (tipo
// 'gasto', subtipo 'efectivo'/'tarjeta'/...). Sirve tanto para lo que manda
// la app al guardar (movimientosRaw) como para un día leído de "Registro"
// con obtenerDiaJSON_ (ahí "movimientos" ya son los internos).
function gastosDelDia_(dia) {
  var gastos = [];
  ['mediodia', 'noche'].forEach(function (k) {
    var t = dia && dia[k];
    if (!t) return;
    (t.movimientosRaw || t.movimientos || []).forEach(function (m) {
      var tipo = String(m.tipo || '').toLowerCase();
      if (tipo === 'gasto' && m.id && m.proveedor) gastos.push(m);
    });
  });
  return gastos;
}

function fechaComoDate_(fechaISO) {
  var p = String(fechaISO).split('-');
  return new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10), 12);
}

function filaVaciaAlbaranes_() {
  var fila = [];
  for (var i = 0; i < ALBARANES_ANCHO_; i++) fila.push('');
  return fila;
}

var FORMA_PAGO_ALBARAN_ = {
  efectivo: 'EFECTIVO', efectivo_antiguo: 'EFECTIVO', tarjeta: 'TARJETA',
  transferencia: 'TRANSFERENCIA', no_pagado: 'NO PAGADO', no_efectivo: 'NO EFECTIVO'
};

// Forma de pago, día de pago y observaciones (Info sin las etiquetas de
// pago) de un gasto. Una factura que estaba "No pagado" y se pagó después
// tiene en Info "PAGADO <MODALIDAD> <fecha> · ..." (ver
// cambiarFormaPagoAlbaran_): el día de pago es esa fecha. Si se pagó en el
// momento, es la fecha de la caja en la que se cargó; si no está pagada,
// queda vacío.
function pagoDeGasto_(m, fechaCajaISO) {
  var info = String(m.info || '');
  var pagado = /^PAGADO(?:\s+([A-Z]+))?(?:\s+(\d{4}-\d{2}-\d{2}))?/.exec(info);
  var forma = FORMA_PAGO_ALBARAN_[m.subtipo] || '';
  var dia = '';
  if (pagado) {
    forma = pagado[1] || forma;
    dia = pagado[2] || '';
  } else if (m.subtipo !== 'no_pagado') {
    dia = fechaCajaISO || '';
  }
  return { forma: forma, dia: dia, obs: quitarTagInfo_(info) };
}

// Arma la fila (A–I) de un gasto según las columnas de ese bloque.
// `fechaISO` es la fecha de la factura; `fechaCajaISO`, la de la caja en la
// que se cargó.
function filaAlbaran_(m, fechaISO, mapa, fechaCajaISO) {
  var fila = filaVaciaAlbaranes_();
  var formaPago = m.subtipo === 'efectivo' ? 'EFECTIVO' : '';
  var info = String(m.info || '');
  // Las formas de pago que no son efectivo ya vienen marcadas en Info
  // (TARJETA, NO PAGADO, TRANSFERENCIA, PAGADO ...); el efectivo no.
  var textoObs = [info, formaPago].filter(String).join(' · ');
  var detalle = m.proveedorDetalle || '';

  // Formato nuevo: la forma y el día de pago van en sus columnas, y en
  // observaciones queda solo lo que se escribió.
  if (mapa.formaPago != null) {
    var pago = pagoDeGasto_(m, fechaCajaISO);
    fila[mapa.formaPago] = pago.forma;
    if (mapa.diaPago != null && pago.dia) fila[mapa.diaPago] = fechaComoDate_(pago.dia);
    info = pago.obs;
    textoObs = pago.obs;
  }

  if (mapa.fecha != null) fila[mapa.fecha] = fechaComoDate_(fechaISO);
  if (mapa.importe != null) fila[mapa.importe] = Number(m.importe) || 0;
  if (mapa.iva != null && m.iva != null && m.iva !== '') fila[mapa.iva] = Number(m.iva) || 0;

  if (mapa.factura != null) fila[mapa.factura] = m.factura || '';
  else if (m.factura) textoObs = ['Fact. ' + m.factura, textoObs].filter(String).join(' · ');

  if (mapa.info != null) {
    fila[mapa.info] = info;
    textoObs = mapa.formaPago != null ? '' : formaPago;
  }

  if (mapa.obs != null) {
    fila[mapa.obs] = textoObs;
  } else if (!detalle) {
    detalle = textoObs; // ej. EXTRAS: solo tiene NOMBRE
  }
  if (mapa.detalle != null) fila[mapa.detalle] = detalle;

  return fila;
}

// Reescribe un bloque de mes: saca los IDs de `idsABorrar`, pone/actualiza
// las filas de `nuevas` ({id, fila}) y deja todo ordenado por fecha, sin
// huecos. Las filas sin ID (cargadas a mano) se conservan.
function reescribirBloqueAlbaranes_(hoja, filaEncabezado, idsABorrar, nuevas) {
  var primera = filaEncabezado + 1;
  var rango = hoja.getRange(primera, 1, ALBARANES_FILAS_DATOS_, ALBARANES_ANCHO_);
  var formulas = rango.getFormulas();
  for (var f = 0; f < formulas.length; f++) {
    for (var c = 0; c < formulas[f].length; c++) {
      if (formulas[f][c]) {
        throw new Error('El bloque de la pestaña "' + hoja.getName() + '" (fila ' + primera + ') todavía tiene fórmulas (IMPORTRANGE del sistema viejo) — hay que borrarlas antes de que la app pueda escribir ahí.');
      }
    }
  }
  var valores = rango.getValues();
  var rangoIds = hoja.getRange(primera, ALBARANES_COL_ID_, ALBARANES_FILAS_DATOS_, 1);
  var ids = rangoIds.getValues();

  var nuevasPorId = {};
  nuevas.forEach(function (n) { nuevasPorId[n.id] = n.fila; });

  var filas = [];
  for (var i = 0; i < valores.length; i++) {
    var id = String(ids[i][0] || '');
    var vacia = valores[i].every(function (v) { return v === '' || v == null; });
    if (id && idsABorrar[id]) continue;
    if (id && nuevasPorId[id]) {
      filas.push({ id: id, fila: nuevasPorId[id] });
      delete nuevasPorId[id];
      continue;
    }
    if (vacia && !id) continue;
    filas.push({ id: id, fila: valores[i] });
  }
  Object.keys(nuevasPorId).forEach(function (id) { filas.push({ id: id, fila: nuevasPorId[id] }); });

  guardarBloqueOrdenado_(hoja, filaEncabezado, filas);
}

// Escribe `filas` ({id, fila}) en el bloque, ordenadas por fecha y sin
// huecos, y borra lo que sobre.
function guardarBloqueOrdenado_(hoja, filaEncabezado, filas) {
  var primera = filaEncabezado + 1;
  var rango = hoja.getRange(primera, 1, ALBARANES_FILAS_DATOS_, ALBARANES_ANCHO_);
  var rangoIds = hoja.getRange(primera, ALBARANES_COL_ID_, ALBARANES_FILAS_DATOS_, 1);
  if (filas.length > ALBARANES_FILAS_DATOS_) {
    throw new Error('El bloque de la pestaña "' + hoja.getName() + '" (fila ' + primera + ') no tiene más lugar (' + ALBARANES_FILAS_DATOS_ + ' filas).');
  }

  // Orden por fecha (las filas sin fecha quedan al final, en su orden).
  var mapa = mapaColumnasAlbaranes_(hoja.getRange(filaEncabezado, 1, 1, ALBARANES_ANCHO_).getValues()[0]);
  var colFecha = mapa.fecha != null ? mapa.fecha : 0;
  filas.forEach(function (r, i) { r.orden = i; });
  filas.sort(function (a, b) {
    var fa = a.fila[colFecha] instanceof Date ? a.fila[colFecha].getTime() : Infinity;
    var fb = b.fila[colFecha] instanceof Date ? b.fila[colFecha].getTime() : Infinity;
    return fa !== fb ? fa - fb : a.orden - b.orden;
  });

  var salida = [], salidaIds = [];
  for (var k = 0; k < ALBARANES_FILAS_DATOS_; k++) {
    var f = k < filas.length ? filas[k].fila.slice(0, ALBARANES_ANCHO_) : [];
    while (f.length < ALBARANES_ANCHO_) f.push('');
    salida.push(f);
    salidaIds.push([k < filas.length ? filas[k].id : '']);
  }
  rango.setValues(salida);
  rangoIds.setValues(salidaIds);
  if (mapa.fecha != null) hoja.getRange(primera, mapa.fecha + 1, ALBARANES_FILAS_DATOS_, 1).setNumberFormat('dd/MM/yyyy');
  if (mapa.diaPago != null) hoja.getRange(primera, mapa.diaPago + 1, ALBARANES_FILAS_DATOS_, 1).setNumberFormat('dd/MM/yyyy');
  if (!hoja.isColumnHiddenByUser(ALBARANES_COL_ID_)) hoja.hideColumns(ALBARANES_COL_ID_);
}

// Fecha de la factura de un gasto (la que se elige en el formulario, que
// puede no ser la de la caja). Si no tiene una válida, la de la caja.
function fechaFacturaGasto_(m, fechaCajaISO) {
  var f = String((m && m.fecha) || '');
  return /^\d{4}-\d{2}-\d{2}$/.test(f) ? f : fechaCajaISO;
}

function textoFechaIds_(v) {
  return v instanceof Date
    ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd')
    : String(v || '');
}

// Deja el archivo de Albaranes igual a lo que tiene la app para la caja de
// `fechaISO`: agrega los gastos nuevos, actualiza los que cambiaron, y borra
// los que ya no están (o que cambiaron de proveedor o de fecha). Cada gasto
// se escribe con la fecha de SU factura (en el bloque de ese mes y en la
// planilla de ese año), aunque se haya cargado en la caja de otro día.
// `gastos` vacío = borrar todo lo de esa caja. Si `soloSiExiste`, no crea
// planillas que no estén en el Índice.
function sincronizarAlbaranesDelDia_(fechaISO, gastos, soloSiExiste) {
  var anioCaja = periodoAnual_(fechaISO);
  var porAnio = {};
  porAnio[anioCaja] = [];
  // Lo cargado en una caja de principios de año puede haber quedado (en un
  // guardado anterior) en la planilla del año anterior, si la factura era
  // de ese año. Más adelante en el año no se mira, para no abrir otra
  // planilla en cada guardado.
  if (parseInt(String(fechaISO).split('-')[1], 10) <= 3) porAnio[String(Number(anioCaja) - 1)] = [];
  gastos.forEach(function (m) {
    var anio = periodoAnual_(fechaFacturaGasto_(m, fechaISO));
    (porAnio[anio] = porAnio[anio] || []).push(m);
  });

  var escritos = 0, errores = [], nadas = [];
  Object.keys(porAnio).sort().forEach(function (anio) {
    var delAnio = porAnio[anio];
    var r = sincronizarAlbaranesDelDiaEnAnio_(anio, fechaISO, delAnio, soloSiExiste || !delAnio.length);
    if (r.nada) { if (delAnio.length) nadas.push(r.nada); return; }
    escritos += r.escritos || 0;
    if (r.error) errores.push(r.error);
  });

  var salida = { ok: errores.length === 0, escritos: escritos };
  if (errores.length) salida.error = errores.join(' ');
  if (!escritos && !errores.length && nadas.length) salida.nada = nadas.join(' ');
  return salida;
}

// Lo mismo, dentro de la planilla de Albaranes de un año.
function sincronizarAlbaranesDelDiaEnAnio_(anio, fechaISO, gastos, soloSiExiste) {
  var idPlanilla = soloSiExiste
    ? buscarEnIndice_('ALBARANES', anio)
    : obtenerOCrearPlanilla_('ALBARANES', anio);
  if (!idPlanilla) return { ok: true, nada: 'No hay planilla de Albaranes ' + anio + ' en el Índice.' };

  var ss = SpreadsheetApp.openById(idPlanilla);
  var hojaIds = getOrCrearHojaIdsAlbaranes_(ss);
  var valoresIds = hojaIds.getDataRange().getValues();

  // "IDs app": ID Movimiento · Fecha (de la caja) · Pestaña · Fecha factura.
  // Las filas de antes no tienen fecha de factura: eran la de la caja.
  function registroIds(fila) {
    var fechaCaja = textoFechaIds_(fila[1]);
    return { fechaCaja: fechaCaja, pestana: String(fila[2] || ''), fechaFactura: textoFechaIds_(fila[3]) || fechaCaja };
  }
  function clave(r) { return r.pestana + '|' + r.fechaFactura.split('-')[1]; }

  // Lo que ya estaba escrito de esta caja: id -> {pestana, fechaFactura}.
  var previos = {};
  for (var i = 1; i < valoresIds.length; i++) {
    var r = registroIds(valoresIds[i]);
    if (r.fechaCaja === fechaISO) previos[String(valoresIds[i][0])] = r;
  }

  // Cambios agrupados por pestaña y mes: { 'NOMBRE|MM': { pestana, mes, borrar: {id:true}, nuevas: [] } }
  var porBloque = {};
  function cambiosDe(r) {
    var k = clave(r);
    if (!porBloque[k]) porBloque[k] = { pestana: r.pestana, fechaFactura: r.fechaFactura, borrar: {}, nuevas: [] };
    return porBloque[k];
  }

  var avisos = [];
  var actuales = {}; // id -> {fechaCaja, pestana, fechaFactura}
  gastos.forEach(function (m) {
    var hoja = pestanaAlbaranesParaProveedor_(ss, m.proveedor);
    if (!hoja) { avisos.push('No hay pestaña para el proveedor "' + m.proveedor + '".'); return; }
    var r = { fechaCaja: fechaISO, pestana: hoja.getName(), fechaFactura: fechaFacturaGasto_(m, fechaISO) };
    actuales[m.id] = r;
    cambiosDe(r).nuevas.push(m);
  });
  Object.keys(previos).forEach(function (id) {
    if (!actuales[id] || clave(actuales[id]) !== clave(previos[id])) cambiosDe(previos[id]).borrar[id] = true;
  });

  Object.keys(porBloque).forEach(function (k) {
    var cambios = porBloque[k];
    var hoja = ss.getSheetByName(cambios.pestana);
    if (!hoja) return;
    var mesIndex = parseInt(cambios.fechaFactura.split('-')[1], 10) - 1;
    var encabezados = filasEncabezadoAlbaranes_(hoja);
    var filaEnc = encabezados[mesIndex];
    var error = null;
    if (!filaEnc) {
      error = 'La pestaña "' + cambios.pestana + '" no tiene el bloque del mes ' + (mesIndex + 1) + '.';
    } else {
      var mapa = mapaColumnasAlbaranes_(hoja.getRange(filaEnc, 1, 1, ALBARANES_ANCHO_).getValues()[0]);
      var nuevas = cambios.nuevas.map(function (m) { return { id: m.id, fila: filaAlbaran_(m, fechaFacturaGasto_(m, fechaISO), mapa, fechaISO) }; });
      try {
        reescribirBloqueAlbaranes_(hoja, filaEnc, cambios.borrar, nuevas);
      } catch (err) {
        error = String(err.message || err);
      }
    }
    if (error) {
      avisos.push(error);
      // No se pudo escribir este bloque: "IDs app" tiene que seguir
      // reflejando lo que de verdad quedó en él (lo de antes).
      cambios.nuevas.forEach(function (m) {
        if (!previos[m.id] || clave(previos[m.id]) !== k) delete actuales[m.id];
        else actuales[m.id] = previos[m.id];
      });
      Object.keys(cambios.borrar).forEach(function (id) { if (!actuales[id]) actuales[id] = previos[id]; });
    }
  });

  // Reescribe "IDs app": saca las filas de esta caja y pone las actuales.
  var resto = [['ID Movimiento', 'Fecha', 'Pestaña', 'Fecha factura']];
  for (var j = 1; j < valoresIds.length; j++) {
    var rj = registroIds(valoresIds[j]);
    if (rj.fechaCaja !== fechaISO) resto.push([valoresIds[j][0], rj.fechaCaja, rj.pestana, rj.fechaFactura]);
  }
  Object.keys(actuales).forEach(function (id) {
    var ra = actuales[id];
    resto.push([id, ra.fechaCaja, ra.pestana, ra.fechaFactura]);
  });
  hojaIds.clearContents();
  hojaIds.getRange(1, 1, resto.length, 4).setNumberFormat('@').setValues(resto);

  var salida = { ok: avisos.length === 0, escritos: Object.keys(actuales).length };
  if (avisos.length) salida.error = avisos.join(' ');
  return salida;
}

// Lo que corre en cada guardado de la app (ver doPost).
function escribirAlbaranes_(data) {
  try {
    return sincronizarAlbaranesDelDia_(data.fecha, gastosDelDia_(data), false);
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

// ============================================================================
// Una sola vez: restaurar los datos viejos de 2026 (menú del Índice).
// ============================================================================
// Los datos de enero a septiembre que traían las fórmulas IMPORTRANGE de la
// planilla de caja vieja están guardados como valores fijos en
// AlbaranesDatos2026.gs (sacados de la copia ALBARANES 2026.xlsx). Esto los
// escribe en cada bloque de mes, junto con lo que ya haya cargado la app
// (filas con ID en la columna Z), todo ordenado por fecha.
//
// No pisa nada: un bloque que ya tiene filas sin ID (cargadas a mano o de
// una restauración anterior) o que tiene fórmulas se saltea y se avisa en
// el resumen. Y no se puede correr dos veces.
function restaurarAlbaranesViejos2026() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('ALBARANES_2026_RESTAURADO')) {
    throw new Error('Los datos viejos de Albaranes 2026 ya se restauraron el ' + props.getProperty('ALBARANES_2026_RESTAURADO') + ' — no se vuelve a hacer para no duplicar.');
  }
  var id = buscarEnIndice_('ALBARANES', '2026');
  if (!id) throw new Error('No hay fila ALBARANES 2026 en la pestaña "Archivos" del Índice.');
  var ss = SpreadsheetApp.openById(id);
  var resumen = [];

  Object.keys(ALBARANES_DATOS_VIEJOS_2026_).forEach(function (nombre) {
    var bloquesDatos = ALBARANES_DATOS_VIEJOS_2026_[nombre];
    var total = bloquesDatos.reduce(function (s, b) { return s + b.length; }, 0);
    if (!total) return;
    var hoja = ss.getSheetByName(nombre);
    if (!hoja) { resumen.push(nombre + ': NO EXISTE la pestaña — no se restauró'); return; }
    var encabezados = filasEncabezadoAlbaranes_(hoja);
    if (encabezados.length !== 12) { resumen.push(nombre + ': tiene ' + encabezados.length + ' bloques "FECHA" (se esperaban 12) — no se restauró'); return; }

    var escritas = 0, salteados = [];
    bloquesDatos.forEach(function (filasViejas, b) {
      if (!filasViejas.length) return;
      var filaEnc = encabezados[b];
      var primera = filaEnc + 1;
      var rango = hoja.getRange(primera, 1, ALBARANES_FILAS_DATOS_, ALBARANES_ANCHO_);
      var tieneFormulas = rango.getFormulas().some(function (f) { return f.some(String); });
      var valores = rango.getValues();
      var ids = hoja.getRange(primera, ALBARANES_COL_ID_, ALBARANES_FILAS_DATOS_, 1).getValues();

      var actuales = [], hayManuales = false;
      for (var i = 0; i < valores.length; i++) {
        var vacia = valores[i].every(function (v) { return v === '' || v == null; });
        if (vacia) continue;
        if (!ids[i][0]) hayManuales = true;
        actuales.push({ id: String(ids[i][0] || ''), fila: valores[i] });
      }
      if (tieneFormulas || hayManuales) { salteados.push(b + 1); return; }

      var viejas = filasViejas.map(function (fila) {
        return {
          id: '',
          fila: fila.map(function (v) {
            return (typeof v === 'string' && v.indexOf('D:') === 0) ? fechaComoDate_(v.substring(2)) : v;
          })
        };
      });
      guardarBloqueOrdenado_(hoja, filaEnc, viejas.concat(actuales));
      escritas += viejas.length;
    });
    resumen.push(nombre + ': ' + escritas + ' filas restauradas' + (salteados.length ? ' — meses salteados porque ya tenían datos: ' + salteados.join(', ') : ''));
  });

  props.setProperty('ALBARANES_2026_RESTAURADO', hoyISO_());
  Logger.log(resumen.join('\n'));
  return resumen;
}

function restaurarAlbaranesViejos2026DesdeMenu() {
  var ui = SpreadsheetApp.getUi();
  var ok = ui.alert(
    'Albaranes: restaurar datos viejos 2026',
    'Esto vuelve a escribir en ALBARANES 2026 los datos de enero a septiembre (los que se borraron), sin tocar lo que haya cargado la app. Se hace una sola vez. ¿Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (ok !== ui.Button.YES) return;
  ui.alert('Listo', restaurarAlbaranesViejos2026().join('\n'), ui.ButtonSet.OK);
}

// Vuelve a escribir en Albaranes todos los días de un mes de Cierre de
// Caja (desde el menú del Índice) — por ejemplo, para cargar lo que se
// guardó en la app antes de que existiera esta función.
function reenviarMesAAlbaranesDesdeMenu() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt('Albaranes: reenviar un mes', 'Mes a reenviar (formato AAAA-MM, ej. 2026-09):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var periodo = String(resp.getResponseText() || '').trim();
  if (!/^\d{4}-\d{2}$/.test(periodo)) { ui.alert('Formato inválido. Tiene que ser AAAA-MM, ej. 2026-09.'); return; }
  ui.alert('Listo', reenviarMesAAlbaranes(periodo).join('\n'), ui.ButtonSet.OK);
}

function reenviarMesAAlbaranes(periodo) {
  var resumen = [];
  planillasDelTipo_('CIERRE_CAJA').forEach(function (p) {
    var ss = SpreadsheetApp.openById(p.id);
    var registro = ss.getSheetByName('Registro');
    if (!registro) return;
    var valores = registro.getDataRange().getValues();
    var idxFecha = valores[0].indexOf('Fecha');
    if (idxFecha === -1) return;
    var fechas = {};
    for (var i = 1; i < valores.length; i++) {
      var raw = valores[i][idxFecha];
      var f = raw instanceof Date ? Utilities.formatDate(raw, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(raw || '');
      if (f.indexOf(periodo + '-') === 0) fechas[f] = true;
    }
    Object.keys(fechas).sort().forEach(function (fecha) {
      var dia = obtenerDiaJSON_(ss, fecha);
      var r = sincronizarAlbaranesDelDia_(fecha, gastosDelDia_(dia), false);
      resumen.push(fecha + ': ' + (r.ok ? r.escritos + ' gastos' : 'ERROR — ' + r.error));
    });
  });
  return resumen.length ? resumen : ['No hay días de ' + periodo + ' en las planillas de Cierre de Caja.'];
}

// ============================================================================
// Una vez: pasar Albaranes al formato con FORMA DE PAGO y DIA DE PAGO (menú).
// ============================================================================
// En cada pestaña de proveedor (menos EXTRAS), en los 12 meses, reordena
// las columnas al formato nuevo moviendo los datos que ya había (nada se
// pierde: en VINO lo de INFO pasa a OBERVACIONES; SUPER queda como VARIOS,
// con el nombre del súper en PROVEEDOR). Arregla la fila de total de cada
// mes (IVA e IMPORTE), la pestaña TOTALES (que sumaba la columna C, donde
// antes estaba el IMPORTE) y crea las pestañas de proveedores nuevos. Al
// final vuelve a escribir lo que cargó la app, para que se llenen FORMA DE
// PAGO y DIA DE PAGO. Un bloque ya pasado no se vuelve a tocar.
var ALBARANES_ENC_NORMAL_ = ['FECHA', 'Nº FACTURA', 'IVA', 'IMPORTE', 'OBERVACIONES', 'FORMA DE PAGO', 'DIA DE PAGO', 'VARIOS'];
var ALBARANES_ENC_CON_PROVEEDOR_ = ['FECHA', 'PROVEEDOR', 'Nº FACTURA', 'IVA', 'IMPORTE', 'OBERVACIONES', 'FORMA DE PAGO', 'DIA DE PAGO', 'VARIOS'];
var ALBARANES_PESTANAS_CON_PROVEEDOR_ = ['VARIOS', 'SUPER'];
var ALBARANES_PESTANAS_SIN_CAMBIO_ = ['EXTRAS'];
var ALBARANES_PROVEEDORES_NUEVOS_ = ['LOS FUENTEÑOS', 'WINEUP'];

function letraColumna_(n) { // 1 -> A
  var s = '';
  while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function encabezadosAlbaranesPara_(nombrePestana) {
  return ALBARANES_PESTANAS_CON_PROVEEDOR_.indexOf(normalizarClave_(nombrePestana)) > -1
    ? ALBARANES_ENC_CON_PROVEEDOR_ : ALBARANES_ENC_NORMAL_;
}

// Pasa un bloque de mes al formato nuevo. Devuelve false si ya estaba.
function pasarBloqueAlbaranesAFormatoNuevo_(hoja, filaEnc) {
  var viejoEnc = hoja.getRange(filaEnc, 1, 1, ALBARANES_ANCHO_).getValues()[0];
  var viejo = mapaColumnasAlbaranes_(viejoEnc);
  if (viejo.formaPago != null) return false;
  var enc = encabezadosAlbaranesPara_(hoja.getName());
  var nuevo = mapaColumnasAlbaranes_(enc);

  var primera = filaEnc + 1;
  var rango = hoja.getRange(primera, 1, ALBARANES_FILAS_DATOS_, ALBARANES_ANCHO_);
  if (rango.getFormulas().some(function (f) { return f.some(String); })) {
    throw new Error('el bloque de la fila ' + filaEnc + ' tiene fórmulas');
  }
  var datos = rango.getValues().map(function (v) {
    var vacia = v.every(function (x) { return x === '' || x == null; });
    if (vacia) return filaVaciaAlbaranes_();
    function de(k) { return viejo[k] != null ? v[viejo[k]] : ''; }
    var fila = filaVaciaAlbaranes_();
    fila[nuevo.fecha] = de('fecha');
    fila[nuevo.factura] = de('factura');
    fila[nuevo.iva] = de('iva');
    fila[nuevo.importe] = de('importe');
    fila[nuevo.obs] = [de('info'), de('obs')].filter(function (x) { return x !== '' && x != null; }).join(' · ');
    if (nuevo.detalle != null) fila[nuevo.detalle] = de('detalle');
    else if (de('detalle') !== '') fila[nuevo.obs] = [de('detalle'), fila[nuevo.obs]].filter(String).join(' · ');
    return fila;
  });

  var encFila = enc.slice();
  while (encFila.length < ALBARANES_ANCHO_) encFila.push('');
  hoja.getRange(filaEnc, 1, 1, ALBARANES_ANCHO_).setValues([encFila]);
  rango.setValues(datos);
  rango.setNumberFormat('General');
  [nuevo.fecha, nuevo.diaPago].forEach(function (c) {
    hoja.getRange(primera, c + 1, ALBARANES_FILAS_DATOS_, 1).setNumberFormat('dd/MM/yyyy');
  });
  [nuevo.iva, nuevo.importe].forEach(function (c) {
    hoja.getRange(primera, c + 1, ALBARANES_FILAS_DATOS_ + 1, 1).setNumberFormat('#,##0.00');
  });

  // Fila de total del mes: suma de IVA y de IMPORTE en sus columnas nuevas.
  var filaTotal = primera + ALBARANES_FILAS_DATOS_;
  var total = hoja.getRange(filaTotal, 2, 1, ALBARANES_ANCHO_ - 1);
  total.clearContent();
  [nuevo.iva, nuevo.importe].forEach(function (c) {
    var L = letraColumna_(c + 1);
    hoja.getRange(filaTotal, c + 1).setFormula('=SUM(' + L + primera + ':' + L + (filaTotal - 1) + ')');
  });
  return true;
}

// Crea la pestaña de un proveedor nuevo copiando una ya pasada al formato
// nuevo, vacía (sin datos ni IDs).
function crearPestanaProveedorAlbaranes_(ss, nombre, modelo) {
  if (pestanaAlbaranesParaProveedor_(ss, nombre)) return false;
  var hoja = modelo.copyTo(ss);
  hoja.setName(nombre);
  ss.setActiveSheet(hoja);
  ss.moveActiveSheet(modelo.getIndex() + 1);
  hoja.getRange('A1').setValue(nombre);
  filasEncabezadoAlbaranes_(hoja).forEach(function (filaEnc) {
    hoja.getRange(filaEnc + 1, 1, ALBARANES_FILAS_DATOS_, ALBARANES_ANCHO_).clearContent();
    hoja.getRange(filaEnc + 1, ALBARANES_COL_ID_, ALBARANES_FILAS_DATOS_, 1).clearContent();
  });
  return true;
}

// TOTALES: cada fila toma, de una pestaña, el total del IMPORTE de cada mes
// (fila 46, 91, 136... de esa pestaña). La pestaña se lee de la fórmula de
// la columna C (='ATLANTA'!C46); la columna del IMPORTE, de sus encabezados.
function arreglarTotalesAlbaranes_(ss, nuevas) {
  var hoja = ss.getSheetByName('TOTALES');
  if (!hoja) return ['No hay pestaña TOTALES.'];
  var resumen = [];
  var ultima = Math.max(hoja.getLastRow(), 3);
  var formulasC = hoja.getRange(1, 3, ultima, 1).getFormulas();
  var valoresB = hoja.getRange(1, 2, ultima, 1).getValues();

  function escribirFila(r, nombre) {
    var pestana = ss.getSheetByName(nombre);
    if (!pestana) return false;
    var filaEnc = filasEncabezadoAlbaranes_(pestana)[0];
    if (!filaEnc) return false;
    var mapa = mapaColumnasAlbaranes_(pestana.getRange(filaEnc, 1, 1, ALBARANES_ANCHO_).getValues()[0]);
    if (mapa.importe == null) return false;
    var L = letraColumna_(mapa.importe + 1);
    var ref = "'" + nombre.replace(/'/g, "''") + "'";
    var filaTotal = filaEnc + 1 + ALBARANES_FILAS_DATOS_;
    var paso = 45; // distancia entre bloques de mes
    var fila = [];
    for (var mes = 0; mes < 12; mes++) fila.push('=' + ref + '!' + L + (filaTotal + paso * mes));
    hoja.getRange(r, 2).setValue(nombre);
    hoja.getRange(r, 3, 1, 12).setFormulas([fila]);
    return true;
  }

  var yaEstan = {};
  for (var r = 2; r <= ultima; r++) {
    var m = /^='?(.+?)'?!\$?[A-Z]+\$?46$/.exec(formulasC[r - 1][0] || '');
    if (!m) continue;
    var nombre = m[1].replace(/''/g, "'");
    yaEstan[normalizarClave_(nombre)] = true;
    if (escribirFila(r, nombre)) resumen.push('TOTALES fila ' + r + ': ' + nombre);
  }
  // Proveedores nuevos: en la primera fila libre antes de la fila TOTAL.
  nuevas.forEach(function (nombre) {
    if (yaEstan[normalizarClave_(nombre)]) return;
    for (var r2 = 3; r2 <= ultima; r2++) {
      var b = String(valoresB[r2 - 1][0] || '');
      if (normalizarClave_(b) === 'TOTAL') break;
      if (!b && !formulasC[r2 - 1][0]) {
        if (escribirFila(r2, nombre)) {
          resumen.push('TOTALES fila ' + r2 + ': ' + nombre + ' (nueva)');
          valoresB[r2 - 1][0] = nombre;
          formulasC[r2 - 1][0] = '=x';
        }
        return;
      }
    }
    resumen.push('TOTALES: no hubo fila libre para ' + nombre + ' — agregala a mano.');
  });
  return resumen;
}

function albaranesPasarAFormatoNuevo() {
  var resumen = [];
  planillasDelTipo_('ALBARANES').forEach(function (p) {
    var ss = SpreadsheetApp.openById(p.id);
    var modelo = null;
    ss.getSheets().forEach(function (hoja) {
      var nombre = hoja.getName();
      if (ALBARANES_PESTANAS_EXCLUIDAS_.indexOf(nombre) > -1) return;
      if (ALBARANES_PESTANAS_SIN_CAMBIO_.indexOf(normalizarClave_(nombre)) > -1) return;
      var encabezados = filasEncabezadoAlbaranes_(hoja);
      if (!encabezados.length) return;
      var pasados = 0, errores = [];
      encabezados.forEach(function (filaEnc) {
        try { if (pasarBloqueAlbaranesAFormatoNuevo_(hoja, filaEnc)) pasados++; }
        catch (err) { errores.push(String(err.message || err)); }
      });
      if (pasados) resumen.push(p.nombre + ' / ' + nombre + ': ' + pasados + ' meses pasados al formato nuevo');
      if (errores.length) resumen.push(p.nombre + ' / ' + nombre + ': NO se pasó — ' + errores.join('; '));
      if (!modelo && ALBARANES_PESTANAS_CON_PROVEEDOR_.indexOf(normalizarClave_(nombre)) === -1 && !errores.length) modelo = hoja;
    });

    var creadas = [];
    if (modelo) {
      ALBARANES_PROVEEDORES_NUEVOS_.forEach(function (nombre) {
        if (crearPestanaProveedorAlbaranes_(ss, nombre, modelo)) creadas.push(nombre);
      });
    }
    if (creadas.length) resumen.push(p.nombre + ': pestañas nuevas ' + creadas.join(', '));
    resumen = resumen.concat(arreglarTotalesAlbaranes_(ss, creadas).map(function (x) { return p.nombre + ' / ' + x; }));

    // Lo que cargó la app, de nuevo (llena FORMA DE PAGO y DIA DE PAGO).
    var hojaIds = ss.getSheetByName(ALBARANES_HOJA_IDS_);
    if (hojaIds) {
      var meses = {};
      hojaIds.getDataRange().getValues().slice(1).forEach(function (fila) {
        var f = textoFechaIds_(fila[1]);
        if (/^\d{4}-\d{2}/.test(f)) meses[f.substring(0, 7)] = true;
      });
      Object.keys(meses).sort().forEach(function (periodo) {
        resumen = resumen.concat(reenviarMesAAlbaranes(periodo).filter(function (x) { return x.indexOf('ERROR') > -1; }));
      });
      resumen.push(p.nombre + ': reescrito lo cargado por la app en ' + Object.keys(meses).sort().join(', '));
    }
  });
  return resumen.length ? resumen : ['No había nada para pasar.'];
}

function albaranesPasarAFormatoNuevoDesdeMenu() {
  var ui = SpreadsheetApp.getUi();
  var ok = ui.alert(
    'Albaranes: formato con forma y día de pago',
    'Esto reordena las columnas de todas las pestañas de proveedor de Albaranes (FECHA · Nº FACTURA · IVA · IMPORTE · OBERVACIONES · FORMA DE PAGO · DIA DE PAGO · VARIOS; en VARIOS y SUPER, con PROVEEDOR después de FECHA), moviendo los datos que ya hay, arregla TOTALES y crea LOS FUENTEÑOS y WINEUP. Mejor hacerlo cuando nadie esté cargando en la app. ¿Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (ok !== ui.Button.YES) return;
  ui.alert('Listo', albaranesPasarAFormatoNuevo().join('\n'), ui.ButtonSet.OK);
}
