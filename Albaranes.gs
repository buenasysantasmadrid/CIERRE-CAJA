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
// posición), porque no todas las pestañas son iguales:
//   FECHA · FACTURA Nº · IMPORTE · IVA · OBERVACIONES   (la mayoría)
//   FECHA · INFO · FACTURA Nº · IMPORTE · IVA · OBERVACIONES   (VINO)
//   FECHA · SUPER · IMPORTE · IVA · OBERVACIONES   (SUPER)
//   FECHA · PROVEEDOR · OBERVACIONES · IMPORTE · IVA   (VARIOS)
//   FECHA · NOMBRE · IMPORTE   (EXTRAS)
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
var ALBARANES_ANCHO_ = 7;   // A–G: la parte visible de cada bloque
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

// Qué columna (0-based dentro de A–G) lleva cada dato en este bloque.
function mapaColumnasAlbaranes_(encabezados) {
  var mapa = {};
  encabezados.forEach(function (h, i) {
    var k = normalizarClave_(h).replace(/[^A-Z]/g, '');
    if (k === 'FECHA') mapa.fecha = i;
    else if (k.indexOf('FACTURA') === 0 || k === 'ALBARAN') mapa.factura = i;
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

// Arma la fila (A–G) de un gasto según las columnas de ese bloque.
function filaAlbaran_(m, fechaISO, mapa) {
  var fila = ['', '', '', '', '', '', ''];
  var formaPago = m.subtipo === 'efectivo' ? 'EFECTIVO' : '';
  var info = String(m.info || '');
  // Las formas de pago que no son efectivo ya vienen marcadas en Info
  // (TARJETA, NO PAGADO, TRANSFERENCIA, PAGADO ...); el efectivo no.
  var textoObs = [info, formaPago].filter(String).join(' · ');
  var detalle = m.proveedorDetalle || '';

  if (mapa.fecha != null) fila[mapa.fecha] = fechaComoDate_(fechaISO);
  if (mapa.importe != null) fila[mapa.importe] = Number(m.importe) || 0;
  if (mapa.iva != null && m.iva != null && m.iva !== '') fila[mapa.iva] = Number(m.iva) || 0;

  if (mapa.factura != null) fila[mapa.factura] = m.factura || '';
  else if (m.factura) textoObs = ['Fact. ' + m.factura, textoObs].filter(String).join(' · ');

  if (mapa.info != null) {
    fila[mapa.info] = info;
    textoObs = formaPago;
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
    salida.push(k < filas.length ? filas[k].fila : ['', '', '', '', '', '', '']);
    salidaIds.push([k < filas.length ? filas[k].id : '']);
  }
  rango.setValues(salida);
  rangoIds.setValues(salidaIds);
  if (mapa.fecha != null) hoja.getRange(primera, mapa.fecha + 1, ALBARANES_FILAS_DATOS_, 1).setNumberFormat('dd/MM/yyyy');
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
      var nuevas = cambios.nuevas.map(function (m) { return { id: m.id, fila: filaAlbaran_(m, fechaFacturaGasto_(m, fechaISO), mapa) }; });
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
