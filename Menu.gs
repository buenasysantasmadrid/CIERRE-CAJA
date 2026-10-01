// ============================================================================
// MENÚ — los archivos del negocio que se abren desde las tres rayas de la
// app (al lado del logo). Cada uno es una planilla que tiene que estar en
// la pestaña "Archivos" del Índice (no se crean solas):
//   STOCK_PRODUCCION | 2026 | <ID>   (Producción de empanadas y quiches)
//   CAMBIO_TINTA     | 2026 | <ID>   (no es por año: si no hay fila de ese
//                                     año, se usa la última que haya)
// ============================================================================

// ID de la planilla de `tipo` para el año de `fechaISO`; si no hay fila de
// ese año, la última de ese tipo que haya en el Índice (para los archivos
// que no se parten por año, como Cambio de tinta).
function planillaDelMenu_(tipo, fechaISO, nombre) {
  var id = buscarEnIndice_(tipo, periodoAnual_(fechaISO || hoyISO_()));
  if (!id) {
    var todas = planillasDelTipo_(tipo);
    if (todas.length) id = todas[todas.length - 1].id;
  }
  if (!id) throw new Error('Falta "' + nombre + '" en la pestaña "Archivos" del Índice (tipo ' + tipo + ').');
  return SpreadsheetApp.openById(id);
}

function respuestaJSON_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Lo que pide la app desde el menú (ver doPost). null si no es de acá.
function accionDelMenu_(data) {
  var fn = null;
  if (data.accionProduccionLeer) fn = function () { return leerProduccion_(data.fecha); };
  else if (data.accionProduccionGuardar) fn = function () { return guardarProduccion_(data.fecha, data.valores || {}); };
  else if (data.accionTintaLeer) fn = function () { return leerTinta_(); };
  else if (data.accionTintaGuardar) fn = function () { return guardarTinta_(data.impresora, data.tinta, data.fecha); };
  if (!fn) return null;
  try {
    return respuestaJSON_(fn());
  } catch (err) {
    return respuestaJSON_({ ok: false, error: String(err.message || err) });
  }
}

// ---------- Producción de empanadas y quiches ----------
// Una pestaña por mes (enero, febrero... en minúsculas). Columna A: los
// sabores; fila 6: los días del mes (B = 1 ... AF = 31). Arriba van las
// empanadas (debajo de la fila "EMPANADAS") y abajo las quiches (debajo de
// "QUICHES"). AG y AH son fórmulas de totales: no se tocan. Los sabores se
// leen de la propia pestaña, así que si se agrega uno nuevo, sale solo.

function pestanaProduccion_(ss, fechaISO) {
  var mes = parseInt(String(fechaISO).split('-')[1], 10);
  var buscado = MESES_MAYUS_[mes - 1];
  var hojas = ss.getSheets();
  for (var i = 0; i < hojas.length; i++) {
    if (normalizarClave_(hojas[i].getName()) === buscado) return hojas[i];
  }
  throw new Error('La planilla de Producción no tiene la pestaña de ' + buscado.toLowerCase() + '.');
}

// { empanadas: [{nombre, fila}], quiches: [...], filaDias }
function filasProduccion_(hoja) {
  var alto = Math.min(hoja.getLastRow(), 80);
  var colA = hoja.getRange(1, 1, alto, 2).getValues();
  var grupos = { empanadas: [], quiches: [] };
  var actual = null, filaDias = 6;
  for (var i = 0; i < alto; i++) {
    var texto = normalizarTexto_(colA[i][0]);
    var clave = normalizarClave_(texto);
    if (!texto && Number(colA[i][1]) === 1) filaDias = i + 1;
    if (clave === 'EMPANADAS') { actual = 'empanadas'; continue; }
    if (clave === 'QUICHES' || clave === 'QUICHE') { actual = 'quiches'; continue; }
    if (!actual) continue;
    if (!texto) {
      if (grupos[actual].length) actual = actual === 'empanadas' ? 'esperandoQuiches' : null;
      continue;
    }
    if (actual === 'esperandoQuiches') continue;
    grupos[actual].push({ nombre: texto, fila: i + 1 });
  }
  return { empanadas: grupos.empanadas, quiches: grupos.quiches, filaDias: filaDias };
}

// Columna (1-based) del día en la fila de los días; si no se encuentra,
// la de siempre (B = día 1).
function columnaDiaProduccion_(hoja, filaDias, dia) {
  var fila = hoja.getRange(filaDias, 1, 1, 40).getValues()[0];
  for (var c = 1; c < fila.length; c++) {
    if (Number(fila[c]) === dia) return c + 1;
  }
  return dia + 1;
}

function leerProduccion_(fechaISO) {
  var ss = planillaDelMenu_('STOCK_PRODUCCION', fechaISO, 'Stock producción');
  var hoja = pestanaProduccion_(ss, fechaISO);
  var f = filasProduccion_(hoja);
  var col = columnaDiaProduccion_(hoja, f.filaDias, parseInt(String(fechaISO).split('-')[2], 10));
  var valores = hoja.getRange(1, col, Math.max(hoja.getLastRow(), 1), 1).getValues();
  function conValor(s) { return { nombre: s.nombre, valor: Number(valores[s.fila - 1][0]) || 0 }; }
  return { ok: true, fecha: fechaISO, pestana: hoja.getName(), empanadas: f.empanadas.map(conValor), quiches: f.quiches.map(conValor) };
}

// `valores`: { empanadas: {CARNE: 8, ...}, quiches: {POLLO: 2, ...} }. Un 0
// deja la celda vacía (como se hacía a mano).
function guardarProduccion_(fechaISO, valores) {
  var ss = planillaDelMenu_('STOCK_PRODUCCION', fechaISO, 'Stock producción');
  var hoja = pestanaProduccion_(ss, fechaISO);
  var f = filasProduccion_(hoja);
  var col = columnaDiaProduccion_(hoja, f.filaDias, parseInt(String(fechaISO).split('-')[2], 10));
  ['empanadas', 'quiches'].forEach(function (grupo) {
    var delGrupo = valores[grupo] || {};
    f[grupo].forEach(function (s) {
      if (!(s.nombre in delGrupo)) return;
      var v = Number(delGrupo[s.nombre]) || 0;
      hoja.getRange(s.fila, col).setValue(v > 0 ? v : '');
    });
  });
  return leerProduccion_(fechaISO);
}

// ---------- Cambio de tinta ----------
// Una sola pestaña con un bloque por impresora, uno debajo del otro. Cada
// bloque empieza con su nombre en la columna A ("BYS", "CLARA"), debajo
// una fila de títulos y después los cambios:
//   A fecha del cambio de tinta negra · B "NEGRA" · C días desde el anterior
//   F fecha del cambio de tinta de color · G "COLOR" · H días desde el anterior
var IMPRESORAS_TINTA_ = [
  { clave: 'BYS', nombre: 'Buenas y Santas' },
  { clave: 'CLARA', nombre: 'Clarita' }
];
var COLUMNAS_TINTA_ = {
  negra: { fecha: 1, etiqueta: 'NEGRA' },  // A, B, C
  color: { fecha: 6, etiqueta: 'COLOR' }   // F, G, H
};

function hojaTinta_() {
  return planillaDelMenu_('CAMBIO_TINTA', hoyISO_(), 'Cambio de tinta').getSheets()[0];
}

// Filas de cada bloque: { BYS: {desde, hasta}, CLARA: {...} } (filas de
// datos, 1-based; `hasta` es la última antes del bloque siguiente).
function bloquesTinta_(hoja) {
  var ultima = Math.max(hoja.getLastRow(), 1);
  var colA = hoja.getRange(1, 1, ultima, 1).getValues();
  var inicios = [];
  for (var i = 0; i < colA.length; i++) {
    var k = normalizarClave_(colA[i][0]);
    IMPRESORAS_TINTA_.forEach(function (imp) { if (k === imp.clave) inicios.push({ clave: imp.clave, fila: i + 1 }); });
  }
  var bloques = {};
  inicios.forEach(function (b, j) {
    var siguiente = inicios[j + 1] ? inicios[j + 1].fila : ultima + 50;
    bloques[b.clave] = { desde: b.fila + 2, hasta: siguiente - 1 };
  });
  return bloques;
}

function comoFecha_(v) {
  if (v instanceof Date) return v;
  if (typeof v === 'number' && v > 30000) return new Date(Math.round((v - 25569) * 86400000)); // número de serie de Sheets
  return null;
}

// Última fila con fecha en esa columna dentro del bloque (o desde-1 si no hay).
function ultimaFilaTinta_(hoja, bloque, col) {
  var alto = bloque.hasta - bloque.desde + 1;
  var vals = hoja.getRange(bloque.desde, col, alto, 1).getValues();
  var ultima = bloque.desde - 1;
  for (var i = 0; i < vals.length; i++) if (comoFecha_(vals[i][0])) ultima = bloque.desde + i;
  return ultima;
}

function leerTinta_() {
  var hoja = hojaTinta_();
  var bloques = bloquesTinta_(hoja);
  var hoy = fechaComoDate_(hoyISO_());
  var tz = Session.getScriptTimeZone();
  return {
    ok: true,
    impresoras: IMPRESORAS_TINTA_.filter(function (imp) { return bloques[imp.clave]; }).map(function (imp) {
      var b = bloques[imp.clave];
      var salida = { clave: imp.clave, nombre: imp.nombre };
      Object.keys(COLUMNAS_TINTA_).forEach(function (tinta) {
        var fila = ultimaFilaTinta_(hoja, b, COLUMNAS_TINTA_[tinta].fecha);
        var fecha = fila >= b.desde ? comoFecha_(hoja.getRange(fila, COLUMNAS_TINTA_[tinta].fecha).getValue()) : null;
        salida[tinta] = fecha ? {
          fecha: Utilities.formatDate(fecha, tz, 'yyyy-MM-dd'),
          dias: Math.round((hoy - fechaComoDate_(Utilities.formatDate(fecha, tz, 'yyyy-MM-dd'))) / 86400000)
        } : null;
      });
      return salida;
    })
  };
}

// Apunta un cambio de tinta debajo del último de esa impresora y ese color,
// con los días desde el anterior. Si el bloque está lleno (se llega al de
// la impresora siguiente), se inserta una fila.
function guardarTinta_(clave, tinta, fechaISO) {
  var cols = COLUMNAS_TINTA_[tinta];
  if (!cols) throw new Error('Tinta desconocida: ' + tinta);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(fechaISO || ''))) throw new Error('Fecha inválida');
  var hoja = hojaTinta_();
  var b = bloquesTinta_(hoja)[clave];
  if (!b) throw new Error('No encontré la impresora ' + clave + ' en Cambio de tinta.');

  var ultima = ultimaFilaTinta_(hoja, b, cols.fecha);
  var fila = ultima + 1;
  // Siempre queda al menos una fila en blanco antes de la impresora siguiente.
  if (fila >= b.hasta) hoja.insertRowsBefore(fila, 1);

  var celdaFecha = hoja.getRange(fila, cols.fecha);
  if (ultima >= b.desde) celdaFecha.setNumberFormat(hoja.getRange(ultima, cols.fecha).getNumberFormat());
  celdaFecha.setValue(fechaComoDate_(fechaISO));
  hoja.getRange(fila, cols.fecha + 1).setValue(cols.etiqueta);
  var letra = letraColumna_(cols.fecha);
  if (ultima >= b.desde) hoja.getRange(fila, cols.fecha + 2).setFormula('=' + letra + fila + '-' + letra + ultima);
  return leerTinta_();
}
