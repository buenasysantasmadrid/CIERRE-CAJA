// ============================================================================
// MENÚ — los archivos del negocio que se abren desde las tres rayas de la
// app (al lado del logo). Cada uno es una planilla que tiene que estar en
// la pestaña "Archivos" del Índice (no se crean solas), y compartida como
// editor con la cuenta que publica el Apps Script:
//   STOCK_PRODUCCION          | 2026 | <ID>  (Producción de empanadas y quiches)
//   COMPARATIVA_CARNE         | 2026 | <ID>  (kilos de pollería y carnicería)
//   CAMBIO_TINTA              | 2026 | <ID>  (no es por año: si no hay fila de
//   ARREGLO_ELECTRODOMESTICOS | 2026 | <ID>   ese año, se usa la última)
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
  var diaria = PLANILLAS_DIARIAS_[data.planilla];
  if (data.accionDiariaLeer && diaria) fn = function () { return leerDiaria_(diaria, data.fecha); };
  else if (data.accionDiariaGuardar && diaria) fn = function () { return guardarDiaria_(diaria, data.fecha, data.valores || {}); };
  else if (data.accionTintaLeer) fn = function () { return leerTinta_(); };
  else if (data.accionTintaGuardar) fn = function () { return guardarTinta_(data.impresora, data.tinta, data.fecha); };
  else if (data.accionArreglosLeer) fn = function () { return leerArreglos_(); };
  else if (data.accionArreglosGuardar) fn = function () { return guardarArreglo_(data.arreglo || {}); };
  if (!fn) return null;
  try {
    return respuestaJSON_(fn());
  } catch (err) {
    return respuestaJSON_({ ok: false, error: String(err.message || err) });
  }
}

// ---------- Planillas "por día": Producción y Comparativa carnicería ----------
// Las dos tienen una pestaña por mes y, en cada una, bloques de filas (los
// sabores o los productos, en la columna A) con una columna por día del
// mes (B = 1 ... AF = 31). Cada bloque empieza debajo de una fila título y
// termina en la primera fila vacía después de sus filas. Las columnas de
// totales (AG, AH) son fórmulas: no se tocan. Las filas se leen de la propia
// pestaña, así que si se agrega un sabor o un producto nuevo, sale solo.
//   Producción: pestañas "enero"... ; bloques "EMPANADAS" y "QUICHES".
//   Carnicería: pestañas "ENE"... "SEPT"...; dos bloques que empiezan con
//   "GÉNERO": el primero es la pollería y el segundo la carnicería.
var PLANILLAS_DIARIAS_ = {
  produccion: {
    tipo: 'STOCK_PRODUCCION', nombre: 'Stock producción',
    titulos: { EMPANADAS: 'empanadas', QUICHES: 'quiches', QUICHE: 'quiches' }
  },
  carniceria: {
    tipo: 'COMPARATIVA_CARNE', nombre: 'Comparativa carnicería',
    titulos: { GENERO: ['polleria', 'carniceria'] } // el mismo título dos veces, en orden
  }
};

// La pestaña del mes: con el nombre entero ("octubre") o abreviado ("OCT",
// "SEPT").
function pestanaDelMes_(ss, fechaISO, nombre) {
  var mes = MESES_MAYUS_[parseInt(String(fechaISO).split('-')[1], 10) - 1];
  var hojas = ss.getSheets();
  for (var i = 0; i < hojas.length; i++) if (normalizarClave_(hojas[i].getName()) === mes) return hojas[i];
  for (var j = 0; j < hojas.length; j++) {
    var n = normalizarClave_(hojas[j].getName());
    if (n.length >= 3 && n.length <= 4 && mes.indexOf(n) === 0) return hojas[j];
  }
  throw new Error('La planilla de ' + nombre + ' no tiene la pestaña de ' + mes.toLowerCase() + '.');
}

// { grupos: { empanadas: [{nombre, fila}], ... }, orden: ['empanadas', ...] }
function filasDiarias_(hoja, cfg) {
  var alto = Math.min(hoja.getLastRow(), 120);
  var colA = hoja.getRange(1, 1, Math.max(alto, 1), 1).getValues();
  var grupos = {}, orden = [], vecesTitulo = {};
  var actual = null;
  for (var i = 0; i < alto; i++) {
    var texto = normalizarTexto_(colA[i][0]);
    var titulo = cfg.titulos[normalizarClave_(texto)];
    if (titulo) {
      var n = vecesTitulo[texto] = (vecesTitulo[texto] || 0) + 1;
      actual = Array.isArray(titulo) ? titulo[n - 1] : titulo;
      if (actual && !grupos[actual]) { grupos[actual] = []; orden.push(actual); }
      continue;
    }
    if (!actual) continue;
    if (!texto) {
      if (grupos[actual].length) actual = null; // fin del bloque
      continue;
    }
    grupos[actual].push({ nombre: texto, fila: i + 1 });
  }
  return { grupos: grupos, orden: orden };
}

// Columna (1-based) del día: se busca la fila con los días (la que tiene
// 1 y 2 en B y C) y en ella el número; si no, la de siempre (B = día 1).
function columnaDelDia_(hoja, dia) {
  var alto = Math.min(hoja.getLastRow(), 20);
  var filas = hoja.getRange(1, 1, Math.max(alto, 1), 40).getValues();
  for (var r = 0; r < filas.length; r++) {
    if (Number(filas[r][1]) === 1 && Number(filas[r][2]) === 2) {
      for (var c = 1; c < filas[r].length; c++) if (Number(filas[r][c]) === dia) return c + 1;
    }
  }
  return dia + 1;
}

function leerDiaria_(cfg, fechaISO) {
  var hoja = pestanaDelMes_(planillaDelMenu_(cfg.tipo, fechaISO, cfg.nombre), fechaISO, cfg.nombre);
  var f = filasDiarias_(hoja, cfg);
  var col = columnaDelDia_(hoja, parseInt(String(fechaISO).split('-')[2], 10));
  var valores = hoja.getRange(1, col, Math.max(hoja.getLastRow(), 1), 1).getValues();
  var grupos = {};
  f.orden.forEach(function (g) {
    grupos[g] = f.grupos[g].map(function (s) { return { nombre: s.nombre, valor: Number(valores[s.fila - 1][0]) || 0 }; });
  });
  return { ok: true, fecha: fechaISO, pestana: hoja.getName(), orden: f.orden, grupos: grupos };
}

// `valores`: { empanadas: {CARNE: 8, ...}, quiches: {...} }. Un 0 deja la
// celda vacía (como se hacía a mano).
function guardarDiaria_(cfg, fechaISO, valores) {
  var hoja = pestanaDelMes_(planillaDelMenu_(cfg.tipo, fechaISO, cfg.nombre), fechaISO, cfg.nombre);
  var f = filasDiarias_(hoja, cfg);
  var col = columnaDelDia_(hoja, parseInt(String(fechaISO).split('-')[2], 10));
  f.orden.forEach(function (g) {
    var delGrupo = valores[g] || {};
    f.grupos[g].forEach(function (s) {
      if (!(s.nombre in delGrupo)) return;
      var v = Math.round((Number(delGrupo[s.nombre]) || 0) * 1000) / 1000;
      hoja.getRange(s.fila, col).setValue(v > 0 ? v : '');
    });
  });
  return leerDiaria_(cfg, fechaISO);
}

// ---------- Arreglo electrodomésticos ----------
// Pestaña "arreglos": una fila de títulos (FECHA, ELECTRODOMESTICO,
// PROVEEDOR, PROBLEMA, ARREGLO, €) y debajo un arreglo por fila. Las
// columnas se buscan por su título. Los nuevos van debajo del último.
var CAMPOS_ARREGLO_ = {
  FECHA: 'fecha', ELECTRODOMESTICO: 'electrodomestico', PROVEEDOR: 'proveedor',
  PROBLEMA: 'problema', ARREGLO: 'arreglo', '€': 'importe', IMPORTE: 'importe'
};

function hojaArreglos_() {
  var ss = planillaDelMenu_('ARREGLO_ELECTRODOMESTICOS', hoyISO_(), 'Arreglo electrodomésticos');
  var hojas = ss.getSheets();
  var hoja = null;
  for (var i = 0; i < hojas.length; i++) if (normalizarClave_(hojas[i].getName()) === 'ARREGLOS') hoja = hojas[i];
  hoja = hoja || hojas.filter(function (h) { return h.getLastRow() > 1; })[0] || hojas[0];
  var alto = Math.min(Math.max(hoja.getLastRow(), 1), 10);
  var arriba = hoja.getRange(1, 1, alto, 15).getValues();
  for (var r = 0; r < arriba.length; r++) {
    var cols = {};
    arriba[r].forEach(function (v, c) { var k = CAMPOS_ARREGLO_[normalizarClave_(v)]; if (k) cols[k] = c + 1; });
    if (cols.fecha && cols.electrodomestico) return { hoja: hoja, filaTitulos: r + 1, cols: cols };
  }
  throw new Error('No encontré la fila de títulos (FECHA, ELECTRODOMESTICO...) en Arreglo electrodomésticos.');
}

function leerArreglos_() {
  var a = hojaArreglos_();
  var ultima = a.hoja.getLastRow();
  var lista = [];
  if (ultima > a.filaTitulos) {
    var vals = a.hoja.getRange(a.filaTitulos + 1, 1, ultima - a.filaTitulos, 15).getValues();
    var tz = Session.getScriptTimeZone();
    vals.forEach(function (fila, i) {
      var item = { fila: a.filaTitulos + 1 + i };
      var algo = false;
      Object.keys(a.cols).forEach(function (k) {
        var v = fila[a.cols[k] - 1];
        if (v instanceof Date) v = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
        item[k] = v === '' || v == null ? '' : v;
        if (item[k] !== '') algo = true;
      });
      if (algo) lista.push(item);
    });
  }
  return { ok: true, arreglos: lista.reverse() };
}

function guardarArreglo_(arr) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(arr.fecha || ''))) throw new Error('Pon la fecha del arreglo');
  if (!String(arr.electrodomestico || '').trim()) throw new Error('Pon qué electrodoméstico se arregló');
  var a = hojaArreglos_();
  var ultima = a.filaTitulos;
  var lr = a.hoja.getLastRow();
  if (lr > a.filaTitulos) {
    var vals = a.hoja.getRange(a.filaTitulos + 1, 1, lr - a.filaTitulos, 15).getValues();
    vals.forEach(function (fila, i) {
      if (Object.keys(a.cols).some(function (k) { return fila[a.cols[k] - 1] !== ''; })) ultima = a.filaTitulos + 1 + i;
    });
  }
  var fila = ultima + 1;
  Object.keys(a.cols).forEach(function (k) {
    var celda = a.hoja.getRange(fila, a.cols[k]);
    if (ultima > a.filaTitulos) celda.setNumberFormat(a.hoja.getRange(ultima, a.cols[k]).getNumberFormat());
    var v = arr[k];
    if (k === 'fecha') v = fechaComoDate_(arr.fecha);
    else if (k === 'importe') v = (v === '' || v == null) ? '' : Number(v) || 0;
    else v = String(v || '').trim();
    celda.setValue(v);
  });
  return leerArreglos_();
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
