// ============================================================================
// Contabilidad: pestañas "años" y "comparativa"
// ============================================================================
// "años" tiene un bloque por año (una fila por mes: DIAS, PROMEDIO DIA,
// TOTAL SIST, MEDIO DIA, NOCHE, UBER/GLOVO, WEB, EMPANADAS, ... TARJETAS,
// TICKETS, EFEVO, DIFERENCIA, RETIRA, PROVEEDORES, NOMINAS, EXTRAS, %), y
// más abajo una tabla ancha (por año: PROMEDIO · TOTAL · TARJETA · TICKET
// de cada mes) y una copia de los bloques de "comparativa".
// "comparativa" tiene por año: el promedio de cada día de la semana de cada
// mes (TOTAL, MEDIODÍA y NOCHE) con la diferencia contra el año anterior, y
// un resumen (PROMEDIO, TOTAL EXTRA, TOTAL FACT, TOTAL PROMEDIO) con su
// diferencia.
//
// Lo que hace esto:
// - 2025: en el archivo de 2026 esos números eran fórmulas a las pestañas
//   del archivo de 2025 y daban #REF!. Se ponen los valores (ver
//   ContabilidadDatos2025.gs, que no va al repo). Una sola vez.
// - El año de la planilla (2026): si no tiene sus bloques, se crean copiando
//   los del año anterior, y se completan con lo que hay en las pestañas de
//   cada mes. Después, cada vez que se guarda un día se actualiza su mes.
// Los bloques se buscan por lo que dicen las celdas (el año en la columna
// A), no por número de fila. NOMINAS y EXTRAS no se tocan (no están en la
// caja); PROVEEDORES sale del TOTAL de TOTALES en Albaranes.
// ============================================================================

var DIAS_SEMANA_COMPARATIVA_ = 7; // lunes a domingo (columnas C a I)

// Después de guardar un día (ver semanasTrasGuardar_).
function actualizarAniosComparativa_(anio, mesIndex) {
  var idCont = buscarEnIndice_('CONTABILIDAD', String(anio));
  if (!idCont) return null;
  var ss = SpreadsheetApp.openById(idCont);
  var hojas = hojasAniosComparativa_(ss);
  if (!hojas.anios && !hojas.comparativa) return null;
  arreglarComparativa2025_(hojas);
  var nuevo = asegurarBloquesDelAnio_(hojas, anio);
  var desde = nuevo ? 0 : mesIndex;
  var proveedores = proveedoresPorMesAlbaranes_(anio);
  for (var m = desde; m <= mesIndex; m++) {
    escribirMesAniosComparativa_(hojas, anio, m, datosMesContabilidad_(ss, m), proveedores[m]);
  }
  return true;
}

// Para correr a mano desde el editor: todo el año actual otra vez.
function completarAniosYComparativa() {
  var hoy = new Date();
  var anio = hoy.getFullYear();
  var idCont = buscarEnIndice_('CONTABILIDAD', String(anio));
  if (!idCont) { Logger.log('No hay Contabilidad ' + anio + ' en el Índice.'); return; }
  var ss = SpreadsheetApp.openById(idCont);
  var hojas = hojasAniosComparativa_(ss);
  var resumen = [];
  resumen.push('2025 en comparativa: ' + arreglarComparativa2025_(hojas));
  asegurarBloquesDelAnio_(hojas, anio);
  var proveedores = proveedoresPorMesAlbaranes_(anio);
  for (var m = 0; m <= hoy.getMonth(); m++) {
    var d = datosMesContabilidad_(ss, m);
    escribirMesAniosComparativa_(hojas, anio, m, d, proveedores[m]);
    resumen.push(MESES_MAYUS_[m] + ': ' + (d ? d.dias + ' días, ' + Math.round(d.sist) + ' € sistema' : 'sin datos'));
  }
  Logger.log(resumen.join('\n'));
}

function hojasAniosComparativa_(ss) {
  var res = { anios: null, comparativa: null };
  ss.getSheets().forEach(function (h) {
    var n = normalizarClave_(h.getName());
    if (n === 'ANOS') res.anios = h;
    else if (n.indexOf('COMPAR') === 0) res.comparativa = h;
  });
  return res;
}

// ---------- lectura de las pestañas de cada mes ----------
// Las columnas se ubican por el encabezado de la fila 2 (los meses de antes
// de septiembre de 2026 tienen UBER EAT y WEB; los de después, GLOVO y TPV 3).
// Los días trabajados: la columna A, que en los meses viejos va acumulando
// (1, 2, 2.5...) y en los nuevos tiene 0,5 / 1 por día.
function datosMesContabilidad_(ss, mesIndex) {
  var hoja = ss.getSheetByName(MESES_MAYUS_[mesIndex]);
  if (!hoja) return null;
  var ancho = Math.min(Math.max(hoja.getLastColumn(), 1), 26);
  var vals = hoja.getRange(1, 1, 34, ancho).getValues();
  var enc = vals[1].map(normalizarClave_);
  function col(nombres) {
    for (var i = 0; i < nombres.length; i++) { var k = enc.indexOf(nombres[i]); if (k > -1) return k; }
    return -1;
  }
  var c = {
    sist: col(['TOTAL SIST']), md: col(['MEDIO DIA']), nc: col(['NOCHE']),
    deliv: col(['UBER EAT', 'UBER', 'GLOVO', 'DELIVEROO']), web: col(['WEB']), emp: col(['EMPANADAS']),
    tarj: col(['TARJETAS']), tick: col(['TICKETS']), efvo: col(['EFEVO', 'EFECTIVO']),
    dif: col(['DIFERENCIA']), ret: col(['RETIRA'])
  };
  if (c.sist === -1) return null;
  var r = { dias: 0, sist: 0, md: 0, nc: 0, deliv: 0, web: 0, emp: 0, tarj: 0, tick: 0, efvo: 0, dif: 0, ret: 0, sem: {}, lista: [] };
  ['total', 'md', 'nc'].forEach(function (k) { r.sem[k] = [0, 0, 0, 0, 0, 0, 0]; r.sem['n' + k] = [0, 0, 0, 0, 0, 0, 0]; });
  var acumulada = false, maxA = 0, sumaA = 0, hayDias = false;
  for (var i = 2; i < vals.length; i++) {
    var f = vals[i];
    if (!(f[1] instanceof Date)) continue;
    var n = function (k) { return c[k] > -1 ? (Number(f[c[k]]) || 0) : 0; };
    var a = Number(f[0]) || 0;
    if (a > 1) acumulada = true;
    if (a > maxA) maxA = a;
    sumaA += a;
    ['sist', 'md', 'nc', 'deliv', 'web', 'emp', 'tarj', 'tick', 'efvo', 'dif', 'ret'].forEach(function (k) { r[k] += n(k); });
    var d = (f[1].getDay() + 6) % 7; // lunes = 0
    var tot = n('sist'), md = n('md'), nc = n('nc');
    if (tot > 0) { r.sem.total[d] += tot; r.sem.ntotal[d]++; hayDias = true; }
    if (md > 0) { r.sem.md[d] += md; r.sem.nmd[d]++; }
    if (nc > 0) { r.sem.nc[d] += nc; r.sem.nnc[d]++; }
    // Día por día (lo usa la página de informes, ver Informes.gs).
    if (tot > 0) r.lista.push({
      fecha: Utilities.formatDate(f[1], Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      sistema: redondear_(tot), mediodia: redondear_(md), noche: redondear_(nc),
      delivery: redondear_(n('deliv') + n('web')), empanadas: redondear_(n('emp')), tarjetas: redondear_(n('tarj'))
    });
  }
  if (!hayDias) return null;
  r.dias = acumulada ? maxA : Math.round(sumaA * 10) / 10;
  return r;
}

// TOTAL de TOTALES (Albaranes del año) de cada mes: lo comprado a
// proveedores.
function proveedoresPorMesAlbaranes_(anio) {
  var res = [];
  try {
    var id = buscarEnIndice_('ALBARANES', String(anio));
    if (!id) return res;
    var hoja = SpreadsheetApp.openById(id).getSheetByName('TOTALES');
    if (!hoja) return res;
    var fila = filaTotalDeTotales_(hoja);
    if (fila === -1) return res;
    res = hoja.getRange(fila, 3, 1, 12).getValues()[0].map(function (v) { return Number(v) || 0; });
  } catch (err) { Logger.log('Proveedores de Albaranes: ' + err); }
  return res;
}

// ---------- 2025: valores en vez de fórmulas rotas ----------
function arreglarComparativa2025_(hojas) {
  if (typeof CONTAB_COMPARATIVA_2025_ === 'undefined') return 'no están los datos (ContabilidadDatos2025.gs)';
  var hechas = [];
  [hojas.comparativa, hojas.anios].forEach(function (h) {
    if (!h) return;
    var S = filaBloqueComparativa_(h, 2025);
    if (!S) return;
    // Ya arreglado: la primera celda ya no es una fórmula a otra pestaña.
    if (!/!/.test(h.getRange(S + 2, 3).getFormula())) return;
    CONTAB_COMPARATIVA_2025_.forEach(function (x) {
      var celda = h.getRange(S + x[0], x[1]);
      if (/!/.test(celda.getFormula())) celda.setValue(x[2]);
    });
    hechas.push(h.getName());
  });
  return hechas.length ? 'puestos en ' + hechas.join(' y ') : 'ya estaba';
}

// Fila del título de un año en comparativa: el año en A y "DIF ..." en J o K.
function filaBloqueComparativa_(hoja, anio) {
  var v = hoja.getRange(1, 1, hoja.getLastRow(), 11).getValues();
  for (var i = 0; i < v.length; i++) {
    if (Number(v[i][0]) !== anio) continue;
    if (/^DIF/i.test(String(v[i][9])) || /^DIF/i.test(String(v[i][10]))) return i + 1;
  }
  return 0;
}

// Fila del resumen de un año en comparativa: el año en A y ENERO en B.
function filaResumenComparativa_(hoja, anio) {
  var v = hoja.getRange(1, 1, hoja.getLastRow(), 2).getValues();
  for (var i = 0; i < v.length; i++) {
    if (Number(v[i][0]) === anio && normalizarClave_(v[i][1]) === 'ENERO') return i + 1;
  }
  return 0;
}

// Bloque de un año en "años": el año en A y DIAS en B de la fila siguiente.
function filaBloqueAnio_(hoja, anio) {
  var v = hoja.getRange(1, 1, hoja.getLastRow(), 2).getValues();
  for (var i = 0; i + 1 < v.length; i++) {
    if (Number(v[i][0]) === anio && normalizarClave_(v[i + 1][1]) === 'DIAS') return i + 1;
  }
  return 0;
}

// Tabla ancha de "años": fila con los años como títulos (cada 4 columnas) y
// PROMEDIO debajo. Devuelve {fila, col} del año, o null.
function celdaAnioTablaAncha_(hoja, anio) {
  var ultima = hoja.getLastRow(), ancho = hoja.getLastColumn();
  var v = hoja.getRange(1, 1, ultima, ancho).getValues();
  for (var i = 0; i + 1 < v.length; i++) {
    for (var j = 2; j < ancho; j++) {
      if (Number(v[i][j]) === anio && normalizarClave_(v[i + 1][j]) === 'PROMEDIO') return { fila: i + 1, col: j + 1 };
    }
  }
  return null;
}

// ---------- crear los bloques del año si no están ----------
function asegurarBloquesDelAnio_(hojas, anio) {
  var creado = false;
  if (hojas.anios) {
    if (crearBloqueAnio_(hojas.anios, anio)) creado = true;
    if (crearColumnaTablaAncha_(hojas.anios, anio)) creado = true;
    if (crearBloqueComparativa_(hojas.anios, anio)) creado = true;
  }
  if (hojas.comparativa && crearBloqueComparativa_(hojas.comparativa, anio)) creado = true;
  if (creado) SpreadsheetApp.flush();
  return creado;
}

function filasVacias_(hoja, desde, cuantas, ancho) {
  if (desde + cuantas - 1 > hoja.getMaxRows()) return true;
  return hoja.getRange(desde, 1, cuantas, ancho).getValues().every(function (f) {
    return f.every(function (x) { return x === '' || x == null; });
  });
}

// Bloque del año en "años": copia del del año anterior, a la misma
// distancia que hay entre los dos anteriores (17 filas). Si ahí hay algo
// (la tabla ancha), se insertan filas.
function crearBloqueAnio_(hoja, anio) {
  if (filaBloqueAnio_(hoja, anio)) return false;
  var S0 = filaBloqueAnio_(hoja, anio - 1);
  if (!S0) return false;
  var Sm = filaBloqueAnio_(hoja, anio - 2);
  var alto = Sm ? S0 - Sm : 17;
  var nuevo = S0 + alto;
  if (!filasVacias_(hoja, nuevo, alto, 24)) hoja.insertRowsBefore(nuevo, alto);
  hoja.getRange(S0, 1, alto, 24).copyTo(hoja.getRange(nuevo, 1, alto, 24));
  hoja.getRange(nuevo, 1).setValue(anio);
  // Sin los números del año anterior (se completan mes por mes).
  hoja.getRange(nuevo + 2, 2, 12, 23).clearContent();
  hoja.getRange(nuevo + 14, 2, alto - 14, 23).clearContent();
  hoja.getRange(nuevo + 1, 7).setValue('UBER / GLOVO');
  // Fila del total del año y fila de promedios por mes.
  var tot = nuevo + 14, pri = nuevo + 2, ult = nuevo + 13;
  var f = [];
  for (var c = 2; c <= 21; c++) {
    var L = letraColumna_(c);
    f.push(c === 3 ? '=IFERROR(D' + tot + '/B' + tot + ',"")'
      : c === 10 ? '=IFERROR((G' + tot + '+H' + tot + '+I' + tot + ')/B' + tot + ',"")'
      : c === 12 ? '=IFERROR(K' + tot + '/B' + tot + ',"")'
      : c === 18 ? '=IFERROR(M' + tot + '/B' + tot + ',"")'
      : '=SUM(' + L + pri + ':' + L + ult + ')');
  }
  f.push('=IFERROR(S' + tot + '/K' + tot + ',"")', '=IFERROR(M' + tot + '/D' + tot + ',"")', '=IFERROR((T' + tot + '+U' + tot + ')/K' + tot + ',"")');
  hoja.getRange(tot, 2, 1, 23).setFormulas([f]);
  var p = [];
  for (var c2 = 4; c2 <= 21; c2++) {
    var L2 = letraColumna_(c2);
    p.push(c2 === 10 || c2 === 12 || c2 === 18 ? '' : '=IFERROR(AVERAGE(' + L2 + pri + ':' + L2 + ult + '),"")');
  }
  hoja.getRange(tot + 1, 4, 1, 18).setFormulas([p]);
  return true;
}

// Columna del año en la tabla ancha: 4 columnas más a la derecha que la del
// año anterior, con el mismo formato.
function crearColumnaTablaAncha_(hoja, anio) {
  if (celdaAnioTablaAncha_(hoja, anio)) return false;
  var prev = celdaAnioTablaAncha_(hoja, anio - 1);
  if (!prev) return false;
  var col = prev.col + 4;
  if (hoja.getMaxColumns() < col + 3) hoja.insertColumnsAfter(hoja.getMaxColumns(), col + 3 - hoja.getMaxColumns());
  hoja.getRange(prev.fila, prev.col, 14, 4).copyTo(hoja.getRange(prev.fila, col, 14, 4));
  hoja.getRange(prev.fila + 2, col, 12, 4).clearContent();
  hoja.getRange(prev.fila, col).setValue(anio);
  return true;
}

// Bloque del año en comparativa (y su resumen): copia del del año anterior,
// a la misma distancia que hay entre los dos anteriores (57 filas), así las
// fórmulas de diferencia (este año menos el anterior) quedan bien.
function crearBloqueComparativa_(hoja, anio) {
  if (filaBloqueComparativa_(hoja, anio)) return false;
  var S0 = filaBloqueComparativa_(hoja, anio - 1);
  var R0 = filaResumenComparativa_(hoja, anio - 1);
  if (!S0 || !R0 || R0 < S0) return false;
  var Sm = filaBloqueComparativa_(hoja, anio - 2);
  var paso = Sm ? S0 - Sm : 57;
  var alto = R0 + 12 - S0; // hasta la última fila de la diferencia del resumen
  var ancho = Math.max(hoja.getLastColumn(), 20);
  var nuevo = S0 + paso;
  if (!filasVacias_(hoja, nuevo, alto, ancho)) hoja.insertRowsBefore(nuevo, alto);
  if (hoja.getMaxRows() < nuevo + alto) hoja.insertRowsAfter(hoja.getMaxRows(), nuevo + alto - hoja.getMaxRows());
  hoja.getRange(S0, 1, alto, ancho).copyTo(hoja.getRange(nuevo, 1, alto, ancho));
  var R = nuevo + (R0 - S0);
  hoja.getRange(nuevo, 1).setValue(anio);
  var celdaDif = /^DIF/i.test(String(hoja.getRange(nuevo, 11).getValue())) ? 11 : 10;
  hoja.getRange(nuevo, celdaDif).setValue('DIF ' + (anio - 1) + '/ ' + anio);
  hoja.getRange(R, 1).setValue(anio);
  hoja.getRange(R + 7, 1).setValue('DIF ' + (anio - 1) + '/' + String(anio).slice(2));
  // Sin los números del año anterior.
  hoja.getRange(nuevo + 2, 3, 36, DIAS_SEMANA_COMPARATIVA_).clearContent();
  hoja.getRange(R + 1, 2, 4, 12).clearContent();
  // Algún número suelto que había entre el bloque y el resumen.
  for (var r = nuevo + 38; r < R; r++) hoja.getRange(r, 1, 1, ancho).clearContent();
  return true;
}

// ---------- escribir un mes ----------
function escribirMesAniosComparativa_(hojas, anio, m, d, proveedores) {
  if (hojas.anios) {
    var S = filaBloqueAnio_(hojas.anios, anio);
    if (S) escribirMesAnios_(hojas.anios, S + 2 + m, d, proveedores);
    var t = celdaAnioTablaAncha_(hojas.anios, anio);
    if (t) {
      hojas.anios.getRange(t.fila + 2 + m, t.col, 1, 4).setValues([d ? [
        redondear_(d.dias ? (d.sist + d.deliv + d.web + d.emp) / d.dias : ''),
        redondear_(d.sist + d.deliv + d.web + d.emp), redondear_(d.tarj), redondear_(d.tick)
      ] : ['', '', '', '']]);
    }
    escribirMesComparativa_(hojas.anios, anio, m, d);
  }
  if (hojas.comparativa) escribirMesComparativa_(hojas.comparativa, anio, m, d);
}

function redondear_(x) { return x === '' ? '' : Math.round(x * 100) / 100; }

function escribirMesAnios_(hoja, fila, d, proveedores) {
  if (!d) { hoja.getRange(fila, 2, 1, 18).clearContent(); hoja.getRange(fila, 22, 1, 3).clearContent(); return; }
  var r = fila;
  hoja.getRange(fila, 2, 1, 18).setValues([[
    d.dias, '=IFERROR(D' + r + '/B' + r + ',"")', redondear_(d.sist), redondear_(d.md), redondear_(d.nc),
    redondear_(d.deliv), redondear_(d.web), redondear_(d.emp),
    '=IFERROR((G' + r + '+H' + r + '+I' + r + ')/B' + r + ',"")', '=D' + r + '+G' + r + '+H' + r + '+I' + r,
    '=IFERROR(K' + r + '/B' + r + ',"")', redondear_(d.tarj), redondear_(d.tick), redondear_(d.efvo),
    redondear_(d.dif), redondear_(d.ret), '=IFERROR(M' + r + '/B' + r + ',"")', proveedores ? redondear_(proveedores) : ''
  ]]);
  // NOMINAS y EXTRAS (T, U) quedan como estén; los % se calculan.
  hoja.getRange(fila, 22, 1, 3).setFormulas([[
    '=IFERROR(S' + r + '/K' + r + ',"")', '=IFERROR(M' + r + '/D' + r + ',"")', '=IFERROR((T' + r + '+U' + r + ')/K' + r + ',"")'
  ]]);
}

function escribirMesComparativa_(hoja, anio, m, d) {
  var S = filaBloqueComparativa_(hoja, anio);
  if (S) {
    var filas = [];
    ['total', 'md', 'nc'].forEach(function (k) {
      var fila = [];
      for (var i = 0; i < DIAS_SEMANA_COMPARATIVA_; i++) {
        fila.push(d && d.sem['n' + k][i] ? redondear_(d.sem[k][i] / d.sem['n' + k][i]) : '');
      }
      filas.push(fila);
    });
    hoja.getRange(S + 2 + 3 * m, 3, 3, DIAS_SEMANA_COMPARATIVA_).setValues(filas);
  }
  var R = filaResumenComparativa_(hoja, anio);
  if (R) {
    var extra = d ? d.deliv + d.web + d.emp : 0;
    hoja.getRange(R + 1, 2 + m, 4, 1).setValues(d && d.dias ? [
      [redondear_(d.sist / d.dias)], [redondear_(extra)], [redondear_(d.sist + extra)], [redondear_((d.sist + extra) / d.dias)]
    ] : [[''], [''], [''], ['']]);
  }
}
