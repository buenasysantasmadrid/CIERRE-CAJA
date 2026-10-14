// ============================================================================
// CAJITA — la caja de seguridad. Su planilla (tipo CAJITA en la pestaña
// "Archivos" del Índice; si no está, la de CAJITA_ID_) tiene una sola
// pestaña: fila de encabezados FECHA · MOTIVO · INGRESO · EGRESO · TOTAL ·
// RESPONSABLE, y debajo una fila por movimiento. La primera fila (FONDO
// FIJO) trae el TOTAL escrito a mano; en las demás TOTAL es una fórmula:
// ese fondo + todos los ingresos − todos los egresos hasta esa fila. Los
// egresos van en negativo.
//
// Se escribe desde dos sitios:
//  - El cierre de caja (en cada guardado del día, ver doPost): un Ingreso
//    "Cajita" sale de la cajita (EGRESO) y un Retiro "Cajita" entra en ella
//    (INGRESO). Esas filas llevan el ID del movimiento en la columna Z
//    (oculta), para actualizarlas o borrarlas sin duplicar.
//  - La pantalla Cajita del menú (Varios → Cajita): ingresos y egresos
//    directos, sin ID (el guardado del día no los toca).
// ============================================================================

var CAJITA_ID_ = '1YvDxuHs5AE2jIHMCe7DjjHtZ6WOgV3w0TDr5NckaEUs';
var CAJITA_COL_ID_ = 26; // Z
var CAJITA_ANCHO_ = 6;   // A:F
var CAJITA_FORMATO_IMPORTE_ = '#,##0.00 "€"';

function hojaCajita_() {
  var id = buscarEnIndice_('CAJITA', periodoAnual_(hoyISO_()));
  if (!id) {
    var todas = planillasDelTipo_('CAJITA');
    if (todas.length) id = todas[todas.length - 1].id;
  }
  return SpreadsheetApp.openById(id || CAJITA_ID_).getSheets()[0];
}

// Fila de los encabezados (la que tiene FECHA en la columna A).
function filaEncabezadosCajita_(hoja) {
  var a = hoja.getRange(1, 1, Math.min(10, Math.max(hoja.getLastRow(), 1)), 1).getValues();
  for (var i = 0; i < a.length; i++) if (normalizarClave_(a[i][0]) === 'FECHA') return i + 1;
  return 2;
}

// Escribe un movimiento en la fila `r` (la fila del fondo fijo es la de
// debajo de los encabezados y no se escribe nunca).
function escribirFilaCajita_(hoja, primera, r, mov, id) {
  var importe = Math.round((Number(mov.importe) || 0) * 100) / 100;
  hoja.getRange(r, 1, 1, CAJITA_ANCHO_).setValues([[
    fechaComoDate_(mov.fecha), String(mov.motivo || ''),
    importe > 0 ? importe : '', importe < 0 ? importe : '', '', String(mov.responsable || '')
  ]]);
  hoja.getRange(r, 1).setNumberFormat(hoja.getRange(primera, 1).getNumberFormat() || 'dd/mm/yyyy');
  hoja.getRange(r, 3, 1, 3).setNumberFormat(CAJITA_FORMATO_IMPORTE_);
  hoja.getRange(r, 5).setFormula('=$E$' + primera + '+SUM(C$' + (primera + 1) + ':C' + r + ')+SUM(D$' + (primera + 1) + ':D' + r + ')');
  hoja.getRange(r, CAJITA_COL_ID_).setValue(id || '');
}

// Lo que muestra la pantalla Cajita: los últimos movimientos (el más nuevo
// primero). El total no va: solo con contraseña (ver totalCajita_).
function leerCajita_() {
  var r = leerCajitaCompleta_();
  return { ok: true, movimientos: r.movimientos };
}

// El total de la cajita, con contraseña. La contraseña está en CAJITA_CLAVE_,
// en CajitaClave.gs: ese archivo no va al repositorio (es público), así que
// solo existe en este ordenador y en Apps Script.
function totalCajita_(clave) {
  if (typeof CAJITA_CLAVE_ === 'undefined' || !CAJITA_CLAVE_) throw new Error('Falta la contraseña de la cajita en Apps Script (CajitaClave.gs)');
  if (String(clave || '') !== CAJITA_CLAVE_) {
    Utilities.sleep(1000); // para que no se puedan probar muchas seguidas
    throw new Error('Contraseña incorrecta');
  }
  return { ok: true, total: leerCajitaCompleta_().total };
}

function leerCajitaCompleta_() {
  var hoja = hojaCajita_();
  SpreadsheetApp.flush();
  var primera = filaEncabezadosCajita_(hoja) + 1;
  var ultima = hoja.getLastRow();
  var movimientos = [], total = 0;
  if (ultima >= primera) {
    var v = hoja.getRange(primera, 1, ultima - primera + 1, CAJITA_ANCHO_).getValues();
    var ids = hoja.getRange(primera, CAJITA_COL_ID_, ultima - primera + 1, 1).getValues();
    v.forEach(function (f, i) {
      if (f[0] === '' && f[1] === '' && f[2] === '' && f[3] === '') return;
      if (typeof f[4] === 'number') total = f[4];
      movimientos.push({
        fila: primera + i, desdeCaja: !!ids[i][0],
        fecha: textoFechaIds_(f[0]), motivo: String(f[1] || ''),
        ingreso: typeof f[2] === 'number' ? f[2] : null, egreso: typeof f[3] === 'number' ? f[3] : null,
        total: typeof f[4] === 'number' ? f[4] : null, responsable: String(f[5] || '')
      });
    });
  }
  return { ok: true, total: total, movimientos: movimientos.reverse().slice(0, 30) };
}

// Ingreso o egreso apuntado desde la pantalla Cajita. Con `m.fila` es una
// corrección de esa fila: antes se comprueba que siga siendo el movimiento
// que se tocó (`m.original`), por si las filas se movieron entretanto.
function guardarCajita_(m) {
  var fecha = String(m.fecha || '');
  var motivo = String(m.motivo || '').trim();
  var importe = Math.abs(Number(m.importe) || 0);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha)) throw new Error('Fecha inválida');
  if (m.tipo !== 'ingreso' && m.tipo !== 'egreso') throw new Error('Elige INGRESO o EGRESO');
  if (!motivo) throw new Error('Falta el motivo');
  if (!importe) throw new Error('Falta el importe');
  var hoja = hojaCajita_();
  var primera = filaEncabezadosCajita_(hoja) + 1;
  var r = Math.max(hoja.getLastRow(), primera) + 1;
  if (m.fila) {
    r = Number(m.fila);
    var o = m.original || {};
    if (!(r > primera && r <= hoja.getLastRow())) throw new Error('Ese movimiento ya no está en la cajita; vuelve a abrirla');
    var f = hoja.getRange(r, 1, 1, CAJITA_ANCHO_).getValues()[0];
    var mismo = textoFechaIds_(f[0]) === o.fecha && String(f[1] || '') === o.motivo &&
      (typeof f[2] === 'number' ? f[2] : null) === o.ingreso && (typeof f[3] === 'number' ? f[3] : null) === o.egreso;
    if (!mismo) throw new Error('La cajita cambió mientras tanto; vuelve a abrirla y tócalo otra vez');
    if (hoja.getRange(r, CAJITA_COL_ID_).getValue()) throw new Error('Este viene de la caja del día: se cambia desde esa caja');
  }
  escribirFilaCajita_(hoja, primera, r, {
    fecha: fecha, motivo: motivo, importe: m.tipo === 'egreso' ? -importe : importe, responsable: m.responsable
  }, '');
  return { ok: true };
}

// Movimientos "Cajita" de la caja del día, ya con el signo de la cajita:
// el Ingreso sale de ella (negativo) y el Retiro entra (positivo). El
// motivo es el del desplegable de la pantalla Cajita ("Ingreso en caja del
// día" / "Retiro caja del día"), con lo escrito en la caja entre paréntesis.
function movimientosCajitaDelDia_(data) {
  var lista = [];
  var motivo = function (base, m) {
    var extra = String(m.motivo || '').trim();
    return extra ? base + ' (' + extra + ')' : base;
  };
  ['mediodia', 'noche'].forEach(function (k) {
    var t = data && data[k];
    (t && t.movimientos || []).forEach(function (m) {
      if (!m || !m.id || m.subtipo !== 'cajita') return;
      var imp = Math.abs(Number(m.importe) || 0);
      if (m.tipo === 'ingreso') lista.push({ id: m.id, importe: -imp, motivo: motivo('Ingreso en caja del día', m), responsable: m.responsable });
      else if (m.tipo === 'egreso') lista.push({ id: m.id, importe: imp, motivo: motivo('Retiro caja del día', m), responsable: m.responsable });
    });
  });
  return lista;
}

// Deja la cajita igual a lo que tiene la app para la caja de `fechaISO`:
// actualiza los que ya estaban, agrega los nuevos al final y borra la fila
// de los que ya no están. `movs` vacío = borrar lo que la app había escrito
// de ese día. Para no abrir la planilla en cada guardado, una marca por día
// dice si la app escribió algo ese día.
function sincronizarCajitaDelDia_(fechaISO, movs) {
  var props = PropertiesService.getScriptProperties();
  var marca = 'cajita_' + fechaISO;
  if (!movs.length && !props.getProperty(marca)) return { ok: true, nada: true };

  var hoja = hojaCajita_();
  var primera = filaEncabezadosCajita_(hoja) + 1;
  var ultima = hoja.getLastRow();
  var filaDeId = {};
  if (ultima > primera) {
    var n = ultima - primera;
    var ids = hoja.getRange(primera + 1, CAJITA_COL_ID_, n, 1).getValues();
    var fechas = hoja.getRange(primera + 1, 1, n, 1).getValues();
    for (var i = 0; i < n; i++) {
      var id = String(ids[i][0] || '');
      if (id && textoFechaIds_(fechas[i][0]) === fechaISO) filaDeId[id] = primera + 1 + i;
    }
  }

  var actuales = {};
  movs.forEach(function (m) {
    actuales[m.id] = true;
    var r = filaDeId[m.id];
    if (!r) r = ++ultima;
    escribirFilaCajita_(hoja, primera, r, { fecha: fechaISO, motivo: m.motivo, importe: m.importe, responsable: m.responsable }, m.id);
  });

  // Los que se borraron en la app: de abajo arriba, para no mover las filas
  // que faltan por borrar.
  Object.keys(filaDeId).filter(function (id) { return !actuales[id]; })
    .map(function (id) { return filaDeId[id]; })
    .sort(function (a, b) { return b - a; })
    .forEach(function (r) { hoja.deleteRow(r); });

  if (movs.length && !hoja.isColumnHiddenByUser(CAJITA_COL_ID_)) hoja.hideColumns(CAJITA_COL_ID_);
  if (movs.length) props.setProperty(marca, '1'); else props.deleteProperty(marca);
  return { ok: true, escritos: movs.length };
}

// Lo que corre en cada guardado de la app (ver doPost).
function escribirCajita_(data) {
  try {
    return sincronizarCajitaDelDia_(data.fecha, movimientosCajitaDelDia_(data));
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}
