// ============================================================================
// CAJITA — la caja de seguridad (en la app, "Sobrecito"). Su planilla
// (tipo CAJITA en la pestaña "Archivos" del Índice; si no está, la de
// CAJITA_ID_) tiene una sola pestaña: fila de encabezados FECHA · MOTIVO ·
// INGRESO · EGRESO · TOTAL · RESPONSABLE, debajo los movimientos, EL MÁS
// NUEVO ARRIBA (cada uno nuevo se inserta justo debajo de los
// encabezados), y al final la fila del FONDO FIJO (MOTIVO "FONDO FIJO"),
// que la app no toca nunca. TOTAL de cada movimiento = el TOTAL de la fila
// de debajo + su ingreso/egreso, así que el de la primera fila es el total
// de ahora. Los egresos van en negativo. Las fórmulas se rehacen en cada
// escritura (ver rehacerTotalesCajita_).
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

// Dónde está cada cosa: `cab` es la fila de los encabezados y `fondo` la
// del FONDO FIJO (si no hay ninguna, la última). Los movimientos van entre
// las dos: de cab + 1 a fondo − 1.
function ubicarCajita_(hoja) {
  var cab = filaEncabezadosCajita_(hoja);
  var ultima = hoja.getLastRow();
  var fondo = ultima;
  if (ultima > cab) {
    var b = hoja.getRange(cab + 1, 2, ultima - cab, 1).getValues();
    for (var i = b.length - 1; i >= 0; i--) {
      if (normalizarClave_(b[i][0]) === 'FONDO FIJO') { fondo = cab + 1 + i; break; }
    }
  }
  if (fondo <= cab) fondo = cab + 1;
  return { hoja: hoja, cab: cab, fondo: fondo };
}

// Escribe un movimiento en la fila `r` (sin el TOTAL: ver
// rehacerTotalesCajita_).
function escribirFilaCajita_(hoja, r, mov, id) {
  var importe = Math.round((Number(mov.importe) || 0) * 100) / 100;
  hoja.getRange(r, 1, 1, CAJITA_ANCHO_).setValues([[
    fechaComoDate_(mov.fecha), String(mov.motivo || ''),
    importe > 0 ? importe : '', importe < 0 ? importe : '', '', String(mov.responsable || '')
  ]]);
  hoja.getRange(r, 3, 1, 3).setNumberFormat(CAJITA_FORMATO_IMPORTE_);
  hoja.getRange(r, CAJITA_COL_ID_).setValue(id || '');
}

// Fila nueva para un movimiento: arriba del todo (justo debajo de los
// encabezados), con el formato de la que tiene debajo (la del más nuevo o,
// si no hay ninguno, la del fondo fijo).
function filaNuevaCajita_(hoja, cab) {
  var r = cab + 1;
  hoja.insertRowBefore(r);
  hoja.getRange(r + 1, 1, 1, CAJITA_ANCHO_).copyTo(hoja.getRange(r, 1, 1, CAJITA_ANCHO_), { formatOnly: true });
  return r;
}

// TOTAL de cada movimiento: el de la fila de debajo + su ingreso/egreso (el
// más viejo parte del fondo fijo, que tiene debajo).
function rehacerTotalesCajita_(hoja) {
  var u = ubicarCajita_(hoja);
  // El TOTAL del fondo fijo es su propio importe. Una versión anterior le
  // escribió una fórmula que apuntaba a otras filas (y daba error): se
  // cambia por su ingreso/egreso.
  var fFondo = hoja.getRange(u.fondo, 5).getFormula();
  var propia = '=N(C' + u.fondo + ')+N(D' + u.fondo + ')';
  var otrasFilas = (fFondo.match(/\$?[A-Z]+\$?\d+/gi) || []).some(function (ref) {
    return Number(ref.replace(/\D/g, '')) !== u.fondo;
  });
  if (fFondo && otrasFilas) hoja.getRange(u.fondo, 5).setFormula(propia);
  var n = u.fondo - u.cab - 1;
  if (n <= 0) return;
  hoja.getRange(u.cab + 1, 5, n, 1).setFormulaR1C1('=R[1]C+N(RC[-2])+N(RC[-1])');
  hoja.getRange(u.cab + 1, 5, n, 1).setNumberFormat(CAJITA_FORMATO_IMPORTE_);
}

// Las escrituras (pantalla y guardado del día) de a una, para que no se
// pisen las filas.
function conCandadoCajita_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return fn(); } finally { lock.releaseLock(); }
}

// Lo que muestra la pantalla Cajita: el total y los últimos movimientos
// (el más nuevo primero). Solo con la contraseña (ver comprobarClaveCajita_).
function leerCajita_(clave) {
  comprobarClaveCajita_(clave);
  return leerCajitaCompleta_();
}

// La contraseña de la Cajita está en CAJITA_CLAVE_, en CajitaClave.gs: ese
// archivo no va al repositorio (es público), así que solo existe en este
// ordenador y en Apps Script.
function comprobarClaveCajita_(clave) {
  if (typeof CAJITA_CLAVE_ === 'undefined' || !CAJITA_CLAVE_) throw new Error('Falta la contraseña del sobrecito en Apps Script (CajitaClave.gs)');
  if (String(clave || '') !== CAJITA_CLAVE_) {
    Utilities.sleep(1000); // para que no se puedan probar muchas seguidas
    throw new Error('Contraseña incorrecta');
  }
}

function leerCajitaCompleta_() {
  var hoja = hojaCajita_();
  // Las fórmulas de TOTAL al día (por si se movieron filas a mano).
  conCandadoCajita_(function () { rehacerTotalesCajita_(hoja); });
  SpreadsheetApp.flush();
  var u = ubicarCajita_(hoja);
  var movimientos = [], total = 0;
  var v = hoja.getRange(u.cab + 1, 1, u.fondo - u.cab, CAJITA_ANCHO_).getValues(); // movimientos + fondo
  var ids = hoja.getRange(u.cab + 1, CAJITA_COL_ID_, u.fondo - u.cab, 1).getValues();
  if (typeof v[0][4] === 'number') total = v[0][4]; // la primera fila: el más nuevo (o el fondo)
  v.forEach(function (f, i) {
    if (i === v.length - 1) return; // el fondo fijo
    if (f[0] === '' && f[1] === '' && f[2] === '' && f[3] === '') return;
    movimientos.push({
      fila: u.cab + 1 + i, desdeCaja: !!ids[i][0],
      fecha: textoFechaIds_(f[0]), motivo: String(f[1] || ''),
      ingreso: typeof f[2] === 'number' ? f[2] : null, egreso: typeof f[3] === 'number' ? f[3] : null,
      total: typeof f[4] === 'number' ? f[4] : null, responsable: String(f[5] || '')
    });
  });
  return { ok: true, total: total, movimientos: movimientos.slice(0, 30), ultimoError: PropertiesService.getScriptProperties().getProperty('cajita_ultimo_error') || '' };
}

// Ingreso o egreso apuntado desde la pantalla Cajita. Con `m.fila` es una
// corrección de esa fila: antes se comprueba que siga siendo el movimiento
// que se tocó (`m.original`), por si las filas se movieron entretanto.
function guardarCajita_(m, clave) {
  comprobarClaveCajita_(clave);
  var fecha = String(m.fecha || '');
  var motivo = String(m.motivo || '').trim();
  var importe = Math.abs(Number(m.importe) || 0);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha)) throw new Error('Fecha inválida');
  if (m.tipo !== 'ingreso' && m.tipo !== 'egreso') throw new Error('Elige INGRESO o EGRESO');
  if (!motivo) throw new Error('Falta el motivo');
  if (!importe) throw new Error('Falta el importe');
  return conCandadoCajita_(function () {
    var hoja = hojaCajita_();
    var u = ubicarCajita_(hoja), r;
    if (m.fila) {
      r = Number(m.fila);
      var o = m.original || {};
      if (!(r > u.cab && r < u.fondo)) throw new Error('Ese movimiento ya no está en el sobrecito; vuelve a abrirlo');
      var f = hoja.getRange(r, 1, 1, CAJITA_ANCHO_).getValues()[0];
      var mismo = textoFechaIds_(f[0]) === o.fecha && String(f[1] || '') === o.motivo &&
        (typeof f[2] === 'number' ? f[2] : null) === o.ingreso && (typeof f[3] === 'number' ? f[3] : null) === o.egreso;
      if (!mismo) throw new Error('El sobrecito cambió mientras tanto; vuelve a abrirlo y tócalo otra vez');
      if (hoja.getRange(r, CAJITA_COL_ID_).getValue()) throw new Error('Este viene de la caja del día: se cambia desde esa caja');
    } else {
      r = filaNuevaCajita_(hoja, u.cab);
    }
    escribirFilaCajita_(hoja, r, {
      fecha: fecha, motivo: motivo, importe: m.tipo === 'egreso' ? -importe : importe, responsable: m.responsable
    }, '');
    rehacerTotalesCajita_(hoja);
    return { ok: true };
  });
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
    // movimientosRaw: con las claves internas (tipo 'ingreso', subtipo
    // 'cajita'); en `movimientos` vienen las etiquetas ("Ingreso", "Sobrecito").
    (t && (t.movimientosRaw || t.movimientos) || []).forEach(function (m) {
      if (!m || !m.id || m.subtipo !== 'cajita') return;
      var imp = Math.abs(Number(m.importe) || 0);
      if (m.tipo === 'ingreso') lista.push({ id: m.id, importe: -imp, motivo: motivo('Ingreso en caja del día', m), responsable: m.responsable });
      else if (m.tipo === 'egreso') lista.push({ id: m.id, importe: imp, motivo: motivo('Retiro caja del día', m), responsable: m.responsable });
    });
  });
  return lista;
}

// Deja la cajita igual a lo que tiene la app para la caja de `fechaISO`:
// borra la fila de los que ya no están, actualiza los que ya estaban y
// agrega los nuevos arriba. `movs` vacío = borrar lo que la app había
// escrito de ese día. Para no abrir la planilla en cada guardado, una marca
// por día dice si la app escribió algo ese día.
function sincronizarCajitaDelDia_(fechaISO, movs) {
  var props = PropertiesService.getScriptProperties();
  var marca = 'cajita_' + fechaISO;
  if (!movs.length && !props.getProperty(marca)) return { ok: true, nada: true };

  return conCandadoCajita_(function () {
    var hoja = hojaCajita_();
    var filasDelDia = function () {
      var filaDeId = {}, u = ubicarCajita_(hoja), n = u.fondo - u.cab - 1;
      if (n > 0) {
        var ids = hoja.getRange(u.cab + 1, CAJITA_COL_ID_, n, 1).getValues();
        var fechas = hoja.getRange(u.cab + 1, 1, n, 1).getValues();
        for (var i = 0; i < n; i++) {
          var id = String(ids[i][0] || '');
          if (id && textoFechaIds_(fechas[i][0]) === fechaISO) filaDeId[id] = u.cab + 1 + i;
        }
      }
      return filaDeId;
    };

    // Primero se borran los que ya no están en la app (de abajo arriba,
    // para no mover las filas que faltan por borrar)...
    var actuales = {};
    movs.forEach(function (m) { actuales[m.id] = true; });
    var antes = filasDelDia();
    Object.keys(antes).filter(function (id) { return !actuales[id]; })
      .map(function (id) { return antes[id]; })
      .sort(function (a, b) { return b - a; })
      .forEach(function (r) { hoja.deleteRow(r); });

    // ...después se corrigen los que ya estaban y los nuevos van arriba.
    var filaDeId = filasDelDia();
    movs.forEach(function (m) {
      if (filaDeId[m.id]) escribirFilaCajita_(hoja, filaDeId[m.id], { fecha: fechaISO, motivo: m.motivo, importe: m.importe, responsable: m.responsable }, m.id);
    });
    movs.filter(function (m) { return !filaDeId[m.id]; }).forEach(function (m) {
      escribirFilaCajita_(hoja, filaNuevaCajita_(hoja, filaEncabezadosCajita_(hoja)), { fecha: fechaISO, motivo: m.motivo, importe: m.importe, responsable: m.responsable }, m.id);
    });
    rehacerTotalesCajita_(hoja);

    if (movs.length && !hoja.isColumnHiddenByUser(CAJITA_COL_ID_)) hoja.hideColumns(CAJITA_COL_ID_);
    if (movs.length) props.setProperty(marca, '1'); else props.deleteProperty(marca);
    return { ok: true, escritos: movs.length };
  });
}

// Lo que corre en cada guardado de la app (ver doPost).
// Si falla, el error queda guardado (con la línea del programa) y la
// pantalla Sobrecito lo muestra hasta el próximo guardado que salga bien.
function escribirCajita_(data) {
  var props = PropertiesService.getScriptProperties();
  try {
    var r = sincronizarCajitaDelDia_(data.fecha, movimientosCajitaDelDia_(data));
    if (!r.nada) props.deleteProperty('cajita_ultimo_error');
    return r;
  } catch (err) {
    var texto = String(err) + (err && err.stack ? ' | ' + String(err.stack).split(/\r?\n/).slice(0, 4).join(' / ') : '');
    try {
      props.setProperty('cajita_ultimo_error', Utilities.formatDate(new Date(), 'Europe/Madrid', 'dd/MM HH:mm') + ' (caja del ' + data.fecha + '): ' + texto.substring(0, 800));
    } catch (e) {}
    return { ok: false, error: String(err) };
  }
}
