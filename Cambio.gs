// ============================================================================
// CAMBIO — el cambio que hay en la caja de seguridad (blister de monedas de
// 0,10 a 2 € y billetes de 5, 10 y 20 €). Su planilla (tipo CAMBIO en la
// pestaña "Archivos" del Índice) la crea sola la primera vez que se abre la
// pantalla. Tiene una sola pestaña:
//   Fila 1: cuánto hay de cada uno (suma de las columnas C a J).
//   Fila 2: cuánto es en euros, y en K el total.
//   Fila 3: encabezados. FECHA · MOVIMIENTO · una columna por cada tipo ·
//           IMPORTE · RESPONSABLE · DETALLE.
//   Desde la 4: los movimientos, EL MÁS NUEVO ARRIBA. Lo que sale va en
//           negativo. Son fórmulas: borrar o corregir una fila a mano en la
//           planilla cambia lo que hay.
//
// Se escribe desde dos sitios:
//  - La pantalla Cambio del menú (Varios → Cambio): lo que se agrega a la caja
//    de seguridad.
//  - El cierre de caja (en cada guardado del día, ver doPost): un Ingreso
//    "Cambio" sale de la caja de seguridad. Esas filas llevan el ID del
//    movimiento en la columna Z (oculta), para actualizarlas o borrarlas sin
//    duplicar. Qué cambio fue se saca del texto que deja la calculadora
//    ("1 blister de 1€"); las monedas sueltas restan la parte de un blister.
//    Si no se entiende, la fila queda con el importe y sin descontar nada.
// ============================================================================

var CAMBIO_DENOMINACIONES_ = [ // unidad: lo que vale un blister o un billete
  { clave: 'blister-0.1', tipo: 'blister', valor: 0.1, unidad: 4, nombre: 'Blister 0,10 €' },
  { clave: 'blister-0.2', tipo: 'blister', valor: 0.2, unidad: 8, nombre: 'Blister 0,20 €' },
  { clave: 'blister-0.5', tipo: 'blister', valor: 0.5, unidad: 20, nombre: 'Blister 0,50 €' },
  { clave: 'blister-1', tipo: 'blister', valor: 1, unidad: 25, nombre: 'Blister 1 €' },
  { clave: 'blister-2', tipo: 'blister', valor: 2, unidad: 50, nombre: 'Blister 2 €' },
  { clave: 'billete-5', tipo: 'billete', valor: 5, unidad: 5, nombre: 'Billetes 5 €' },
  { clave: 'billete-10', tipo: 'billete', valor: 10, unidad: 10, nombre: 'Billetes 10 €' },
  { clave: 'billete-20', tipo: 'billete', valor: 20, unidad: 20, nombre: 'Billetes 20 €' }
];
var CAMBIO_FILA_ENCABEZADOS_ = 3;
var CAMBIO_PRIMERA_COL_ = 3; // C
var CAMBIO_COL_IMPORTE_ = CAMBIO_PRIMERA_COL_ + CAMBIO_DENOMINACIONES_.length; // K
var CAMBIO_ANCHO_ = CAMBIO_COL_IMPORTE_ + 2; // A:M
var CAMBIO_COL_ID_ = 26; // Z
// Las cajas de antes no descuentan: lo que había se cuenta y se agrega a mano.
var CAMBIO_DESDE_ = '2026-11-01';

function letraCambio_(i) {
  return String.fromCharCode(64 + CAMBIO_PRIMERA_COL_ + i);
}

// Lo que vale una fila: cantidad × lo que vale cada uno.
function formulaFilaCambio_(r) {
  return '=' + CAMBIO_DENOMINACIONES_.map(function (d, i) { return letraCambio_(i) + r + '*' + d.unidad; }).join('+');
}

function hojaCambio_() {
  var lista = planillasDelTipo_('CAMBIO');
  if (lista.length) return SpreadsheetApp.openById(lista[0].id).getSheets()[0];
  return crearPlanillaCambio_();
}

// La planilla nueva queda en la misma carpeta que la del sobrecito y con los
// mismos editores.
function crearPlanillaCambio_() {
  var ss = SpreadsheetApp.create('Cambio');
  var hoja = ss.getSheets()[0];
  hoja.setName('Cambio');
  var filaHay = ['', 'HAY AHORA'], filaEuros = ['', 'EN EUROS'], enc = ['FECHA', 'MOVIMIENTO'];
  CAMBIO_DENOMINACIONES_.forEach(function (d, i) {
    var l = letraCambio_(i);
    filaHay.push('=SUM(' + l + '$' + CAMBIO_FILA_ENCABEZADOS_ + ':' + l + ')');
    filaEuros.push('=' + l + '1*' + d.unidad);
    enc.push(d.nombre);
  });
  var uK = letraCambio_(CAMBIO_DENOMINACIONES_.length - 1);
  filaHay = filaHay.concat(['', '', '']);
  filaEuros = filaEuros.concat(['=SUM(C2:' + uK + '2)', '', '']);
  enc = enc.concat(['IMPORTE', 'RESPONSABLE', 'DETALLE']);
  hoja.getRange(1, 1, 3, CAMBIO_ANCHO_).setValues([filaHay, filaEuros, enc]);
  hoja.getRange(1, 1, 3, CAMBIO_ANCHO_).setFontWeight('bold').setHorizontalAlignment('center');
  hoja.getRange(1, 1, 1, CAMBIO_ANCHO_).setBackground('#d9ead3');
  hoja.getRange(3, 1, 1, CAMBIO_ANCHO_).setBackground('#efefef').setWrap(true);
  hoja.getRange(2, CAMBIO_PRIMERA_COL_, 1, CAMBIO_DENOMINACIONES_.length + 1).setNumberFormat(CAJITA_FORMATO_IMPORTE_);
  hoja.setFrozenRows(CAMBIO_FILA_ENCABEZADOS_);
  hoja.setColumnWidth(2, 200);
  hoja.setColumnWidth(CAMBIO_ANCHO_, 320);
  hoja.hideColumns(CAMBIO_COL_ID_);

  var archivo = DriveApp.getFileById(ss.getId());
  try {
    var cajita = DriveApp.getFileById(hojaCajita_().getParent().getId());
    var yo = Session.getEffectiveUser().getEmail();
    var personas = cajita.getEditors().concat([cajita.getOwner()]);
    personas.forEach(function (p) {
      if (p && p.getEmail() && p.getEmail() !== yo) archivo.addEditor(p.getEmail());
    });
    var carpetas = cajita.getParents();
    if (carpetas.hasNext()) archivo.moveTo(carpetas.next());
  } catch (err) {
    Logger.log('Planilla de Cambio, carpeta y editores: ' + err);
  }
  registrarEnIndice_('CAMBIO', 'TODOS', ss.getId(), 'Cambio');
  return hoja;
}

// Fila nueva para un movimiento: justo debajo de los encabezados, dentro de
// las sumas de la fila 1 (que empiezan en los encabezados).
function filaNuevaCambio_(hoja) {
  var r = CAMBIO_FILA_ENCABEZADOS_ + 1;
  hoja.insertRowBefore(r);
  hoja.getRange(r, 1, 1, CAMBIO_ANCHO_).setFontWeight('normal').setBackground(null);
  return r;
}

// mov.cantidades: lista alineada con CAMBIO_DENOMINACIONES_, o null si no se
// sabe qué fue (la fila lleva solo mov.importe y no cambia lo que hay).
function escribirFilaCambio_(hoja, r, mov, id) {
  var celdas = CAMBIO_DENOMINACIONES_.map(function (d, i) {
    return mov.cantidades && mov.cantidades[i] ? mov.cantidades[i] : '';
  });
  hoja.getRange(r, 1, 1, CAMBIO_ANCHO_).setValues([
    [fechaComoDate_(mov.fecha), String(mov.movimiento || '')].concat(celdas, [
      mov.cantidades ? formulaFilaCambio_(r) : mov.importe, String(mov.responsable || ''), String(mov.detalle || '')
    ])
  ]);
  hoja.getRange(r, 1).setNumberFormat('dd/MM/yyyy');
  hoja.getRange(r, CAMBIO_COL_IMPORTE_).setNumberFormat(CAJITA_FORMATO_IMPORTE_);
  hoja.getRange(r, CAMBIO_COL_ID_).setValue(id || '');
}

function leerCambio_() {
  var hoja = conCandadoCajita_(hojaCambio_);
  SpreadsheetApp.flush();
  var n = CAMBIO_DENOMINACIONES_.length;
  var arriba = hoja.getRange(1, CAMBIO_PRIMERA_COL_, 2, n + 1).getValues();
  var stock = CAMBIO_DENOMINACIONES_.map(function (d, i) {
    return { clave: d.clave, hay: Number(arriba[0][i]) || 0, euros: Number(arriba[1][i]) || 0 };
  });
  var movimientos = [];
  var cuantas = Math.min(30, hoja.getLastRow() - CAMBIO_FILA_ENCABEZADOS_);
  if (cuantas > 0) {
    var v = hoja.getRange(CAMBIO_FILA_ENCABEZADOS_ + 1, 1, cuantas, CAMBIO_ANCHO_).getValues();
    var ids = hoja.getRange(CAMBIO_FILA_ENCABEZADOS_ + 1, CAMBIO_COL_ID_, cuantas, 1).getValues();
    v.forEach(function (f, i) {
      if (f.every(function (x) { return x === ''; })) return;
      movimientos.push({
        fecha: textoFechaIds_(f[0]), movimiento: String(f[1] || ''), desdeCaja: !!ids[i][0],
        cantidades: textoCantidadesCambio_(f.slice(CAMBIO_PRIMERA_COL_ - 1, CAMBIO_COL_IMPORTE_ - 1)),
        importe: Number(f[CAMBIO_COL_IMPORTE_ - 1]) || 0,
        responsable: String(f[CAMBIO_COL_IMPORTE_] || ''), detalle: String(f[CAMBIO_COL_IMPORTE_ + 1] || '')
      });
    });
  }
  return {
    ok: true, stock: stock, total: Number(arriba[1][n]) || 0, movimientos: movimientos,
    ultimoError: PropertiesService.getScriptProperties().getProperty('cambio_ultimo_error') || ''
  };
}

// "2 × Blister 1 €, 10 × Billetes 5 €", sin el signo (ya va en el importe).
function textoCantidadesCambio_(celdas) {
  return celdas.map(function (c, i) {
    var q = Math.abs(Number(c) || 0);
    return q ? String(Math.round(q * 100) / 100).replace('.', ',') + ' × ' + CAMBIO_DENOMINACIONES_[i].nombre : '';
  }).filter(String).join(', ');
}

// Cambio agregado a la caja de seguridad desde la pantalla.
function guardarCambio_(m) {
  var fecha = String(m.fecha || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha)) throw new Error('Fecha inválida');
  if (!String(m.responsable || '').trim()) throw new Error('Falta el responsable');
  var pedidas = m.cantidades || {};
  var cantidades = CAMBIO_DENOMINACIONES_.map(function (d) {
    var q = Number(pedidas[d.clave]) || 0;
    if (q < 0 || q !== Math.floor(q)) throw new Error('Las cantidades tienen que ser números enteros');
    return q;
  });
  if (!cantidades.some(Boolean)) throw new Error('Pon la cantidad de algo');
  return conCandadoCajita_(function () {
    var hoja = hojaCambio_();
    var r = filaNuevaCambio_(hoja);
    escribirFilaCambio_(hoja, r, {
      fecha: fecha, movimiento: 'Agregado a la caja de seguridad', cantidades: cantidades,
      responsable: m.responsable, detalle: String(m.detalle || '').trim()
    }, '');
    return { ok: true, fila: r };
  });
}

// Del texto de la calculadora a la lista de cantidades que salen, o null si
// no se entiende o si no da el importe del movimiento.
function cantidadesDeTextoCambio_(texto, importe) {
  var patron = /(\d+)\s*(billetes?|monedas?|bl[ií]ster(?:es)?)\s*de\s*([\d.,]+)\s*€?/gi;
  var cantidades = CAMBIO_DENOMINACIONES_.map(function () { return 0; });
  var suma = 0, hubo = false, t;
  while ((t = patron.exec(String(texto || '')))) {
    var n = Number(t[1]), valor = Number(t[3].replace(',', '.')), esBillete = /^billete/i.test(t[2]), esMoneda = /^moneda/i.test(t[2]);
    var i = -1;
    CAMBIO_DENOMINACIONES_.forEach(function (d, k) {
      if (d.valor === valor && d.tipo === (esBillete ? 'billete' : 'blister')) i = k;
    });
    if (i === -1) return null;
    var euros = esMoneda ? n * valor : n * CAMBIO_DENOMINACIONES_[i].unidad;
    cantidades[i] += Math.round(euros / CAMBIO_DENOMINACIONES_[i].unidad * 10000) / 10000;
    suma += euros;
    hubo = true;
  }
  return hubo && Math.abs(suma - importe) <= 0.01 ? cantidades : null;
}

// Los Ingresos "Cambio" de la caja del día: lo que sale de la caja de seguridad.
function movimientosCambioDelDia_(data) {
  var lista = [];
  ['mediodia', 'noche'].forEach(function (k) {
    var t = data && data[k];
    (t && (t.movimientosRaw || t.movimientos) || []).forEach(function (m) {
      if (!m || !m.id || m.tipo !== 'ingreso' || m.subtipo !== 'cambio') return;
      var importe = Math.abs(Number(m.importe) || 0);
      var cantidades = cantidadesDeTextoCambio_(m.motivo, importe);
      lista.push({
        id: m.id, movimiento: 'Entregado a la caja del día', responsable: m.responsable,
        importe: -importe,
        cantidades: cantidades && cantidades.map(function (q) { return -q; }),
        detalle: cantidades ? String(m.motivo) : '⚠️ No se sabe qué cambio fue, no se descontó: corregirlo a mano. Dice: ' + String(m.motivo || '(nada)')
      });
    });
  });
  return lista;
}

// Deja la planilla igual a lo que tiene la app para la caja de `fechaISO`
// (igual que sincronizarCajitaDelDia_ en Cajita.gs). `movs` vacío = borrar lo
// que la app había escrito de ese día.
function sincronizarCambioDelDia_(fechaISO, movs) {
  if (fechaISO < CAMBIO_DESDE_) return { ok: true, nada: true };
  var props = PropertiesService.getScriptProperties();
  var marca = 'cambio_' + fechaISO;
  if (!movs.length && !props.getProperty(marca)) return { ok: true, nada: true };

  return conCandadoCajita_(function () {
    var hoja = hojaCambio_();
    var filasDelDia = function () {
      var filaDeId = {}, n = hoja.getLastRow() - CAMBIO_FILA_ENCABEZADOS_;
      if (n > 0) {
        var ids = hoja.getRange(CAMBIO_FILA_ENCABEZADOS_ + 1, CAMBIO_COL_ID_, n, 1).getValues();
        var fechas = hoja.getRange(CAMBIO_FILA_ENCABEZADOS_ + 1, 1, n, 1).getValues();
        for (var i = 0; i < n; i++) {
          var id = String(ids[i][0] || '');
          if (id && textoFechaIds_(fechas[i][0]) === fechaISO) filaDeId[id] = CAMBIO_FILA_ENCABEZADOS_ + 1 + i;
        }
      }
      return filaDeId;
    };

    var actuales = {};
    movs.forEach(function (m) { actuales[m.id] = true; });
    var antes = filasDelDia();
    Object.keys(antes).filter(function (id) { return !actuales[id]; })
      .map(function (id) { return antes[id]; })
      .sort(function (a, b) { return b - a; })
      .forEach(function (r) { hoja.deleteRow(r); });

    var filaDeId = filasDelDia();
    movs.forEach(function (m) {
      var mov = { fecha: fechaISO, movimiento: m.movimiento, cantidades: m.cantidades, importe: m.importe, responsable: m.responsable, detalle: m.detalle };
      escribirFilaCambio_(hoja, filaDeId[m.id] || filaNuevaCambio_(hoja), mov, m.id);
    });

    if (movs.length) props.setProperty(marca, '1'); else props.deleteProperty(marca);
    return { ok: true, escritos: movs.length };
  });
}

// Lo que corre en cada guardado de la app (ver doPost). Si falla, la pantalla
// Cambio lo muestra hasta el próximo guardado que salga bien.
function escribirCambio_(data) {
  var props = PropertiesService.getScriptProperties();
  try {
    var r = sincronizarCambioDelDia_(data.fecha, movimientosCambioDelDia_(data));
    if (!r.nada) props.deleteProperty('cambio_ultimo_error');
    return r;
  } catch (err) {
    try {
      props.setProperty('cambio_ultimo_error', Utilities.formatDate(new Date(), 'Europe/Madrid', 'dd/MM HH:mm') + ' (caja del ' + data.fecha + '): ' + String(err).substring(0, 800));
    } catch (e) {}
    return { ok: false, error: String(err) };
  }
}

// Prueba completa con un día que no existe (GET ?probarCambio=1, ver doGet):
// escribe, lee y deja la planilla como estaba.
function probarCambio_() {
  var fecha = '2099-01-01';
  var foto = function () {
    var d = leerCambio_();
    return { total: d.total, hay: d.stock.map(function (s) { return s.hay; }), filas: d.movimientos.length, primera: d.movimientos[0] };
  };
  var salida = { antes: foto() };
  var cantidades = {};
  CAMBIO_DENOMINACIONES_.forEach(function (d) { cantidades[d.clave] = 2; });
  var agregado = guardarCambio_({ fecha: fecha, responsable: 'Prueba', detalle: 'prueba', cantidades: cantidades });
  salida.conAgregado = foto();
  var movs = movimientosCambioDelDia_({ mediodia: { movimientosRaw: [
    { id: 'prueba-1', tipo: 'ingreso', subtipo: 'cambio', importe: 25, motivo: '1 blister de 1€', responsable: 'Prueba' },
    { id: 'prueba-2', tipo: 'ingreso', subtipo: 'cambio', importe: 20, motivo: '40 monedas de 0,5€', responsable: 'Prueba' },
    { id: 'prueba-3', tipo: 'ingreso', subtipo: 'cambio', importe: 50, motivo: '10 billetes de 5€', responsable: 'Prueba' },
    { id: 'prueba-4', tipo: 'ingreso', subtipo: 'cambio', importe: 7, motivo: 'a mano', responsable: 'Prueba' }
  ] } });
  sincronizarCambioDelDia_(fecha, movs);
  salida.conCaja = foto();
  sincronizarCambioDelDia_(fecha, movs.slice(0, 1));
  salida.quitandoTres = foto();
  sincronizarCambioDelDia_(fecha, []);
  conCandadoCajita_(function () { hojaCambio_().deleteRow(agregado.fila); });
  salida.despues = foto();
  salida.ok = JSON.stringify(salida.antes) === JSON.stringify(salida.despues);
  return salida;
}
