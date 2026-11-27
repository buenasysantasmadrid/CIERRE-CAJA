// ============================================================================
// COMPARATIVA CARNICERÍA desde los albaranes (27/11/2026, Luciana)
// ============================================================================
// Cuando se lee un albarán / ticket / factura del día de Pollería,
// Mercandalucía o Las 3B (la misma lectura de PRECIOS, PreciosLectura.gs),
// sus kilos se PROPONEN para la planilla "Comparativa carnicería" (fila del
// producto, columna del día del documento). NADA se apunta sin confirmar:
//   · Desde la app de caja: después de guardar el gasto sale la ventana
//     "¿Cargamos esto en la Comparativa...?" con cada línea y su producto
//     (lo dudoso en amarillo, "— Elegir —"). Aceptar → se apunta
//     (confirmarComparativa_). Cancelar → no se apunta (se hace a mano).
//   · Si nadie contesta en la app, o el documento llegó solo por email, las
//     líneas esperan en amarillo en la pestaña "POR CONFIRMAR" de la propia
//     planilla (y en PENDIENTES): al elegir el PRODUCTO allí, se apunta
//     (activador alEditarComparativa).
//   · Cómo se reconoce cada línea (lo dijo Luciana) → columna "Sugerido":
//       Pollería: "Filetes" → Filete; "Huevos" → Huevos (se apuntan CARTONES
//         de 30: las docenas × 12 / 30); "Carne de pollo" → Pollo picado.
//       Las 3B: 9,49 €/kg → Picada vacuno (ternera); 5,49 €/kg → Picada cerdo.
//       Mercandalucía: 18,9 €/kg → Lomo bajo (= bife de chorizo); 13,50 €/kg
//         → Filete de cadera o Entraña: SIN sugerencia, se elige siempre (si
//         el texto dice "filetes" o "entraña", no).
//   · La pestaña guarda todas las líneas: la celda de un día y un producto es
//     la suma de sus líneas confirmadas (se escribe encima de lo que hubiera
//     a mano, sin sumar dos veces).
//   · El mismo documento otra vez desde la app (escaneado de nuevo) se vuelve
//     a preguntar; desde el email (el PDF que reemplaza la foto) no se toca.
//   · La factura del mes NO se cuenta (sería contar dos veces).
// ============================================================================

var COMP_DESDE_ = '2026-11-27';
var COMP_PESTANA_ = 'POR CONFIRMAR';
var COMP_VARIOS_ = 'VARIOS (no se apunta)';
var COMP_A_MANO_ = 'NO (se apunta a mano)';
var COMP_CAB_ = ['Fecha', 'Proveedor', 'Nº', 'En el albarán', 'Cantidad', 'Se cobra por', 'Precio €', 'Sugerido', 'PRODUCTO', 'Se apunta', 'clave', 'id', 'apuntado en'];
var COMP_N_ = COMP_CAB_.length;
var COMP_FILA_CAB_ = 2;
var COMP_COL_PRODUCTO_ = 9;
// Proveedor (como su carpeta, normalizado) → bloques de la Comparativa que se ofrecen.
var COMP_PROVEEDORES_ = { POLLERIA: ['polleria'], MERCANDALUCIA: ['carniceria'], 'LAS 3B': ['carniceria', 'polleria'] };

function proveedorComparativa_(proveedor) {
  var k = normalizarClave_(String(proveedor || '').split(' — ')[0]);
  if (k === '3B') k = 'LAS 3B';
  return COMP_PROVEEDORES_[k] ? k : '';
}
function cuentaEnComparativa_(prod) { return !!prod && prod !== COMP_VARIOS_ && prod !== COMP_A_MANO_; }

// El producto de la Comparativa de una línea, o '' si hay que preguntar.
// Los nombres se buscan en las filas de la planilla (si se cambia un
// nombre allí, sigue funcionando).
function productoComparativa_(prov, texto, precio, grupos) {
  var t = sinAcentos_(texto);
  var buscar = function (g, re) {
    var l = grupos[g] || [];
    for (var i = 0; i < l.length; i++) if (re.test(sinAcentos_(l[i].nombre))) return l[i].nombre;
    return '';
  };
  var cerca = function (v, d) { return precio > 0 && Math.abs(precio - v) <= d; };
  // Texto que no dice qué es ("3,05 C PESO", "5.010 kg x 9.49 EURO/kg", "CARNICERIA POLLERIA 3B").
  var generico = !t.replace(/\b(c|peso|kg|kgs|kilo|kilos|eur|euro|euros|x|a|de|carniceria|polleria|3b|carne|np)\b/g, ' ').replace(/\d+/g, ' ').trim();
  if (prov === 'POLLERIA') {
    if (/huevo/.test(t)) return buscar('polleria', /huevo/);
    if (/filete|pechuga/.test(t)) return buscar('polleria', /filete/);
    if (/carne de pollo|picad/.test(t)) return buscar('polleria', /picad/);
    if (/cuarto/.test(t)) return buscar('polleria', /cuarto/);
    return '';
  }
  if (prov === 'MERCANDALUCIA') {
    if (/entra/.test(t)) return buscar('carniceria', /entrana anojo/) || buscar('carniceria', /entrana/);
    if (/filete/.test(t)) return buscar('carniceria', /cadera/) || buscar('carniceria', /filete/);
    if (/lomo|bife/.test(t)) return buscar('carniceria', /lomo/);
    if (generico && cerca(18.4, 1.6)) return buscar('carniceria', /lomo/); // 17,9 / 18,9
    return ''; // 13,50 (filetes o entraña) y lo demás: se pregunta
  }
  if (prov === 'LAS 3B') {
    if (/cerdo/.test(t)) return buscar('carniceria', /picada cerdo/);
    if (/ternera|vacuno/.test(t)) return buscar('carniceria', /picada vacuno/);
    if (generico && cerca(9.49, 0.5)) return buscar('carniceria', /picada vacuno/) || buscar('carniceria', /picada/);
    if (generico && cerca(5.49, 0.3)) return buscar('carniceria', /picada cerdo/);
    return '';
  }
  return '';
}

// Lo que se apunta: kilos, salvo los huevos (cartones de 30).
function cantidadComparativa_(producto, cant, cobro, precio) {
  if (!(cant > 0)) return 0;
  if (/huevo/.test(sinAcentos_(producto))) {
    var c = sinAcentos_(cobro);
    if (/carton|bandeja/.test(c)) return cant;
    if (/^(ud|uds|u|unidad|unidades|huevo|huevos)$/.test(c) && !(precio > 1)) return redondeo_(cant / 30, 2);
    return redondeo_(cant * 12 / 30, 2); // docenas (precio de la docena ~2,60 €)
  }
  return redondeo_(cant, 3);
}

// Los bloques (pollería, carnicería) de la pestaña del mes de `fecha`.
function gruposComparativa_(ss, fecha, cache) {
  var k = fecha.slice(0, 7);
  if (!cache[k]) {
    var hoja = pestanaDelMes_(ss, fecha, 'Comparativa carnicería');
    cache[k] = { hoja: hoja, filas: filasDiarias_(hoja, PLANILLAS_DIARIAS_.carniceria) };
  }
  return cache[k];
}

function opcionesComparativa_(prov, grupos, conAMano) {
  var l = [];
  (COMP_PROVEEDORES_[prov] || []).forEach(function (g) { (grupos[g] || []).forEach(function (x) { l.push(x.nombre); }); });
  l.push(COMP_VARIOS_);
  if (conAMano) l.push(COMP_A_MANO_);
  return l;
}

function hojaPorConfirmar_(ss) {
  var h = ss.getSheetByName(COMP_PESTANA_);
  if (!h) {
    h = ss.insertSheet(COMP_PESTANA_, 0);
    h.getRange(1, 1, 1, 9).merge();
    h.setRowHeight(1, 48);
    h.setFrozenRows(COMP_FILA_CAB_);
  }
  var cab = h.getRange(COMP_FILA_CAB_, 1, 1, COMP_N_).getValues()[0];
  if (cab.join('|') !== COMP_CAB_.join('|')) {
    h.getRange(1, 1).setValue('Kilos de Pollería, Mercandalucía y Las 3B leídos de los albaranes. ' +
      'ARRIBA EN AMARILLO: falta confirmar. Elegir en PRODUCTO qué es y los kilos se apuntan solos en su día. ' +
      'VARIOS o NO = no se apunta. Si algo de abajo está mal, se puede cambiar. Los huevos se apuntan en cartones de 30.')
      .setWrap(true).setFontWeight('bold');
    h.showColumns(1, h.getMaxColumns());
    h.getRange(COMP_FILA_CAB_, 1, 1, h.getMaxColumns()).clearContent();
    h.getRange(COMP_FILA_CAB_, 1, 1, COMP_N_).setValues([COMP_CAB_]).setFontWeight('bold').setBackground('#f4cccc');
    h.setColumnWidth(4, 260); h.setColumnWidth(8, 200); h.setColumnWidth(9, 230);
    h.hideColumns(11, 3);
  }
  return h;
}

function leerPorConfirmar_(h) {
  var n = h.getLastRow() - COMP_FILA_CAB_;
  if (n < 1) return [];
  return h.getRange(COMP_FILA_CAB_ + 1, 1, n, COMP_N_).getValues().filter(function (f) { return f[0] && f[1]; }).map(function (f) {
    return { fecha: fechaIsoDe_(f[0]), proveedor: String(f[1]), numero: String(f[2]), texto: String(f[3]), cant: Number(f[4]) || 0,
      cobro: String(f[5]), precio: Number(f[6]) || 0, sugerido: String(f[7]), producto: String(f[8]), apunta: Number(f[9]) || 0,
      clave: String(f[10]), id: String(f[11]), apuntadoEn: String(f[12]) };
  });
}

// Ordena (lo que falta confirmar arriba) y pinta la pestaña entera.
function pintarPorConfirmar_(ss, h, filas) {
  var cache = {};
  filas.sort(function (a, b) {
    var pa = a.producto ? 1 : 0, pb = b.producto ? 1 : 0;
    return pa - pb || b.fecha.localeCompare(a.fecha) || a.proveedor.localeCompare(b.proveedor) || a.id.localeCompare(b.id);
  });
  var antes = Math.max(h.getLastRow() - COMP_FILA_CAB_, 0);
  if (antes) {
    var viejo = h.getRange(COMP_FILA_CAB_ + 1, 1, antes, COMP_N_);
    viejo.clearContent(); viejo.clearDataValidations(); viejo.setBackground(null);
  }
  if (!filas.length) return;
  var r = h.getRange(COMP_FILA_CAB_ + 1, 1, filas.length, COMP_N_);
  r.setValues(filas.map(function (x) {
    return [fechaComoDate_(x.fecha), x.proveedor, x.numero, x.texto, x.cant || '', x.cobro, x.precio || '', x.sugerido, x.producto,
      cuentaEnComparativa_(x.producto) ? x.apunta : '', x.clave, x.id, x.apuntadoEn];
  }));
  h.getRange(COMP_FILA_CAB_ + 1, 1, filas.length, 1).setNumberFormat('dd/mm/yyyy');
  h.getRange(COMP_FILA_CAB_ + 1, COMP_COL_PRODUCTO_, filas.length, 1).setDataValidations(filas.map(function (x) {
    var g = gruposComparativa_(ss, x.fecha, cache).filas.grupos;
    return [SpreadsheetApp.newDataValidation().requireValueInList(opcionesComparativa_(proveedorComparativa_(x.proveedor), g, true), true).setAllowInvalid(false).build()];
  }));
  r.setBackgrounds(filas.map(function (x) { var c = x.producto ? null : '#fff2cc'; return COMP_CAB_.map(function () { return c; }); }));
}

// Escribe en la Comparativa la suma de las líneas confirmadas de cada (día, producto) tocado.
// tocados: { 'AAAA-MM-DD': { 'Nombre producto': true } }
function recalcularComparativa_(ss, filas, tocados) {
  var cache = {};
  Object.keys(tocados).forEach(function (fecha) {
    var g = gruposComparativa_(ss, fecha, cache);
    var fila = {};
    g.filas.orden.forEach(function (k) { g.filas.grupos[k].forEach(function (s) { fila[s.nombre] = s.fila; }); });
    var col = columnaDelDia_(g.hoja, parseInt(fecha.slice(8, 10), 10));
    Object.keys(tocados[fecha]).forEach(function (prod) {
      if (!cuentaEnComparativa_(prod) || !fila[prod]) return;
      var suma = 0;
      filas.forEach(function (x) { if (x.fecha === fecha && x.producto === prod) suma += x.apunta; });
      suma = Math.round(suma * 100) / 100;
      g.hoja.getRange(fila[prod], col).setValue(suma > 0 ? suma : '');
    });
    filas.forEach(function (x) { if (x.fecha === fecha) x.apuntadoEn = cuentaEnComparativa_(x.producto) ? x.producto : ''; });
  });
}

function tocar_(tocados, fecha, prod) {
  if (!cuentaEnComparativa_(prod)) return;
  (tocados[fecha] = tocados[fecha] || {})[prod] = true;
}

function planillaComparativa_(fecha) { return planillaDelMenu_('COMPARATIVA_CARNE', fecha, 'Comparativa carnicería'); }

// Desde PreciosLectura.gs, con las líneas de un documento ya leído. NO apunta
// nada: deja las líneas sin confirmar en "POR CONFIRMAR" y devuelve la
// propuesta para la ventana de la app.
// lineas: [{ texto, cant, cobro, precio }]; d: { proveedor, fecha, numero, idArchivo, desdeApp }.
function apuntarComparativa_(lineas, d) {
  var prov = proveedorComparativa_(d.proveedor);
  if (!prov || !(String(d.fecha) >= COMP_DESDE_)) return null;
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = planillaComparativa_(d.fecha);
    var h = hojaPorConfirmar_(ss);
    var filas = leerPorConfirmar_(h);
    var clave = prov + '|' + d.fecha + '|' + (claveNumeroPrecio_(d.numero) || d.idArchivo || '');
    var esDeEste = function (x) { return x.clave === clave || (d.idArchivo && x.id.indexOf(d.idArchivo + '#') === 0); };
    // Desde el email (el PDF que reemplaza la foto) o leído otra vez: si ya está, no se toca.
    if (!d.desdeApp && filas.some(esDeEste)) return { ok: true, yaEstaba: true };
    var tocados = {}, elegido = {};
    filas = filas.filter(function (x) {
      if (!esDeEste(x)) return true;
      tocar_(tocados, x.fecha, x.apuntadoEn);
      if (x.producto && x.producto !== COMP_A_MANO_) elegido[sinAcentos_(x.texto) + '|' + x.cant] = x.producto;
      return false;
    });
    var g = gruposComparativa_(ss, d.fecha, {}).filas.grupos;
    var opciones = opcionesComparativa_(prov, g, false);
    var nuevas = [];
    lineas.forEach(function (l, i) {
      if (!String(l.texto || '').trim() || !(l.cant > 0)) return;
      var sug = elegido[sinAcentos_(l.texto) + '|' + l.cant] || productoComparativa_(prov, l.texto, l.precio, g);
      if (sug && opciones.indexOf(sug) === -1) sug = '';
      nuevas.push({ fecha: d.fecha, proveedor: String(d.proveedor).split(' — ')[0], numero: String(d.numero || ''), texto: String(l.texto).trim(),
        cant: l.cant, cobro: String(l.cobro || ''), precio: l.precio > 0 ? redondeo_(l.precio, 2) : 0, sugerido: sug, producto: '',
        apunta: 0, clave: clave, id: (d.idArchivo || clave) + '#' + i, apuntadoEn: '' });
    });
    filas = filas.concat(nuevas);
    if (Object.keys(tocados).length) recalcularComparativa_(ss, filas, tocados);
    pintarPorConfirmar_(ss, h, filas);
    try { asegurarActivadorComparativa_(ss.getId()); } catch (err) { Logger.log('Comparativa, activador: ' + err); }
    if (!nuevas.length) return { ok: true, lineas: 0 };
    return {
      ok: true, lineas: nuevas.length,
      propuesta: {
        clave: clave, fecha: d.fecha, proveedor: nuevas[0].proveedor, numero: nuevas[0].numero, opciones: opciones,
        lineas: nuevas.map(function (x) {
          return { id: x.id, texto: x.texto, cant: x.cant, cobro: x.cobro, precio: x.precio, sugerido: x.sugerido };
        })
      }
    };
  } finally {
    lock.releaseLock();
  }
}

// La respuesta de la ventana de la app: data = { clave, fecha, cancelar, lineas: [{ id, producto }] }.
function confirmarComparativa_(data) {
  var fecha = String(data.fecha || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha) || !data.clave) return { ok: false, error: 'Faltan datos' };
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = planillaComparativa_(fecha);
    var h = hojaPorConfirmar_(ss);
    var filas = leerPorConfirmar_(h);
    var elegidos = {};
    (data.lineas || []).forEach(function (l) { elegidos[String(l.id)] = String(l.producto || ''); });
    var cache = {}, tocados = {}, apuntadas = 0, faltan = 0;
    filas.forEach(function (x) {
      if (x.clave !== String(data.clave)) return;
      var nuevo = data.cancelar ? COMP_A_MANO_ : (elegidos[x.id] || '');
      var g = gruposComparativa_(ss, x.fecha, cache).filas.grupos;
      if (nuevo && opcionesComparativa_(proveedorComparativa_(x.proveedor), g, true).indexOf(nuevo) === -1) nuevo = '';
      if (!nuevo) { faltan++; return; }
      tocar_(tocados, x.fecha, x.apuntadoEn);
      x.producto = nuevo;
      x.apunta = cuentaEnComparativa_(nuevo) ? cantidadComparativa_(nuevo, x.cant, x.cobro, x.precio) : 0;
      tocar_(tocados, x.fecha, nuevo);
      if (cuentaEnComparativa_(nuevo)) apuntadas++;
    });
    if (Object.keys(tocados).length) recalcularComparativa_(ss, filas, tocados);
    pintarPorConfirmar_(ss, h, filas);
    return { ok: true, apuntadas: apuntadas, faltan: faltan, cancelado: !!data.cancelar };
  } finally {
    lock.releaseLock();
  }
}

function asegurarActivadorComparativa_(idPlanilla) {
  var yaEsta = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === 'alEditarComparativa' && t.getTriggerSourceId() === idPlanilla;
  });
  if (!yaEsta) ScriptApp.newTrigger('alEditarComparativa').forSpreadsheet(idPlanilla).onEdit().create();
}

// Al elegir (o cambiar) un PRODUCTO en "POR CONFIRMAR": se apuntan los kilos.
function alEditarComparativa(e) {
  try {
    if (!e || !e.range) return;
    var h = e.range.getSheet();
    if (h.getName() !== COMP_PESTANA_) return;
    var c1 = e.range.getColumn(), c2 = c1 + e.range.getNumColumns() - 1;
    if (c1 > COMP_COL_PRODUCTO_ || c2 < COMP_COL_PRODUCTO_ || e.range.getLastRow() <= COMP_FILA_CAB_) return;
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      var ss = h.getParent();
      var filas = leerPorConfirmar_(h);
      var tocados = {};
      filas.forEach(function (x) {
        var cuenta = cuentaEnComparativa_(x.producto);
        x.apunta = cuenta ? cantidadComparativa_(x.producto, x.cant, x.cobro, x.precio) : 0;
        if ((x.apuntadoEn || '') !== (cuenta ? x.producto : '')) { tocar_(tocados, x.fecha, x.apuntadoEn); tocar_(tocados, x.fecha, x.producto); }
      });
      if (Object.keys(tocados).length) recalcularComparativa_(ss, filas, tocados);
      // Sin reordenar (para no mover las filas mientras se eligen): solo J, M y el color.
      var porId = {};
      filas.forEach(function (x) { porId[x.id] = x; });
      var n = h.getLastRow() - COMP_FILA_CAB_;
      h.getRange(COMP_FILA_CAB_ + 1, 1, n, COMP_N_).getValues().forEach(function (f, i) {
        var x = porId[String(f[11])];
        if (!f[0] || !x) return;
        var fila = COMP_FILA_CAB_ + 1 + i;
        h.getRange(fila, 10).setValue(cuentaEnComparativa_(x.producto) ? x.apunta : '');
        h.getRange(fila, 13).setValue(x.apuntadoEn);
        h.getRange(fila, 1, 1, COMP_N_).setBackground(x.producto ? null : '#fff2cc');
      });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    Logger.log('Comparativa, al editar: ' + err);
  }
}

// Para PENDIENTES: las líneas sin confirmar (de la planilla de este año y,
// en enero, también la del anterior).
function carnePorConfirmar_() {
  var lista = [];
  var hoy = hoyISO_();
  [hoy, (Number(hoy.slice(0, 4)) - 1) + '-12-31'].forEach(function (f, i) {
    if (i === 1 && hoy.slice(5, 7) !== '01') return;
    var id = buscarEnIndice_('COMPARATIVA_CARNE', periodoAnual_(f));
    if (!id) return;
    var ss = SpreadsheetApp.openById(id);
    var h = ss.getSheetByName(COMP_PESTANA_);
    if (!h) return;
    var url = ss.getUrl() + '#gid=' + h.getSheetId();
    leerPorConfirmar_(h).forEach(function (x) {
      if (x.producto) return;
      lista.push({ proveedor: x.proveedor, fecha: x.fecha, numero: x.numero, url: url,
        que: x.texto + ' — ' + x.cant + ' ' + x.cobro + (x.precio ? ' a ' + String(x.precio).replace('.', ',') + ' €' : '') +
          (x.sugerido ? ' (¿' + x.sugerido + '?)' : '') + ': confirmar el producto en la pestaña POR CONFIRMAR de la Comparativa',
        clave: 'carne|' + x.id });
    });
  });
  return lista;
}
