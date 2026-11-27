// ============================================================================
// COMPARATIVA CARNICERÍA desde los albaranes (27/11/2026, Luciana)
// ============================================================================
// Cuando se lee un albarán / ticket / factura del día de Pollería,
// Mercandalucía o Las 3B (la misma lectura de PRECIOS, PreciosLectura.gs),
// sus kilos se apuntan solos en la planilla "Comparativa carnicería", en la
// fila del producto y la columna del día del documento.
//   · Cómo se reconoce cada línea (lo dijo Luciana):
//       Pollería: "Filetes" → Filete; "Huevos" → Huevos (se apuntan CARTONES
//         de 30: las docenas × 12 / 30); "Carne de pollo" → Pollo picado.
//       Las 3B: 9,49 €/kg → Picada vacuno (ternera); 5,49 €/kg → Picada cerdo.
//       Mercandalucía: 18,9 €/kg → Lomo bajo (= bife de chorizo); 13,50 €/kg
//         → Filete de cadera o Entraña: SE PREGUNTA SIEMPRE (si el texto dice
//         "filetes" o "entraña", no).
//     Lo que no se reconoce también se pregunta.
//   · Las preguntas van en la pestaña "POR CONFIRMAR" de la propia planilla,
//     arriba y en amarillo, con un desplegable con los productos de ese
//     proveedor y "VARIOS" (no se apunta). Al elegir, los kilos se apuntan
//     solos (activador alEditarComparativa). También sale en PENDIENTES.
//   · Esa pestaña guarda TODAS las líneas leídas: la celda de un día y un
//     producto es la suma de sus líneas (se escribe encima de lo que hubiera
//     a mano, sin sumar dos veces). Si un producto está mal, se cambia ahí.
//   · El mismo documento leído otra vez (p. ej. el PDF del email después de
//     la foto) reemplaza sus líneas y conserva lo que ya se eligió.
//   · La factura del mes NO se cuenta (sería contar dos veces).
//   · Solo desde COMP_DESDE_: lo anterior se apuntó a mano y no se toca.
// ============================================================================

var COMP_DESDE_ = '2026-11-28';
var COMP_PESTANA_ = 'POR CONFIRMAR';
var COMP_VARIOS_ = 'VARIOS (no se apunta)';
var COMP_CAB_ = ['Fecha', 'Proveedor', 'Nº', 'En el albarán', 'Cantidad', 'Se cobra por', 'Precio €', 'PRODUCTO', 'Se apunta', 'clave', 'archivo', 'apuntado en'];
var COMP_N_ = COMP_CAB_.length;
var COMP_FILA_CAB_ = 2;
var COMP_COL_PRODUCTO_ = 8;
// Proveedor (como su carpeta, normalizado) → bloques de la Comparativa que se ofrecen.
var COMP_PROVEEDORES_ = { POLLERIA: ['polleria'], MERCANDALUCIA: ['carniceria'], 'LAS 3B': ['carniceria', 'polleria'] };

function proveedorComparativa_(proveedor) {
  var k = normalizarClave_(String(proveedor || '').split(' — ')[0]);
  if (k === '3B') k = 'LAS 3B';
  return COMP_PROVEEDORES_[k] ? k : '';
}

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

function opcionesComparativa_(prov, grupos) {
  var l = [];
  (COMP_PROVEEDORES_[prov] || []).forEach(function (g) { (grupos[g] || []).forEach(function (x) { l.push(x.nombre); }); });
  l.push(COMP_VARIOS_);
  return l;
}

function hojaPorConfirmar_(ss) {
  var h = ss.getSheetByName(COMP_PESTANA_);
  if (!h) {
    h = ss.insertSheet(COMP_PESTANA_, 0);
    h.getRange(1, 1).setValue('Kilos de Pollería, Mercandalucía y Las 3B leídos de los albaranes (desde el 28/11/2026). ' +
      'ARRIBA EN AMARILLO: elegir en PRODUCTO qué es y los kilos se apuntan solos en su día. VARIOS = no se apunta. ' +
      'Si algo de abajo está mal, se puede cambiar el producto. Los huevos se apuntan en cartones de 30.')
      .setWrap(true).setFontWeight('bold');
    h.getRange(1, 1, 1, 9).merge();
    h.setRowHeight(1, 48);
    h.getRange(COMP_FILA_CAB_, 1, 1, COMP_N_).setValues([COMP_CAB_]).setFontWeight('bold').setBackground('#f4cccc');
    h.setFrozenRows(COMP_FILA_CAB_);
    h.setColumnWidth(4, 260); h.setColumnWidth(8, 230);
    h.hideColumns(10, 3);
  }
  return h;
}

function leerPorConfirmar_(h) {
  var n = h.getLastRow() - COMP_FILA_CAB_;
  if (n < 1) return [];
  return h.getRange(COMP_FILA_CAB_ + 1, 1, n, COMP_N_).getValues().filter(function (f) { return f[0] && f[1]; }).map(function (f) {
    return { fecha: fechaIsoDe_(f[0]), proveedor: String(f[1]), numero: String(f[2]), texto: String(f[3]), cant: Number(f[4]) || 0,
      cobro: String(f[5]), precio: Number(f[6]) || 0, producto: String(f[7]), apunta: Number(f[8]) || 0,
      clave: String(f[9]), archivo: String(f[10]), apuntadoEn: String(f[11]) };
  });
}

// Lee de nuevo la pestaña, la ordena (lo que falta elegir arriba) y pinta.
function pintarPorConfirmar_(ss, h, filas) {
  var cache = {};
  filas.sort(function (a, b) {
    var pa = a.producto ? 1 : 0, pb = b.producto ? 1 : 0;
    return pa - pb || b.fecha.localeCompare(a.fecha) || a.proveedor.localeCompare(b.proveedor) || a.numero.localeCompare(b.numero);
  });
  var antes = Math.max(h.getLastRow() - COMP_FILA_CAB_, 0);
  if (antes) {
    var viejo = h.getRange(COMP_FILA_CAB_ + 1, 1, antes, COMP_N_);
    viejo.clearContent(); viejo.clearDataValidations(); viejo.setBackground(null);
  }
  if (!filas.length) return;
  var r = h.getRange(COMP_FILA_CAB_ + 1, 1, filas.length, COMP_N_);
  r.setValues(filas.map(function (x) {
    return [fechaComoDate_(x.fecha), x.proveedor, x.numero, x.texto, x.cant || '', x.cobro, x.precio || '', x.producto,
      x.producto && x.producto !== COMP_VARIOS_ ? x.apunta : '', x.clave, x.archivo, x.apuntadoEn];
  }));
  h.getRange(COMP_FILA_CAB_ + 1, 1, filas.length, 1).setNumberFormat('dd/mm/yyyy');
  h.getRange(COMP_FILA_CAB_ + 1, COMP_COL_PRODUCTO_, filas.length, 1).setDataValidations(filas.map(function (x) {
    var g = gruposComparativa_(ss, x.fecha, cache).filas.grupos;
    var prov = proveedorComparativa_(x.proveedor);
    return [SpreadsheetApp.newDataValidation().requireValueInList(opcionesComparativa_(prov, g), true).setAllowInvalid(false).build()];
  }));
  r.setBackgrounds(filas.map(function (x) { var c = x.producto ? null : '#fff2cc'; return COMP_CAB_.map(function () { return c; }); }));
}

// Escribe en la Comparativa la suma de las líneas de cada (día, producto) tocado.
// tocados: { 'AAAA-MM-DD': { 'Nombre producto': true } }
function recalcularComparativa_(ss, filas, tocados) {
  var cache = {};
  Object.keys(tocados).forEach(function (fecha) {
    var g = gruposComparativa_(ss, fecha, cache);
    var fila = {};
    g.filas.orden.forEach(function (k) { g.filas.grupos[k].forEach(function (s) { fila[s.nombre] = s.fila; }); });
    var col = columnaDelDia_(g.hoja, parseInt(fecha.slice(8, 10), 10));
    Object.keys(tocados[fecha]).forEach(function (prod) {
      if (!prod || prod === COMP_VARIOS_ || !fila[prod]) return;
      var suma = 0;
      filas.forEach(function (x) { if (x.fecha === fecha && x.producto === prod) suma += x.apunta; });
      suma = Math.round(suma * 100) / 100;
      g.hoja.getRange(fila[prod], col).setValue(suma > 0 ? suma : '');
    });
    filas.forEach(function (x) { if (x.fecha === fecha) x.apuntadoEn = x.producto && x.producto !== COMP_VARIOS_ ? x.producto : ''; });
  });
}

function tocar_(tocados, fecha, prod) {
  if (!prod || prod === COMP_VARIOS_) return;
  (tocados[fecha] = tocados[fecha] || {})[prod] = true;
}

// Desde PreciosLectura.gs, con las líneas de un documento ya leído.
// lineas: [{ texto, cant, cobro, precio }]; d: { proveedor, fecha, numero, idArchivo }.
function apuntarComparativa_(lineas, d) {
  var prov = proveedorComparativa_(d.proveedor);
  if (!prov || !(String(d.fecha) >= COMP_DESDE_)) return null;
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = planillaDelMenu_('COMPARATIVA_CARNE', d.fecha, 'Comparativa carnicería');
    var h = hojaPorConfirmar_(ss);
    var filas = leerPorConfirmar_(h);
    var clave = prov + '|' + d.fecha + '|' + (claveNumeroPrecio_(d.numero) || d.idArchivo || '');
    var tocados = {}, elegido = {};
    // Las líneas de este documento de otra vez: fuera (se recuerda lo elegido).
    filas = filas.filter(function (x) {
      var esta = x.clave === clave || (d.idArchivo && x.archivo === d.idArchivo);
      if (esta) {
        tocar_(tocados, x.fecha, x.apuntadoEn);
        if (x.producto) elegido[sinAcentos_(x.texto) + '|' + x.cant] = x.producto;
      }
      return !esta;
    });
    var cache = {}, preguntas = 0;
    var g = gruposComparativa_(ss, d.fecha, cache).filas.grupos;
    var opciones = opcionesComparativa_(prov, g);
    lineas.forEach(function (l) {
      if (!String(l.texto || '').trim() || !(l.cant > 0)) return;
      var prod = elegido[sinAcentos_(l.texto) + '|' + l.cant] || productoComparativa_(prov, l.texto, l.precio, g);
      if (prod && opciones.indexOf(prod) === -1) prod = '';
      if (!prod) preguntas++;
      filas.push({ fecha: d.fecha, proveedor: String(d.proveedor).split(' — ')[0], numero: String(d.numero || ''), texto: String(l.texto).trim(),
        cant: l.cant, cobro: String(l.cobro || ''), precio: l.precio > 0 ? redondeo_(l.precio, 2) : 0, producto: prod,
        apunta: prod ? cantidadComparativa_(prod, l.cant, l.cobro, l.precio) : 0, clave: clave, archivo: d.idArchivo || '', apuntadoEn: '' });
      tocar_(tocados, d.fecha, prod);
    });
    recalcularComparativa_(ss, filas, tocados);
    pintarPorConfirmar_(ss, h, filas);
    try { asegurarActivadorComparativa_(ss.getId()); } catch (err) { Logger.log('Comparativa, activador: ' + err); }
    return { ok: true, lineas: lineas.length, preguntas: preguntas };
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
        var deVerdad = x.producto && x.producto !== COMP_VARIOS_;
        x.apunta = deVerdad ? cantidadComparativa_(x.producto, x.cant, x.cobro, x.precio) : 0;
        if ((x.apuntadoEn || '') !== (deVerdad ? x.producto : '')) { tocar_(tocados, x.fecha, x.apuntadoEn); tocar_(tocados, x.fecha, x.producto); }
      });
      if (Object.keys(tocados).length) recalcularComparativa_(ss, filas, tocados);
      // Sin reordenar (para no mover las filas mientras se eligen): solo columnas I y L y el color.
      var n = h.getLastRow() - COMP_FILA_CAB_;
      var v = h.getRange(COMP_FILA_CAB_ + 1, 1, n, COMP_N_).getValues();
      var porClave = {};
      filas.forEach(function (x) { porClave[x.clave + '|' + x.texto + '|' + x.cant + '|' + x.producto] = x; });
      v.forEach(function (f, i) {
        if (!f[0]) return;
        var x = porClave[f[9] + '|' + f[3] + '|' + (Number(f[4]) || 0) + '|' + f[7]];
        if (!x) return;
        var fila = COMP_FILA_CAB_ + 1 + i;
        h.getRange(fila, 9).setValue(x.producto && x.producto !== COMP_VARIOS_ ? x.apunta : '');
        h.getRange(fila, 12).setValue(x.apuntadoEn);
        h.getRange(fila, 1, 1, COMP_N_).setBackground(x.producto ? null : '#fff2cc');
      });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    Logger.log('Comparativa, al editar: ' + err);
  }
}

// Para PENDIENTES: las líneas sin elegir (de las planillas de este año y el anterior).
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
        que: x.texto + ' — ' + x.cant + ' ' + x.cobro + (x.precio ? ' a ' + String(x.precio).replace('.', ',') + ' €' : '') + ': elegir el producto en la pestaña POR CONFIRMAR de la Comparativa',
        clave: 'carne|' + x.clave + '|' + sinAcentos_(x.texto) + '|' + x.cant });
    });
  });
  return lista;
}
