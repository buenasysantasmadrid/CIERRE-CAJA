// ============================================================================
// ALBARANES: pestaña "CÓMO TRABAJA CADA PROVEEDOR" (24/11/2026, Luciana)
// ============================================================================
// Una fila por proveedor (cada pestaña de ALBARANES) con dos desplegables
// que Luciana puede cambiar:
//   · QUÉ ENTRA EN LA CAJA: qué papel se carga cuando llega la compra.
//   · CÓMO LLEGA LA FACTURA: qué pasa después.
// La columna "QUÉ HACE EL SISTEMA" la escribe el script (textoTratamiento_)
// y "QUIÉN LO PUSO" dice si es una suposición de Claude o lo eligió ella.
// Lo elegido se guarda en la propiedad "tratamientos" ({clave: {c, f}}) y,
// al cargar el script, cambia las listas que usan los demás archivos:
//   INFORME_MES_FACTURA_MENSUAL_ (Albaranes.gs): f = MENSUAL.
//   FALTAN_ALBARAN_LUEGO_FACTURA_ (FacturasQueFaltan.gs): f = POR_ALBARAN.
// Y además (Bandejas.gs, DocumentosAlbaranes.gs):
//   f = SIN_GASTO: la factura del email se carga enseguida, sin buscar
//     albarán ni esperar.
//   c = DEFINITIVA o f = NO_LLEGA: nunca "falta factura".
// Se guarda al editar la pestaña (alEditarAlbaranes) y cada 2 horas
// (actualizarDocumentosAlbaranes), que también añade los proveedores nuevos.
// Claude: GET ?tratamientos=AAAA crea/pone al día la pestaña.
// ============================================================================

var TRAT_PESTANA_ = 'CÓMO TRABAJA CADA PROVEEDOR';
ALBARANES_PESTANAS_EXCLUIDAS_.push(TRAT_PESTANA_);
var TRAT_PROP_ = 'tratamientos';
var TRAT_QUITADOS_ = ['ALIMENTOS AUSTRALES', 'GABARRERA']; // ya no se usan (Gabarrera = Bailandera)
var TRAT_FILA_CAB_ = 4;

var TRAT_CAJA_ = {
  ALBARAN: 'Albarán o ticket',
  FACTURA: 'Factura (después llega otra por email)',
  DEFINITIVA: 'Factura definitiva (no llega otra)',
  GASTO: 'Solo el gasto, sin papel',
  NADA: 'Nada: no pasa por la caja'
};
var TRAT_FACTURA_ = {
  MENSUAL: 'A fin de mes, una factura con todos los albaranes',
  TRIMESTRAL: 'Cada 3 meses, una factura con todos los albaranes',
  POR_ALBARAN: 'Poco a poco, una factura por cada albarán',
  DE_FACTURA: 'Llega la factura de la factura',
  SIN_GASTO: 'Llega la factura sin que se haya cargado el gasto',
  NO_LLEGA: 'No llega ninguna factura'
};

// Lo guardado: { CLAVE: { c: 'ALBARAN', f: 'MENSUAL' } }.
var TRATAMIENTOS_ = (function () {
  try {
    var t = JSON.parse(PropertiesService.getScriptProperties().getProperty(TRAT_PROP_) || 'null');
    if (!t || !Object.keys(t).length) return {};
    var mensual = [], porAlbaran = [];
    Object.keys(t).forEach(function (k) {
      if (t[k].f === 'MENSUAL' || t[k].f === 'TRIMESTRAL') mensual.push(k);
      if (t[k].f === 'POR_ALBARAN' && (t[k].c === 'ALBARAN' || t[k].c === 'GASTO') && BANDEJA_PROVEEDORES_GENERICOS_.indexOf(k) === -1) porAlbaran.push(k);
    });
    INFORME_MES_FACTURA_MENSUAL_.length = 0;
    mensual.forEach(function (k) { INFORME_MES_FACTURA_MENSUAL_.push(k); });
    FACTURA_TRIMESTRAL_.length = 0;
    Object.keys(t).forEach(function (k) { if (t[k].f === 'TRIMESTRAL') FACTURA_TRIMESTRAL_.push(k); });
    FALTAN_ALBARAN_LUEGO_FACTURA_.length = 0;
    porAlbaran.forEach(function (k) { FALTAN_ALBARAN_LUEGO_FACTURA_.push(k); });
    return t;
  } catch (e) { return {}; }
})();

function tratamientoDe_(nombre) { return TRATAMIENTOS_[normalizarClave_(nombre)] || null; }

// Lo que hace el sistema con cada combinación, en palabras.
function textoTratamiento_(c, f) {
  var caja = {
    ALBARAN: 'En la caja se carga el albarán o ticket (foto). ',
    FACTURA: 'En la caja se carga la foto de la factura. ',
    DEFINITIVA: 'En la caja se carga la foto de la factura, y esa es la buena. ',
    GASTO: 'En la caja se carga solo el gasto, sin foto. ',
    NADA: 'No se carga nada en la caja. '
  }[c] || '';
  var fac = {
    MENSUAL: 'A fin de mes llega UNA factura con todos los albaranes: guardo el PDF como la factura del mes y comparo su total con la suma de los albaranes (si no coincide, sale en PENDIENTES). Si a principios del mes siguiente no llegó, sale en PENDIENTES.',
    POR_ALBARAN: 'Después va llegando la factura de cada albarán: la busco por importe y fecha y el albarán pasa a factura. Los albaranes que siguen sin factura salen en PENDIENTES.',
    DE_FACTURA: 'Después llega por email la misma factura escaneada: cambio la foto por el PDF del email y no la cargo dos veces.',
    TRIMESTRAL: 'Cada 3 meses llega UNA factura con todos los albaranes del trimestre: guardo el PDF como la factura y no la espero hasta que acaba el trimestre (marzo, junio, septiembre, diciembre). Si a principios del mes siguiente no llegó, sale en PENDIENTES.',
    SIN_GASTO: 'La factura llega sin que haya nada en la caja: la cargo enseguida en su pestaña (como NO PAGADO; si la factura dice tarjeta o domiciliación, como pagada). Si un mes no llega cuando suele llegar, sale en PENDIENTES.',
    NO_LLEGA: 'No llega ninguna factura: no la espero y nunca sale como "falta factura".'
  }[f] || '';
  if (c === 'DEFINITIVA' && f !== 'DE_FACTURA') fac = 'No llega otra factura: la foto es la definitiva. ' + (f === 'NO_LLEGA' ? '' : fac);
  return (caja + fac).trim() || 'Elegir las dos columnas.';
}

function codigoDe_(lista, texto) {
  var t = String(texto || '');
  for (var k in lista) if (lista[k] === t) return k;
  return '';
}

// Suposición para un proveedor nuevo, por lo que hay en su pestaña
// (columna DOCUMENTO).
function suponerTratamiento_(clave, docs) {
  if (['POLLERIA', 'LA JAULA', 'MERCANDALUCIA', 'LAS 3B'].indexOf(clave) > -1) return { c: 'ALBARAN', f: 'MENSUAL' };
  if (['ATLANTA', 'DISBESA', 'IBEROLACTAL', 'CBG', 'DIA CARNE', 'SUPER'].indexOf(clave) > -1) return { c: 'ALBARAN', f: 'POR_ALBARAN' };
  if (clave === 'ENVAPRO') return { c: 'FACTURA', f: 'DE_FACTURA' };
  if (['SERVICIOS', 'GLOVO', 'FC VYR'].indexOf(clave) > -1) return { c: 'NADA', f: 'SIN_GASTO' };
  if (['VARIOS', 'EXTRAS'].indexOf(clave) > -1) return { c: 'GASTO', f: 'NO_LLEGA' };
  var d = docs || {};
  var alb = d['ALBARÁN'] || 0, fac = (d.FACTURA || 0) + (d.TICKET || 0), nada = d[''] || 0;
  if (alb + fac + nada < 3) return { c: '', f: '' }; // casi sin datos: que lo elija Luciana
  if (alb > fac && alb >= nada) return { c: 'ALBARAN', f: 'POR_ALBARAN' };
  if (fac >= nada && fac) return { c: 'DEFINITIVA', f: 'NO_LLEGA' };
  return { c: 'GASTO', f: 'NO_LLEGA' };
}

// Crea o pone al día la pestaña (añade los proveedores nuevos) y guarda lo
// elegido. Devuelve cuántos hay.
function actualizarTratamientos_(ss, rehacer) {
  var h = ss.getSheetByName(TRAT_PESTANA_);
  if (h && rehacer) { ss.deleteSheet(h); h = null; }
  var nueva = !h;
  if (nueva) h = ss.insertSheet(TRAT_PESTANA_, 1);
  var filas = {};
  if (!nueva && h.getLastRow() > TRAT_FILA_CAB_) {
    h.getRange(TRAT_FILA_CAB_ + 1, 1, h.getLastRow() - TRAT_FILA_CAB_, 5).getValues().forEach(function (v) {
      if (v[0]) filas[normalizarClave_(v[0])] = { nombre: v[0], c: codigoDe_(TRAT_CAJA_, v[1]), f: codigoDe_(TRAT_FACTURA_, v[2]), quien: v[4] };
    });
  }
  // Proveedores que ya no se usan (25/11/2026, Luciana): su pestaña sigue en
  // ALBARANES con lo viejo, pero no salen aquí.
  var quitados = false;
  TRAT_QUITADOS_.forEach(function (k) { if (filas[k]) { delete filas[k]; quitados = true; } });
  // Las pestañas de proveedor de la planilla (las que no están, se añaden).
  var docs = null;
  ss.getSheets().forEach(function (s) {
    var nombre = s.getName(), clave = normalizarClave_(nombre);
    if (ALBARANES_PESTANAS_EXCLUIDAS_.indexOf(nombre) > -1 || nombre === INFORME_MES_PESTANA_ || filas[clave]) return;
    if (/^(FALTA FACTURA|PENDIENTES|NOO|VINO)$/.test(clave)) return;
    if (TRAT_QUITADOS_.indexOf(clave) > -1) return;
    if (!docs) docs = cuentaDocumentosPorPestana_(ss);
    var sup = suponerTratamiento_(clave, docs[nombre]);
    filas[clave] = { nombre: nombre, c: sup.c, f: sup.f, quien: sup.c ? 'Claude (suposición: revisar)' : 'Falta elegir' };
  });
  var lista = Object.keys(filas).sort().map(function (k) { return filas[k]; });
  // Sin proveedores nuevos: no se reescribe (por si Luciana la está mirando).
  if (!nueva && !docs && !quitados) { guardarTratamientos_(lista); return { ok: true, proveedores: lista.length, sinCambios: true }; }

  h.getRange(1, 1, h.getMaxRows(), h.getMaxColumns()).clearDataValidations();
  h.clear();
  h.getRange(1, 1).setValue('Cómo trabaja cada proveedor').setFontWeight('bold').setFontSize(14);
  h.getRange(2, 1).setValue('Elige en las columnas B y C cómo llegan los papeles de cada proveedor. La columna D dice qué hará el sistema con eso (se escribe sola). ' +
    'Las filas en amarillo las supuse yo (Claude): revísalas y, si están bien, vuelve a elegir lo mismo para que pase a "Luciana". Las rojas no las sé: elígelas tú.').setFontColor('#666').setWrap(true);
  h.getRange(TRAT_FILA_CAB_, 1, 1, 5).setValues([['PROVEEDOR', 'QUÉ ENTRA EN LA CAJA', 'CÓMO LLEGA LA FACTURA', 'QUÉ HACE EL SISTEMA', 'QUIÉN LO PUSO']])
    .setFontWeight('bold').setBackground('#2f4f3a').setFontColor('#fff');
  var n = lista.length;
  if (n) {
    var ini = TRAT_FILA_CAB_ + 1;
    h.getRange(ini, 1, n, 5).setValues(lista.map(function (x) {
      return [x.nombre, TRAT_CAJA_[x.c] || '', TRAT_FACTURA_[x.f] || '', textoTratamiento_(x.c, x.f), x.quien || ''];
    }));
    var valores = function (o) { return Object.keys(o).map(function (k) { return o[k]; }); };
    h.getRange(ini, 2, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(valores(TRAT_CAJA_), true).setAllowInvalid(false).build());
    h.getRange(ini, 3, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(valores(TRAT_FACTURA_), true).setAllowInvalid(false).build());
    h.getRange(ini, 1, n, 1).setFontWeight('bold');
    h.getRange(ini, 4, n, 1).setWrap(true).setFontColor('#444');
    h.getRange(ini, 1, n, 5).setBackgrounds(lista.map(function (x) {
      var c = /Claude/.test(x.quien) ? '#fff2cc' : /Falta/.test(x.quien) ? '#fde8e4' : null;
      return [c, c, c, c, c];
    }));
    h.getRange(ini, 1, n, 5).setVerticalAlignment('middle');
  }
  h.setFrozenRows(TRAT_FILA_CAB_);
  h.setColumnWidth(1, 170); h.setColumnWidth(2, 260); h.setColumnWidth(3, 330); h.setColumnWidth(4, 520); h.setColumnWidth(5, 190);
  h.setRowHeight(2, 40);
  guardarTratamientos_(lista);
  try { instalarTriggerInformeMes_(ss.getId()); } catch (err) { Logger.log(err); }
  return { ok: true, proveedores: n, supuestos: lista.filter(function (x) { return /Claude/.test(x.quien); }).map(function (x) { return x.nombre + ': ' + TRAT_CAJA_[x.c] + ' / ' + TRAT_FACTURA_[x.f]; }) };
}

function guardarTratamientos_(lista) {
  var t = {};
  lista.forEach(function (x) { if (x.c && x.f) t[normalizarClave_(x.nombre)] = { c: x.c, f: x.f }; });
  PropertiesService.getScriptProperties().setProperty(TRAT_PROP_, JSON.stringify(t));
}

// Cuántas filas de cada tipo (columna DOCUMENTO) tiene cada pestaña, desde
// septiembre (antes no había PDF en Drive y no dice nada).
function cuentaDocumentosPorPestana_(ss) {
  var r = {};
  ss.getSheets().forEach(function (h) {
    var nombre = h.getName();
    if (ALBARANES_PESTANAS_EXCLUIDAS_.indexOf(nombre) > -1) return;
    var c = r[nombre] = {};
    try {
      var filas = filasDelAnioAlbaranes_(h);
      if (!filas.length) return;
      var j = h.getRange(1, ALBARANES_COL_DOC_, h.getLastRow(), 1).getValues();
      filas.forEach(function (f) {
        if (f.mes < 8 || !f.importe) return;
        var v = String(j[f.fila - 1][0] || '');
        c[v] = (c[v] || 0) + 1;
      });
    } catch (e) { /* sin bloques */ }
  });
  return r;
}

// Desde alEditarAlbaranes: al cambiar B o C, se reescribe D, "Luciana" en E
// y se guarda.
function alEditarTratamientos_(e) {
  var h = e.range.getSheet();
  var fila = e.range.getRow(), col = e.range.getColumn();
  if (fila <= TRAT_FILA_CAB_ || col + e.range.getNumColumns() - 1 < 2 || col > 3) return;
  for (var r = fila; r < fila + e.range.getNumRows(); r++) {
    var v = h.getRange(r, 1, 1, 3).getValues()[0];
    if (!v[0]) continue;
    var c = codigoDe_(TRAT_CAJA_, v[1]), f = codigoDe_(TRAT_FACTURA_, v[2]);
    h.getRange(r, 4, 1, 2).setValues([[textoTratamiento_(c, f), 'Luciana']]);
    h.getRange(r, 1, 1, 5).setBackground(null);
  }
  var lista = h.getRange(TRAT_FILA_CAB_ + 1, 1, Math.max(1, h.getLastRow() - TRAT_FILA_CAB_), 3).getValues()
    .filter(function (v) { return v[0]; })
    .map(function (v) { return { nombre: v[0], c: codigoDe_(TRAT_CAJA_, v[1]), f: codigoDe_(TRAT_FACTURA_, v[2]) }; });
  guardarTratamientos_(lista);
}

// Solo Claude (GET ?tratamientos=AAAA&poner=CLAVE:C:F,CLAVE:C:F): pone lo que
// Luciana le dijo por el chat (códigos de TRAT_CAJA_ y TRAT_FACTURA_) en las
// filas de esos proveedores, como si lo hubiera elegido ella.
function ponerTratamientos_(ss, texto) {
  var h = ss.getSheetByName(TRAT_PESTANA_);
  if (!h) return { ok: false, error: 'no está la pestaña' };
  var pedidos = {};
  String(texto).split(',').forEach(function (p) {
    var x = p.split(':');
    if (x.length === 3 && TRAT_CAJA_[x[1]] && TRAT_FACTURA_[x[2]]) pedidos[normalizarClave_(x[0])] = { c: x[1], f: x[2] };
  });
  var n = Math.max(1, h.getLastRow() - TRAT_FILA_CAB_);
  // El desplegable de la columna C, con las opciones de ahora (por si hay nuevas).
  h.getRange(TRAT_FILA_CAB_ + 1, 3, n, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.keys(TRAT_FACTURA_).map(function (k) { return TRAT_FACTURA_[k]; }), true).setAllowInvalid(false).build());
  var datos = h.getRange(TRAT_FILA_CAB_ + 1, 1, n, 5).getValues();
  var hechos = [];
  datos.forEach(function (v, i) {
    var p = pedidos[normalizarClave_(v[0])];
    if (!p) return;
    h.getRange(TRAT_FILA_CAB_ + 1 + i, 2, 1, 4).setValues([[TRAT_CAJA_[p.c], TRAT_FACTURA_[p.f], textoTratamiento_(p.c, p.f), 'Luciana (por chat)']]);
    h.getRange(TRAT_FILA_CAB_ + 1 + i, 1, 1, 5).setBackground(null);
    hechos.push(v[0]);
    delete pedidos[normalizarClave_(v[0])];
  });
  var lista = h.getRange(TRAT_FILA_CAB_ + 1, 1, n, 3).getValues()
    .filter(function (v) { return v[0]; })
    .map(function (v) { return { nombre: v[0], c: codigoDe_(TRAT_CAJA_, v[1]), f: codigoDe_(TRAT_FACTURA_, v[2]) }; });
  guardarTratamientos_(lista);
  return { ok: true, puestos: hechos, noEncontrados: Object.keys(pedidos) };
}
