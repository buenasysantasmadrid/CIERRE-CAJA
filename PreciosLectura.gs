// ============================================================================
// PRECIOS de todas las facturas, albaranes y tickets (26/11/2026, Luciana:
// "que tengas en cuenta las cantidades que trae el producto para saber el
// precio unitario").
// ============================================================================
// Una sola tabla, pestaña "PRECIOS" del Índice, con cada producto de cada
// documento: cuántas unidades se cobran, qué es cada unidad (caja, kg…),
// cuántas piezas trae, cuánto trae cada pieza (5 l, 0,33 l, 1 kg…) y de ahí
// el precio de la pieza y el precio por kg / litro / unidad, que es el que se
// compara con la vez anterior.
//   · Lee las facturas escaneadas en la caja (preciosFactura_, al momento) y
//     TODO lo que hay en Drive "Facturas proveedores" (también lo que llega
//     por email) con el activador leerPreciosPendientes, cada 30 minutos,
//     desde PRECIOS_DESDE_. No lee Servicios ni Glovo ni las subcarpetas
//     "Albaranes" (ya está su factura).
//   · Comprueba las cuentas: cantidad × precio = importe de la línea y la
//     suma de las líneas = base de la factura. Lo que no cuadra, lo que no se
//     sabe cuántas piezas trae o un cambio de precio muy grande va en la
//     columna "Revisar", sale en PENDIENTES y NO cambia ningún precio.
//   · De la misma lectura salen los precios para la app de Pedidos
//     (APP_Precios) y, para Frutapro, la pestaña "Precios proveedores":
//     siguen igual que antes para el escandallo y el informe de los viernes.
//   · La misma factura leída otra vez (otro PDF con el mismo número, p. ej.
//     el del email después de la foto) reemplaza sus líneas.
// La pestaña "PRECIOS archivos leídos" dice qué archivo se leyó y cómo fue
// (si falla, se reintenta hasta 3 veces).
// ============================================================================

var PRECIOS_PESTANA_ = 'PRECIOS';
var PRECIOS_ARCHIVOS_PESTANA_ = 'PRECIOS archivos leídos';
var PRECIOS_CAB_ = ['Fecha', 'Proveedor', 'Documento', 'Nº', 'En la factura', 'Producto', 'ID pedidos',
  'Cantidad', 'Se cobra por', 'Piezas por bulto', 'Cada pieza trae', 'Medida',
  'Precio (de lo que se cobra)', 'Importe línea', 'Precio pieza', 'Precio por kg / l / ud',
  'Precio anterior', 'Cambio %', 'Revisar', 'PDF', 'Id archivo', 'Leído'];
var PRECIOS_N_ = PRECIOS_CAB_.length;
var PRECIOS_ARCH_CAB_ = ['Id archivo', 'Archivo', 'Carpeta', 'Resultado', 'Líneas', 'Dudosas', 'Intentos', 'Fecha'];
var PRECIOS_DESDE_ = '2026-10';
// FC VYR (Pidebio, Madrid Rural, Dia de FC VYR) no se lee (26/11/2026, Luciana): no es compra del restaurante.
var PRECIOS_CARPETAS_NO_ = ['servicios', 'glovo', 'fc vyr'];
function esFcVyrPrecios_(proveedor, detalle) {
  return PRECIOS_CARPETAS_NO_.indexOf(sinAcentos_(proveedor)) > -1 || !!proveedorFcVyr_(proveedor || '', detalle || '');
}
var PRECIOS_CAMBIO_GRANDE_ = 40; // % de cambio que se marca para revisar
var PRECIOS_MAX_INTENTOS_ = 3;

var ESQUEMA_PRECIOS_ = {
  type: 'object',
  properties: {
    tipo_documento: { type: 'string' },
    numero: { type: 'string' },
    fecha: { type: 'string' },
    proveedor_pedidos: { type: 'string' },
    base_imponible: { type: 'string' },
    otros_importes: { type: 'string' },
    lineas: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          texto: { type: 'string' }, producto: { type: 'string' },
          cantidad: { type: 'string' }, se_cobra_por: { type: 'string' },
          piezas_por_bulto: { type: 'string' }, contenido_pieza: { type: 'string' }, medida: { type: 'string' },
          precio: { type: 'string' }, importe: { type: 'string' },
          id_producto: { type: 'string' }, formato_factura: { type: 'string' },
          formato_pedidos: { type: 'string' }, precio_formato_pedidos: { type: 'string' }
        },
        required: ['texto', 'producto', 'cantidad', 'se_cobra_por', 'piezas_por_bulto', 'contenido_pieza', 'medida',
          'precio', 'importe', 'id_producto', 'formato_factura', 'formato_pedidos', 'precio_formato_pedidos'],
        additionalProperties: false
      }
    }
  },
  required: ['tipo_documento', 'numero', 'fecha', 'proveedor_pedidos', 'base_imponible', 'otros_importes', 'lineas'],
  additionalProperties: false
};

function numPrecio_(s) {
  var t = String(s == null ? '' : s).replace(/[€\s]/g, '');
  if (!t) return NaN;
  if (t.indexOf(',') > -1 && t.indexOf('.') > -1) t = t.replace(/\./g, '').replace(',', '.');
  else t = t.replace(',', '.');
  var n = Number(t);
  return isFinite(n) ? n : NaN;
}
function redondeo_(n, d) { var f = Math.pow(10, d || 4); return Math.round(n * f) / f; }
function claveNumeroPrecio_(n) { return String(n || '').toUpperCase().replace(/[^0-9A-Z]/g, ''); }

function instruccionesPrecios_(proveedor, deCual, catalogo) {
  return 'Este documento es una factura, albarán o ticket del proveedor "' + proveedor + '" del restaurante "Buenas y Santas" (el cliente). ' +
    'Hay que sacar el precio de cada producto para comparar precios entre compras, también cuando el mismo producto viene en formatos distintos ' +
    '(caja de 24 o pack de 12, garrafa de 5 l o caja de 3 garrafas, cubo de 3,6 kg o de 5 kg).\n\n' +
    '- tipo_documento: "factura", "albaran", "ticket" u "otro".\n' +
    '- numero: el número del documento tal como sale. fecha: la del documento, AAAA-MM-DD.\n' +
    '- base_imponible: total del documento SIN IVA. otros_importes: suma, con signo, de lo que no es producto (portes, envases o cascos, ' +
    'recargos; los descuentos generales en negativo). Vacíos si no hay.\n' +
    '- lineas: una por cada producto, en el orden del documento. NO incluyas portes, envases o cascos retornables, recargos, ' +
    'descuentos generales, subtotales ni totales. Si el documento solo resume otros albaranes (lista de albaranes con su importe, sin productos), lineas vacío.\n' +
    '  · texto: la descripción tal como sale, con su código si lo tiene (igual cada vez).\n' +
    '  · producto: nombre claro y corto (ej. "Coca-Cola lata 33 cl", "Queso feta", "Aceite de girasol").\n' +
    '  · cantidad: cuántas unidades COBRADAS hay en la línea (la columna de cantidad: cajas, kg, unidades…).\n' +
    '  · se_cobra_por: qué es UNA unidad cobrada: "caja", "pack", "kg", "unidad", "garrafa", "bolsa", "bandeja"…\n' +
    '  · piezas_por_bulto: cuántas piezas sueltas trae UNA unidad cobrada ("x24" → 24; "5Lx3" → 3; "2X3.78 LITROS" → 2; "4BBx2.5KGS" → 4). ' +
    'Si se cobra por kg o por pieza suelta → 1. Míralo en la descripción, en la columna de unidades por bulto o deducido del importe. ' +
    'Vacío si no se puede saber con seguridad.\n' +
    '  · contenido_pieza y medida: cuánto trae UNA pieza suelta y en qué. medida "kg" para pesos (500 g → 0.5), "l" para volúmenes ' +
    '(33 cl → 0.33; 750 ml → 0.75) o "ud" cuando no tiene peso ni volumen claro (lechuga, huevo, servilletas → contenido 1). ' +
    'Si se cobra por kg: contenido 1 y medida "kg". Vacíos si no se sabe.\n' +
    '  · precio: precio de UNA unidad cobrada, SIN IVA, con los descuentos de la propia línea aplicados. Si solo trae el importe, divídelo entre la cantidad.\n' +
    '  · importe: importe de la línea sin IVA, como sale en el documento (vacío si no lo trae).\n' +
    'Ejemplos: "Coca Cola Lata 33cl x24", 3 CAJ, 33,40, importe 100,20 → cantidad 3, se_cobra_por "caja", piezas_por_bulto 24, contenido_pieza 0.33, medida "l", precio 33.40, importe 100.20. ' +
    '"PROVOLONE 5 KGS", 5,528 Kg a 9,85 → cantidad 5.528, se_cobra_por "kg", piezas_por_bulto 1, contenido_pieza 1, medida "kg", precio 9.85. ' +
    '"Aceite Girasol 5Lx3", 1 CAJ a 28,57 → piezas_por_bulto 3, contenido_pieza 5, medida "l".\n' +
    (deCual
      ? '\nAbajo están los productos de ese proveedor en la app de pedidos del restaurante (ID | producto | unidad).\n' +
        '- proveedor_pedidos: escribe "' + deCual + '"; si el documento claramente NO es de ese proveedor, vacío.\n' +
        '  · id_producto: el ID del producto de la lista que sea el mismo artículo, aunque esté escrito distinto. Si no corresponde a ninguno o no estás seguro, vacío.\n' +
        '  · formato_factura: lo que es UNA unidad cobrada con cuántas piezas trae (ej. "pack 12 latas", "caja 24 latas", "kg").\n' +
        '  · formato_pedidos: el formato del producto de la lista (lo que dice su nombre o unidad, ej. "COCA-COLA 24 lata" = "caja 24 latas"); vacío si no hay id_producto.\n' +
        '  · precio_formato_pedidos: precio convertido a formato_pedidos (pack 12 latas a 9 → caja 24 latas = 18). Igual a precio si es el mismo formato. ' +
        'Si no puedes saber con seguridad cuántas piezas trae alguno de los dos, VACÍO.\n'
      : '\n- proveedor_pedidos, id_producto, formato_factura, formato_pedidos y precio_formato_pedidos: vacíos.\n') +
    'Números con punto decimal y sin símbolo de euro (ej. 12.5). Si un dato no se lee con seguridad, déjalo vacío: nunca lo inventes.' +
    (deCual ? '\n\nPRODUCTOS DE ' + deCual + ':\n' + catalogo.join('\n') : '');
}

// A qué proveedor de Pedidos va (se recuerda por nombre para no preguntar cada vez).
function proveedorPedidosRecordado_(proveedor, detalle, nombres) {
  var props = PropertiesService.getScriptProperties();
  var mapa = {};
  try { mapa = JSON.parse(props.getProperty('precios_mapa_pedidos') || '{}'); } catch (e) { mapa = {}; }
  var k = sinAcentos_(proveedor + ' ' + (detalle || ''));
  if (k in mapa && (mapa[k] === '' || nombres.indexOf(mapa[k]) > -1)) return mapa[k];
  var r = proveedorDePedidos_({ detalle: detalle || '' }, proveedor, nombres);
  // Solo se recuerdan los nombres de carpeta (sin detalle: el detalle cambia con cada ticket).
  if (!detalle) { mapa[k] = r; try { props.setProperty('precios_mapa_pedidos', JSON.stringify(mapa)); } catch (e) { /* lleno */ } }
  return r;
}

// Lee un documento y guarda sus precios. meta: { proveedor, detalle, fecha,
// numero, url, idArchivo, proveedorPedidos }. contenido: las hojas para Claude.
function leerYGuardarPrecios_(contenido, meta) {
  var proveedor = String(meta.proveedor || '').trim() || 'Sin proveedor';
  if (esFcVyrPrecios_(proveedor, meta.detalle)) return { ok: true, lineas: 0, dudosas: 0, sinProductos: true, noSeLee: 'FC VYR' };
  var cat = null, nombres = [], deCual = '';
  try {
    cat = catalogoPedidos_();
    nombres = Object.keys(cat.proveedores);
    deCual = meta.proveedorPedidos && nombres.indexOf(meta.proveedorPedidos) > -1
      ? meta.proveedorPedidos : proveedorPedidosRecordado_(proveedor, meta.detalle, nombres);
    if (deCual && !cat.proveedores[deCual].length) deCual = '';
  } catch (err) { deCual = ''; }
  var quien = proveedor + (meta.detalle ? ' (' + meta.detalle + ')' : '');
  GASTO_MOTIVO_ = 'precios';
  var r;
  try {
    r = pedirJsonAClaude_(contenido.concat([{ type: 'text', text: instruccionesPrecios_(quien, deCual, deCual ? cat.proveedores[deCual] : []) }]),
      ESQUEMA_PRECIOS_, 'medium', MODELO_ESCANEO);
  } finally { GASTO_MOTIVO_ = 'escaneo'; }
  if (!r.ok) return { ok: false, error: r.error };

  var fecha = /^\d{4}-\d{2}-\d{2}$/.test(String(meta.fecha || '')) ? meta.fecha
    : (/^\d{4}-\d{2}-\d{2}$/.test(String(r.fecha || '')) ? r.fecha : Utilities.formatDate(new Date(), 'Europe/Madrid', 'yyyy-MM-dd'));
  var numero = String(r.numero || '').trim() || String(meta.numero || '').trim();
  var tipo = String(r.tipo_documento || '').toLowerCase().replace('albarán', 'albaran');
  var nombreProv = proveedor + (meta.detalle ? ' — ' + meta.detalle : '');

  // Cada línea con sus cuentas.
  // Un albarán sin valorar (sin precios ni importes) no sirve para precios.
  var lineas = (r.lineas || []).filter(function (l) {
    return String(l.texto || '').trim() && (numPrecio_(l.precio) > 0 || numPrecio_(l.importe) > 0);
  }).map(function (l) {
    var cant = numPrecio_(l.cantidad), precio = numPrecio_(l.precio), imp = numPrecio_(l.importe);
    var piezas = numPrecio_(l.piezas_por_bulto), cont = numPrecio_(l.contenido_pieza);
    var medida = sinAcentos_(l.medida);
    medida = /^(kg|kilo|kilos|kgs)$/.test(medida) ? 'kg' : /^(l|litro|litros|lt|lts)$/.test(medida) ? 'l' : /^(ud|uds|unidad|unidades|u)$/.test(medida) ? 'ud' : '';
    if (!(precio > 0) && imp > 0 && cant > 0) precio = imp / cant;
    var revisar = [];
    // Muchos tickets (Mercadona, Mercandalucía) traen el importe de la línea
    // CON IVA y el precio sin IVA: si cuadra con algún IVA, está bien.
    var conIva = [1, 1.02, 1.04, 1.05, 1.10, 1.21].some(function (k) { return Math.abs(cant * precio * k - imp) <= Math.max(0.06, imp * 0.03); });
    if (cant > 0 && precio > 0 && imp > 0 && !conIva) {
      revisar.push('cantidad × precio da ' + redondeo_(cant * precio, 2) + ' y la línea dice ' + imp);
    }
    // Lo que se cobra por kg o por pieza suelta trae 1 pieza.
    var cobro = sinAcentos_(l.se_cobra_por);
    if (!(piezas > 0) && (!cobro || /^(kg|kgs|kilo|kilos|unidad|unidades|ud|uds|u|un|uni|pieza|piezas|botella|lata|litro|litros|l)$/.test(cobro))) piezas = 1;
    var precioPieza = piezas > 0 && precio > 0 ? precio / piezas : '';
    if (!(piezas > 0)) revisar.push('no se sabe cuántas piezas trae cada ' + (String(l.se_cobra_por || '').trim() || 'bulto'));
    // Sin peso ni volumen claro se compara por pieza (no es un fallo).
    var precioBase = precioPieza && cont > 0 && medida ? precioPieza / cont : '';
    return { l: l, cant: cant, precio: precio, imp: imp, piezas: piezas, cont: cont, medida: medida,
      precioPieza: precioPieza, precioBase: precioBase, revisar: revisar };
  });

  // ¿La suma de las líneas da la base de la factura?
  var base = numPrecio_(r.base_imponible), otros = numPrecio_(r.otros_importes) || 0;
  var suma = lineas.reduce(function (s, x) { return s + (x.imp > 0 ? x.imp : (x.cant > 0 && x.precio > 0 ? x.cant * x.precio : 0)); }, 0);
  var notaDoc = '';
  // (las líneas pueden venir con IVA: entonces suman entre la base y la base + 21 %)
  var total = suma + otros;
  var cuadra = Math.abs(total - base) <= Math.max(1, base * 0.02) || (total >= base * 0.99 && total <= base * 1.22);
  if (lineas.length && base > 0 && !cuadra) {
    notaDoc = 'las líneas suman ' + redondeo_(suma, 2) + (otros ? ' (+ otros ' + redondeo_(otros, 2) + ')' : '') + ' y la base del documento es ' + base + ': puede faltar o sobrar alguna línea';
  }

  var guardado = guardarFilasPrecios_(lineas, {
    fecha: fecha, proveedor: nombreProv, tipo: tipo, numero: numero, url: meta.url || '', idArchivo: meta.idArchivo || '', notaDoc: notaDoc
  });

  var res = { ok: true, lineas: lineas.length, dudosas: guardado.dudosas, tipo: tipo, numero: numero, sinProductos: !lineas.length };

  // Pollería, Mercandalucía y Las 3B: los kilos a la Comparativa carnicería
  // (ComparativaCarne.gs). La factura del mes no (sería contar dos veces).
  if (!meta.noComparativa && proveedorComparativa_(proveedor) && lineas.length) {
    try {
      res.comparativa = apuntarComparativa_(lineas.map(function (x) {
        return { texto: x.l.texto, cant: x.cant, cobro: x.l.se_cobra_por, precio: x.precio };
      }), { proveedor: proveedor, fecha: fecha, numero: numero, idArchivo: meta.idArchivo || '', desdeApp: !!meta.desdeApp });
    } catch (err) { res.errorComparativa = String(err && err.message || err); Logger.log('Comparativa: ' + err); }
  }

  // App de Pedidos: las dudosas van sin emparejar (no cambian el precio del producto).
  if (deCual && r.proveedor_pedidos === deCual && lineas.length) {
    try {
      res.pedidos = pedirAPedidos_({
        accion: 'precios', proveedor: deCual, factura: numero, fecha: fecha, foto: meta.url || '', reemplazar: true,
        lineas: lineas.map(function (x) {
          var p = lineaEnFormatoPedidos_({
            texto: x.l.texto, cantidad: x.l.cantidad, unidad: x.l.se_cobra_por, precio_unitario: x.precio > 0 ? String(redondeo_(x.precio)) : '',
            id_producto: x.revisar.length ? '' : x.l.id_producto, formato_factura: x.l.formato_factura,
            formato_pedidos: x.l.formato_pedidos, precio_formato_pedidos: x.l.precio_formato_pedidos
          });
          return p;
        })
      });
    } catch (err) { res.errorPedidos = err.message; }
  }
  // Frutapro: también en "Precios proveedores" (por kg o pieza suelta), sin las dudosas.
  var propio = proveedorConPreciosPropios_(proveedor + ' ' + (meta.detalle || ''));
  if (propio && lineas.length) {
    try {
      res.propios = escribirPreciosPropios_(propio, numero, fecha, meta.url || '', lineas.filter(function (x) { return !x.revisar.length && x.precio > 0; }).map(function (x) {
        var conBase = x.precioBase > 0;
        return {
          texto: x.l.texto, producto: x.l.producto,
          cantidad: conBase ? redondeo_(x.cant * x.piezas * x.cont, 3) : x.cant,
          unidad: conBase ? (x.medida === 'ud' ? 'unidad' : x.medida) : x.l.se_cobra_por,
          precio_unitario: conBase ? x.precioBase : x.precio
        };
      }));
    } catch (err) { res.errorPropios = err.message; }
  }
  return res;
}

function hojaPrecios_() {
  var ss = SpreadsheetApp.openById(INDICE_SHEET_ID_);
  var h = ss.getSheetByName(PRECIOS_PESTANA_);
  if (!h) {
    h = ss.insertSheet(PRECIOS_PESTANA_);
    h.getRange(1, 1, 1, PRECIOS_N_).setValues([PRECIOS_CAB_]).setFontWeight('bold').setBackground('#e8eaf6').setWrap(true);
    h.setFrozenRows(1);
    h.setColumnWidth(5, 260); h.setColumnWidth(6, 200); h.setColumnWidth(19, 320);
    h.hideColumns(21, 2);
  }
  return h;
}

function fechaIsoDe_(v) {
  return v instanceof Date ? Utilities.formatDate(v, 'Europe/Madrid', 'yyyy-MM-dd') : String(v || '').slice(0, 10);
}

// Escribe las líneas de un documento (quitando antes las suyas de otra vez)
// y deja la tabla ordenada de la más nueva a la más vieja.
function guardarFilasPrecios_(lineas, d) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var h = hojaPrecios_();
    var filas = h.getLastRow() > 1 ? h.getRange(2, 1, h.getLastRow() - 1, PRECIOS_N_).getValues() : [];
    var antes = filas.length;
    var kNum = claveNumeroPrecio_(d.numero), kProv = sinAcentos_(d.proveedor);
    filas = filas.filter(function (f) {
      if (d.idArchivo && String(f[20]) === d.idArchivo) return false;
      if (kNum.length >= 3 && sinAcentos_(f[1]) === kProv && claveNumeroPrecio_(f[3]) === kNum) return false;
      return true;
    });
    // Precio anterior: el más reciente de antes de esta fecha, del mismo producto y medida.
    var anterior = {};
    // Se compara por kg / l / ud si se sabe; si no, por pieza; si no, por lo que se cobra.
    var clave = function (prov, id, texto, modo, medida, cobro) {
      return sinAcentos_(prov) + '|' + (id || sinAcentos_(texto)) + '|' + (modo === 'b' ? 'b|' + medida : modo === 'p' ? 'p' : 'c|' + sinAcentos_(cobro));
    };
    filas.forEach(function (f) {
      var fe = fechaIsoDe_(f[0]);
      if (fe > d.fecha || String(f[18])) return;
      var modo = Number(f[15]) > 0 ? 'b' : Number(f[14]) > 0 ? 'p' : 'c';
      var k = clave(f[1], f[6], f[4], modo, f[11], f[8]);
      if (!anterior[k] || anterior[k].fecha <= fe) anterior[k] = { fecha: fe, precio: Number(f[modo === 'b' ? 15 : modo === 'p' ? 14 : 12]) };
    });
    var ahora = new Date(), dudosas = 0;
    var nuevas = lineas.map(function (x, i) {
      var modo = x.precioBase > 0 ? 'b' : x.precioPieza > 0 ? 'p' : 'c';
      var actual = modo === 'b' ? x.precioBase : modo === 'p' ? x.precioPieza : x.precio;
      var prev = anterior[clave(d.proveedor, x.l.id_producto, x.l.texto, modo, x.medida, x.l.se_cobra_por)];
      var cambio = prev && prev.precio > 0 && actual > 0 ? redondeo_((actual - prev.precio) / prev.precio * 100, 1) : '';
      var revisar = x.revisar.slice();
      if (cambio !== '' && Math.abs(cambio) >= PRECIOS_CAMBIO_GRANDE_) revisar.push('cambio de ' + cambio + '% respecto a la vez anterior: comprobar que se leyó bien');
      if (i === 0 && d.notaDoc) revisar.push(d.notaDoc);
      if (revisar.length) dudosas++;
      x.revisar = revisar;
      return [fechaComoDate_(d.fecha), d.proveedor, d.tipo, d.numero, String(x.l.texto).trim(), String(x.l.producto || '').trim(),
        String(x.l.id_producto || '').trim(), x.cant > 0 ? x.cant : '', String(x.l.se_cobra_por || '').trim(),
        x.piezas > 0 ? x.piezas : '', x.cont > 0 ? x.cont : '', x.medida,
        x.precio > 0 ? redondeo_(x.precio) : '', x.imp > 0 ? x.imp : '', x.precioPieza ? redondeo_(x.precioPieza) : '',
        x.precioBase ? redondeo_(x.precioBase) : '', prev ? redondeo_(prev.precio) : '', cambio,
        revisar.join(' · '), d.url, d.idArchivo, ahora];
    });
    filas = filas.concat(nuevas);
    filas.sort(function (a, b) {
      var fa = fechaIsoDe_(a[0]), fb = fechaIsoDe_(b[0]);
      return fb.localeCompare(fa) || String(a[1]).localeCompare(String(b[1])) || String(a[3]).localeCompare(String(b[3]));
    });
    if (antes > filas.length) h.getRange(filas.length + 2, 1, antes - filas.length, PRECIOS_N_).clearContent();
    if (filas.length) {
      h.getRange(2, 1, filas.length, PRECIOS_N_).setValues(filas);
      h.getRange(2, 1, filas.length, 1).setNumberFormat('dd/mm/yyyy');
      h.getRange(2, 22, filas.length, 1).setNumberFormat('dd/mm/yyyy hh:mm');
    }
    return { dudosas: dudosas };
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------------------
// Archivos de Drive
// ---------------------------------------------------------------------------
function hojaArchivosPrecios_() {
  var ss = SpreadsheetApp.openById(INDICE_SHEET_ID_);
  var h = ss.getSheetByName(PRECIOS_ARCHIVOS_PESTANA_);
  if (!h) {
    h = ss.insertSheet(PRECIOS_ARCHIVOS_PESTANA_);
    h.getRange(1, 1, 1, PRECIOS_ARCH_CAB_.length).setValues([PRECIOS_ARCH_CAB_]).setFontWeight('bold');
    h.setFrozenRows(1);
  }
  return h;
}
function leerArchivosPrecios_() {
  var h = hojaArchivosPrecios_();
  var mapa = {};
  if (h.getLastRow() < 2) return mapa;
  h.getRange(2, 1, h.getLastRow() - 1, PRECIOS_ARCH_CAB_.length).getValues().forEach(function (f, i) {
    if (f[0]) mapa[String(f[0])] = { fila: i + 2, resultado: String(f[3]), intentos: Number(f[6]) || 0 };
  });
  return mapa;
}
function apuntarArchivoPrecios_(id, nombre, carpeta, resultado, lineas, dudosas, mapa) {
  var h = hojaArchivosPrecios_();
  mapa = mapa || leerArchivosPrecios_();
  var previo = mapa[id];
  var intentos = (previo ? previo.intentos : 0) + (resultado.indexOf('error') === 0 ? 1 : 0);
  var fila = [id, nombre, carpeta, resultado, lineas, dudosas, intentos, new Date()];
  if (previo) h.getRange(previo.fila, 1, 1, fila.length).setValues([fila]);
  else { h.appendRow(fila); previo = { fila: h.getLastRow() }; }
  mapa[id] = { fila: previo.fila, resultado: resultado, intentos: intentos };
}

// Proveedor, detalle y fecha de un archivo por su carpeta y su nombre
// ("AAAA-MM-DD Proveedor [detalle] número").
function metaDeArchivo_(archivo) {
  var padre = archivo.getParents().hasNext() ? archivo.getParents().next() : null;
  var carpeta = padre ? padre.getName() : '';
  if (carpeta === 'Albaranes' && padre.getParents().hasNext()) carpeta = padre.getParents().next().getName();
  var nombre = archivo.getName().replace(/\.(pdf|jpe?g|png)$/i, '');
  var m = /^(\d{4}-\d{2}-\d{2})\s*(.*)$/.exec(nombre);
  var fecha = m ? m[1] : Utilities.formatDate(archivo.getDateCreated(), 'Europe/Madrid', 'yyyy-MM-dd');
  var resto = (m ? m[2] : nombre).trim();
  if (carpeta && resto.indexOf(carpeta + ' ') === 0) resto = resto.slice(carpeta.length + 1);
  var fija = CARPETAS_FIJAS_FACTURAS_.indexOf(carpeta) > -1;
  // En Super / Varios el nombre de la tienda va en el archivo, antes del número.
  var detalle = fija ? resto.replace(/\s+\S*\d\S*$/, '').trim() : '';
  return { proveedor: carpeta || 'Sin proveedor', detalle: detalle, fecha: fecha, numero: '', url: archivo.getUrl(), idArchivo: archivo.getId(), carpeta: carpeta };
}

function contenidoDeArchivo_(archivo) {
  var blob = archivo.getBlob(), tipo = blob.getContentType();
  var b64 = Utilities.base64Encode(blob.getBytes());
  if (tipo === 'application/pdf') return [{ type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: b64 } }];
  return [{ type: 'image', source: { type: 'base64', media_type: tipo, data: b64 } }];
}

function leerPreciosDeArchivo_(archivo, mapa, extra) {
  var meta = metaDeArchivo_(archivo);
  if (extra) Object.keys(extra).forEach(function (k) { if (extra[k]) meta[k] = extra[k]; });
  // La factura del mes (Bandejas.gs, nota "factura") no va a la Comparativa carnicería.
  if (String(archivo.getDescription() || '').split('\n')[0].toLowerCase() === 'factura') meta.noComparativa = true;
  var r;
  try { r = leerYGuardarPrecios_(contenidoDeArchivo_(archivo), meta); }
  catch (err) { r = { ok: false, error: String(err && err.message || err) }; }
  apuntarArchivoPrecios_(archivo.getId(), archivo.getName(), meta.carpeta,
    r.ok ? (r.sinProductos ? 'sin productos' : 'leído') : 'error: ' + String(r.error).slice(0, 200),
    r.ok ? r.lineas : '', r.ok ? r.dudosas : '', mapa);
  return r;
}

// Los archivos de Drive que faltan por leer, del más nuevo al más viejo.
function archivosSinPrecios_(mapa) {
  var raiz = carpetaFacturas_(), lista = [];
  var hoy = new Date(), meses = [];
  for (var d = new Date(Number(PRECIOS_DESDE_.slice(0, 4)), Number(PRECIOS_DESDE_.slice(5, 7)) - 1, 1); d <= hoy; d = new Date(d.getFullYear(), d.getMonth() + 1, 1)) {
    meses.push([String(d.getFullYear()), ('0' + (d.getMonth() + 1)).slice(-2)]);
  }
  meses.forEach(function (am) {
    var itA = raiz.getFoldersByName(am[0]);
    if (!itA.hasNext()) return;
    var itM = itA.next().getFoldersByName(am[1]);
    if (!itM.hasNext()) return;
    var provs = itM.next().getFolders();
    while (provs.hasNext()) {
      var prov = provs.next(), nombreProv = prov.getName();
      if (/^[📥⚠🗑]/.test(nombreProv) || esFcVyrPrecios_(nombreProv, '')) continue;
      if (typeof SERVICIOS_CONOCIDOS_ !== 'undefined' && esServicioPrecios_(nombreProv)) continue;
      var fs = prov.getFiles();
      while (fs.hasNext()) {
        var f = fs.next();
        if (!/\.(pdf|jpe?g|png)$/i.test(f.getName())) continue;
        var ya = mapa[f.getId()];
        if (ya && (ya.resultado.indexOf('error') !== 0 || ya.intentos >= PRECIOS_MAX_INTENTOS_)) continue;
        lista.push({ archivo: f, nombre: f.getName() });
      }
    }
  });
  lista.sort(function (a, b) { return b.nombre.localeCompare(a.nombre); });
  return lista;
}
function esServicioPrecios_(nombre) {
  var n = sinAcentos_(nombre);
  return (SERVICIOS_CONOCIDOS_ || []).some(function (s) {
    var t = sinAcentos_(typeof s === 'string' ? s : (s && (s.nombre || s[0])) || '');
    return t && t === n;
  });
}

// Activador cada 30 minutos: lee lo que falte hasta llegar a ~4,5 minutos.
function leerPreciosPendientes() {
  var props = PropertiesService.getScriptProperties();
  // Pausa (26/11/2026: Luciana vio un cobro de créditos de Claude). ?precios=pausar / reanudar.
  if (props.getProperty('precios_pausa')) return { ok: true, pausado: true };
  var corriendo = Number(props.getProperty('precios_corriendo') || 0);
  if (corriendo && Date.now() - corriendo < 7 * 60 * 1000) return { ok: true, yaCorriendo: true };
  props.setProperty('precios_corriendo', String(Date.now()));
  var t0 = Date.now(), hechos = 0, errores = 0, lineas = 0, dudosas = 0, quedan = 0;
  try {
    var mapa = leerArchivosPrecios_();
    var lista = archivosSinPrecios_(mapa);
    for (var i = 0; i < lista.length; i++) {
      if (Date.now() - t0 > 200000) { quedan = lista.length - i; break; }
      var r = leerPreciosDeArchivo_(lista[i].archivo, mapa);
      if (r.ok) { hechos++; lineas += r.lineas || 0; dudosas += r.dudosas || 0; } else errores++;
    }
  } finally {
    props.deleteProperty('precios_corriendo');
  }
  return { ok: true, leidos: hechos, errores: errores, lineas: lineas, dudosas: dudosas, quedan: quedan, segundos: Math.round((Date.now() - t0) / 1000) };
}

function asegurarActivadorPrecios_() {
  var hay = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'leerPreciosPendientes'; });
  if (!hay) ScriptApp.newTrigger('leerPreciosPendientes').timeBased().everyMinutes(30).create();
  return !hay;
}

// Para PENDIENTES: los documentos con líneas dudosas de los últimos 45 días.
function preciosDudosos_() {
  var h = SpreadsheetApp.openById(INDICE_SHEET_ID_).getSheetByName(PRECIOS_PESTANA_);
  if (!h || h.getLastRow() < 2) return [];
  var desde = Utilities.formatDate(new Date(Date.now() - 45 * 86400000), 'Europe/Madrid', 'yyyy-MM-dd');
  var docs = {}, orden = [];
  h.getRange(2, 1, h.getLastRow() - 1, PRECIOS_N_).getValues().forEach(function (f) {
    if (!String(f[18]) || fechaIsoDe_(f[0]) < desde) return;
    var k = String(f[20] || (f[1] + '|' + f[3]));
    if (!docs[k]) { docs[k] = { clave: k, proveedor: f[1], fecha: fechaIsoDe_(f[0]), numero: f[3], url: f[19], que: [] }; orden.push(k); }
    docs[k].que.push((f[5] || f[4]) + ': ' + f[18]);
  });
  return orden.map(function (k) { return docs[k]; });
}

// Vuelve a leer los archivos que dejaron líneas dudosas (después de mejorar las comprobaciones).
function releerPreciosDudosos_() {
  var h = hojaArchivosPrecios_(), n = 0;
  if (h.getLastRow() < 2) return { ok: true, marcados: 0 };
  var v = h.getRange(2, 1, h.getLastRow() - 1, PRECIOS_ARCH_CAB_.length).getValues();
  v.forEach(function (f, i) { if (Number(f[5]) > 0) { h.getRange(i + 2, 4, 1, 4).setValues([['error: releer', f[4], f[5], 0]]); n++; } });
  return { ok: true, marcados: n };
}

// Para que Claude compruebe la tabla: las últimas N filas.
function muestraPrecios_(n) {
  var h = SpreadsheetApp.openById(INDICE_SHEET_ID_).getSheetByName(PRECIOS_PESTANA_);
  if (!h || h.getLastRow() < 2) return [];
  return h.getRange(2, 1, Math.min(Number(n) || 60, h.getLastRow() - 1), 19).getDisplayValues();
}

// Quita de PRECIOS lo que no se tiene que leer (FC VYR), por si se leyó antes.
function quitarPreciosNoLeer_() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var h = hojaPrecios_();
    if (h.getLastRow() < 2) return { ok: true, quitadas: 0 };
    var filas = h.getRange(2, 1, h.getLastRow() - 1, PRECIOS_N_).getValues();
    var quedan = filas.filter(function (f) { var p = String(f[1]).split(' — '); return !esFcVyrPrecios_(p[0], p[1] || ''); });
    if (quedan.length < filas.length) {
      h.getRange(2, 1, filas.length, PRECIOS_N_).clearContent();
      if (quedan.length) h.getRange(2, 1, quedan.length, PRECIOS_N_).setValues(quedan);
    }
    return { ok: true, quitadas: filas.length - quedan.length };
  } finally {
    lock.releaseLock();
  }
}

// Pausar o reanudar la lectura automática de precios (cada lectura gasta créditos de Claude).
function pausarPrecios_(pausar) {
  var props = PropertiesService.getScriptProperties();
  if (pausar) props.setProperty('precios_pausa', String(new Date())); else props.deleteProperty('precios_pausa');
  var mapa = leerArchivosPrecios_();
  var leidos = Object.keys(mapa).length;
  return { ok: true, pausado: !!pausar, archivosLeidos: leidos, quedan: archivosSinPrecios_(mapa).length };
}
