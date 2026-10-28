// ============================================================================
// Fotos de facturas y precios para la app de Pedidos a proveedores.
// ============================================================================
// Después de "Escanear" una factura, la app manda la misma foto otra vez, en
// segundo plano (nadie espera esta respuesta), a accionPreciosFactura:
//   1. La foto se guarda en Drive: "Facturas proveedores" / <proveedor> / <AAAA-MM>.
//   2. Claude lee cada línea de producto de la factura y la empareja con los
//      productos de la app de Pedidos (si el proveedor es uno de los suyos).
//   3. Las líneas se mandan a la hoja de Pedidos (pestaña APP_Precios), que
//      guarda el historial de precios y avisa en la app de los cambios.
// Si algo falla aquí, el cierre de caja no se entera: es todo aparte.
// ============================================================================

var PEDIDOS_URL_ = 'https://script.google.com/macros/s/AKfycbyMwaUKlKn46HqJtZ8vmceOulgwoXaF-qGCzHZp4w32uV2Spvk-4q-Cf2qY6u7B0bur/exec';
var CARPETA_FACTURAS_ = 'Facturas proveedores';

function preciosFactura_(data) {
  if (!data.imageBase64) return { ok: false, error: 'No llegó la foto' };
  var proveedor = String(data.proveedor || '').trim() || 'Sin proveedor';
  var fecha = /^\d{4}-\d{2}-\d{2}$/.test(String(data.fecha || ''))
    ? data.fecha : Utilities.formatDate(new Date(), 'Europe/Madrid', 'yyyy-MM-dd');
  var res = { ok: true };
  try { res.foto = guardarFotoFactura_(data, proveedor, fecha); } catch (err) { res.errorFoto = String(err); }
  // soloPdf: se añadió una hoja a una factura ya leída — solo se vuelve a
  // guardar el PDF (los precios ya se mandaron).
  var cual = proveedor + ', nº ' + (String(data.numero_factura || '').trim() || '?') + ', fecha ' + fecha;
  if (res.errorFoto) {
    avisarPorEmail_('una factura NO se guardó en Drive', 'La factura escaneada (' + cual + ') no se guardó en Drive: ' + res.errorFoto + '\nHay que volver a escanearla.');
  }
  if (data.soloPdf) return res;
  try {
    // Frutapro (no está en la app de Pedidos): sus precios van a la pestaña
    // "Precios proveedores" del Índice (ver más abajo).
    res.precios = proveedorConPreciosPropios_(proveedor)
      ? guardarPreciosPropios_(data, proveedorConPreciosPropios_(proveedor), fecha, res.foto || '')
      : mandarPreciosAPedidos_(data, proveedor, fecha, res.foto || '');
  } catch (err) { res.errorPrecios = String(err); }
  if (!res.errorPrecios && res.precios && res.precios.ok === false) res.errorPrecios = res.precios.error || 'error desconocido';
  if (res.errorPrecios) {
    avisarPorEmail_('los precios de una factura no llegaron a Pedidos', 'De la factura escaneada (' + cual + ') no se pudieron pasar los precios a la app de Pedidos: ' + res.errorPrecios +
      (res.foto ? '\nLa factura sí quedó guardada en Drive: ' + res.foto : ''));
  }
  return res;
}

function carpetaHija_(padre, nombre) {
  var it = padre.getFoldersByName(nombre);
  return it.hasNext() ? it.next() : padre.createFolder(nombre);
}

function sinCaracteresRaros_(s) {
  return String(s).replace(/[\\\/:*?"<>|#%]/g, '-').replace(/\s+/g, ' ').trim();
}

// Devuelve el enlace de la factura en Drive: "Facturas proveedores" / <AAAA>
// / <MM> / <proveedor>. La app manda la factura escaneada como PDF (todas
// sus hojas juntas: es el que se presenta para los impuestos); si no llega
// el PDF (app vieja), se guarda la foto. La misma factura escaneada otra vez
// (mismo proveedor, fecha y número) reemplaza a la anterior (por si se le
// añadió una hoja).
function guardarFotoFactura_(data, proveedor, fecha) {
  var raiz = carpetaFacturas_();
  var dir = carpetaHija_(carpetaHija_(carpetaHija_(raiz, fecha.slice(0, 4)), fecha.slice(5, 7)), sinCaracteresRaros_(proveedor));
  var esPdf = !!data.pdfBase64;
  var tipo = esPdf ? 'application/pdf' : (data.mediaType || 'image/jpeg');
  // Sin número de factura, la hora del primer escaneo (la manda la app).
  var numero = String(data.numero_factura || '').trim() || String(data.idEscaneo || '').trim();
  // detalle: el nombre del súper cuando la carpeta es "Super" (ej. "Carrefour Market").
  var detalle = String(data.detalle || '').trim();
  var nombre = sinCaracteresRaros_(fecha + ' ' + proveedor + (detalle ? ' ' + detalle : '') + ' ' +
    (numero || Utilities.formatDate(new Date(), 'Europe/Madrid', 'HH.mm.ss'))) + (esPdf ? '.pdf' : (tipo === 'image/png' ? '.png' : '.jpg'));
  if (numero) {
    var existe = dir.getFilesByName(nombre);
    while (existe.hasNext()) existe.next().setTrashed(true);
  }
  // El mismo escaneo guardado antes con otro nombre de proveedor (la segunda
  // lectura puede leerlo distinto, ej. "Sports Direct" / "Sports Direct
  // Xanadu ES"): se borra el anterior, para que no quede dos veces.
  var cache = CacheService.getScriptCache();
  var claveEscaneo = data.idEscaneo ? 'escaneo_' + fecha + '_' + data.idEscaneo : '';
  var anterior = claveEscaneo ? cache.get(claveEscaneo) : null;
  if (anterior) {
    try { DriveApp.getFileById(anterior).setTrashed(true); } catch (err) { /* ya no está */ }
  }
  var blob = Utilities.newBlob(Utilities.base64Decode(esPdf ? data.pdfBase64 : data.imageBase64), tipo, nombre);
  var archivo = dir.createFile(blob);
  if (claveEscaneo) cache.put(claveEscaneo, archivo.getId(), 6 * 60 * 60);
  return archivo.getUrl();
}

// La carpeta de las facturas es la de buenasysantasmadrid ("Mi unidad /
// Facturas proveedores"): la app corre como lfisbein, así que esa carpeta
// tiene que estar compartida con lfisbein como editor. Si no se puede abrir
// (todavía no compartida), se guardan en la de lfisbein y llega un email.
var CARPETA_FACTURAS_ID_ = '15R08k4qgm84msoGqxpUnMtU7sLeczWCz';
function carpetaFacturas_() {
  try {
    var c = DriveApp.getFolderById(CARPETA_FACTURAS_ID_);
    c.getName(); // falla aquí si no hay acceso
    return c;
  } catch (err) {
    avisarPorEmail_('las facturas se están guardando en el Drive de lfisbein',
      'No se puede abrir la carpeta "Facturas proveedores" de buenasysantasmadrid (' + err + ').\n' +
      'Hay que compartirla con lfisbein@gmail.com como editor. Mientras tanto las facturas van a "Facturas proveedores" en el Drive de lfisbein, y se pasarán solas a la buena cuando esté compartida.');
    return carpetaHija_(DriveApp.getRootFolder(), CARPETA_FACTURAS_);
  }
}

// Desde cada guardado de la caja (doPost): si la carpeta de buenasysantas ya
// se puede usar y quedaron facturas en la de lfisbein, se pasan a la buena
// (juntando año / mes / proveedor) y la vieja se tira a la papelera.
function asegurarCarpetaFacturasCompartida_() {
  try {
    var props = PropertiesService.getScriptProperties();
    if (props.getProperty('facturas_unificadas') === CARPETA_FACTURAS_ID_) return;
    var buena = DriveApp.getFolderById(CARPETA_FACTURAS_ID_);
    buena.getName();
    var viejas = DriveApp.getRootFolder().getFoldersByName(CARPETA_FACTURAS_);
    while (viejas.hasNext()) {
      var vieja = viejas.next();
      if (vieja.getId() === CARPETA_FACTURAS_ID_) continue;
      juntarCarpetas_(vieja, buena);
      vieja.setTrashed(true);
    }
    props.setProperty('facturas_unificadas', CARPETA_FACTURAS_ID_);
  } catch (err) { /* todavía no compartida: se intenta en el próximo guardado */ }
}
function juntarCarpetas_(origen, destino) {
  var archivos = origen.getFiles();
  while (archivos.hasNext()) archivos.next().moveTo(destino);
  var subs = origen.getFolders();
  while (subs.hasNext()) {
    var sub = subs.next();
    var igual = destino.getFoldersByName(sub.getName());
    if (igual.hasNext()) { juntarCarpetas_(sub, igual.next()); sub.setTrashed(true); }
    else sub.moveTo(destino);
  }
}

// Las hojas de la factura para Claude: todas las que mandó la app (o la
// única foto, en la app vieja).
function imagenesFactura_(data) {
  // Factura releída desde su PDF de Drive (botón de la app de Pedidos).
  if (data.pdfParaClaude) {
    return [{ type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: data.pdfParaClaude } }];
  }
  var lista = (data.imagenes && data.imagenes.length) ? data.imagenes : [data.imageBase64];
  return lista.map(function (b64) {
    return { type: 'image', source: { type: 'base64', media_type: data.mediaType || 'image/jpeg', data: b64 } };
  });
}

// POST a la app de Pedidos. Si devuelve una página de error (no JSON), se
// prueba otra vez y, si sigue, el error dice qué ponía la página.
function pedirAPedidos_(cuerpo) {
  var texto = '', resp;
  for (var intento = 0; intento < 2; intento++) {
    if (intento) Utilities.sleep(3000);
    resp = UrlFetchApp.fetch(PEDIDOS_URL_, {
      method: 'post', contentType: 'text/plain;charset=utf-8', muteHttpExceptions: true,
      payload: JSON.stringify(Object.assign({ claveServidor: CLAVE_PEDIDOS_ }, cuerpo))
    });
    texto = resp.getContentText();
    try { return JSON.parse(texto); } catch (err) { /* página de error */ }
  }
  var limpio = texto.replace(/<style[\s\S]*?<\/style>|<script[\s\S]*?<\/script>/gi, ' ').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
  throw new Error('la app de Pedidos respondió una página de error (' + resp.getResponseCode() + '): ' + limpio.slice(0, 400));
}

// Proveedores y productos de la app de Pedidos (se guardan una hora).
function catalogoPedidos_() {
  var cache = CacheService.getScriptCache();
  var guardado = cache.get('catalogoPedidos');
  if (guardado) return JSON.parse(guardado);
  var j = pedirAPedidos_({ accion: 'catalogo' });
  if (!j.ok) throw new Error(j.error || 'La hoja de Pedidos no responde');
  var cat = { proveedores: {}, total: 0 };
  j.productos.forEach(function (p) {
    (cat.proveedores[p.proveedor] = cat.proveedores[p.proveedor] || []).push(
      p.id + ' | ' + p.producto + (p.unidad ? ' | ' + p.unidad : ''));
    cat.total++;
  });
  j.proveedores.forEach(function (p) { cat.proveedores[p.proveedor] = cat.proveedores[p.proveedor] || []; });
  try { cache.put('catalogoPedidos', JSON.stringify(cat), 3600); } catch (err) { /* demasiado grande: no pasa nada */ }
  return cat;
}

var ESQUEMA_LINEAS_FACTURA = {
  type: 'object',
  properties: {
    proveedor_pedidos: { type: 'string' },
    lineas: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          texto: { type: 'string' },
          cantidad: { type: 'string' },
          unidad: { type: 'string' },
          precio_unitario: { type: 'string' },
          id_producto: { type: 'string' }
        },
        required: ['texto', 'cantidad', 'unidad', 'precio_unitario', 'id_producto'],
        additionalProperties: false
      }
    }
  },
  required: ['proveedor_pedidos', 'lineas'],
  additionalProperties: false
};

var ESQUEMA_PROVEEDOR_PEDIDOS_ = {
  type: 'object',
  properties: { proveedor_pedidos: { type: 'string' } },
  required: ['proveedor_pedidos'],
  additionalProperties: false
};

// A cuál proveedor de Pedidos corresponde la factura, sin mandar la foto: el
// que diga quien la manda, el que comparta una palabra con el nombre leído
// o, si no, se pregunta a la IA solo con los nombres (casi gratis). '' si
// no es de ninguno: entonces no se lee la factura (es lo que más gasta).
function proveedorDePedidos_(data, proveedor, nombres) {
  if (data.proveedorPedidos && nombres.indexOf(data.proveedorPedidos) >= 0) return data.proveedorPedidos;
  var leido = (proveedor === 'Sin proveedor' ? '' : proveedor) + ' ' + String(data.detalle || '');
  var palabras = sinAcentos_(leido).split(' ').filter(function (w) { return w.length >= 3; });
  if (!palabras.length) return '';
  var porTexto = nombres.filter(function (n) {
    var texto = ' ' + sinAcentos_(n) + ' ';
    return palabras.some(function (w) { return texto.indexOf(' ' + w + ' ') >= 0; });
  });
  if (porTexto.length === 1) return porTexto[0];
  var r = pedirJsonAClaude_([{ type: 'text', text:
    'Una factura de un proveedor de un restaurante se leyó con este emisor: "' + leido.trim() + '".\n' +
    'proveedor_pedidos: a cuál de estos proveedores corresponde, escrito EXACTAMENTE como en la lista ' +
    '(vale aunque uno sea la razón social, esté abreviado o sea otro nombre comercial). Si no es ninguno o no estás seguro, vacío.\n\n' +
    nombres.join('\n') }], ESQUEMA_PROVEEDOR_PEDIDOS_, 'low', MODELO_BARATO_);
  if (!r.ok) throw new Error(r.error);
  return nombres.indexOf(r.proveedor_pedidos) >= 0 ? r.proveedor_pedidos : '';
}

function mandarPreciosAPedidos_(data, proveedor, fecha, foto) {
  var cat = catalogoPedidos_();
  var nombres = Object.keys(cat.proveedores);
  var deCual = proveedorDePedidos_(data, proveedor, nombres);
  if (!deCual || !cat.proveedores[deCual].length) return { ok: true, noEsDePedidos: true };
  var instrucciones =
    'Esta es la foto de una factura, albarán o ticket del proveedor "' + deCual + '" del restaurante "Buenas y Santas" (el cliente). ' +
    'Abajo están los productos de ese proveedor en la app de pedidos del restaurante (ID | producto | unidad).\n\n' +
    '- proveedor_pedidos: escribe "' + deCual + '"; si la factura claramente NO es de ese proveedor, déjalo vacío y devuelve lineas vacío.\n' +
    '- lineas: una por cada producto facturado, en el orden de la factura. NO incluyas portes, envases o cascos retornables, ' +
    'recargos, descuentos generales, subtotales ni totales.\n' +
    '  · texto: la descripción tal como sale en la factura.\n' +
    '  · cantidad: la cantidad facturada en esa línea.\n' +
    '  · unidad: la unidad en la que se cobra (unidad, caja, kg, botella, paquete...), como lo indique la factura; vacío si no lo dice.\n' +
    '  · precio_unitario: precio de UNA unidad de esa línea, SIN IVA y con los descuentos de la propia línea ya aplicados. ' +
    'Si la factura solo trae el importe de la línea, divídelo entre la cantidad.\n' +
    '  · id_producto: el ID del producto de la lista que sea el mismo artículo, aunque esté escrito distinto ' +
    '(abreviaturas, marca, formato). Si no corresponde a ninguno o no estás seguro, vacío.\n' +
    'Números con punto decimal y sin símbolo de euro (ej. 12.5). Si un dato no se lee con seguridad, déjalo vacío en vez de inventarlo.\n\n' +
    'PRODUCTOS DE ' + deCual + ':\n' + cat.proveedores[deCual].join('\n');
  var r = pedirJsonAClaude_(imagenesFactura_(data).concat([
    { type: 'text', text: instrucciones }
  ]), ESQUEMA_LINEAS_FACTURA, 'medium', MODELO_BARATO_);
  if (!r.ok) return { ok: false, error: r.error };
  if (r.proveedor_pedidos !== deCual || !(r.lineas || []).length) {
    return { ok: true, noEsDePedidos: true };
  }
  try {
    return pedirAPedidos_({
      accion: 'precios', proveedor: deCual, factura: String(data.numero_factura || '').trim(),
      fecha: fecha, foto: foto, reemplazar: !!data.reemplazar,
      lineas: r.lineas.map(function (l) {
        return { texto: l.texto, id: l.id_producto, cantidad: l.cantidad, unidad: l.unidad, precio: l.precio_unitario };
      })
    });
  } catch (err) {
    return { ok: false, error: err.message };
  }
}


// ============================================================================
// Precios de proveedores que NO están en la app de Pedidos (Frutapro): para
// el escandallo y para ver sus cambios de precio en "Compras de la semana"
// (InformeSemanalPedidos.gs). Van a la pestaña "Precios proveedores" del
// Índice, con las mismas columnas que APP_Precios de Pedidos:
//   Fecha · Proveedor · Factura · En la factura · ID producto · Producto ·
//   Cantidad · Unidad · Precio · Precio anterior · Cambio % · Foto · Registrado
// Sin catálogo: el mismo producto se reconoce por cómo sale escrito en la
// factura (sin mayúsculas, tildes ni signos). La misma factura escaneada
// otra vez reemplaza sus líneas.
// ============================================================================
var PRECIOS_PROPIOS_PESTANA_ = 'Precios proveedores';
// Proveedor → palabras que lo reconocen (en el proveedor que manda la app).
var PRECIOS_PROPIOS_PROVEEDORES_ = { 'Frutapro': ['frutapro', 'mediterranean onions'] };

function proveedorConPreciosPropios_(proveedor) {
  var n = ' ' + sinAcentos_(proveedor) + ' ';
  var nombres = Object.keys(PRECIOS_PROPIOS_PROVEEDORES_);
  for (var i = 0; i < nombres.length; i++) {
    if (PRECIOS_PROPIOS_PROVEEDORES_[nombres[i]].some(function (w) { return n.indexOf(' ' + w + ' ') >= 0; })) return nombres[i];
  }
  return '';
}

function clavePrecioPropio_(texto) { return sinAcentos_(texto); }

var ESQUEMA_LINEAS_SIN_CATALOGO_ = {
  type: 'object',
  properties: {
    lineas: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          texto: { type: 'string' }, producto: { type: 'string' }, cantidad: { type: 'string' },
          unidad: { type: 'string' }, precio_unitario: { type: 'string' }
        },
        required: ['texto', 'producto', 'cantidad', 'unidad', 'precio_unitario'],
        additionalProperties: false
      }
    }
  },
  required: ['lineas'],
  additionalProperties: false
};

function hojaPreciosPropios_() {
  var ss = SpreadsheetApp.openById(INDICE_SHEET_ID_);
  var hoja = ss.getSheetByName(PRECIOS_PROPIOS_PESTANA_);
  if (!hoja) {
    hoja = ss.insertSheet(PRECIOS_PROPIOS_PESTANA_);
    hoja.appendRow(['Fecha', 'Proveedor', 'Factura', 'En la factura', 'ID producto', 'Producto', 'Cantidad', 'Unidad',
      'Precio', 'Precio anterior', 'Cambio %', 'Foto', 'Registrado']);
    hoja.setFrozenRows(1);
    hoja.getRange('1:1').setFontWeight('bold');
  }
  return hoja;
}

// Filas de la pestaña (para "Compras de la semana"); [] si todavía no existe.
function filasPreciosPropios_() {
  var hoja = SpreadsheetApp.openById(INDICE_SHEET_ID_).getSheetByName(PRECIOS_PROPIOS_PESTANA_);
  if (!hoja || hoja.getLastRow() < 2) return [];
  return hoja.getRange(2, 1, hoja.getLastRow() - 1, 13).getValues();
}

function guardarPreciosPropios_(data, proveedor, fecha, foto) {
  var instrucciones =
    'Esta es la foto de una factura o albarán del proveedor "' + proveedor + '" del restaurante "Buenas y Santas" (el cliente).\n' +
    '- lineas: una por cada producto facturado, en el orden de la factura. NO incluyas portes, envases o cajas retornables, ' +
    'recargos, descuentos generales, subtotales ni totales.\n' +
    '  · texto: la descripción tal como sale en la factura (igual cada vez, para poder comparar precios entre facturas).\n' +
    '  · producto: el nombre del producto claro y corto, en minúsculas con la primera en mayúscula (ej. "Tomate pera").\n' +
    '  · cantidad: la cantidad facturada en esa línea.\n' +
    '  · unidad: la unidad en la que se cobra (kg, unidad, caja, manojo...), como lo indique la factura; vacío si no lo dice.\n' +
    '  · precio_unitario: precio de UNA unidad de esa línea, SIN IVA y con los descuentos de la propia línea ya aplicados. ' +
    'Si la factura solo trae el importe de la línea, divídelo entre la cantidad.\n' +
    'Números con punto decimal y sin símbolo de euro (ej. 12.5). Si un dato no se lee con seguridad, déjalo vacío en vez de inventarlo.';
  var r = pedirJsonAClaude_(imagenesFactura_(data).concat([{ type: 'text', text: instrucciones }]),
    ESQUEMA_LINEAS_SIN_CATALOGO_, 'medium', MODELO_BARATO_);
  if (!r.ok) return { ok: false, error: r.error };
  var lineas = (r.lineas || []).filter(function (l) { return String(l.texto || '').trim() && Number(l.precio_unitario) > 0; });
  if (!lineas.length) return { ok: true, lineas: 0 };

  var factura = String(data.numero_factura || '').trim();
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var hoja = hojaPreciosPropios_();
    var filas = hoja.getLastRow() > 1 ? hoja.getRange(2, 1, hoja.getLastRow() - 1, 13).getValues() : [];
    // La misma factura otra vez: se borran sus líneas de antes.
    if (factura) {
      for (var i = filas.length - 1; i >= 0; i--) {
        if (filas[i][1] === proveedor && String(filas[i][2]).trim() === factura) { hoja.deleteRow(i + 2); filas.splice(i, 1); }
      }
    }
    // Precio anterior: el de la factura más reciente de antes de esta fecha.
    var anterior = {};
    filas.forEach(function (f) {
      if (f[1] !== proveedor) return;
      var fe = f[0] instanceof Date ? Utilities.formatDate(f[0], 'Europe/Madrid', 'yyyy-MM-dd') : String(f[0]);
      if (fe > fecha) return;
      var k = clavePrecioPropio_(f[3]);
      if (!anterior[k] || anterior[k].fecha <= fe) anterior[k] = { fecha: fe, precio: Number(f[8]) || 0 };
    });
    var ahora = new Date();
    var nuevas = lineas.map(function (l) {
      var precio = Math.round(Number(l.precio_unitario) * 10000) / 10000;
      var antes = (anterior[clavePrecioPropio_(l.texto)] || {}).precio || '';
      return [fechaComoDate_(fecha), proveedor, factura, String(l.texto).trim(), '', String(l.producto || l.texto).trim(),
        Number(l.cantidad) || l.cantidad, l.unidad, precio, antes,
        antes ? Math.round((precio - antes) / antes * 1000) / 10 : '', foto, ahora];
    });
    hoja.getRange(hoja.getLastRow() + 1, 1, nuevas.length, 13).setValues(nuevas);
    return { ok: true, lineas: nuevas.length };
  } finally {
    lock.releaseLock();
  }
}


// ============================================================================
// Releer los precios de una factura que ya está en Drive (botón "Leer precios
// de una factura" de la app de Pedidos). La app de Pedidos no tiene sesión del
// cierre: por eso solo se pueden listar facturas de la carpeta de facturas y
// releer como mucho RELECTURAS_AL_DIA_ al día.
// ============================================================================
var RELECTURAS_AL_DIA_ = 40;

function sinAcentos_(s) {
  return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
}

// Facturas de los últimos ~2 meses: { facturas: [{ id, nombre, carpeta, fecha,
// url, deEste, conPrecios }] }. deEste: la carpeta o el nombre llevan el nombre
// del proveedor de Pedidos. conPrecios: sus precios ya están en Pedidos.
function facturasParaPedidos_(data) {
  var proveedor = String(data.proveedor || '').trim();
  var palabras = sinAcentos_(proveedor).split(' ').filter(function (w) { return w.length >= 3; });
  var raiz = carpetaFacturas_();
  var hoy = new Date(), limite = Utilities.formatDate(new Date(hoy.getTime() - 62 * 86400000), 'Europe/Madrid', 'yyyy-MM-dd');
  var meses = {};
  for (var d = 0; d <= 62; d += 15) {
    meses[Utilities.formatDate(new Date(hoy.getTime() - d * 86400000), 'Europe/Madrid', 'yyyy/MM')] = 1;
  }
  meses[Utilities.formatDate(new Date(hoy.getTime() - 62 * 86400000), 'Europe/Madrid', 'yyyy/MM')] = 1;
  var lista = [];
  Object.keys(meses).forEach(function (am) {
    var anios = raiz.getFoldersByName(am.slice(0, 4));
    if (!anios.hasNext()) return;
    var mesesIt = anios.next().getFoldersByName(am.slice(5, 7));
    if (!mesesIt.hasNext()) return;
    var provs = mesesIt.next().getFolders();
    while (provs.hasNext()) {
      var carpeta = provs.next();
      var archivos = carpeta.getFiles();
      while (archivos.hasNext()) {
        var f = archivos.next();
        if (f.isTrashed()) continue;
        var nombre = f.getName();
        var fecha = /^\d{4}-\d{2}-\d{2}/.test(nombre) ? nombre.slice(0, 10) : Utilities.formatDate(f.getDateCreated(), 'Europe/Madrid', 'yyyy-MM-dd');
        if (fecha < limite) continue;
        var texto = ' ' + sinAcentos_(carpeta.getName() + ' ' + nombre) + ' ';
        lista.push({
          id: f.getId(), nombre: nombre.replace(/\.(pdf|jpg|png)$/i, ''), carpeta: carpeta.getName(), fecha: fecha, url: f.getUrl(),
          deEste: palabras.some(function (w) { return texto.indexOf(' ' + w + ' ') >= 0; })
        });
      }
    }
  });
  // Cuáles ya tienen sus precios en Pedidos: por el enlace de la factura o,
  // si el PDF se volvió a guardar (otro enlace), por su número en el nombre.
  try {
    var j = pedirAPedidos_({ accion: 'facturasConPrecios' });
    var hechas = {}, numeros = [];
    (j.facturas || []).forEach(function (x) {
      if (x.fotoId) hechas[x.fotoId] = x.filas;
      var n = sinAcentos_(x.factura);
      if (n.replace(/\D/g, '').length >= 3) numeros.push({ n: n, filas: x.filas });
    });
    lista.forEach(function (x) {
      var nombre = ' ' + sinAcentos_(x.nombre);
      var porNumero = numeros.filter(function (k) { return nombre.slice(-k.n.length - 1) === ' ' + k.n; })[0];
      x.conPrecios = hechas[x.id] || (porNumero ? porNumero.filas : 0);
    });
  } catch (err) { /* sin marcar: no pasa nada */ }
  lista.sort(function (a, b) { return b.fecha.localeCompare(a.fecha) || a.nombre.localeCompare(b.nombre); });
  return { ok: true, facturas: lista };
}

function dentroDeFacturas_(archivo) {
  var raizId = carpetaFacturas_().getId();
  var nivel = [archivo];
  for (var i = 0; i < 5 && nivel.length; i++) {
    var siguientes = [];
    for (var k = 0; k < nivel.length; k++) {
      var padres = nivel[k].getParents();
      while (padres.hasNext()) {
        var p = padres.next();
        if (p.getId() === raizId) return true;
        siguientes.push(p);
      }
    }
    nivel = siguientes;
  }
  return false;
}

// { idArchivo, proveedor (el de Pedidos), reemplazar } → lee la factura y
// manda sus precios a Pedidos, como si se acabara de escanear.
function preciosDesdeDrive_(data) {
  var archivo;
  try { archivo = DriveApp.getFileById(String(data.idArchivo || '')); } catch (err) { return { ok: false, error: 'No encuentro esa factura en Drive' }; }
  if (archivo.isTrashed() || !dentroDeFacturas_(archivo)) return { ok: false, error: 'Esa factura no está en la carpeta de facturas' };
  var props = PropertiesService.getScriptProperties();
  var claveDia = 'relecturas_' + Utilities.formatDate(new Date(), 'Europe/Madrid', 'yyyy-MM-dd');
  var hechas = Number(props.getProperty(claveDia) || 0);
  if (hechas >= RELECTURAS_AL_DIA_) return { ok: false, error: 'Hoy ya se han leído muchas facturas. Prueba mañana.' };
  props.setProperty(claveDia, String(hechas + 1));

  var blob = archivo.getBlob(), tipo = blob.getContentType();
  var b64 = Utilities.base64Encode(blob.getBytes());
  var carpeta = archivo.getParents().hasNext() ? archivo.getParents().next().getName() : '';
  // Nombre: "<AAAA-MM-DD> <proveedor> [<súper>] <nº factura>.pdf"
  var nombre = archivo.getName().replace(/\.(pdf|jpg|png)$/i, '');
  var fecha = /^\d{4}-\d{2}-\d{2}/.test(nombre) ? nombre.slice(0, 10) : Utilities.formatDate(archivo.getDateCreated(), 'Europe/Madrid', 'yyyy-MM-dd');
  var resto = nombre.replace(/^\d{4}-\d{2}-\d{2}\s*/, '');
  if (carpeta && resto.indexOf(carpeta + ' ') === 0) resto = resto.slice(carpeta.length + 1);
  var lectura = {
    detalle: carpeta, numero_factura: resto.trim(), proveedorPedidos: String(data.proveedor || '').trim(),
    reemplazar: !!data.reemplazar
  };
  if (tipo === 'application/pdf') lectura.pdfParaClaude = b64;
  else { lectura.imageBase64 = b64; lectura.mediaType = tipo; }
  var r = mandarPreciosAPedidos_(lectura, carpeta || 'Sin proveedor', fecha, archivo.getUrl());
  if (r && r.noEsDePedidos) return { ok: false, error: 'En esa factura no se ha reconocido ningún producto de este proveedor' };
  return r;
}

// ============================================================================
// Acceso de la app de Pedidos: solo las cuentas de CUENTAS_PERMITIDAS_
// (CierreCaja.gs). La app inicia sesión con Google (el mismo ID de cliente que
// la caja) y manda aquí el "credential". Si la cuenta vale, se le devuelve un
// "acceso" firmado con CLAVE_PEDIDOS_ (ClavePedidos.gs) que dura DIAS_SESION_
// días. El script de Pedidos, que no puede hablar con Google, solo comprueba la
// firma. Quitar una cuenta de la lista la deja fuera aquí al momento; en
// Pedidos, cuando venza su acceso (o enseguida, si se cambia la clave).
// ============================================================================
function firmaPedidos_(texto) {
  return Utilities.computeHmacSha256Signature(texto, CLAVE_PEDIDOS_).map(function (b) {
    return ('0' + (b & 255).toString(16)).slice(-2);
  }).join('');
}

function sesionPedidos_(idToken) {
  if (!idToken) return { ok: false, error: 'No llegó el inicio de sesión de Google.' };
  var resp = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken), { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) return { ok: false, error: 'Google no dio por válido el inicio de sesión. Prueba de nuevo.' };
  var info = JSON.parse(resp.getContentText());
  if (info.aud !== GOOGLE_CLIENT_ID_) return { ok: false, error: 'El inicio de sesión no es de esta app.' };
  var email = String(info.email || '').toLowerCase();
  if (String(info.email_verified) !== 'true') return { ok: false, error: 'La cuenta ' + email + ' no está verificada por Google.' };
  if (CUENTAS_PERMITIDAS_.indexOf(email) === -1) return { ok: false, error: 'La cuenta ' + email + ' no tiene acceso a los pedidos.' };
  var expira = Date.now() + DIAS_SESION_ * 24 * 60 * 60 * 1000;
  var texto = email + '|' + expira;
  return { ok: true, acceso: texto + '|' + firmaPedidos_(texto), email: email, expira: expira };
}

// Devuelve la cuenta del acceso, o null si no vale.
function accesoPedidosValido_(acceso) {
  var p = String(acceso || '').split('|');
  if (p.length !== 3 || !(Number(p[1]) > Date.now())) return null;
  if (firmaPedidos_(p[0] + '|' + p[1]) !== p[2]) return null;
  return CUENTAS_PERMITIDAS_.indexOf(p[0]) === -1 ? null : p[0];
}