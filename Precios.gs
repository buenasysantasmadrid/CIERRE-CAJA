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
  try { res.precios = mandarPreciosAPedidos_(data, proveedor, fecha, res.foto || ''); } catch (err) { res.errorPrecios = String(err); }
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
  var nombre = sinCaracteresRaros_(fecha + ' ' + proveedor + ' ' +
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
  var lista = (data.imagenes && data.imagenes.length) ? data.imagenes : [data.imageBase64];
  return lista.map(function (b64) {
    return { type: 'image', source: { type: 'base64', media_type: data.mediaType || 'image/jpeg', data: b64 } };
  });
}

// Proveedores y productos de la app de Pedidos (se guardan una hora).
function catalogoPedidos_() {
  var cache = CacheService.getScriptCache();
  var guardado = cache.get('catalogoPedidos');
  if (guardado) return JSON.parse(guardado);
  var resp = UrlFetchApp.fetch(PEDIDOS_URL_, {
    method: 'post', contentType: 'text/plain;charset=utf-8',
    payload: JSON.stringify({ accion: 'catalogo' }), muteHttpExceptions: true
  });
  var j = JSON.parse(resp.getContentText());
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

function mandarPreciosAPedidos_(data, proveedor, fecha, foto) {
  var cat = catalogoPedidos_();
  var nombres = Object.keys(cat.proveedores);
  var lista = nombres.map(function (n) {
    return '### ' + n + '\n' + (cat.proveedores[n].length ? cat.proveedores[n].join('\n') : '(sin productos)');
  }).join('\n\n');
  var instrucciones =
    'Esta es la foto de una factura, albarán o ticket de un proveedor del restaurante "Buenas y Santas" (el cliente). ' +
    'En la primera lectura salió como proveedor: "' + proveedor + '".\n' +
    'Abajo está la lista de proveedores de la app de pedidos del restaurante, cada uno con sus productos (ID | producto | unidad).\n\n' +
    '- proveedor_pedidos: a cuál de esos proveedores corresponde el emisor de la factura, escrito EXACTAMENTE como en la lista ' +
    '(vale aunque en la factura salga con su razón social, abreviado o con otro nombre comercial). ' +
    'Si no es ninguno de la lista o no estás seguro, déjalo vacío y devuelve lineas vacío.\n' +
    '- lineas: una por cada producto facturado, en el orden de la factura. NO incluyas portes, envases o cascos retornables, ' +
    'recargos, descuentos generales, subtotales ni totales.\n' +
    '  · texto: la descripción tal como sale en la factura.\n' +
    '  · cantidad: la cantidad facturada en esa línea.\n' +
    '  · unidad: la unidad en la que se cobra (unidad, caja, kg, botella, paquete...), como lo indique la factura; vacío si no lo dice.\n' +
    '  · precio_unitario: precio de UNA unidad de esa línea, SIN IVA y con los descuentos de la propia línea ya aplicados. ' +
    'Si la factura solo trae el importe de la línea, divídelo entre la cantidad.\n' +
    '  · id_producto: el ID del producto de ESE proveedor en la lista que sea el mismo artículo, aunque esté escrito distinto ' +
    '(abreviaturas, marca, formato). Si no corresponde a ninguno o no estás seguro, vacío.\n' +
    'Números con punto decimal y sin símbolo de euro (ej. 12.5). Si un dato no se lee con seguridad, déjalo vacío en vez de inventarlo.\n\n' +
    'PROVEEDORES Y PRODUCTOS DE LA APP DE PEDIDOS:\n\n' + lista;
  var r = pedirJsonAClaude_(imagenesFactura_(data).concat([
    { type: 'text', text: instrucciones }
  ]), ESQUEMA_LINEAS_FACTURA, 'medium');
  if (!r.ok) return { ok: false, error: r.error };
  if (!r.proveedor_pedidos || nombres.indexOf(r.proveedor_pedidos) < 0 || !(r.lineas || []).length) {
    return { ok: true, noEsDePedidos: true };
  }
  var resp = UrlFetchApp.fetch(PEDIDOS_URL_, {
    method: 'post', contentType: 'text/plain;charset=utf-8', muteHttpExceptions: true,
    payload: JSON.stringify({
      accion: 'precios', proveedor: r.proveedor_pedidos, factura: String(data.numero_factura || '').trim(),
      fecha: fecha, foto: foto,
      lineas: r.lineas.map(function (l) {
        return { texto: l.texto, id: l.id_producto, cantidad: l.cantidad, unidad: l.unidad, precio: l.precio_unitario };
      })
    })
  });
  return JSON.parse(resp.getContentText());
}
