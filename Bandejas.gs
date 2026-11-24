// ============================================================================
// BANDEJAS DE FACTURAS — las facturas que llegan en PDF (por email) entran
// sin imprimirlas ni escanearlas. Dentro de "Facturas proveedores" está la
// carpeta de entrada "📥 Facturas nuevas" (desde el móvil, "Guardar en
// Drive"): ahí van todas y se decide sola de cuál de las dos clases es
// (ver decidirTipoFactura_ más abajo). Lo que no está claro (mercadería sin
// su albarán en la caja y sin decir que se paga por transferencia) no se
// carga: queda en su subcarpeta "⚠️ Revisar" y se avisa.
// Antes eran dos carpetas, una por clase; si todavía existen, lo que se
// guarde en ellas se sigue procesando como antes:
//   · "📥 Facturas de albarán": compras a un proveedor cuyo albarán ya se
//     cargó en la caja cuando llegó la mercadería. No se carga otra vez: se
//     busca en ALBARANES (mismo número, un albarán que la factura menciona,
//     o el mismo importe en esas fechas) y ese gasto pasa a ser la factura
//     (número, fecha, IVA e importe de la factura), como la factura de un
//     ticket. Si no aparece, se avisa.
//   · "📥 Facturas pagadas por transferencia": servicios y demás que se
//     pagan desde el banco. Se escriben directo en ALBARANES (la pestaña del
//     proveedor, o SERVICIOS si no tiene), con forma TRANSFERENCIA y día de
//     pago = fecha de la factura. No tocan ninguna caja. En la columna Z
//     llevan el ID "pdf-<archivo>" (no están en "IDs app", así que los
//     guardados de la caja no las borran).
// Cada hora (activador procesarBandejasFacturas, se pone solo con el primer
// guardado de la caja) Claude lee lo que haya y el archivo se mueve a
// "Facturas proveedores / AAAA / MM / Proveedor", con el mismo nombre que
// las escaneadas: es lo que se entrega a los contables. Todo se apunta para
// el resumen de los lunes (ResumenFacturas.gs); los problemas van además al
// resumen de avisos. Lo que no se puede
// leer queda en la subcarpeta "⚠️ No se pudo leer" de su bandeja.
// ============================================================================

var BANDEJA_UNICA_ = '📥 Facturas nuevas';
// Facturas externas: a nuestro nombre pero no son gasto del restaurante. Van
// a FC VYR (no suman en las cuentas del mes, sí van a los contables).
// (En la app y en Drive se llaman "NO USAR". Antes era "📥 Facturas
// externas": si todavía está, se le cambia el nombre.)
var BANDEJA_EXTERNAS_ = '📥 Facturas NO USAR';
var BANDEJA_EXTERNAS_VIEJA_ = '📥 Facturas externas';
var BANDEJA_ALBARAN_ = '📥 Facturas de albarán'; // vieja: solo se procesa si existe
var BANDEJA_TRANSFERENCIA_ = '📥 Facturas pagadas por transferencia'; // vieja: ídem
var BANDEJA_NO_LEIDAS_ = '⚠️ No se pudo leer';
var BANDEJA_REVISAR_ = '⚠️ Revisar';
var BANDEJA_REPETIDAS_ = '🗑️ Repetidas (ya estaban, se pueden borrar)';

// Lo mismo que se lee al escanear, y además cómo dice la factura que se paga
// (para saber, en "Facturas nuevas", si es de transferencia).
function esquemaFacturaBandeja_() {
  var e = JSON.parse(JSON.stringify(ESQUEMA_FACTURA));
  e.properties.forma_pago = { type: 'string', enum: ['transferencia', 'domiciliacion', 'efectivo', 'tarjeta', 'contado', ''] };
  e.required = e.required.concat(['forma_pago']);
  return e;
}
var INSTRUCCIONES_FORMA_PAGO_ =
  '\n- forma_pago: cómo dice la factura que se paga: "transferencia" (transferencia bancaria, o un IBAN/cuenta del proveedor para ingresar), ' +
  '"domiciliacion" (domiciliación, recibo domiciliado, adeudo, cargo en cuenta), "efectivo", "tarjeta", o "contado" (al contado, sin decir cómo). ' +
  'Vacío si la factura no lo dice.';
var BANDEJA_TIEMPO_MAX_MS_ = 4.5 * 60 * 1000; // el activador corta a los 6 minutos
var BANDEJA_INTENTOS_ = 3; // fallos (ej. Claude saturado) antes de dejarla en "No se pudo leer"
var BANDEJA_PROVEEDORES_GENERICOS_ = ['VARIOS', 'SUPER', 'EXTRAS', 'SERVICIOS', 'GLOVO', 'FC VYR'];

// Desde cada guardado de la caja (doPost): crea las dos carpetas y el
// activador de cada hora, si no están.
function asegurarBandejasFacturas_() {
  try { asegurarActivadorPaqueteContables_(); } catch (err) { Logger.log('Paquete contables: ' + err); } // ver Contables.gs
  try {
    var cache = CacheService.getScriptCache();
    if (cache.get('bandejas_facturas')) return;
    var raiz = DriveApp.getFolderById(CARPETA_FACTURAS_ID_);
    carpetaHija_(raiz, BANDEJA_UNICA_);
    renombrarBandejaExternasVieja_(raiz);
    carpetaHija_(raiz, BANDEJA_EXTERNAS_);
    var hay = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'procesarBandejasFacturas'; });
    if (!hay) ScriptApp.newTrigger('procesarBandejasFacturas').timeBased().everyHours(1).create();
    asegurarActivadorFacturasQueFaltan_(); // el email del día 3 (FacturasQueFaltan.gs)
    cache.put('bandejas_facturas', '1', 21600);
  } catch (err) {
    Logger.log('Bandejas de facturas: ' + err);
  }
}

function renombrarBandejaExternasVieja_(raiz) {
  var vieja = raiz.getFoldersByName(BANDEJA_EXTERNAS_VIEJA_);
  if (vieja.hasNext() && !raiz.getFoldersByName(BANDEJA_EXTERNAS_).hasNext()) vieja.next().setName(BANDEJA_EXTERNAS_);
}

// La lista de proveedores que manda la app al escanear (con los nombres
// con que salen en la factura): se guarda para leer las de las bandejas
// igual que las escaneadas.
function guardarProveedoresParaLeer_(lista) {
  try {
    if (!lista || !lista.length) return;
    var texto = JSON.stringify(lista);
    var props = PropertiesService.getScriptProperties();
    if (props.getProperty('proveedores_ia') !== texto) props.setProperty('proveedores_ia', texto);
  } catch (err) { /* la próxima vez */ }
}

// Si todavía no se escaneó nada desde que existe esto, los nombres de las
// pestañas de ALBARANES del año.
function proveedoresParaLeer_() {
  try {
    var guardada = PropertiesService.getScriptProperties().getProperty('proveedores_ia');
    if (guardada) return JSON.parse(guardada);
  } catch (err) { /* abajo */ }
  var id = buscarEnIndice_('ALBARANES', Utilities.formatDate(new Date(), 'Europe/Madrid', 'yyyy'));
  if (!id) return [];
  return SpreadsheetApp.openById(id).getSheets().map(function (h) { return h.getName(); }).filter(function (n) {
    return ALBARANES_PESTANAS_EXCLUIDAS_.indexOf(n) === -1 && BANDEJA_PROVEEDORES_GENERICOS_.indexOf(normalizarClave_(n)) === -1 && n !== INFORME_MES_PESTANA_;
  });
}

// El nombre de la lista (sin "(en la factura: ...)") que eligió Claude, o ''.
function proveedorDeLaListaLeida_(nombre, lista) {
  var buscado = normalizarClave_(nombre);
  if (!buscado) return '';
  for (var i = 0; i < lista.length; i++) {
    var base = String(lista[i]).split(' (en la factura:')[0].trim();
    if (normalizarClave_(base) === buscado) return base;
  }
  return '';
}

// Lo que corre cada hora (también lo puede lanzar Claude: ?procesarBandejas=1,
// ver doGet). Devuelve lo hecho, para verlo.
function procesarBandejasFacturas() {
  var cache = CacheService.getScriptCache();
  if (cache.get('bandejas_corriendo')) return { ok: false, error: 'ya se están procesando' }; // la vuelta anterior todavía no terminó
  cache.put('bandejas_corriendo', '1', 400);
  try {
    var inicio = Date.now();
    var raiz = DriveApp.getFolderById(CARPETA_FACTURAS_ID_);
    var lista = proveedoresParaLeer_();
    var hechas = [], salida ={ ok: true, hechas: hechas, problemas: [], seReintentan: [] };
    renombrarBandejaExternasVieja_(raiz);
    [[BANDEJA_UNICA_, 'auto'], [BANDEJA_EXTERNAS_, 'externa'], [BANDEJA_ALBARAN_, 'albaran'], [BANDEJA_TRANSFERENCIA_, 'transferencia']].forEach(function (b) {
      // Las dos viejas no se vuelven a crear: solo si todavía están.
      var it = raiz.getFoldersByName(b[0]);
      var bandeja = it.hasNext() ? it.next() : (b[1] === 'auto' || b[1] === 'externa' ? carpetaHija_(raiz, b[0]) : null);
      if (!bandeja) return;
      if (b[1] === 'auto' || b[1] === 'externa') { carpetaHija_(bandeja, BANDEJA_NO_LEIDAS_); carpetaHija_(bandeja, BANDEJA_REVISAR_); } // que se vean desde el principio
      var archivos = bandeja.getFiles();
      // Y las que esperan en "⚠️ Revisar" a que se cargue su foto o albarán
      // en la caja (un día; después se cargan solas como NO PAGADO).
      var esperan = [];
      if (b[1] === 'auto') {
        var enRevisar = carpetaHija_(bandeja, BANDEJA_REVISAR_).getFiles();
        while (enRevisar.hasNext()) {
          var e = enRevisar.next(), le = leerEsperando_(e);
          if (!le) continue;
          esperan.push(e);
        }
      }
      var siguiente = function () { return archivos.hasNext() ? archivos.next() : esperan.shift(); };
      var archivo;
      while ((archivos.hasNext() || esperan.length) && Date.now() - inicio < BANDEJA_TIEMPO_MAX_MS_) {
        archivo = siguiente();
        var r;
        try {
          r = procesarFacturaDeBandeja_(archivo, b[1], raiz, lista, bandeja);
        } catch (err) {
          salida.seReintentan.push(archivo.getName() + ': ' + (err.message || err));
          r = reintentarDespues_(archivo, bandeja, String(err.message || err));
        }
        if (r.hecha) {
          hechas.push(r.hecha);
          apuntarMovFactura_('cargada', r.proveedor, r.importe, 'Por la carpeta: ' + r.hecha, r.enlace); // para el resumen de los lunes
        }
        // Para el resumen de los lunes: solo lo que hay que mirar (ver ResumenFacturas.gs).
        (r.avisos || []).forEach(function (a) { apuntarMovFactura_(a.tipo, r.proveedor, r.importe, a.que, r.enlace); });
        if (r.problema) {
          salida.problemas.push(r.problema);
          avisarPorEmail_('una factura de la carpeta "' + b[0] + '" ' + (r.hecha ? 'tiene algo que revisar' : 'no se cargó'),
            r.problema + (r.enlace ? '\nEl PDF: ' + r.enlace : ''));
        }
      }
    });
    return salida;
  } finally {
    cache.remove('bandejas_corriendo');
  }
}

// Facturas en "⚠️ Revisar" que esperan a que su foto o su albarán se cargue
// en la caja: en la descripción llevan "esperando: {desde, r}" (r = lo leído).
var ESPERANDO_ = 'esperando: ';
// 24/11/2026 (Luciana): hay facturas que no tienen albarán. Se espera un día
// por si se carga la foto o el albarán en la caja (Envapro llegó por email la
// noche antes de cargar la foto); después se carga sola como NO PAGADO.
var ESPERANDO_HORAS_ = 24;
function esperaVencida_(e) { return !!e && Date.now() - new Date(e.desde).getTime() > ESPERANDO_HORAS_ * 3600000; }
function leerEsperando_(archivo) {
  var d = String(archivo.getDescription() || '');
  if (d.indexOf(ESPERANDO_) !== 0) return null;
  try { return JSON.parse(d.slice(ESPERANDO_.length)); } catch (e) { return null; }
}

// ¿Ya está en ALBARANES una fila con este mismo nº de factura (la foto de la
// factura cargada en la caja)? Si el PDF de esa fila se escaneó como
// albarán, no: entonces es una factura de albarán normal.
function facturaYaCargada_(ss, raiz, proveedor, numero) {
  var k = claveNumeroFactura_(numero);
  if (k.length < 4 || !proveedor) return null;
  var hoja = pestanaAlbaranesParaProveedor_(ss, proveedor);
  if (!hoja) return null;
  var filas = filasDelAnioAlbaranes_(hoja);
  var f = filas.filter(function (x) { return claveNumeroFactura_(x.numero) === k; })[0];
  if (!f) return null;
  var carpeta = sinCaracteresRaros_(proveedor), dirFoto = null;
  if (f.fecha) {
    var fe = Utilities.formatDate(f.fecha, 'Europe/Madrid', 'yyyy-MM-dd');
    var itA = raiz.getFoldersByName(fe.slice(0, 4)), itM = itA.hasNext() ? itA.next().getFoldersByName(fe.slice(5, 7)) : null;
    var itP = itM && itM.hasNext() ? itM.next().getFoldersByName(carpeta) : null;
    if (itP && itP.hasNext()) {
      dirFoto = itP.next();
      var fs = dirFoto.getFiles();
      while (fs.hasNext()) {
        var a = fs.next();
        var nombre = a.getName().replace(/\.[a-z0-9]{2,4}$/i, '').replace(/^\d{4}-\d{2}-\d{2}\s*/, '');
        if (claveNumeroFactura_(nombre).endsWith(k) && String(a.getDescription() || '').toLowerCase() === DOC_ALBARAN_) return null;
      }
    }
  }
  return { pestana: hoja.getName(), importe: f.importe, carpeta: carpeta, dirFoto: dirFoto };
}

// Un fallo que puede ser pasajero (Claude saturado, Google lento): se
// prueba en la próxima vuelta; al tercero, a "No se pudo leer".
function reintentarDespues_(archivo, bandeja, error) {
  var m = /intentos: (\d+)/.exec(archivo.getDescription() || '');
  var intentos = (m ? Number(m[1]) : 0) + 1;
  if (intentos < BANDEJA_INTENTOS_) {
    archivo.setDescription('intentos: ' + intentos);
    return {};
  }
  archivo.setDescription('');
  archivo.moveTo(carpetaHija_(bandeja, BANDEJA_NO_LEIDAS_));
  return {
    problema: '"' + archivo.getName() + '" falló ' + intentos + ' veces (' + error + '). Quedó en "' + BANDEJA_NO_LEIDAS_ + '": se puede volver a poner en la carpeta para probar otra vez, o cargarla a mano.',
    proveedor: archivo.getName(), avisos: [{ tipo: 'fallo', que: 'No se pudo cargar: está en "' + BANDEJA_NO_LEIDAS_ + '". Cargarla a mano o volver a ponerla en la carpeta.' }],
    enlace: archivo.getUrl()
  };
}

function noSePudoLeer_(archivo, bandeja, motivo) {
  archivo.moveTo(carpetaHija_(bandeja, BANDEJA_NO_LEIDAS_));
  return {
    problema: '"' + archivo.getName() + '": ' + motivo + '. Quedó en "' + BANDEJA_NO_LEIDAS_ + '" para cargarla a mano.',
    proveedor: archivo.getName(), avisos: [{ tipo: 'fallo', que: 'No se pudo leer (' + motivo + '): cargarla a mano.' }],
    enlace: archivo.getUrl()
  };
}

// La factura del mes de Pollería, La Jaula, Mercandalucía o Las 3B. Va a la
// carpeta del mes que cubre (el de sus albaranes: si es de los primeros 10
// días, el mes anterior), con la nota "factura": así la ven el INFORME MES y
// el paquete de los contables (facturasMensualesEnDrive_ en Albaranes.gs).
function guardarFacturaMensual_(archivo, raiz, ss, proveedor, fecha, numero, importe, referencias, iva, base) {
  var anio = Number(fecha.slice(0, 4)), mes = Number(fecha.slice(5, 7)) - 1;
  var fechasRef = referencias.map(function (x) { return x.fecha; }).filter(function (f) { return /^\d{4}-\d{2}-\d{2}$/.test(f); }).sort();
  if (fechasRef.length) {
    var ultima = fechasRef[fechasRef.length - 1];
    anio = Number(ultima.slice(0, 4)); mes = Number(ultima.slice(5, 7)) - 1;
  } else if (Number(fecha.slice(8, 10)) <= 10) {
    mes--; if (mes < 0) { mes = 11; anio--; }
  }
  var mm = ('0' + (mes + 1)).slice(-2);
  var nombreMes = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'][mes];

  // La suma de los albaranes de ese mes en ALBARANES.
  var suma = null, n = 0;
  try {
    var ssMes = String(anio) === fecha.slice(0, 4) ? ss : SpreadsheetApp.openById(buscarEnIndice_('ALBARANES', String(anio)));
    var hoja = pestanaAlbaranesParaProveedor_(ssMes, proveedor);
    var b = hoja ? filasDelMesAlbaranes_(hoja, mes) : null;
    if (b) {
      suma = Math.round(b.filas.reduce(function (s, f) { return s + f.importe; }, 0) * 100) / 100;
      n = b.filas.filter(function (f) { return f.importe; }).length;
    }
  } catch (err) { Logger.log('Factura mensual, suma: ' + err); }

  var dir = carpetaHija_(carpetaHija_(carpetaHija_(raiz, String(anio)), mm), sinCaracteresRaros_(proveedor));
  var ext = (/\.[a-z0-9]{2,4}$/i.exec(archivo.getName()) || ['.pdf'])[0].toLowerCase();
  var nombre = nombreArchivoFactura_(fecha, proveedor, '', numero) + ext;
  if (dir.getFilesByName(nombre).hasNext()) {
    archivo.setName(nombre); archivo.setDescription('');
    archivo.moveTo(carpetaHija_(archivo.getParents().next(), BANDEJA_REPETIDAS_));
    return { hecha: proveedor + ' nº ' + numero + ': la factura de ' + nombreMes + ' ya estaba guardada (la copia nueva está en "' + BANDEJA_REPETIDAS_ + '")', proveedor: proveedor, importe: importe, avisos: [] };
  }
  archivo.setName(nombre);
  // Los datos de la factura (lo que va a los contables): la primera línea
  // sigue siendo "factura"; debajo, total, IVA y base leídos.
  if (iva == null && base != null) iva = Math.round((importe - base) * 100) / 100;
  if (base == null && iva != null) base = Math.round((importe - iva) * 100) / 100;
  archivo.setDescription('factura\n' + JSON.stringify({ numero: numero, fecha: fecha, total: importe, iva: iva, base: base }));
  archivo.moveTo(dir);

  var cual = proveedor + (numero ? ' nº ' + numero : '') + ', factura de ' + nombreMes + ' ' + anio + ', ' + euros_(importe);
  var r = { hecha: '', proveedor: proveedor, importe: importe, avisos: [], enlace: archivo.getUrl() };
  if (suma == null || !n) {
    r.hecha = cual + ': guardada como la factura del mes';
    r.problema = cual + ': no encontré albaranes de ' + proveedor + ' en ' + nombreMes + ' en ALBARANES para comparar.';
    r.avisos.push({ tipo: 'fallo', que: 'Factura del mes guardada, pero no hay albaranes de ' + nombreMes + ' en ALBARANES' });
  } else if (Math.abs(suma - importe) > 0.05) {
    r.hecha = cual + ': guardada como la factura del mes (los ' + n + ' albaranes suman ' + euros_(suma) + ')';
    r.avisos.push({ tipo: 'diferencia', que: 'Factura del mes ' + euros_(importe) + ' y sus ' + n + ' albaranes de ' + nombreMes + ' suman ' + euros_(suma) + ' (' + (importe > suma ? '+' : '') + euros_(importe - suma) + '): ¿falta cargar algún albarán?' });
  } else {
    r.hecha = cual + ': guardada como la factura del mes; coincide con sus ' + n + ' albaranes';
  }
  return r;
}

// "Facturas nuevas" con dudas: no se carga nada; queda en "⚠️ Revisar".
function aRevisar_(archivo, bandeja, cual, motivo, lecturaParaEsperar) {
  if (lecturaParaEsperar) {
    // Se vuelve a probar cada hora (procesarBandejasFacturas) un día.
    var antes = leerEsperando_(archivo);
    var desde = antes ? antes.desde : new Date().toISOString();
    archivo.setDescription(ESPERANDO_ + JSON.stringify({ desde: desde, r: lecturaParaEsperar }));
    archivo.moveTo(carpetaHija_(bandeja, BANDEJA_REVISAR_));
    return {}; // no se avisa: se ve en la pestaña PENDIENTES de ALBARANES
  } else {
    archivo.setDescription('');
    archivo.moveTo(carpetaHija_(bandeja, BANDEJA_REVISAR_));
  }
  return {
    problema: cual + ': ' + motivo + '. Quedó en "' + BANDEJA_REVISAR_ + '", sin cargar.',
    proveedor: cual, avisos: [{ tipo: 'fallo', que: 'No sé si es de albarán o de transferencia: está en "' + BANDEJA_REVISAR_ + '", sin cargar' }],
    enlace: archivo.getUrl()
  };
}

function procesarFacturaDeBandeja_(archivo, tipo, raiz, lista, bandeja) {
  var mime = archivo.getMimeType();
  var esPdf = mime === MimeType.PDF;
  if (!esPdf && !/^image\/(jpeg|png|webp|gif)$/.test(mime)) return noSePudoLeer_(archivo, bandeja, 'no es un PDF ni una foto');
  if (archivo.getSize() > 25 * 1024 * 1024) return noSePudoLeer_(archivo, bandeja, 'es demasiado grande para leerla');

  // Una que está esperando a que se cargue en la caja (ver aRevisar_) ya se
  // leyó: no se vuelve a pagar a Claude.
  var esperando = leerEsperando_(archivo);
  var vencida = esperaVencida_(esperando);
  var r = esperando && esperando.r;
  if (!r) {
    var b64 = Utilities.base64Encode(archivo.getBlob().getBytes());
    var hoja = esPdf
      ? { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: b64 } }
      : { type: 'image', source: { type: 'base64', media_type: mime, data: b64 } };
    r = pedirJsonAClaude_([hoja, {
      type: 'text', text: INSTRUCCIONES_FACTURA.replace('Esta es la foto de', 'Este es el documento de').replace(/\nLos importes con punto/, INSTRUCCIONES_FORMA_PAGO_ + '\nLos importes con punto') +
        '\n\nProveedores habituales:\n' + (lista.join('\n') || '(no hay lista)')
    }], esquemaFacturaBandeja_());
    if (!r || r.ok === false) throw new Error('Claude no pudo leerla: ' + ((r && r.error) || 'sin respuesta'));
    corregirTotalFactura_(r);
  }

  var fecha = String(r.fecha || '');
  var importe = parseFloat(String(r.importe_total || '').replace(',', '.')) || 0;
  if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha)) return noSePudoLeer_(archivo, bandeja, 'no se leyó la fecha de la factura');
  if (!importe) return noSePudoLeer_(archivo, bandeja, 'no se leyó el importe');
  var numero = String(r.numero_factura || '').trim();
  var leido = String(r.proveedor || '').trim();
  var deLista = proveedorDeLaListaLeida_(r.proveedor_lista, lista);
  var iva = parseFloat(String(r.importe_iva || '').replace(',', '.'));

  var anio = fecha.slice(0, 4);
  var idPlanilla = buscarEnIndice_('ALBARANES', anio);
  if (!idPlanilla) throw new Error('No hay planilla de ALBARANES ' + anio + ' en el Índice');
  var ss = SpreadsheetApp.openById(idPlanilla);

  // Pidebio y Madrid Rural, y toda factura externa (carpeta "📥 Facturas
  // externas" o "Factura externa" en la app: a nuestro nombre pero no es
  // gasto del restaurante): a la pestaña FC VYR (fuera del total de compras,
  // pero va a los contables), sin buscar albaranes ni mirar la forma de pago.
  var vyr = proveedorFcVyr_(leido, deLista, r.proveedor_lista) || (tipo === 'externa' ? (leido || deLista || 'Sin proveedor') : '');
  if (vyr) {
    var resVyr = escribirTransferenciaEnAlbaranes_(ss, ALBARANES_PESTANA_FC_VYR_, vyr, fecha, numero, importe, isNaN(iva) ? null : iva, 'pdf-' + archivo.getId());
    var cualVyr = 'FC VYR — ' + vyr + (numero ? ' nº ' + numero : '') + ' del ' + fechaCorta_(fecha) + ', ' + euros_(importe);
    return {
      hecha: cualVyr + ': ' + resVyr.texto, problema: '',
      proveedor: vyr, importe: importe, avisos: [],
      enlace: moverFacturaACarpetaProveedor_(archivo, raiz, fecha, vyr, '', numero)
    };
  }

  var proveedor = deLista || leido || 'Sin proveedor', detalle = '';
  var cual0 = proveedor + (numero ? ' nº ' + numero : '') + ' del ' + fechaCorta_(fecha) + ', ' + euros_(importe);

  // "Facturas nuevas": primero se busca su albarán en la caja; si está, es
  // factura de albarán. Si no, transferencia solo si es un servicio (sin
  // pestaña propia en ALBARANES) o si la factura dice transferencia o
  // domiciliación. Si no está claro, a "⚠️ Revisar" sin cargar nada.
  var resultado;
  var forma = String(r.forma_pago || '');
  // Tarjeta y transferencia son lo mismo (las dos salen del banco, no de la caja).
  var porBanco = forma === 'transferencia' || forma === 'domiciliacion' || forma === 'tarjeta';

  // Pollería, La Jaula, Mercandalucía, Las 3B: lo de cada día son albaranes
  // (cargados en la caja) y una vez al mes llega LA factura. Esa no se carga
  // en ningún lado (sería pagar dos veces): solo se guarda el PDF como la
  // factura del mes y se compara su total con la suma de los albaranes.
  if (tipo !== 'transferencia' && deLista && INFORME_MES_FACTURA_MENSUAL_.indexOf(normalizarClave_(deLista)) > -1) {
    var baseLeida = parseFloat(String(r.base_imponible || '').replace(',', '.'));
    return guardarFacturaMensual_(archivo, raiz, ss, deLista, fecha, numero, importe, r.referencias || [], isNaN(iva) ? null : iva, isNaN(baseLeida) ? null : baseLeida);
  }

  // 24/11/2026 (Luciana): la factura ya estaba cargada en la caja con la
  // FOTO de la factura (Envapro y otros que no mandan albarán): no se busca
  // albarán, solo se cambia la foto por el PDF del email.
  if (tipo === 'auto' || tipo === 'albaran') {
    var ya = facturaYaCargada_(ss, raiz, deLista || leido, numero);
    if (ya) {
      var enlaceYa = moverFacturaACarpetaProveedor_(archivo, raiz, fecha, ya.carpeta, '', numero);
      if (ya.dirFoto) quitarFotosDeLaMismaFactura_(ya.dirFoto, archivo, numero);
      var difYa = Math.abs(ya.importe - importe) > 0.02;
      return {
        hecha: cual0 + ': ya estaba cargada en la caja con la foto; la foto se cambió por el PDF del email',
        problema: difYa ? cual0 + ': en ALBARANES (' + ya.pestana + ') pone ' + euros_(ya.importe) + ' y la factura del email dice ' + euros_(importe) + '. Revisar cuál está bien.' : '',
        proveedor: proveedor, importe: importe,
        avisos: difYa ? [{ tipo: 'diferencia', que: 'En la caja pone ' + euros_(ya.importe) + ' y la factura del email dice ' + euros_(importe) }] : [],
        enlace: enlaceYa
      };
    }
  }

  // "Cómo trabaja cada proveedor" (TratamientoProveedores.gs): la factura
  // llega sin que se haya cargado el gasto → se carga enseguida.
  var trat = deLista ? tratamientoDe_(deLista) : null;
  if (tipo === 'auto' && trat && trat.f === 'SIN_GASTO' && BANDEJA_PROVEEDORES_GENERICOS_.indexOf(normalizarClave_(deLista)) === -1 && pestanaAlbaranesParaProveedor_(ss, deLista)) {
    return cargarFacturaSinAlbaran_(ss, archivo, raiz, deLista, cual0, fecha, numero, importe, iva, forma);
  }

  if (tipo === 'auto') {
    var servicio = !deLista || BANDEJA_PROVEEDORES_GENERICOS_.indexOf(normalizarClave_(deLista)) > -1 || !pestanaAlbaranesParaProveedor_(ss, deLista);
    var esDia = / DIA /.test(' ' + normalizarClave_(proveedor).replace(/[^A-Z]+/g, ' ') + ' ');
    if (!servicio || esDia) {
      resultado = pasarAlbaranesAFactura_(ss, proveedor, fecha, numero, importe, isNaN(iva) ? null : iva, r.referencias || []);
      // 24/11/2026 (Luciana): una factura de Dia sin su ticket en la caja es
      // de FC VYR (a nuestro nombre, pero no es gasto del restaurante).
      if (esDia && resultado.noEncontrado) {
        var resDia = escribirTransferenciaEnAlbaranes_(ss, ALBARANES_PESTANA_FC_VYR_, 'Dia', fecha, numero, importe, isNaN(iva) ? null : iva, 'pdf-' + archivo.getId());
        var cualDia = 'FC VYR — Dia' + (numero ? ' nº ' + numero : '') + ' del ' + fechaCorta_(fecha) + ', ' + euros_(importe);
        return {
          hecha: cualDia + ' (no tiene ticket en la caja): ' + resDia.texto, problema: '',
          proveedor: 'Dia', importe: importe, avisos: [],
          enlace: moverFacturaACarpetaProveedor_(archivo, raiz, fecha, 'FC VYR', 'Dia', numero)
        };
      }
      if (!resultado.noEncontrado && !resultado.sinPestana) tipo = 'albaran';
      // Proveedor con albaranes y sin el albarán en la caja: no se carga
      // nunca sola (antes, si decía transferencia, se cargaba como ya
      // pagada; que llegue la factura no quiere decir que esté pagada).
      else if (porBanco && resultado.sinPestana) { tipo = 'transferencia'; resultado = null; }
      else if (!vencida) return aRevisar_(archivo, bandeja, cual0, 'esperando un día por si se carga su albarán o su foto en la caja', r);
      // Pasado el día sin albarán ni foto: es una factura sin albarán. Se
      // carga en su pestaña como NO PAGADO (que llegue no quiere decir que
      // esté pagada; el banco la marca cuando se pague). Si la factura dice
      // que se pagó con tarjeta (compra que no pasó por el local), ya está
      // pagada: TARJETA/TRANSFERENCIA.
      else return cargarFacturaSinAlbaran_(ss, archivo, raiz, deLista || proveedor, cual0, fecha, numero, importe, iva, forma);
    } else if (forma === 'efectivo') {
      return aRevisar_(archivo, bandeja, cual0, 'no es de un proveedor con albaranes y la factura dice que se pagó en efectivo. Cargarla a mano');
    } else {
      tipo = 'transferencia';
    }
  }

  // Proveedor (carpeta y pestaña) y detalle (el nombre real, en SERVICIOS o
  // VARIOS). Solo los servicios conocidos van a SERVICIOS; lo que no se sabe,
  // a VARIOS y se pregunta (ServiciosVarios.gs).
  var dudaServicio = false;
  // Glovo tiene su pestaña (GLOVO) y su carpeta: nunca a SERVICIOS ni VARIOS.
  if (/glovo/i.test(leido + ' ' + (deLista || ''))) { proveedor = 'Glovo'; detalle = ''; }
  else if (tipo === 'transferencia' && (!deLista || BANDEJA_PROVEEDORES_GENERICOS_.indexOf(normalizarClave_(deLista)) > -1 || !pestanaAlbaranesParaProveedor_(ss, deLista))) {
    var nombreReal = leido || deLista;
    var servicio = servicioConocido_(leido) || servicioConocido_(deLista);
    if (servicio) { proveedor = 'Servicios'; detalle = servicio; }
    else {
      proveedor = 'Varios';
      detalle = nombreReal;
      dudaServicio = !variosConocido_(nombreReal);
    }
  }
  var cual = (detalle ? proveedor + ' — ' + detalle : proveedor) + (numero ? ' nº ' + numero : '') + ' del ' + fechaCorta_(fecha) + ', ' + euros_(importe);

  if (resultado) {
    // ya pasada a factura arriba ("Facturas nuevas")
    if (resultado.proveedor) { proveedor = resultado.proveedor; detalle = resultado.detalle || ''; }
  } else if (tipo === 'albaran') {
    resultado = pasarAlbaranesAFactura_(ss, proveedor, fecha, numero, importe, isNaN(iva) ? null : iva, r.referencias || []);
    if (resultado.proveedor) { proveedor = resultado.proveedor; detalle = resultado.detalle || ''; } // Dia: donde estaba el ticket
  } else {
    resultado = escribirTransferenciaEnAlbaranes_(ss, proveedor, detalle, fecha, numero, importe, isNaN(iva) ? null : iva, 'pdf-' + archivo.getId());
  }

  var enlace = moverFacturaACarpetaProveedor_(archivo, raiz, fecha, proveedor, detalle, numero);
  var avisos = (resultado.avisos || []).slice();
  if (dudaServicio) avisos.push({ tipo: 'varios', que: 'No sé si es un servicio: la puse en VARIOS. ¿Va en VARIOS o en SERVICIOS? Díselo a Claude y se acordará para las próximas' });
  else if (!deLista) avisos.push({ tipo: 'varios', que: 'No es de ningún proveedor de la lista: cargada en ' + (proveedor === 'Servicios' ? 'Servicios' : '"' + proveedor + '"') });
  return {
    hecha: cual + ': ' + resultado.texto,
    problema: resultado.problema ? cual + ': ' + resultado.problema : '',
    proveedor: detalle || proveedor, importe: importe, avisos: avisos,
    enlace: enlace
  };
}

// Factura sin albarán ni foto en la caja: a su pestaña como NO PAGADO; si
// dice tarjeta o domiciliación, ya está pagada (TARJETA/TRANSFERENCIA).
function cargarFacturaSinAlbaran_(ss, archivo, raiz, proveedor, cual, fecha, numero, importe, iva, forma) {
  var pagada = forma === 'tarjeta' || forma === 'domiciliacion';
  var res = escribirTransferenciaEnAlbaranes_(ss, proveedor, '', fecha, numero, importe, isNaN(iva) ? null : iva, 'pdf-' + archivo.getId(), pagada ? 'tarjeta' : true);
  return {
    hecha: cual + ': ' + res.texto, problema: '',
    proveedor: proveedor, importe: importe, avisos: [],
    enlace: moverFacturaACarpetaProveedor_(archivo, raiz, fecha, proveedor, '', numero)
  };
}

// La misma carpeta y el mismo nombre que una escaneada (ver
// guardarFotoFactura_ en Precios.gs). Si ya estaba escaneada con ese
// nombre, queda la del email (es la original).
function moverFacturaACarpetaProveedor_(archivo, raiz, fecha, proveedor, detalle, numero) {
  var dir = carpetaHija_(carpetaHija_(carpetaHija_(raiz, fecha.slice(0, 4)), fecha.slice(5, 7)), sinCaracteresRaros_(proveedor));
  var ext = (/\.[a-z0-9]{2,4}$/i.exec(archivo.getName()) || ['.pdf'])[0].toLowerCase();
  var nombre = nombreArchivoFactura_(fecha, proveedor, detalle, numero) + ext;
  if (numero) {
    var existe = dir.getFilesByName(nombre);
    while (existe.hasNext()) {
      var otro = existe.next();
      if (otro.getId() === archivo.getId()) continue;
      try {
        otro.setTrashed(true);
      } catch (err) {
        // La app no puede tirar a la papelera un archivo de otra cuenta
        // (buenasysantasmadrid): se queda el que ya estaba y la nueva copia
        // va a "🗑️ Repetidas" dentro de la bandeja, para borrarla a mano.
        archivo.setName(nombre);
        archivo.setDescription('');
        archivo.moveTo(carpetaHija_(archivo.getParents().next(), BANDEJA_REPETIDAS_));
        return otro.getUrl();
      }
    }
  }
  archivo.setName(nombre);
  archivo.setDescription('');
  archivo.moveTo(dir);
  if (numero) quitarFotosDeLaMismaFactura_(dir, archivo, numero);
  return archivo.getUrl();
}

// 23/11/2026 (Luciana): una factura que se sacó en foto y después llega por
// email: queda la del email (es la buena) y la foto va a la papelera (se
// puede recuperar 30 días). Se busca por número aunque el nombre del
// archivo no sea idéntico (otra fecha leída, el número escrito con
// guiones...), en la carpeta del proveedor y en su subcarpeta Albaranes
// (ahí la pudo apartar pasarAlbaranesAFactura_). Los albaranes de verdad
// (descripción "albaran") no se tocan. Números de menos de 4 cifras no
// (Las 3B "000002": se repiten).
function quitarFotosDeLaMismaFactura_(dir, archivo, numero) {
  var clave = claveNumeroFactura_(numero);
  if (clave.length < 4) return 0;
  var quitadas = 0;
  var revisar = function (carpeta) {
    var fs = carpeta.getFiles();
    while (fs.hasNext()) {
      var f = fs.next();
      if (f.getId() === archivo.getId()) continue;
      if (String(f.getDescription() || '').toLowerCase() === 'albaran') continue;
      var sinFecha = f.getName().replace(/\.[a-z0-9]{2,4}$/i, '').replace(/^\d{4}-\d{2}-\d{2}\s*/, '');
      if (!claveNumeroFactura_(sinFecha).endsWith(clave)) continue;
      try { f.setTrashed(true); quitadas++; }
      catch (err) { Logger.log('No se pudo quitar la foto ' + f.getName() + ': ' + err); }
    }
  };
  revisar(dir);
  var subs = dir.getFoldersByName('Albaranes');
  if (subs.hasNext()) revisar(subs.next());
  if (quitadas) archivo.setDescription('del email (reemplaza la foto)');
  return quitadas;
}

// Número de factura o albarán para comparar: sin espacios, barras ni ceros
// delante ("A-2026/0153" y "A20260153" son el mismo).
function claveNumeroFactura_(v) {
  return String(v == null ? '' : v).toUpperCase().replace(/[^0-9A-Z]/g, '').replace(/^0+/, '');
}

// Las filas de ALBARANES (A–I y el ID de la Z) de un mes de una pestaña.
function filasDelMesAlbaranes_(hoja, mesIndex) {
  var filaEnc = filasEncabezadoAlbaranes_(hoja)[mesIndex];
  if (!filaEnc) return null;
  var mapa = mapaColumnasAlbaranes_(hoja.getRange(filaEnc, 1, 1, ALBARANES_ANCHO_).getValues()[0]);
  var valores = hoja.getRange(filaEnc + 1, 1, ALBARANES_FILAS_DATOS_, ALBARANES_ANCHO_).getValues();
  var ids = hoja.getRange(filaEnc + 1, ALBARANES_COL_ID_, ALBARANES_FILAS_DATOS_, 1).getValues();
  var filas = [];
  valores.forEach(function (v, i) {
    if (v.every(function (x) { return x === '' || x == null; })) return;
    filas.push({
      fecha: mapa.fecha != null && v[mapa.fecha] instanceof Date ? v[mapa.fecha] : null,
      numero: mapa.factura != null ? String(v[mapa.factura] || '') : '',
      importe: mapa.importe != null ? Number(v[mapa.importe]) || 0 : 0,
      iva: mapa.iva != null ? Number(v[mapa.iva]) || 0 : 0,
      forma: mapa.formaPago != null ? normalizarClave_(v[mapa.formaPago]) : '',
      detalle: mapa.detalle != null ? String(v[mapa.detalle] || '') : '',
      obs: mapa.obs != null ? String(v[mapa.obs] || '') : '',
      id: String(ids[i][0] || '')
    });
  });
  return { filaEnc: filaEnc, mapa: mapa, filas: filas };
}

// Factura de albarán: se buscan los albaranes cargados en la caja (en
// ALBARANES: el mes de la factura, el anterior y los de los albaranes que
// menciona, del mismo año) y se pasan a la factura igual que un ticket (ver
// aplicarFacturaATickets_ en FacturaTicket.gs): el gasto se queda con el
// número, la fecha, el IVA y el importe de la factura (manda la factura,
// aunque el albarán dijera otro total), y el PDF del albarán se aparta a
// <proveedor>/Albaranes. La forma de pago no cambia. Si se pagó en efectivo,
// el importe se queda en lo que salió de la caja (si no, cambiaría la
// diferencia de esa caja) y, si la factura dice otro, se avisa.
function pasarAlbaranesAFactura_(ss, proveedor, fecha, numero, importe, iva, referencias) {
  // Dia: el ticket puede estar en DIA CARNE o en SUPER (como "Dia"); la
  // factura se queda donde esté el ticket (en la app se pregunta, aquí no
  // hay a quién).
  var esDia = / DIA /.test(' ' + normalizarClave_(proveedor).replace(/[^A-Z]+/g, ' ') + ' ');
  // Disbesa: el albarán puede estar en DISBESA o en DISBESA CAFE (ver ALBARANES_BUSCAR_JUNTOS_).
  var juntos = ALBARANES_BUSCAR_JUNTOS_.filter(function (g) { return g.indexOf(normalizarClave_(proveedor)) > -1; })[0];
  var hojas = (esDia ? ['Dia Carne', 'Super'] : juntos || [proveedor]).map(function (p) { return pestanaAlbaranesParaProveedor_(ss, p); }).filter(Boolean);
  if (!hojas.length) return { sinPestana: true, avisos: [{ tipo: 'fallo', que: 'No tiene pestaña en ALBARANES: no se pudo buscar su albarán' }], texto: 'guardada en Drive', problema: 'no hay pestaña de "' + proveedor + '" en ALBARANES, así que no se pudo buscar el albarán para ponerle los datos de la factura.' };
  var hoja = hojas[0];

  var numeros = [numero].concat(referencias.map(function (x) { return x.numero; })).map(claveNumeroFactura_).filter(String);
  var fechasRef = referencias.map(function (x) { return x.fecha; }).filter(function (f) { return /^\d{4}-\d{2}-\d{2}$/.test(f); });
  var meses = {};
  var mesFactura = Number(fecha.slice(5, 7)) - 1;
  meses[mesFactura] = true;
  if (mesFactura > 0) meses[mesFactura - 1] = true;
  fechasRef.forEach(function (f) { if (f.slice(0, 4) === fecha.slice(0, 4)) meses[Number(f.slice(5, 7)) - 1] = true; });

  var todas = [];
  hojas.forEach(function (h) {
    var esSuper = normalizarClave_(h.getName()) === 'SUPER';
    Object.keys(meses).forEach(function (m) {
      var b = filasDelMesAlbaranes_(h, Number(m));
      if (!b) return;
      b.filas.forEach(function (f) {
        if (esSuper && !/ DIA /.test(' ' + normalizarClave_(f.detalle).replace(/[^A-Z]+/g, ' ') + ' ')) return; // en SUPER, solo los de Dia
        f.pestana = h.getName();
        todas.push(f);
      });
    });
  });

  var dia = function (d) { return d ? Utilities.formatDate(d, 'Europe/Madrid', 'yyyy-MM-dd') : ''; };
  var porNumero = todas.filter(function (f) {
    var k = claveNumeroFactura_(f.numero) || claveNumeroFactura_(f.obs);
    return k && numeros.some(function (n) { return n === k || (n.length >= 4 && k.indexOf(n) > -1); });
  });
  var encontradas = porNumero;
  if (!encontradas.length) {
    // Sin número que coincida: un importe parecido (la factura puede ser algo
    // menos que el albarán por una devolución: hasta un 10 %, o 5 €) en el día
    // de un albarán que menciona o hasta 10 días antes o después de la
    // factura; el más parecido. Los albaranes que ya pasaron a factura
    // ("Ticket ..." en observaciones) no cuentan.
    var t = fechaComoDate_(fecha).getTime();
    var margen = Math.max(5, importe * 0.10);
    encontradas = todas.filter(function (f) {
      if (Math.abs(f.importe - importe) > margen || !f.fecha) return false;
      if (/(^|·\s*)TICKET\s/.test(normalizarClave_(f.obs))) return false;
      return fechasRef.indexOf(dia(f.fecha)) > -1 || Math.abs(f.fecha.getTime() - t) <= 10 * 86400000;
    }).sort(function (a, b) {
      return Math.abs(a.importe - importe) - Math.abs(b.importe - importe) || Math.abs(a.fecha.getTime() - t) - Math.abs(b.fecha.getTime() - t);
    }).slice(0, 1);
  }
  if (!encontradas.length) {
    return { noEncontrado: true, avisos: [{ tipo: 'fallo', que: 'No encontré su albarán en la caja: ¿se cargó?' }], texto: 'guardada en Drive', problema: 'no encontré en ALBARANES (' + (esDia ? 'pestañas DIA CARNE y SUPER' : 'pestaña ' + hoja.getName()) + ') el albarán cargado en la caja. Si no se cargó, hay que cargarlo.' };
  }
  // La carpeta de Drive: la del proveedor donde está cargado el ticket.
  var enSuper = esDia && normalizarClave_(encontradas[0].pestana) === 'SUPER';
  var donde = esDia ? { proveedor: enSuper ? 'Super' : 'Dia Carne', detalle: enSuper ? 'Dia' : '' } : {};
  var cuales = encontradas.map(function (f) { return (f.numero ? 'nº ' + f.numero + ' ' : '') + 'del ' + fechaCorta_(dia(f.fecha)) + ' (' + euros_(f.importe) + ')'; }).join(', ');

  // De qué caja es cada uno ("IDs app"). Los cargados a mano en la planilla
  // (sin ID) no se pueden cambiar desde aquí.
  var cajaDe = {};
  getOrCrearHojaIdsAlbaranes_(ss).getDataRange().getValues().slice(1).forEach(function (f) { cajaDe[String(f[0])] = textoFechaIds_(f[1]); });
  var conCaja = encontradas.filter(function (f) { return f.id && f.id.indexOf('pdf-') !== 0 && cajaDe[f.id]; });
  if (conCaja.length < encontradas.length) {
    return { proveedor: donde.proveedor, detalle: donde.detalle, avisos: [{ tipo: 'fallo', que: 'Su albarán está escrito a mano en ALBARANES: ponerle la factura a mano' }], texto: 'coincide con ' + cuales, problema: 'el albarán (' + cuales + ') está escrito a mano en ALBARANES, no viene de una caja: hay que ponerle los datos de la factura a mano.' };
  }

  var suma = Math.round(encontradas.reduce(function (s, f) { return s + f.importe; }, 0) * 100) / 100;
  var enEfectivo = encontradas.some(function (f) { return f.forma === 'EFECTIVO'; });
  var r = aplicarFacturaATickets_({
    tickets: conCaja.map(function (f) {
      return { idMovimiento: f.id, fecha: cajaDe[f.id], importe: f.importe, factura: f.numero, fechaFactura: dia(f.fecha), proveedorTexto: proveedor };
    }),
    factura: { numero: numero, fecha: fecha, importe: enEfectivo ? suma : importe, iva: iva == null ? 0 : iva },
    proveedor: proveedor, quien: 'Factura PDF de Drive', sinDrive: true
  });
  if (!r.ok) {
    return { proveedor: donde.proveedor, detalle: donde.detalle, avisos: [{ tipo: 'fallo', que: 'No se pudo pasar su albarán a factura (va en el resumen de avisos)' }], texto: 'coincide con ' + cuales, problema: 'no se pudo pasar el albarán (' + cuales + ') a la factura: ' + r.dias.filter(function (d) { return !d.ok; }).map(function (d) { return 'caja del ' + fechaCorta_(d.fecha) + ': ' + d.error; }).join('; ') };
  }
  // El PDF del albarán (escaneado con su número) se aparta a <proveedor>/Albaranes.
  try {
    // (Dia: se mira en las dos carpetas, Super y Dia Carne.)
    conCaja.forEach(function (f) { moverTicketAAlbaranes_({ factura: f.numero, fecha: cajaDe[f.id], fechaFactura: dia(f.fecha), proveedorTexto: esDia ? 'Super' : proveedor }, esDia ? 'Dia Carne' : proveedor); });
  } catch (err) { Logger.log('PDF del albarán: ' + err); }
  var distinto = Math.abs(suma - importe) > 0.02;
  return {
    proveedor: donde.proveedor, detalle: donde.detalle,
    texto: 'el albarán ' + cuales + ' ahora es la factura' + (numero ? ' nº ' + numero : '') +
      (distinto && !enEfectivo ? ' (el importe cambió de ' + euros_(suma) + ' a ' + euros_(importe) + ')' : ''),
    // Una diferencia no es un fallo: va solo al resumen de los lunes.
    avisos: !distinto ? [] : [{
      tipo: 'diferencia',
      que: enEfectivo
        ? 'Se pagaron ' + euros_(suma) + ' en efectivo y la factura es de ' + euros_(importe) + '. Se dejó lo pagado: hablar con el proveedor'
        : 'El albarán era de ' + euros_(suma) + '; se cambió a lo de la factura'
    }]
  };
}

// Factura pagada por transferencia: una fila nueva en ALBARANES. Si ya
// estaba (mismo número, o mismo nombre, día e importe), no se repite.
function escribirTransferenciaEnAlbaranes_(ss, proveedor, detalle, fecha, numero, importe, iva, id, noPagado) {
  var hoja = proveedor === 'Servicios' ? asegurarPestanaServiciosAlbaranes_(ss)
    : proveedor === ALBARANES_PESTANA_FC_VYR_ ? asegurarPestanaFcVyrAlbaranes_(ss)
    : pestanaAlbaranesParaProveedor_(ss, proveedor);
  if (!hoja) throw new Error('No hay pestaña "' + (proveedor === 'Servicios' ? ALBARANES_PESTANA_SERVICIOS_ : proveedor) + '" en ALBARANES');
  var mesIndex = Number(fecha.slice(5, 7)) - 1;
  var b = filasDelMesAlbaranes_(hoja, mesIndex);
  if (!b) throw new Error('La pestaña "' + hoja.getName() + '" no tiene el bloque del mes ' + (mesIndex + 1));

  var kNumero = claveNumeroFactura_(numero);
  var ya = b.filas.filter(function (f) {
    if (f.id === id) return false;
    if (kNumero && claveNumeroFactura_(f.numero) === kNumero) return true;
    return f.fecha && Utilities.formatDate(f.fecha, 'Europe/Madrid', 'yyyy-MM-dd') === fecha && Math.abs(f.importe - importe) <= 0.02 &&
      (!detalle || normalizarClave_(f.detalle) === normalizarClave_(detalle));
  })[0];
  if (ya) return { texto: 'ya estaba cargada en ALBARANES (pestaña ' + hoja.getName() + '), no se repitió' };

  var m = {
    proveedor: proveedor, proveedorDetalle: detalle, importe: importe, iva: iva, factura: numero,
    subtipo: noPagado === 'tarjeta' ? 'tarjeta' : noPagado ? 'no_pagado' : 'transferencia',
    info: noPagado === 'tarjeta' ? FORMA_PAGO_ALBARAN_.tarjeta + ' · PDF recibido por email, sin albarán'
      : noPagado ? FORMA_PAGO_ALBARAN_.no_pagado + ' · PDF recibido por email, sin albarán' : FORMA_PAGO_ALBARAN_.transferencia + ' · PDF recibido por email'
  };
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    reescribirBloqueAlbaranes_(hoja, b.filaEnc, {}, [{ id: id, fila: filaAlbaran_(m, fecha, b.mapa, fecha) }]);
  } finally {
    lock.releaseLock();
  }
  try { refrescarInformeMesSiToca_(ss, [mesIndex]); } catch (err) { Logger.log('INFORME MES: ' + err); }
  try { actualizarProveedoresContabilidad_(fecha.slice(0, 4)); } catch (err) { Logger.log('PROVEEDORES Contabilidad: ' + err); }
  if (noPagado === 'tarjeta') return { texto: 'cargada como factura en ALBARANES (pestaña ' + hoja.getName() + '), ya pagada (tarjeta o domiciliación)' };
  if (noPagado) return { texto: 'cargada como factura en ALBARANES (pestaña ' + hoja.getName() + ') como NO PAGADO' };
  return { texto: 'cargada en ALBARANES (pestaña ' + hoja.getName() + '), pagada por transferencia el ' + fechaCorta_(fecha) };
}

// Solo Claude (POST con claveClaude, ver doPost): mueve y/o renombra un
// archivo dentro de "Facturas proveedores", para ordenar a mano lo que quedó
// en otra carpeta. Rutas como "2026/10/Dia Carne". Si en el destino ya hay
// uno con ese nombre, se le añade " (2)".
function moverArchivoFacturas_(data) {
  var raiz = DriveApp.getFolderById(CARPETA_FACTURAS_ID_);
  var ir = function (ruta, crear) {
    var dir = raiz;
    String(ruta || '').split('/').filter(String).forEach(function (n) {
      var it = dir.getFoldersByName(n);
      if (it.hasNext()) dir = it.next();
      else if (crear) dir = dir.createFolder(n);
      else throw new Error('No existe la carpeta ' + n);
    });
    return dir;
  };
  var origen = ir(data.desde, false);
  var it = origen.getFilesByName(String(data.nombre || ''));
  if (!it.hasNext()) return { ok: false, error: 'No está "' + data.nombre + '" en ' + data.desde };
  var archivo = it.next();
  // papelera: true → a la papelera de Drive (se puede recuperar 30 días).
  if (data.papelera) { archivo.setTrashed(true); return { ok: true, papelera: true, nombre: archivo.getName(), carpeta: data.desde }; }
  var destino = data.hasta ? ir(data.hasta, true) : origen;
  var nombre = String(data.nuevoNombre || archivo.getName());
  if (destino.getFilesByName(nombre).hasNext() && !(destino.getId() === origen.getId() && nombre === archivo.getName())) {
    nombre = nombre.replace(/(\.[a-z0-9]+)?$/i, ' (2)$1');
  }
  archivo.setName(nombre);
  if (destino.getId() !== origen.getId()) archivo.moveTo(destino);
  return { ok: true, nombre: nombre, carpeta: data.hasta || data.desde };
}

function fechaCorta_(fechaISO) {
  var p = String(fechaISO || '').split('-');
  return p.length === 3 ? p[2] + '/' + p[1] + '/' + p[0] : String(fechaISO || '');
}

function euros_(n) {
  return (Math.round(Number(n) * 100) / 100).toFixed(2).replace('.', ',') + ' €';
}

// Desde la app (menú Varios → Facturas): la foto de una factura en papel que
// no se carga en la caja (sobre todo la del mes de Pollería, La Jaula,
// Mercandalucía y Las 3B). Se guarda en "Facturas nuevas" y se pasa ya, como
// si se hubiera puesto en la carpeta; devuelve qué se hizo con ella.
function guardarFacturaPapel_(data) {
  if (!data.pdfBase64) return { ok: false, error: 'No llegó la foto' };
  var raiz = DriveApp.getFolderById(CARPETA_FACTURAS_ID_);
  // data.externa: "Factura externa" en la app (va a FC VYR, ver BANDEJA_EXTERNAS_).
  var tipo = data.externa ? 'externa' : 'auto';
  var bandeja = carpetaHija_(raiz, data.externa ? BANDEJA_EXTERNAS_ : BANDEJA_UNICA_);
  var nombre = (data.externa ? 'Factura NO USAR ' : 'Factura en papel ') + Utilities.formatDate(new Date(), 'Europe/Madrid', 'yyyy-MM-dd HH.mm.ss') + '.pdf';
  var archivo = bandeja.createFile(Utilities.newBlob(Utilities.base64Decode(data.pdfBase64), 'application/pdf', nombre));
  var r;
  try {
    r = procesarFacturaDeBandeja_(archivo, tipo, raiz, proveedoresParaLeer_(), bandeja);
  } catch (err) {
    // Se queda en la carpeta y se prueba otra vez en la vuelta de cada hora.
    return { ok: true, guardada: true, hecha: 'Guardada. No se pudo leer ahora (' + (err.message || err) + '): se vuelve a intentar sola en menos de una hora.' };
  }
  if (r.hecha) apuntarMovFactura_('cargada', r.proveedor, r.importe, 'Foto en la app (' + (data.externa ? 'Factura NO USAR' : 'Facturas') + '): ' + r.hecha, r.enlace);
  (r.avisos || []).forEach(function (a) { apuntarMovFactura_(a.tipo, r.proveedor, r.importe, a.que, r.enlace); });
  return { ok: true, guardada: true, hecha: r.hecha || '', problema: r.problema || '' };
}

// Solo Claude (GET ?pasarAFcVyr=AAAA): las facturas de Pidebio y Madrid
// Rural que se cargaron en SERVICIOS antes de que existiera FC VYR pasan a
// esa pestaña (la fila tal cual, con el nombre corto en PROVEEDOR) y su PDF
// a la carpeta del proveedor.
function pasarServiciosAFcVyr_(anio) {
  var ss = SpreadsheetApp.openById(buscarEnIndice_('ALBARANES', String(anio)));
  var serv = pestanaAlbaranesParaProveedor_(ss, ALBARANES_PESTANA_SERVICIOS_);
  if (!serv) return { ok: true, hechos: ['No hay pestaña SERVICIOS'] };
  var vyr = asegurarPestanaFcVyrAlbaranes_(ss);
  if (!vyr) return { ok: false, error: 'No se pudo crear la pestaña FC VYR (no hay VARIOS para copiar)' };
  var raiz = DriveApp.getFolderById(CARPETA_FACTURAS_ID_);
  var hechos = [], meses = [];
  var encServ = filasEncabezadoAlbaranes_(serv), encVyr = filasEncabezadoAlbaranes_(vyr);
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    encServ.forEach(function (filaEnc, mes) {
      var mapa = mapaColumnasAlbaranes_(serv.getRange(filaEnc, 1, 1, ALBARANES_ANCHO_).getValues()[0]);
      if (mapa.detalle == null) return;
      var valores = serv.getRange(filaEnc + 1, 1, ALBARANES_FILAS_DATOS_, ALBARANES_ANCHO_).getValues();
      var ids = serv.getRange(filaEnc + 1, ALBARANES_COL_ID_, ALBARANES_FILAS_DATOS_, 1).getValues();
      var quedan = [], pasan = [];
      valores.forEach(function (v, i) {
        var id = String(ids[i][0] || '');
        if (v.every(function (x) { return x === '' || x == null; }) && !id) return;
        var nombre = proveedorFcVyr_(v[mapa.detalle]);
        if (!nombre) { quedan.push({ id: id, fila: v }); return; }
        if (id.indexOf('pdf-') !== 0) { // cargada desde una caja: no se toca
          quedan.push({ id: id, fila: v });
          hechos.push(nombre + ' (mes ' + (mes + 1) + '): viene de una caja, se queda en SERVICIOS');
          return;
        }
        var fila = v.slice();
        fila[mapa.detalle] = nombre;
        pasan.push({ id: id, fila: fila, nombre: nombre, numero: mapa.factura != null ? String(v[mapa.factura] || '') : '',
          fecha: mapa.fecha != null && v[mapa.fecha] instanceof Date ? Utilities.formatDate(v[mapa.fecha], 'Europe/Madrid', 'yyyy-MM-dd') : '',
          importe: mapa.importe != null ? Number(v[mapa.importe]) || 0 : 0 });
      });
      if (!pasan.length) return;
      if (!encVyr[mes]) throw new Error('FC VYR no tiene el bloque del mes ' + (mes + 1));
      reescribirBloqueAlbaranes_(vyr, encVyr[mes], {}, pasan.map(function (p, k) { return { id: p.id || ('fcvyr-' + mes + '-' + k), fila: p.fila }; }));
      guardarBloqueOrdenado_(serv, filaEnc, quedan);
      meses.push(mes);
      pasan.forEach(function (p) {
        var pdf = '';
        if (p.id.indexOf('pdf-') === 0 && p.fecha) {
          try { pdf = ' · PDF a ' + p.fecha.slice(0, 7).replace('-', '/') + '/' + p.nombre + ': ' + moverFacturaACarpetaProveedor_(DriveApp.getFileById(p.id.slice(4)), raiz, p.fecha, p.nombre, '', p.numero); }
          catch (err) { pdf = ' · el PDF no se pudo mover: ' + err; }
        }
        hechos.push(p.nombre + ' ' + p.numero + ' del ' + fechaCorta_(p.fecha) + ' (' + euros_(p.importe) + '): de SERVICIOS a FC VYR' + pdf);
      });
    });
  } finally {
    lock.releaseLock();
  }
  try { if (meses.length) refrescarInformeMesSiToca_(ss, meses); } catch (err) { hechos.push('INFORME MES: ' + err); }
  try { sacarDelTotalAlbaranes_(ss, ALBARANES_COPIA_VIEJA_.de); } catch (err) { hechos.push('TOTAL: ' + err); }
  invalidarCacheListados_();
  return { ok: true, hechos: hechos };
}
