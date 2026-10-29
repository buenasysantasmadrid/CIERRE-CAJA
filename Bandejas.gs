// ============================================================================
// BANDEJAS DE FACTURAS — las facturas que llegan en PDF (por email) entran
// sin imprimirlas ni escanearlas. Dentro de "Facturas proveedores" hay dos
// carpetas de entrada; desde el móvil, "Guardar en Drive" y elegir una:
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
// las escaneadas: es lo que se entrega a los contables. Llega un email con
// lo cargado; los problemas van al resumen de avisos. Lo que no se puede
// leer queda en la subcarpeta "⚠️ No se pudo leer" de su bandeja.
// ============================================================================

var BANDEJA_ALBARAN_ = '📥 Facturas de albarán';
var BANDEJA_TRANSFERENCIA_ = '📥 Facturas pagadas por transferencia';
var BANDEJA_NO_LEIDAS_ = '⚠️ No se pudo leer';
var BANDEJA_TIEMPO_MAX_MS_ = 4.5 * 60 * 1000; // el activador corta a los 6 minutos
var BANDEJA_INTENTOS_ = 3; // fallos (ej. Claude saturado) antes de dejarla en "No se pudo leer"
var BANDEJA_PROVEEDORES_GENERICOS_ = ['VARIOS', 'SUPER', 'EXTRAS', 'SERVICIOS', 'GLOVO'];

// Desde cada guardado de la caja (doPost): crea las dos carpetas y el
// activador de cada hora, si no están.
function asegurarBandejasFacturas_() {
  try {
    var cache = CacheService.getScriptCache();
    if (cache.get('bandejas_facturas')) return;
    var raiz = DriveApp.getFolderById(CARPETA_FACTURAS_ID_);
    carpetaHija_(raiz, BANDEJA_ALBARAN_);
    carpetaHija_(raiz, BANDEJA_TRANSFERENCIA_);
    var hay = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'procesarBandejasFacturas'; });
    if (!hay) ScriptApp.newTrigger('procesarBandejasFacturas').timeBased().everyHours(1).create();
    cache.put('bandejas_facturas', '1', 21600);
  } catch (err) {
    Logger.log('Bandejas de facturas: ' + err);
  }
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

// Lo que corre cada hora.
function procesarBandejasFacturas() {
  var cache = CacheService.getScriptCache();
  if (cache.get('bandejas_corriendo')) return; // la vuelta anterior todavía no terminó
  cache.put('bandejas_corriendo', '1', 400);
  try {
    var inicio = Date.now();
    var raiz = DriveApp.getFolderById(CARPETA_FACTURAS_ID_);
    var lista = proveedoresParaLeer_();
    var hechas = [], problemas = 0;
    [[BANDEJA_ALBARAN_, 'albaran'], [BANDEJA_TRANSFERENCIA_, 'transferencia']].forEach(function (b) {
      var bandeja = carpetaHija_(raiz, b[0]);
      var archivos = bandeja.getFiles();
      while (archivos.hasNext() && Date.now() - inicio < BANDEJA_TIEMPO_MAX_MS_) {
        var archivo = archivos.next();
        var r;
        try {
          r = procesarFacturaDeBandeja_(archivo, b[1], raiz, lista);
        } catch (err) {
          r = reintentarDespues_(archivo, bandeja, String(err.message || err));
        }
        if (r.hecha) hechas.push(r.hecha);
        if (r.problema) {
          problemas++;
          avisarPorEmail_('una factura de la carpeta "' + b[0] + '" ' + (r.hecha ? 'tiene algo que revisar' : 'no se cargó'),
            r.problema + (r.enlace ? '\nEl PDF: ' + r.enlace : ''));
        }
      }
    });
    if (hechas.length) mandarResumenBandejas_(hechas, problemas);
  } finally {
    cache.remove('bandejas_corriendo');
  }
}

function mandarResumenBandejas_(hechas, problemas) {
  var para = Session.getEffectiveUser().getEmail();
  if (!para) return;
  MailApp.sendEmail(para, 'Cierre de caja: ' + hechas.length + (hechas.length === 1 ? ' factura cargada' : ' facturas cargadas') + ' desde Drive',
    hechas.map(function (h, i) { return (i + 1) + '. ' + h; }).join('\n\n') +
    (problemas ? '\n\n' + problemas + (problemas === 1 ? ' tuvo algo para revisar: va' : ' tuvieron algo para revisar: van') + ' en el resumen de avisos.' : ''));
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
  return { problema: '"' + archivo.getName() + '" falló ' + intentos + ' veces (' + error + '). Quedó en "' + BANDEJA_NO_LEIDAS_ + '": se puede volver a poner en la carpeta para probar otra vez, o cargarla a mano.', enlace: archivo.getUrl() };
}

function noSePudoLeer_(archivo, bandeja, motivo) {
  archivo.moveTo(carpetaHija_(bandeja, BANDEJA_NO_LEIDAS_));
  return { problema: '"' + archivo.getName() + '": ' + motivo + '. Quedó en "' + BANDEJA_NO_LEIDAS_ + '" para cargarla a mano.', enlace: archivo.getUrl() };
}

function procesarFacturaDeBandeja_(archivo, tipo, raiz, lista) {
  var bandeja = carpetaHija_(raiz, tipo === 'albaran' ? BANDEJA_ALBARAN_ : BANDEJA_TRANSFERENCIA_);
  var mime = archivo.getMimeType();
  var esPdf = mime === MimeType.PDF;
  if (!esPdf && !/^image\/(jpeg|png|webp|gif)$/.test(mime)) return noSePudoLeer_(archivo, bandeja, 'no es un PDF ni una foto');
  if (archivo.getSize() > 25 * 1024 * 1024) return noSePudoLeer_(archivo, bandeja, 'es demasiado grande para leerla');

  var b64 = Utilities.base64Encode(archivo.getBlob().getBytes());
  var hoja = esPdf
    ? { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: b64 } }
    : { type: 'image', source: { type: 'base64', media_type: mime, data: b64 } };
  var r = pedirJsonAClaude_([hoja, {
    type: 'text', text: INSTRUCCIONES_FACTURA.replace('Esta es la foto de', 'Este es el documento de') + '\n\nProveedores habituales:\n' + (lista.join('\n') || '(no hay lista)')
  }], ESQUEMA_FACTURA);
  if (!r || r.ok === false) throw new Error('Claude no pudo leerla: ' + ((r && r.error) || 'sin respuesta'));
  corregirTotalFactura_(r);

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

  // Proveedor (carpeta y pestaña) y detalle (el nombre real, en SERVICIOS).
  var proveedor = deLista || leido || 'Sin proveedor', detalle = '';
  if (tipo === 'transferencia' && (!deLista || BANDEJA_PROVEEDORES_GENERICOS_.indexOf(normalizarClave_(deLista)) > -1 || !pestanaAlbaranesParaProveedor_(ss, deLista))) {
    proveedor = 'Servicios';
    detalle = leido || deLista;
  }
  var cual = (detalle ? 'Servicios — ' + detalle : proveedor) + (numero ? ' nº ' + numero : '') + ' del ' + fechaCorta_(fecha) + ', ' + euros_(importe);

  var resultado;
  if (tipo === 'albaran') {
    resultado = pasarAlbaranesAFactura_(ss, proveedor, fecha, numero, importe, isNaN(iva) ? null : iva, r.referencias || []);
  } else {
    resultado = escribirTransferenciaEnAlbaranes_(ss, proveedor, detalle, fecha, numero, importe, isNaN(iva) ? null : iva, 'pdf-' + archivo.getId());
  }

  var enlace = moverFacturaACarpetaProveedor_(archivo, raiz, fecha, proveedor, detalle, numero);
  return {
    hecha: cual + ': ' + resultado.texto + '\n   ' + enlace,
    problema: resultado.problema ? cual + ': ' + resultado.problema : '',
    enlace: enlace
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
      if (otro.getId() !== archivo.getId()) otro.setTrashed(true);
    }
  }
  archivo.setName(nombre);
  archivo.setDescription('');
  archivo.moveTo(dir);
  return archivo.getUrl();
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
  var hoja = pestanaAlbaranesParaProveedor_(ss, proveedor);
  if (!hoja) return { texto: 'guardada en Drive', problema: 'no hay pestaña de "' + proveedor + '" en ALBARANES, así que no se pudo buscar el albarán para ponerle los datos de la factura.' };

  var numeros = [numero].concat(referencias.map(function (x) { return x.numero; })).map(claveNumeroFactura_).filter(String);
  var fechasRef = referencias.map(function (x) { return x.fecha; }).filter(function (f) { return /^\d{4}-\d{2}-\d{2}$/.test(f); });
  var meses = {};
  var mesFactura = Number(fecha.slice(5, 7)) - 1;
  meses[mesFactura] = true;
  if (mesFactura > 0) meses[mesFactura - 1] = true;
  fechasRef.forEach(function (f) { if (f.slice(0, 4) === fecha.slice(0, 4)) meses[Number(f.slice(5, 7)) - 1] = true; });

  var todas = [];
  Object.keys(meses).forEach(function (m) {
    var b = filasDelMesAlbaranes_(hoja, Number(m));
    if (b) todas = todas.concat(b.filas);
  });

  var dia = function (d) { return d ? Utilities.formatDate(d, 'Europe/Madrid', 'yyyy-MM-dd') : ''; };
  var porNumero = todas.filter(function (f) {
    var k = claveNumeroFactura_(f.numero) || claveNumeroFactura_(f.obs);
    return k && numeros.some(function (n) { return n === k || (n.length >= 4 && k.indexOf(n) > -1); });
  });
  var encontradas = porNumero;
  if (!encontradas.length) {
    // Sin número que coincida: el mismo importe en el día de un albarán que
    // menciona o hasta 10 días antes o después de la factura.
    var t = fechaComoDate_(fecha).getTime();
    encontradas = todas.filter(function (f) {
      if (Math.abs(f.importe - importe) > 0.02 || !f.fecha) return false;
      return fechasRef.indexOf(dia(f.fecha)) > -1 || Math.abs(f.fecha.getTime() - t) <= 10 * 86400000;
    }).slice(0, 1);
  }
  if (!encontradas.length) {
    return { texto: 'guardada en Drive', problema: 'no encontré en ALBARANES (pestaña ' + hoja.getName() + ') el albarán cargado en la caja. Si no se cargó, hay que cargarlo.' };
  }
  var cuales = encontradas.map(function (f) { return (f.numero ? 'nº ' + f.numero + ' ' : '') + 'del ' + fechaCorta_(dia(f.fecha)) + ' (' + euros_(f.importe) + ')'; }).join(', ');

  // De qué caja es cada uno ("IDs app"). Los cargados a mano en la planilla
  // (sin ID) no se pueden cambiar desde aquí.
  var cajaDe = {};
  getOrCrearHojaIdsAlbaranes_(ss).getDataRange().getValues().slice(1).forEach(function (f) { cajaDe[String(f[0])] = textoFechaIds_(f[1]); });
  var conCaja = encontradas.filter(function (f) { return f.id && f.id.indexOf('pdf-') !== 0 && cajaDe[f.id]; });
  if (conCaja.length < encontradas.length) {
    return { texto: 'coincide con ' + cuales, problema: 'el albarán (' + cuales + ') está escrito a mano en ALBARANES, no viene de una caja: hay que ponerle los datos de la factura a mano.' };
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
    return { texto: 'coincide con ' + cuales, problema: 'no se pudo pasar el albarán (' + cuales + ') a la factura: ' + r.dias.filter(function (d) { return !d.ok; }).map(function (d) { return 'caja del ' + fechaCorta_(d.fecha) + ': ' + d.error; }).join('; ') };
  }
  // El PDF del albarán (escaneado con su número) se aparta a <proveedor>/Albaranes.
  try {
    conCaja.forEach(function (f) { moverTicketAAlbaranes_({ factura: f.numero, fecha: cajaDe[f.id], fechaFactura: dia(f.fecha), proveedorTexto: proveedor }, proveedor); });
  } catch (err) { Logger.log('PDF del albarán: ' + err); }
  var distinto = Math.abs(suma - importe) > 0.02;
  return {
    texto: 'el albarán ' + cuales + ' ahora es la factura' + (numero ? ' nº ' + numero : '') +
      (distinto && !enEfectivo ? ' (el importe cambió de ' + euros_(suma) + ' a ' + euros_(importe) + ')' : ''),
    problema: distinto && enEfectivo
      ? 'se pagó en efectivo ' + euros_(suma) + ' pero la factura' + (numero ? ' nº ' + numero : '') + ' es de ' + euros_(importe) +
        '. Se dejó lo pagado (' + euros_(suma) + ') para no cambiar la caja: revisar con el proveedor.'
      : ''
  };
}

// Factura pagada por transferencia: una fila nueva en ALBARANES. Si ya
// estaba (mismo número, o mismo nombre, día e importe), no se repite.
function escribirTransferenciaEnAlbaranes_(ss, proveedor, detalle, fecha, numero, importe, iva, id) {
  var hoja = proveedor === 'Servicios' ? asegurarPestanaServiciosAlbaranes_(ss) : pestanaAlbaranesParaProveedor_(ss, proveedor);
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
    subtipo: 'transferencia', info: 'TRANSFERENCIA · PDF recibido por email'
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
  return { texto: 'cargada en ALBARANES (pestaña ' + hoja.getName() + '), pagada por transferencia el ' + fechaCorta_(fecha) };
}

function fechaCorta_(fechaISO) {
  var p = String(fechaISO || '').split('-');
  return p.length === 3 ? p[2] + '/' + p[1] + '/' + p[0] : String(fechaISO || '');
}

function euros_(n) {
  return (Math.round(Number(n) * 100) / 100).toFixed(2).replace('.', ',') + ' €';
}
