// ============================================================================
// FACTURAS POR EMAIL (proyecto aparte, de la cuenta buenasysantasmadrid).
// Se reenvía el email del proveedor con la factura a
// buenasysantasmadrid+facturas@gmail.com (es la misma cuenta: lo de "+facturas"
// solo sirve para encontrarlos). Cada hora los PDF y fotos de esos emails se
// guardan en "Facturas proveedores / 📥 Facturas nuevas", y desde ahí la app
// de la caja los carga como siempre (Bandejas.gs del proyecto principal).
//   · Solo se toman los reenviados desde REMITENTES (para que nadie de fuera
//     meta facturas); los demás quedan con la etiqueta "Factura ignorada".
//   · Solo adjuntos PDF o fotos (las fotos pequeñas, logos y firmas, no).
//   · El email ya pasado queda con la etiqueta "Factura procesada".
//   · Si no trae ningún PDF (ej. la factura es un enlace), llega un email.
//   · Los emails no se quedan en la bandeja de entrada: se archivan y quedan
//     solo con su etiqueta. Los avisos de arriba llegan a lfisbein@gmail.com.
//   · Pasa cada 10 minutos (el activador se cambia solo, ver activadorCada10_).
// Instalar: ejecutar a mano "instalar" desde el editor (pide los permisos y
// pone el activador de cada 10 minutos).
// ============================================================================

var DIRECCION = 'buenasysantasmadrid+facturas@gmail.com';
var REMITENTES = ['info@buenasysantas.es', 'luciana@buenasysantas.es', 'clara@buenasysantas.es', 'yamila@buenasysantas.es', 'lfisbein@gmail.com', 'yamilafisbein@gmail.com', 'efronclara@gmail.com', 'buenasysantasmadrid@gmail.com'];
var CARPETA_FACTURAS_ID = '15R08k4qgm84msoGqxpUnMtU7sLeczWCz'; // "Facturas proveedores"
var BANDEJA = '📥 Facturas nuevas';
var ETIQUETA_OK = 'Factura procesada';
var ETIQUETA_NO = 'Factura ignorada';
var FOTO_MINIMA = 30 * 1024;

function instalar() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'traerFacturasDelCorreo') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('traerFacturasDelCorreo').timeBased().everyMinutes(10).create();
  PropertiesService.getScriptProperties().setProperty('activador', 'cada10');
  var r = traerFacturasDelCorreo();
  Logger.log('Instalado. ' + JSON.stringify(r));
  return r;
}

// El activador era de cada hora: se cambia a cada 10 minutos una sola vez.
function activadorCada10_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('activador') === 'cada10') return;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'traerFacturasDelCorreo') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('traerFacturasDelCorreo').timeBased().everyMinutes(10).create();
  props.setProperty('activador', 'cada10');
}

// ----------------------------------------------------------------------------
// VIGILANCIA DE LA APP DE CAJA. Cada 15 minutos se prueba si la app responde.
// Va en este proyecto (de otra cuenta) a propósito: si la app de caja se cae
// (el 18/11 Google le quitó el permiso a lfisbein y daba "Acceso denegado" a
// todos), sus propios avisos tampoco salen. Avisa a lfisbein cuando falla 2
// veces seguidas (para no avisar por un corte de un momento) y otra vez cuando
// vuelve a funcionar. Instalar: ejecutar a mano "instalarVigilancia".
// ----------------------------------------------------------------------------
var URL_APP_CAJA = 'https://script.google.com/macros/s/AKfycbyrfOE73JqgjTXAEaLetRz8aJg1Ldb-Du4dsqGgW34H-zgwLSJMrAkSKApiTCoz6t7m_A/exec';

function instalarVigilancia() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'vigilarAppCaja') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('vigilarAppCaja').timeBased().everyMinutes(15).create();
  var r = probarAppCaja_();
  Logger.log('Vigilancia instalada. Ahora la app ' + (r.ok ? 'funciona bien.' : 'NO responde: ' + r.detalle));
}

function probarAppCaja_() {
  try {
    var resp = UrlFetchApp.fetch(URL_APP_CAJA + '?salud=1', { muteHttpExceptions: true, followRedirects: true });
    var codigo = resp.getResponseCode();
    if (codigo !== 200) return { ok: false, detalle: 'Google respondió con el código ' + codigo + (codigo === 403 ? ' (Acceso denegado)' : '') };
    try { JSON.parse(resp.getContentText()); } catch (e) { return { ok: false, detalle: 'respondió una página de Google en vez de la app (seguramente pide autorización)' }; }
    return { ok: true };
  } catch (err) {
    return { ok: false, detalle: 'no se pudo conectar: ' + err };
  }
}

function vigilarAppCaja() {
  try { return vigilarAppCaja_(); } catch (err) { avisarError_('Vigilancia de la app de caja', err); }
}
function vigilarAppCaja_() {
  var props = PropertiesService.getScriptProperties();
  var r = probarAppCaja_();
  var fallos = Number(props.getProperty('appFallos') || 0);
  var avisado = props.getProperty('appAvisado') === 'si';
  var yo = 'lfisbein@gmail.com';
  if (r.ok) {
    if (avisado) MailApp.sendEmail(yo, 'App de caja: vuelve a funcionar', 'La app de caja vuelve a responder bien.');
    props.setProperty('appFallos', '0');
    props.setProperty('appAvisado', 'no');
    return;
  }
  fallos++;
  props.setProperty('appFallos', String(fallos));
  if (fallos >= 2 && !avisado) {
    MailApp.sendEmail(yo, '⚠️ App de caja CAÍDA: nadie puede entrar',
      'La app de caja no responde (' + r.detalle + ').\n\n' +
      'Lo que lo arregló la otra vez (18/11):\n' +
      '1. Entrar en el navegador con lfisbein@gmail.com.\n' +
      '2. Abrir https://script.google.com/home/projects/1INf8MRY3syralbQwrwSHKRQ8wy6BCnWZ9IK9joh2Uw3l5bnWSXgmBJUW/edit\n' +
      '3. Abrir Salud.gs, elegir "revisarSalud" y pulsar Ejecutar. Si pide permisos, aceptarlos.\n\n' +
      'Cuando vuelva a funcionar llega otro email.');
    props.setProperty('appAvisado', 'si');
  }
}

function traerFacturasDelCorreo() {
  try { return traerFacturasDelCorreo_(); } catch (err) { avisarError_('Facturas por email', err); }
}
function traerFacturasDelCorreo_() {
  try { activadorCada10_(); } catch (err) { Logger.log('Activador: ' + err); }
  var salida = { traidas: [], ignorados: [], sinAdjunto: [] };
  var raiz = DriveApp.getFolderById(CARPETA_FACTURAS_ID);
  var it = raiz.getFoldersByName(BANDEJA);
  var bandeja = it.hasNext() ? it.next() : raiz.createFolder(BANDEJA);
  var ok = GmailApp.getUserLabelByName(ETIQUETA_OK) || GmailApp.createLabel(ETIQUETA_OK);
  var no = GmailApp.getUserLabelByName(ETIQUETA_NO) || GmailApp.createLabel(ETIQUETA_NO);
  var hilos = GmailApp.search('to:' + DIRECCION + ' -label:"' + ETIQUETA_OK + '" -label:"' + ETIQUETA_NO + '" newer_than:30d', 0, 30);
  hilos.forEach(function (hilo) {
    var algunoBueno = false;
    hilo.getMessages().forEach(function (msj) {
      var para = (msj.getTo() + ',' + msj.getCc()).toLowerCase();
      if (para.indexOf(DIRECCION) === -1) return; // otro mensaje del hilo
      var de = String(msj.getFrom() || '').toLowerCase();
      var asunto = msj.getSubject() || '(sin asunto)';
      if (!REMITENTES.some(function (r) { return de.indexOf(r) > -1; })) {
        salida.ignorados.push(asunto + ' (de ' + msj.getFrom() + ')');
        return;
      }
      algunoBueno = true;
      var adjuntos = msj.getAttachments({ includeInlineImages: false }).filter(function (a) {
        if (a.getContentType() === 'application/pdf' || /\.pdf$/i.test(a.getName())) return true;
        return /^image\/(jpeg|png|webp|gif)$/.test(a.getContentType()) && a.getSize() >= FOTO_MINIMA;
      });
      if (!adjuntos.length) { salida.sinAdjunto.push(asunto); return; }
      adjuntos.forEach(function (a) {
        var blob = a.copyBlob();
        if (/\.pdf$/i.test(a.getName())) blob.setContentType('application/pdf');
        var archivo = bandeja.createFile(blob);
        archivo.setName(a.getName() || ('Factura por email ' + Utilities.formatDate(new Date(), 'Europe/Madrid', 'yyyy-MM-dd HH.mm') + '.pdf'));
        salida.traidas.push(asunto + ': ' + archivo.getName());
      });
    });
    hilo.addLabel(algunoBueno ? ok : no);
    hilo.moveToArchive();
  });
  // Los que ya estaban pasados y siguen en la bandeja de entrada.
  GmailApp.search('to:' + DIRECCION + ' in:inbox (label:"' + ETIQUETA_OK + '" OR label:"' + ETIQUETA_NO + '")', 0, 100).forEach(function (h) { h.moveToArchive(); });
  var yo = 'lfisbein@gmail.com'; // los avisos de que algo salió mal van a la dueña
  if (salida.ignorados.length) {
    MailApp.sendEmail(yo, 'Facturas por email: un email no se tomó',
      'Solo se toman las facturas reenviadas desde ' + REMITENTES.join(', ') + '. Estos llegaron de otra dirección y quedaron con la etiqueta "' + ETIQUETA_NO + '":\n' + salida.ignorados.join('\n'));
  }
  if (salida.sinAdjunto.length) {
    MailApp.sendEmail(yo, 'Facturas por email: un email no traía el PDF',
      'Se reenviaron a ' + DIRECCION + ' pero no traían ningún PDF ni foto (puede que la factura venga como enlace: hay que descargarla y guardarla en "' + BANDEJA + '"):\n' + salida.sinAdjunto.join('\n'));
  }
  return salida;
}

// 26/11/2026 (Luciana): si algo falla, el email va a lfisbein, no a buenasysantasmadrid.
// Los activadores son de esta cuenta (buenasysantasmadrid): si un error se escapa, Google
// manda su aviso automático a esta cuenta. Por eso se atrapa aquí y se avisa a lfisbein,
// como mucho una vez cada 6 horas por el mismo problema (los cortes de Google son frecuentes).
function avisarError_(que, err) {
  var texto = String(err && err.message || err);
  Logger.log(que + ': ' + texto);
  var props = PropertiesService.getScriptProperties();
  var clave = 'aviso_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, que + texto)).slice(0, 20);
  var antes = Number(props.getProperty(clave) || 0);
  if (Date.now() - antes < 6 * 3600 * 1000) return;
  props.setProperty(clave, String(Date.now()));
  try {
    MailApp.sendEmail('lfisbein@gmail.com', 'Problema: ' + que,
      'Ha fallado "' + que + '" (cuenta buenasysantasmadrid):\n\n' + texto +
      '\n\nSi es un corte de Google, se arregla solo en la siguiente vuelta. Si este email se repite, avisar a Claude.');
  } catch (e) { Logger.log('No se pudo mandar el aviso: ' + e); }
}
