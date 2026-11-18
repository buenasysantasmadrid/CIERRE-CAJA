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
var REMITENTES = ['info@buenasysantas.es', 'luciana@buenasysantas.es', 'clara@buenasysantas.es', 'yamila@buenasysantas.es', 'lfisbein@gmail.com', 'yamilafisbein@gmail.com', 'efronclara@gmail.com'];
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

function traerFacturasDelCorreo() {
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
