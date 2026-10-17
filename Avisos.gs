// AVISOS POR EMAIL — la app ya no muestra en pantalla los fallos al
// guardar la caja: cuando algo no se guarda, le llega un email a la cuenta
// que publica la app (la dueña del script). El mismo aviso se manda como
// mucho una vez cada 6 horas, para no llenar el correo si se repite.
var AVISOS_SILENCIO_SEG_ = 6 * 60 * 60;

function avisarPorEmail_(asunto, detalle) {
  try {
    var huella = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, asunto + '|' + detalle);
    var clave = 'aviso_' + Utilities.base64EncodeWebSafe(huella);
    var cache = CacheService.getScriptCache();
    if (cache.get(clave)) return;
    var para = Session.getEffectiveUser().getEmail();
    if (!para) return;
    MailApp.sendEmail(para, 'Cierre de caja: ' + asunto,
      detalle + '\n\n(' + Utilities.formatDate(new Date(), 'Europe/Madrid', 'dd/MM/yyyy HH:mm') + ')');
    cache.put(clave, '1', AVISOS_SILENCIO_SEG_);
  } catch (err) {
    Logger.log('No se pudo mandar el aviso por email: ' + err);
  }
}

// Después de guardar la caja: si alguna planilla no se pudo actualizar.
function avisarFallosDelGuardado_(data, partes) {
  var fallos = Object.keys(partes).filter(function (k) {
    return partes[k] && partes[k].ok === false;
  }).map(function (k) { return '- ' + k + ': ' + partes[k].error; });
  if (!fallos.length) return;
  avisarPorEmail_('algo no se guardó (caja del ' + data.fecha + ')',
    'Al guardar la caja del ' + data.fecha + ' no se pudo actualizar:\n' + fallos.join('\n') +
    '\n\nLa caja quedó guardada en el Registro; lo de arriba hay que revisarlo.');
}

// Lo que avisa la app (ej. un dispositivo que estuvo mucho rato sin poder
// guardar en el Sheet).
function avisoDesdeLaApp_(data) {
  avisarPorEmail_(String(data.asunto || 'aviso de la app').slice(0, 120), String(data.detalle || '').slice(0, 2000));
  return { ok: true };
}

// Para correr a mano desde el editor la primera vez: pide el permiso de
// mandar emails y manda uno de prueba.
function probarAvisoPorEmail() {
  MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'Cierre de caja: prueba de avisos',
    'Si te llegó este email, los avisos de errores de la caja funcionan.');
}
