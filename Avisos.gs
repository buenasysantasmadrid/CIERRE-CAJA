// AVISOS POR EMAIL — la app ya no muestra en pantalla los fallos al
// guardar la caja: cuando algo no se guarda, le llega un email a la cuenta
// que publica la app (la dueña del script). El mismo aviso se manda como
// mucho una vez cada 6 horas, para no llenar el correo si se repite.
var AVISOS_SILENCIO_SEG_ = 6 * 60 * 60;

function avisarPorEmail_(asunto, detalle) {
  registrarError_(asunto, detalle);
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

// REGISTRO DE ERRORES — cada aviso queda también en la pestaña "Errores"
// del Índice (fecha, asunto, detalle), para que Claude los lea solo desde el
// ordenador (GET ?errores=1, ver doGet) y los arregle sin que haya que
// pasarle los emails. El mismo aviso repetido se apunta como mucho una vez
// cada 10 minutos. Se guardan los últimos 500.
var ERRORES_PESTANA_ = 'Errores';
function registrarError_(asunto, detalle) {
  try {
    var huella = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, asunto + '|' + detalle));
    var cache = CacheService.getScriptCache();
    if (cache.get('registro_' + huella)) return;
    cache.put('registro_' + huella, '1', 600);
    var ss = SpreadsheetApp.openById(INDICE_SHEET_ID_);
    var hoja = ss.getSheetByName(ERRORES_PESTANA_);
    if (!hoja) {
      hoja = ss.insertSheet(ERRORES_PESTANA_);
      hoja.appendRow(['Fecha', 'Asunto', 'Detalle']);
      hoja.setFrozenRows(1);
    }
    hoja.appendRow([new Date(), String(asunto).slice(0, 300), String(detalle || '').slice(0, 5000)]);
    var sobran = hoja.getLastRow() - 1 - 500;
    if (sobran > 50) hoja.deleteRows(2, sobran);
  } catch (err) {
    Logger.log('No se pudo apuntar el error en el registro: ' + err);
  }
}

// Los errores apuntados desde `desde` (fecha ISO; sin ella, los de los
// últimos 3 días), del más viejo al más nuevo.
function leerRegistroErrores_(desde) {
  try {
    var hoja = SpreadsheetApp.openById(INDICE_SHEET_ID_).getSheetByName(ERRORES_PESTANA_);
    if (!hoja || hoja.getLastRow() < 2) return { ok: true, errores: [] };
    var limite = desde ? new Date(desde) : new Date(Date.now() - 3 * 24 * 3600 * 1000);
    var filas = hoja.getRange(2, 1, hoja.getLastRow() - 1, 3).getValues();
    return {
      ok: true,
      errores: filas.filter(function (f) { return f[0] instanceof Date && f[0] > limite; }).map(function (f) {
        return { fecha: f[0].toISOString(), asunto: f[1], detalle: f[2] };
      })
    };
  } catch (err) {
    return { ok: false, error: String(err) };
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
