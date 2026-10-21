// AVISOS POR EMAIL — la app ya no muestra en pantalla los fallos al
// guardar la caja: cada aviso se apunta en la pestaña "Errores" del Índice
// y le llega a la cuenta que publica la app (la dueña del script) en un
// solo email de resumen, dos veces al día (ver mandarResumenAvisos), en vez
// de un email por aviso. Solo si no se pudo apuntar se manda en el momento
// (el mismo, como mucho una vez cada 6 horas).
var AVISOS_SILENCIO_SEG_ = 6 * 60 * 60;
var AVISOS_HORAS_RESUMEN_ = [9, 18];

function avisarPorEmail_(asunto, detalle) {
  if (registrarError_(asunto, detalle)) {
    asegurarActivadorResumenAvisos_();
    return;
  }
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
// cada 10 minutos. Se guardan los últimos 500. Devuelve false si no se pudo
// apuntar.
var ERRORES_PESTANA_ = 'Errores';
function registrarError_(asunto, detalle) {
  try {
    var huella = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, asunto + '|' + detalle));
    var cache = CacheService.getScriptCache();
    if (cache.get('registro_' + huella)) return true;
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
    return true;
  } catch (err) {
    Logger.log('No se pudo apuntar el error en el registro: ' + err);
    return false;
  }
}

// RESUMEN DE AVISOS — a las 9:00 y a las 18:00 (AVISOS_HORAS_RESUMEN_) un
// solo email con lo apuntado en "Errores" desde el resumen anterior; el
// mismo aviso repetido sale una vez, con cuántas veces pasó. Si no hubo
// nada, no se manda. Los activadores se ponen solos con el primer aviso.
function asegurarActivadorResumenAvisos_() {
  try {
    var cache = CacheService.getScriptCache();
    if (cache.get('activador_resumen_avisos')) return;
    var hay = ScriptApp.getProjectTriggers().filter(function (t) { return t.getHandlerFunction() === 'mandarResumenAvisos'; }).length;
    if (hay < AVISOS_HORAS_RESUMEN_.length) {
      ScriptApp.getProjectTriggers().forEach(function (t) {
        if (t.getHandlerFunction() === 'mandarResumenAvisos') ScriptApp.deleteTrigger(t);
      });
      // Lo de antes ya llegó por email, uno por uno.
      var props = PropertiesService.getScriptProperties();
      if (!props.getProperty('avisos_resumen_desde')) props.setProperty('avisos_resumen_desde', new Date(Date.now() - 1000).toISOString());
      AVISOS_HORAS_RESUMEN_.forEach(function (h) {
        ScriptApp.newTrigger('mandarResumenAvisos').timeBased().everyDays(1).atHour(h).nearMinute(0).inTimezone('Europe/Madrid').create();
      });
    }
    cache.put('activador_resumen_avisos', '1', 21600);
  } catch (err) {
    Logger.log('No se pudo poner el activador del resumen de avisos: ' + err);
  }
}

function mandarResumenAvisos() {
  var props = PropertiesService.getScriptProperties();
  var desde = props.getProperty('avisos_resumen_desde') || new Date(Date.now() - 24 * 3600 * 1000).toISOString();
  var ahora = new Date();
  var r = leerRegistroErrores_(desde);
  if (!r.ok) { Logger.log('Resumen de avisos: ' + r.error); return; }
  // Agrupados por aviso (sin la línea "Cuándo:", que cambia cada vez).
  var grupos = [], porClave = {};
  r.errores.forEach(function (e) {
    var detalle = String(e.detalle || '').replace(/\n*Cuándo:[^\n]*/g, '').replace(/\n*Dispositivo:[^\n]*/g, '').trim();
    var k = e.asunto + '|' + detalle;
    if (!porClave[k]) { porClave[k] = { asunto: e.asunto, detalle: detalle, veces: 0, ultima: e.fecha }; grupos.push(porClave[k]); }
    porClave[k].veces++;
    porClave[k].ultima = e.fecha;
  });
  if (grupos.length) {
    var hora = function (iso) { return Utilities.formatDate(new Date(iso), 'Europe/Madrid', 'dd/MM HH:mm'); };
    var cuerpo = grupos.map(function (g, i) {
      return (i + 1) + '. ' + g.asunto + ' (' + (g.veces > 1 ? g.veces + ' veces, la última el ' : 'el ') + hora(g.ultima) + ')\n' + g.detalle;
    }).join('\n\n');
    var para = Session.getEffectiveUser().getEmail();
    if (para) {
      MailApp.sendEmail(para, 'Cierre de caja: ' + grupos.length + (grupos.length === 1 ? ' aviso' : ' avisos') + ' desde el ' + hora(desde),
        cuerpo + '\n\n(Resumen de las ' + Utilities.formatDate(ahora, 'Europe/Madrid', 'HH:mm') + '. Llega a las ' + AVISOS_HORAS_RESUMEN_.join(':00 y a las ') + ':00 si hubo algo.)');
    }
  }
  props.setProperty('avisos_resumen_desde', ahora.toISOString());
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
