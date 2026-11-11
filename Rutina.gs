// ============================================================================
// RUTINA DE CLAUDE EN LA NUBE — cada 3 horas (9 a 24 h) una sesión de Claude
// en la nube (sin ordenador encendido) lee los errores nuevos, arregla lo que
// puede en el repositorio de GitHub, lo publica y solo le escribe a la dueña
// cuando hace falta que ella haga algo.
//
// La rutina entra con CLAVE_RUTINA_ (ClaveRutina.gs, fuera del repositorio
// público; la clave va solo en las instrucciones privadas de la rutina). Con
// esa clave solo puede hacer lo de aquí abajo: leer errores y cajas, la
// comprobación de salud, mandarle un email a la dueña y publicar lo que esté
// en GitHub (o volver a una versión anterior).
//
//   GET ?rutina=CLAVE&errores=1                → errores desde la última revisión
//   GET ?rutina=CLAVE&marcarRevisado=ISO       → la próxima vez, solo los de después
//   GET ?rutina=CLAVE&fecha=AAAA-MM-DD         → una caja
//   GET ?rutina=CLAVE&mes=AAAA-MM              → todas las cajas del mes
//   GET ?rutina=CLAVE&salud=1                  → ¿la app puede guardar?
//   GET ?rutina=CLAVE&avisar=1&asunto=..&detalle=..[&html=1]  → email a la dueña
//   GET ?rutina=CLAVE&publicar=SHA&descripcion=..[&nuevos=Archivo1,Archivo2]
//   GET ?rutina=CLAVE&volverAVersion=N         → la app vuelve a la versión N
//   GET ?rutina=CLAVE&versiones=1              → últimas versiones publicadas
// ============================================================================

var RUTINA_REPO_RAW_ = 'https://raw.githubusercontent.com/buenasysantasmadrid/CIERRE-CAJA/';
var RUTINA_IMPLEMENTACION_ = 'AKfycbyrfOE73JqgjTXAEaLetRz8aJg1Ldb-Du4dsqGgW34H-zgwLSJMrAkSKApiTCoz6t7m_A';

function esRutina_(parametros) {
  return typeof CLAVE_RUTINA_ !== 'undefined' && !!CLAVE_RUTINA_ && parametros.rutina === CLAVE_RUTINA_;
}

function respuestaRutina_(p) {
  var r;
  try {
    if (p.errores) {
      var props = PropertiesService.getScriptProperties();
      var desde = p.desde || props.getProperty('rutina_revisado_hasta') || new Date(Date.now() - 24 * 3600 * 1000).toISOString();
      r = leerRegistroErrores_(desde);
      r.desde = desde;
    } else if (p.marcarRevisado) {
      var d = new Date(String(p.marcarRevisado));
      if (isNaN(d.getTime())) throw new Error('marcarRevisado: fecha no válida');
      PropertiesService.getScriptProperties().setProperty('rutina_revisado_hasta', d.toISOString());
      r = { ok: true, revisadoHasta: d.toISOString() };
    } else if (p.mes) {
      r = obtenerMesJSON_(String(p.mes));
    } else if (p.fecha) {
      r = obtenerDiaJSON_(planillaCierreCaja_(String(p.fecha)), String(p.fecha));
    } else if (p.salud) {
      r = comprobarSalud_();
    } else if (p.avisar) {
      r = avisarALaDuenia_(String(p.asunto || 'aviso'), String(p.detalle || ''), !!p.html);
    } else if (p.publicar) {
      r = publicarDesdeGitHub_(String(p.publicar), String(p.descripcion || 'Arreglo de la rutina'), String(p.nuevos || ''));
    } else if (p.volverAVersion) {
      r = desplegarVersion_(parseInt(p.volverAVersion, 10), 'Vuelta a la versión ' + p.volverAVersion + ' (rutina)');
    } else if (p.versiones) {
      r = ultimasVersiones_();
    } else {
      r = { ok: false, error: 'Nada que hacer. Ver Rutina.gs.' };
    }
  } catch (err) {
    r = { ok: false, error: String(err), donde: String(err && err.stack || '') };
  }
  return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
}

// Email a la dueña (solo cuando hace falta que ella haga algo).
function avisarALaDuenia_(asunto, detalle, esHtml) {
  var opciones = { to: EDITOR_UNICO_, subject: 'Cierre de caja: ' + asunto.slice(0, 150) };
  if (esHtml) { opciones.htmlBody = detalle; opciones.body = detalle.replace(/<[^>]+>/g, ' '); }
  else opciones.body = detalle;
  MailApp.sendEmail(opciones);
  return { ok: true, enviadoA: EDITOR_UNICO_ };
}

// ---------- publicar desde GitHub ----------
// Trae de GitHub (en el commit SHA) cada archivo del proyecto: los que están
// en el repositorio se reemplazan; los que no (claves y datos, que no van al
// repositorio público) se dejan como están. Nunca se borra nada. Después se
// crea una versión y la app pasa a usarla (misma URL), corriendo con la
// cuenta de la dueña (la que corre esto).
function apiScript_(metodo, ruta, cuerpo) {
  var opciones = {
    method: metodo, muteHttpExceptions: true, contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
  };
  if (cuerpo) opciones.payload = JSON.stringify(cuerpo);
  var resp = UrlFetchApp.fetch('https://script.googleapis.com/v1/projects/' + ScriptApp.getScriptId() + ruta, opciones);
  var codigo = resp.getResponseCode();
  if (codigo >= 300) throw new Error('Apps Script API ' + metodo + ' ' + ruta + ' → ' + codigo + ': ' + resp.getContentText().slice(0, 500));
  return JSON.parse(resp.getContentText() || '{}');
}

function publicarDesdeGitHub_(sha, descripcion, nuevos) {
  if (!/^[0-9a-f]{7,40}$/.test(sha)) throw new Error('publicar: hace falta el SHA del commit (7 a 40 caracteres hexadecimales).');
  var actual = apiScript_('get', '/content');
  var archivos = (actual.files || []).map(function (f) { return { name: f.name, type: f.type, source: f.source }; });
  nuevos.split(',').map(function (n) { return n.trim().replace(/\.gs$/, ''); }).filter(String).forEach(function (n) {
    if (!/^[A-Za-z0-9_]+$/.test(n)) throw new Error('Nombre de archivo nuevo no válido: ' + n);
    if (!archivos.some(function (f) { return f.name === n; })) archivos.push({ name: n, type: 'SERVER_JS', source: '' });
  });
  var cambiados = [], sinCambios = [], fueraDelRepo = [];
  archivos.forEach(function (f) {
    var ruta = f.type === 'JSON' ? f.name + '.json' : f.type === 'SERVER_JS' ? f.name + '.gs' : null;
    if (!ruta) { fueraDelRepo.push(f.name); return; }
    var resp = UrlFetchApp.fetch(RUTINA_REPO_RAW_ + sha + '/' + encodeURIComponent(ruta), { muteHttpExceptions: true });
    if (resp.getResponseCode() === 404) { fueraDelRepo.push(f.name); return; }
    if (resp.getResponseCode() !== 200) throw new Error('No se pudo leer ' + ruta + ' de GitHub (' + resp.getResponseCode() + ').');
    var texto = resp.getContentText('UTF-8');
    if (!texto.trim()) throw new Error(ruta + ' está vacío en GitHub: no se publica.');
    if (texto === f.source) { sinCambios.push(f.name); return; }
    f.source = texto;
    cambiados.push(f.name);
  });
  archivos = archivos.filter(function (f) { return f.source !== ''; });
  if (!cambiados.length) return { ok: true, sinCambios: true, nota: 'Google ya tenía lo mismo que GitHub en ' + sha + '.' };
  apiScript_('put', '/content', { files: archivos });
  var version = apiScript_('post', '/versions', { description: (descripcion + ' (' + sha.slice(0, 7) + ')').slice(0, 200) });
  desplegarVersion_(version.versionNumber, descripcion);
  return { ok: true, version: version.versionNumber, cambiados: cambiados, fueraDelRepo: fueraDelRepo };
}

function desplegarVersion_(numero, descripcion) {
  if (!(numero > 0)) throw new Error('Número de versión no válido.');
  apiScript_('put', '/deployments/' + RUTINA_IMPLEMENTACION_, {
    deploymentConfig: { scriptId: ScriptApp.getScriptId(), versionNumber: numero, manifestFileName: 'appsscript', description: String(descripcion).slice(0, 200) }
  });
  return { ok: true, version: numero };
}

function ultimasVersiones_() {
  var todas = [], token = '';
  do {
    var r = apiScript_('get', '/versions?pageSize=50' + (token ? '&pageToken=' + encodeURIComponent(token) : ''));
    todas = todas.concat(r.versions || []);
    token = r.nextPageToken || '';
  } while (token);
  var imp = apiScript_('get', '/deployments/' + RUTINA_IMPLEMENTACION_);
  return {
    ok: true,
    enUso: imp.deploymentConfig && imp.deploymentConfig.versionNumber,
    ultimas: todas.sort(function (a, b) { return b.versionNumber - a.versionNumber; }).slice(0, 8).map(function (v) { return { version: v.versionNumber, descripcion: v.description || '', creada: v.createTime }; })
  };
}

// Para correr a mano UNA VEZ desde el editor de Apps Script (con lfisbein):
// pide los permisos nuevos (publicar el proyecto) y comprueba que funcionan.
function autorizarRutina() {
  var r = ultimasVersiones_();
  Logger.log('Permisos OK. Versión en uso: ' + r.enUso);
  return r;
}
