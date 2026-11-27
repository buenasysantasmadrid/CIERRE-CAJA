// Lo que cuesta usar Claude (26/11/2026: Luciana vio una recarga de créditos y quiere saber el coste exacto).
// Cada llamada a Claude apunta sus tokens y su coste en dólares (sin IVA) por día, motivo y modelo,
// en las propiedades del script. ?precios=gasto lo devuelve. ?precios=medir&n=5 cuenta, SIN GASTAR,
// los tokens que manda una lectura de precios (con el contador gratuito de Anthropic).

var MEDIR_TOKENS_ = false;      // true solo dentro de medirPrecios_
var MEDIDO_ = 0;
var GASTO_MOTIVO_ = 'escaneo';  // 'precios' mientras lee precios (PreciosLectura.gs)

// Dólares por millón de tokens: [entrada, salida, lectura de caché]. Precios oficiales
// (platform.claude.com/docs/en/about-claude/pricing, consultado el 26/11/2026).
var PRECIOS_CLAUDE_USD_ = {
  'claude-sonnet-5': [2, 10, 0.2],
  'claude-sonnet-5-5': [2, 10, 0.1],
  'claude-haiku-5-5': [0.10, 0.50, 0.01],
  'claude-opus-5-5': [4, 20, 0.2]
};

function costeUsd_(modelo, entrada, salida, cacheLee, cacheEscribe) {
  var p = PRECIOS_CLAUDE_USD_[modelo];
  if (!p) return null;
  return (entrada * p[0] + salida * p[1] + (cacheLee || 0) * p[2] + (cacheEscribe || 0) * p[0] * 1.25) / 1e6;
}

function apuntarGastoClaude_(modelo, usage) {
  if (!usage) return;
  var ent = Number(usage.input_tokens) || 0, sal = Number(usage.output_tokens) || 0;
  var cl = Number(usage.cache_read_input_tokens) || 0, ce = Number(usage.cache_creation_input_tokens) || 0;
  var usd = costeUsd_(modelo, ent, sal, cl, ce) || 0;
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return; }
  try {
    var props = PropertiesService.getScriptProperties();
    var clave = 'gasto_' + Utilities.formatDate(new Date(), 'Europe/Madrid', 'yyyy-MM-dd');
    var d = {};
    try { d = JSON.parse(props.getProperty(clave) || '{}'); } catch (e) { d = {}; }
    var k = GASTO_MOTIVO_ + ' · ' + modelo;
    var x = d[k] || { lecturas: 0, entrada: 0, salida: 0, usd: 0 };
    x.lecturas++; x.entrada += ent + cl + ce; x.salida += sal; x.usd = Math.round((x.usd + usd) * 10000) / 10000;
    d[k] = x;
    props.setProperty(clave, JSON.stringify(d));
  } finally {
    lock.releaseLock();
  }
}

/** Gasto apuntado por día (desde el 26/11/2026), en dólares sin IVA. */
function gastoClaude_() {
  var props = PropertiesService.getScriptProperties().getProperties();
  var dias = Object.keys(props).filter(function (k) { return /^gasto_\d{4}-\d{2}-\d{2}$/.test(k); }).sort();
  var total = 0, out = {};
  dias.forEach(function (k) {
    var d = JSON.parse(props[k]);
    out[k.slice(6)] = d;
    Object.keys(d).forEach(function (m) { total += d[m].usd; });
  });
  return { ok: true, nota: 'Dólares sin IVA (el IVA español es un 21 % más)', totalUsd: Math.round(total * 100) / 100, dias: out };
}

/** Cuenta, sin gastar, los tokens que manda la lectura de precios de n documentos ya leídos. */
function medirPrecios_(n) {
  n = Math.max(1, Math.min(Number(n) || 5, 15));
  var h = hojaArchivosPrecios_();
  var filas = h.getLastRow() > 1 ? h.getRange(2, 1, h.getLastRow() - 1, PRECIOS_ARCH_CAB_.length).getValues() : [];
  var leidos = filas.filter(function (f) { return f[0] && String(f[3]) === 'leído'; });
  var muestra = [];
  for (var i = 0; i < n && i < leidos.length; i++) muestra.push(leidos[Math.floor(i * leidos.length / n)]);
  var t0 = Date.now(), res = [];
  MEDIR_TOKENS_ = true;
  try {
    muestra.forEach(function (f) {
      if (Date.now() - t0 > 240000) return;
      MEDIDO_ = 0;
      var archivo = DriveApp.getFileById(String(f[0]));
      try { leerYGuardarPrecios_(contenidoDeArchivo_(archivo), metaDeArchivo_(archivo)); } catch (e) {}
      res.push({ archivo: String(f[1]), carpeta: String(f[2]), lineas: Number(f[4]) || 0, tokensEntrada: MEDIDO_ });
    });
  } finally {
    MEDIR_TOKENS_ = false;
  }
  var lineasTot = 0;
  leidos.forEach(function (f) { lineasTot += Number(f[4]) || 0; });
  return { ok: true, modelo: MODELO_ESCANEO, documentosLeidos: leidos.length, lineasLeidas: lineasTot,
    totalArchivos: filas.length, muestra: res };
}
