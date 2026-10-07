// ============================================================================
// Página de informes (informes.html) — solo EDITOR_UNICO_
// ============================================================================
// Una página aparte de la caja, para ver estadísticas y hacer preguntas
// sobre la Contabilidad y los Albaranes. Solo entra EDITOR_UNICO_ (aunque
// las otras cuentas puedan iniciar sesión en la caja, acá se les responde
// que no tienen acceso).
//
// De dónde sale cada cosa:
// - Ventas por mes de 2009 en adelante: pestaña "años" de la Contabilidad
//   del año actual (ver AniosComparativa.gs).
// - Promedio por día de la semana de 2016 en adelante: pestaña "comparativa".
// - Día por día del año actual: las pestañas de cada mes de la Contabilidad.
// - Albaranes por proveedor y mes: 2016-2025 de InformesDatos.gs (sacado de
//   los Excel viejos; no va al repo) y el año actual de TOTALES en la
//   planilla de Albaranes.
// Las preguntas libres las responde la IA (Claude) con esos mismos datos.
// ============================================================================

var INFORMES_CACHE_SEGUNDOS_ = 600;
var MODELO_INFORMES_ = 'claude-opus-5-5';

// Desde doPost. Devuelve la respuesta, o null si el pedido no es de informes.
function accionDeInformes_(data, email) {
  if (!data.accionInformesDatos && !data.accionInformesPregunta) return null;
  var res;
  if (email !== EDITOR_UNICO_) {
    res = { ok: false, sinPermiso: true, error: 'Esta página es solo para ' + EDITOR_UNICO_ + '.' };
  } else {
    try {
      res = data.accionInformesDatos ? { ok: true, datos: datosInformes_(!!data.fresco) } : preguntaInformes_(data);
    } catch (err) {
      res = { ok: false, error: String(err && err.message || err) };
    }
  }
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- datos (con memoria rápida de 10 minutos, en trozos de 90 KB) ----------
function datosInformes_(fresco) {
  var cache = CacheService.getScriptCache();
  if (!fresco) {
    var n = Number(cache.get('informes_n') || 0);
    if (n) {
      var partes = cache.getAll(Array.apply(null, Array(n)).map(function (_, i) { return 'informes_' + i; }));
      var texto = '';
      for (var i = 0; i < n; i++) { if (!partes['informes_' + i]) { texto = ''; break; } texto += partes['informes_' + i]; }
      if (texto) return JSON.parse(texto);
    }
  }
  var datos = armarDatosInformes_();
  try {
    var json = JSON.stringify(datos), trozos = {}, k = 0;
    for (var p = 0; p < json.length; p += 90000) trozos['informes_' + (k++)] = json.slice(p, p + 90000);
    trozos.informes_n = String(k);
    cache.putAll(trozos, INFORMES_CACHE_SEGUNDOS_);
  } catch (err) { Logger.log('Informes, memoria rápida: ' + err); }
  return datos;
}

function armarDatosInformes_() {
  var hoy = new Date(), anio = hoy.getFullYear();
  var datos = { generado: hoy.toISOString(), anioActual: anio, ventas: {}, semana: {}, dias: [], albaranes: {}, avisos: [] };
  var idCont = buscarEnIndice_('CONTABILIDAD', String(anio));
  if (idCont) {
    var ss = SpreadsheetApp.openById(idCont);
    var hojas = hojasAniosComparativa_(ss);
    if (hojas.anios) datos.ventas = ventasDeAnios_(hojas.anios); else datos.avisos.push('No está la pestaña "años".');
    if (hojas.comparativa) datos.semana = semanaDeComparativa_(hojas.comparativa); else datos.avisos.push('No está la pestaña "comparativa".');
    for (var m = 0; m <= hoy.getMonth(); m++) {
      var d = datosMesContabilidad_(ss, m);
      if (d) datos.dias = datos.dias.concat(d.lista);
    }
  } else {
    datos.avisos.push('No hay Contabilidad ' + anio + ' en el Índice.');
  }
  if (typeof INFORMES_ALBARANES_HISTORICO_ !== 'undefined') {
    Object.keys(INFORMES_ALBARANES_HISTORICO_).forEach(function (a) { datos.albaranes[a] = INFORMES_ALBARANES_HISTORICO_[a]; });
  } else {
    datos.avisos.push('Faltan los albaranes de 2016-2025 (InformesDatos.gs).');
  }
  var actual = albaranesDelAnio_(anio);
  if (actual) datos.albaranes[String(anio)] = actual;
  return datos;
}

// "años": un bloque por año (el año en A, los encabezados en la fila de
// abajo, una fila por mes). Las columnas cambian de un año a otro, así que se
// leen por su nombre.
var COLUMNAS_VENTAS_ = {
  'DIAS': 'dias', 'TOTAL SIST': 'sistema', 'MEDIO DIA': 'mediodia', 'NOCHE': 'noche',
  'DELIVEROO': 'delivery', 'UBER': 'delivery', 'UBER EAT': 'delivery', 'GLOVO': 'delivery', 'UBER / GLOVO': 'delivery',
  'WEB': 'web', 'EMPANADAS': 'empanadas', 'TOTAL': 'total', 'TARJETAS': 'tarjetas', 'TICKETS': 'tickets',
  'EFEVO': 'efectivo', 'DIFERENCIA': 'diferencia', 'RETIRA': 'retira', 'PROVEEDORES': 'proveedores',
  'NOMINAS': 'nominas', 'EXTRAS': 'extras'
};

function ventasDeAnios_(hoja) {
  var v = hoja.getRange(1, 1, hoja.getLastRow(), Math.min(hoja.getLastColumn(), 26)).getValues();
  var res = {};
  for (var i = 0; i + 2 < v.length; i++) {
    var anio = Number(v[i][0]);
    if (!(anio > 2000 && anio < 2100) || normalizarClave_(v[i + 1][1]) !== 'DIAS') continue;
    var cols = {};
    v[i + 1].forEach(function (h, c) { var k = COLUMNAS_VENTAS_[normalizarClave_(h)]; if (k && cols[k] == null) cols[k] = c; });
    var meses = [];
    for (var r = i + 2; r < Math.min(i + 16, v.length); r++) {
      var mes = MESES_MAYUS_.indexOf(normalizarClave_(v[r][0]));
      if (mes === -1) continue;
      var fila = { mes: mes + 1 };
      Object.keys(cols).forEach(function (k) { var x = v[r][cols[k]]; if (typeof x === 'number' && isFinite(x)) fila[k] = Math.round(x * 100) / 100; });
      if (!(fila.sistema > 0)) continue;
      if (fila.total == null) fila.total = Math.round(((fila.sistema || 0) + (fila.delivery || 0) + (fila.web || 0) + (fila.empanadas || 0)) * 100) / 100;
      meses.push(fila);
    }
    if (meses.length) res[String(anio)] = meses;
  }
  return res;
}

// "comparativa": por año, cada mes con el promedio de cada día de la semana
// (lunes a domingo) en TOTAL, MEDIODÍA y NOCHE.
function semanaDeComparativa_(hoja) {
  var v = hoja.getRange(1, 1, hoja.getLastRow(), 9).getValues();
  var res = {};
  for (var i = 0; i + 1 < v.length; i++) {
    var anio = Number(v[i][0]);
    if (!(anio > 2000 && anio < 2100) || normalizarClave_(v[i + 1][2]) !== 'LUNES' || res[String(anio)]) continue;
    var meses = {};
    for (var r = i + 2; r < Math.min(i + 40, v.length); r++) {
      var mes = MESES_MAYUS_.indexOf(normalizarClave_(v[r][0]));
      if (mes === -1) continue;
      var bloque = {};
      ['total', 'mediodia', 'noche'].forEach(function (k, j) {
        var f = v[r + j] || [];
        bloque[k] = f.slice(2, 9).map(function (x) { return typeof x === 'number' && isFinite(x) ? Math.round(x * 100) / 100 : null; });
      });
      if (bloque.total.some(function (x) { return x; })) meses[String(mes + 1)] = bloque;
    }
    if (Object.keys(meses).length) res[String(anio)] = meses;
  }
  return res;
}

// TOTALES de la planilla de Albaranes del año: una fila por proveedor (B) y
// los meses en C..N, hasta la fila TOTAL. NOO no cuenta (copia de FRUTAPRO).
function albaranesDelAnio_(anio) {
  var id = buscarEnIndice_('ALBARANES', String(anio));
  if (!id) return null;
  var hoja = SpreadsheetApp.openById(id).getSheetByName('TOTALES');
  if (!hoja) return null;
  var filaTotal = filaTotalDeTotales_(hoja);
  if (filaTotal === -1) return null;
  var v = hoja.getRange(2, 2, filaTotal - 2, 13).getValues();
  var res = {};
  v.forEach(function (f) {
    var nombre = normalizarClave_(f[0]);
    if (!nombre || nombre === 'NOO') return;
    var meses = f.slice(1, 13).map(function (x) { return typeof x === 'number' && isFinite(x) ? Math.round(x * 100) / 100 : 0; });
    if (meses.some(function (x) { return x; })) res[nombre] = meses;
  });
  return res;
}

// ---------- preguntas a la IA ----------
var ESQUEMA_INFORME_ = {
  type: 'object',
  additionalProperties: false,
  required: ['respuesta', 'grafico'],
  properties: {
    respuesta: { type: 'string' },
    grafico: {
      type: 'object',
      additionalProperties: false,
      required: ['tipo', 'titulo', 'etiquetas', 'series'],
      properties: {
        tipo: { type: 'string', enum: ['ninguno', 'barras', 'lineas'] },
        titulo: { type: 'string' },
        etiquetas: { type: 'array', items: { type: 'string' } },
        series: {
          type: 'array',
          items: {
            type: 'object',
            additionalProperties: false,
            required: ['nombre', 'valores'],
            properties: { nombre: { type: 'string' }, valores: { type: 'array', items: { type: 'number' } } }
          }
        }
      }
    }
  }
};

var INSTRUCCIONES_INFORMES_ =
  'Eres el analista de datos de "Buenas y Santas", un restaurante argentino de empanadas en Madrid. ' +
  'Respondes a la dueña, en castellano de España, preguntas sobre las ventas y las compras del negocio, usando SOLO los datos de abajo.\n\n' +
  'Cómo leer los datos (JSON):\n' +
  '- ventas[año] = una fila por mes (mes 1-12): dias (días trabajados; medio día = 0,5), sistema (lo facturado en el local, "TOTAL SIST"), ' +
  'mediodia y noche (sistema partido por turno), delivery (Deliveroo/Uber Eats/Glovo según el año), web, empanadas (venta de empanadas a otros negocios), ' +
  'total (sistema + delivery + web + empanadas), tarjetas, tickets (tickets restaurante), efectivo, diferencia (descuadre de caja), retira, ' +
  'proveedores (compras del mes), nominas, extras. Los campos que faltan en un año es porque ese año no se anotaban.\n' +
  '- semana[año][mes] = promedio por día de la semana [lunes..domingo] de lo facturado en TOTAL, mediodia y noche (null = ese día no se abrió).\n' +
  '- dias = cada día trabajado del año actual: f (fecha), s (sistema), md, nc (mediodía, noche), dl (delivery), em (empanadas), tj (tarjetas).\n' +
  '- albaranes[año][PROVEEDOR] = compras a ese proveedor por mes [enero..diciembre]. Algunos proveedores cambiaron de nombre con los años; ' +
  'VARIOS, SUPER, EXTRAS y SERVICIOS agrupan varios.\n' +
  '- El año actual está en curso: al comparar con años anteriores, compara los mismos meses (y avisa si el mes actual va por la mitad).\n\n' +
  'Cómo responder:\n' +
  '- Ve al grano: primero la respuesta (con las cifras), después lo que haga falta para entenderla. Importes en euros con formato español (1.234,56 €).\n' +
  '- Haz las cuentas con cuidado; si una cifra es una estimación o falta un dato, dilo.\n' +
  '- respuesta: texto plano con saltos de línea; puedes usar listas con "- " y **negrita**. Nada de tablas en texto.\n' +
  '- grafico: si un gráfico ayuda (evolución, comparación entre meses/años/proveedores), rellénalo: "lineas" para evolución en el tiempo, ' +
  '"barras" para comparar; etiquetas = el eje X y cada serie con un valor por etiqueta (máximo 4 series). Si no hace falta, tipo "ninguno" con listas vacías.';

function preguntaInformes_(data) {
  var pregunta = String(data.pregunta || '').trim();
  if (!pregunta) return { ok: false, error: 'Escribe una pregunta.' };
  var clave = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!clave) return { ok: false, error: 'Falta la clave ANTHROPIC_API_KEY en las Propiedades del script' };

  // Los datos van en el system con cache_control: si se hacen varias
  // preguntas seguidas, la IA no los vuelve a leer enteros (más rápido y barato).
  var datos = datosInformes_(false);
  var compacto = {
    ventas: datos.ventas, semana: datos.semana, albaranes: datos.albaranes,
    dias: datos.dias.map(function (d) { return { f: d.fecha, s: d.sistema, md: d.mediodia, nc: d.noche, dl: d.delivery, em: d.empanadas, tj: d.tarjetas }; })
  };
  var mensajes = [];
  (data.historial || []).slice(-6).forEach(function (h) {
    if (!h || !h.pregunta || !h.respuesta) return;
    mensajes.push({ role: 'user', content: String(h.pregunta) });
    mensajes.push({ role: 'assistant', content: String(h.respuesta) });
  });
  mensajes.push({ role: 'user', content: pregunta + '\n\n(Hoy es ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd') + '.)' });

  var cuerpo = {
    model: MODELO_INFORMES_,
    max_tokens: 16000,
    output_config: { effort: 'medium', format: { type: 'json_schema', schema: ESQUEMA_INFORME_ } },
    // Si la IA no puede responder con este modelo, la API prueba sola con otro.
    fallbacks: 'default',
    system: [
      { type: 'text', text: INSTRUCCIONES_INFORMES_ },
      { type: 'text', text: 'DATOS:\n' + JSON.stringify(compacto), cache_control: { type: 'ephemeral' } }
    ],
    messages: mensajes
  };
  var resp, codigo, json;
  for (var intento = 0; intento < 3; intento++) {
    resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-api-key': clave, 'anthropic-version': '2023-06-01', 'anthropic-beta': 'server-side-fallback-2026-07-01' },
      payload: JSON.stringify(cuerpo),
      muteHttpExceptions: true
    });
    codigo = resp.getResponseCode();
    if (codigo !== 429 && codigo < 500) break;
    Utilities.sleep(2000 * (intento + 1));
  }
  try { json = JSON.parse(resp.getContentText()); } catch (err) { json = null; }
  if (codigo !== 200 || !json) return { ok: false, error: (json && json.error && json.error.message) || ('Error de la IA (' + codigo + ')') };
  if (json.stop_reason === 'refusal') return { ok: false, error: 'La IA no quiso responder a esto.' };
  if (json.stop_reason === 'max_tokens') return { ok: false, error: 'La respuesta salió demasiado larga. Prueba con una pregunta más concreta.' };
  var texto = (json.content || []).filter(function (b) { return b.type === 'text'; }).map(function (b) { return b.text; }).join('');
  try {
    var r = JSON.parse(texto);
    return { ok: true, respuesta: r.respuesta, grafico: r.grafico };
  } catch (err) {
    return { ok: false, error: 'No se pudo entender la respuesta de la IA.' };
  }
}
