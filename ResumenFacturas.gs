// ============================================================================
// RESUMEN SEMANAL DE FACTURAS — lo que pasa con las facturas sin que nadie
// lo mire (las de las carpetas de entrada de Drive, ver Bandejas.gs, y los
// PDF que se mueven de carpeta al cambiar el proveedor de un gasto, ver
// moverPdfDeGasto_ en Precios.gs) se apunta en la pestaña "Movimientos
// facturas" del Índice, y los lunes a las 8:00 llega un solo email con lo de
// la semana (a la cuenta que publica la app). Si no hubo nada, no se manda.
// Lo que falla además va al resumen de avisos (Avisos.gs), como siempre.
// ============================================================================

var MOV_FACTURAS_PESTANA_ = 'Movimientos facturas';
// Tipos de movimiento, en el orden en que salen en el email.
var MOV_FACTURAS_TIPOS_ = [
  ['sin proveedor', 'FACTURAS SIN PROVEEDOR DE LA LISTA', 'Se cargaron igual (las de transferencia en SERVICIOS). Si son de un proveedor habitual, hay que añadirlo a la lista de la app.'],
  ['revisar', 'PARA REVISAR', ''],
  ['cargada', 'FACTURAS CARGADAS DESDE LAS CARPETAS DE DRIVE', ''],
  ['movido', 'PDF MOVIDOS DE CARPETA', '']
];

function apuntarMovFactura_(tipo, texto, enlace) {
  try {
    var ss = SpreadsheetApp.openById(INDICE_SHEET_ID_);
    var hoja = ss.getSheetByName(MOV_FACTURAS_PESTANA_);
    if (!hoja) {
      hoja = ss.insertSheet(MOV_FACTURAS_PESTANA_);
      hoja.appendRow(['Fecha', 'Tipo', 'Detalle', 'Enlace']);
      hoja.setFrozenRows(1);
    }
    hoja.appendRow([new Date(), tipo, String(texto || '').slice(0, 5000), enlace || '']);
    var sobran = hoja.getLastRow() - 1 - 2000;
    if (sobran > 100) hoja.deleteRows(2, sobran);
    asegurarActivadorResumenFacturas_();
  } catch (err) {
    Logger.log('No se pudo apuntar el movimiento de factura: ' + err);
  }
}

function asegurarActivadorResumenFacturas_() {
  var cache = CacheService.getScriptCache();
  if (cache.get('activador_resumen_facturas')) return;
  var hay = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'mandarResumenFacturas'; });
  if (!hay) ScriptApp.newTrigger('mandarResumenFacturas').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).nearMinute(0).inTimezone('Europe/Madrid').create();
  cache.put('activador_resumen_facturas', '1', 21600);
}

// Lo que corre los lunes a las 8:00.
function mandarResumenFacturas() {
  var props = PropertiesService.getScriptProperties();
  var desde = props.getProperty('resumen_facturas_desde') || new Date(Date.now() - 7 * 86400000).toISOString();
  var ahora = new Date();
  var r = resumenFacturas_(new Date(desde), ahora);
  if (r.cuantos) MailApp.sendEmail(Session.getEffectiveUser().getEmail(), r.asunto, r.cuerpo);
  props.setProperty('resumen_facturas_desde', ahora.toISOString());
}

// Una muestra (Claude: GET ?muestraResumenFacturas=1): lo de los últimos 7
// días, sin cambiar desde cuándo cuenta el resumen de verdad.
function mandarMuestraResumenFacturas_() {
  var r = resumenFacturas_(new Date(Date.now() - 7 * 86400000), new Date());
  MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'MUESTRA — ' + r.asunto, r.cuerpo);
  return { ok: true, cuantos: r.cuantos };
}

function resumenFacturas_(desde, hasta) {
  var hoja = SpreadsheetApp.openById(INDICE_SHEET_ID_).getSheetByName(MOV_FACTURAS_PESTANA_);
  var filas = hoja && hoja.getLastRow() > 1 ? hoja.getRange(2, 1, hoja.getLastRow() - 1, 4).getValues() : [];
  filas = filas.filter(function (f) { return f[0] instanceof Date && f[0] >= desde && f[0] <= hasta; });
  var dia = function (d) { return Utilities.formatDate(d, 'Europe/Madrid', 'dd/MM'); };
  var partes = [];
  MOV_FACTURAS_TIPOS_.forEach(function (t) {
    var deEste = filas.filter(function (f) { return f[1] === t[0]; });
    if (!deEste.length) return;
    partes.push(t[1] + ' (' + deEste.length + ')' + (t[2] ? '\n' + t[2] : '') + '\n\n' + deEste.map(function (f) {
      return '• ' + dia(f[0]) + ' — ' + f[2] + (f[3] ? '\n  ' + f[3] : '');
    }).join('\n'));
  });
  var cuantos = filas.length;
  var periodo = 'del ' + dia(desde) + ' al ' + dia(hasta);
  return {
    cuantos: cuantos,
    asunto: 'Cierre de caja: resumen de facturas ' + periodo,
    cuerpo: cuantos
      ? 'Resumen de facturas ' + periodo + '.\n\n' + partes.join('\n\n────────────\n\n') +
        '\n\n(Llega los lunes a las 8:00 si hubo algo. Lo que falla también llega en el resumen de avisos.)'
      : 'No hubo movimientos de facturas ' + periodo + '.'
  };
}
