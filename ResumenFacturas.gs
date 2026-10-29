// ============================================================================
// RESUMEN SEMANAL DE FACTURAS — los lunes a las 8:00 llega un email (a la
// cuenta que publica la app) solo con lo que hay que mirar, no con lo que
// salió bien:
//   · Fallos: facturas de las carpetas de Drive (Bandejas.gs) que no se
//     pudieron cargar o cuyo albarán no apareció.
//   · Diferencias: la factura dice otro importe que el albarán.
//   · En Varios o sin proveedor: gastos cargados en Varios en la semana (de
//     las cajas) y facturas de las carpetas que no son de ningún proveedor
//     de la lista, para ver si hay que pasarlos a uno.
//   · Y cómo está todo hoy (RevisionAlbaranes.gs): errores de la semana,
//     facturas repetidas en ALBARANES, lo que falta pagar y los albaranes
//     que todavía no tienen su factura.
// De cada uno, solo proveedor, total y qué pasa. Los fallos y diferencias se
// apuntan en la pestaña "Movimientos facturas" del Índice; lo de Varios se
// saca de las cajas al armar el email. Si no hay nada, no se manda.
// ============================================================================

var MOV_FACTURAS_PESTANA_ = 'Movimientos facturas';
var MOV_FACTURAS_ENCABEZADOS_ = ['Fecha', 'Tipo', 'Proveedor', 'Total', 'Qué pasa', 'Enlace'];
// Secciones del email: tipo, título, explicación, color.
var MOV_FACTURAS_SECCIONES_ = [
  ['fallo', 'Fallos', 'No se pudieron hacer solas: hay que mirarlas.', '#c0392b'],
  ['error', 'Errores al guardar', 'Lo que falló esta semana (también llegó en los avisos de las 9 y las 18).', '#c0392b'],
  ['repetido', 'Repetidas en ALBARANES', 'La misma factura está dos veces: una sobra.', '#8e44ad'],
  ['diferencia', 'Diferencias de importe', 'La factura no dice lo mismo que el albarán.', '#d35400'],
  ['sinpagar', 'Falta pagar', 'Cargadas como "No pagado" y sin día de pago.', '#2471a3'],
  ['albaran', 'Albaranes sin factura', 'Escaneados como albarán; su factura todavía no llegó.', '#16a085'],
  ['varios', 'En Varios o sin proveedor', '¿Hay que pasarlas a un proveedor?', '#b7950b']
];

// tipo: 'fallo', 'diferencia' o 'varios'.
function apuntarMovFactura_(tipo, proveedor, total, que, enlace) {
  try {
    var ss = SpreadsheetApp.openById(INDICE_SHEET_ID_);
    var hoja = ss.getSheetByName(MOV_FACTURAS_PESTANA_);
    if (!hoja) hoja = ss.insertSheet(MOV_FACTURAS_PESTANA_);
    if (hoja.getRange(1, 3).getValue() !== 'Proveedor') { // nueva, o de la primera versión (otras columnas)
      hoja.clear();
      hoja.appendRow(MOV_FACTURAS_ENCABEZADOS_);
      hoja.setFrozenRows(1);
    }
    hoja.appendRow([new Date(), tipo, String(proveedor || '—'), Number(total) || '', String(que || '').slice(0, 1000), enlace || '']);
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
  if (r.cuantos) MailApp.sendEmail({ to: Session.getEffectiveUser().getEmail(), subject: r.asunto, htmlBody: r.html, body: r.texto });
  props.setProperty('resumen_facturas_desde', ahora.toISOString());
}

// Una muestra (Claude: GET ?muestraResumenFacturas=1): lo de los últimos 7
// días, sin cambiar desde cuándo cuenta el resumen de verdad.
function mandarMuestraResumenFacturas_() {
  var r = resumenFacturas_(new Date(Date.now() - 7 * 86400000), new Date());
  MailApp.sendEmail({ to: Session.getEffectiveUser().getEmail(), subject: 'MUESTRA — ' + r.asunto, htmlBody: r.html, body: r.texto });
  return { ok: true, cuantos: r.cuantos };
}

// Los gastos cargados en Varios en las cajas de esas fechas.
function gastosEnVarios_(desde, hasta) {
  var d = Utilities.formatDate(desde, 'Europe/Madrid', 'yyyy-MM-dd'), h = Utilities.formatDate(hasta, 'Europe/Madrid', 'yyyy-MM-dd');
  var r = listarFacturasProveedores_('Varios', '');
  return (r.facturas || []).filter(function (f) { return f.fecha >= d && f.fecha <= h; }).map(function (f) {
    return {
      fecha: fechaComoDate_(f.fecha), proveedor: String(f.proveedorTexto || '').replace(/^Varios\s*—\s*/, '') || 'Varios',
      total: Number(f.importe) || 0, que: 'Cargada en Varios (caja del ' + fechaCorta_(f.fecha).slice(0, 5) + ')', enlace: ''
    };
  });
}

function resumenFacturas_(desde, hasta) {
  var hoja = SpreadsheetApp.openById(INDICE_SHEET_ID_).getSheetByName(MOV_FACTURAS_PESTANA_);
  var filas = [];
  if (hoja && hoja.getLastRow() > 1 && hoja.getRange(1, 3).getValue() === 'Proveedor') {
    filas = hoja.getRange(2, 1, hoja.getLastRow() - 1, 6).getValues()
      .filter(function (f) { return f[0] instanceof Date && f[0] >= desde && f[0] <= hasta; })
      .map(function (f) { return { fecha: f[0], tipo: f[1], proveedor: f[2], total: f[3], que: f[4], enlace: f[5] }; });
  }
  try {
    gastosEnVarios_(desde, hasta).forEach(function (g) { g.tipo = 'varios'; filas.push(g); });
  } catch (err) { Logger.log('Resumen de facturas, Varios: ' + err); }
  try { estadoAlbaranes_(hasta).forEach(function (f) { filas.push(f); }); } catch (err) { Logger.log('Resumen de facturas, ALBARANES: ' + err); }
  try {
    erroresDeLaSemana_(desde).forEach(function (e) {
      filas.push({ tipo: 'error', fecha: new Date(e.ultima), proveedor: e.asunto, total: '', que: (e.veces > 1 ? e.veces + ' veces. ' : '') + String(e.detalle || '').split('\n')[0].slice(0, 200) });
    });
  } catch (err) { Logger.log('Resumen de facturas, errores: ' + err); }

  var esc = function (s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); };
  var dia = function (d) { return Utilities.formatDate(d, 'Europe/Madrid', 'dd/MM'); };
  var periodo = 'del ' + dia(desde) + ' al ' + dia(hasta);
  var html = [], texto = [];
  MOV_FACTURAS_SECCIONES_.forEach(function (s) {
    var deEsta = filas.filter(function (f) { return f.tipo === s[0]; });
    if (!deEsta.length) return;
    deEsta.sort(function (a, b) { return a.fecha - b.fecha; });
    var suma = deEsta.reduce(function (t, f) { return t + (Number(f.total) || 0); }, 0);
    var cuenta = deEsta.length + (s[0] === 'sinpagar' && suma ? ' · ' + euros_(suma) : '');
    html.push('<h2 style="margin:28px 0 4px;font-size:17px;color:' + s[3] + '">' + s[1] + ' <span style="font-weight:normal;color:#888">(' + cuenta + ')</span></h2>' +
      '<p style="margin:0 0 10px;color:#666;font-size:13px">' + s[2] + '</p>' +
      '<table style="border-collapse:collapse;width:100%;font-size:14px;border-left:4px solid ' + s[3] + '">' +
      deEsta.map(function (f, i) {
        return '<tr style="background:' + (i % 2 ? '#fff' : '#f7f7f7') + '">' +
          '<td style="padding:8px 10px;font-weight:bold;white-space:nowrap">' + esc(f.proveedor) + '</td>' +
          '<td style="padding:8px 10px;text-align:right;white-space:nowrap">' + (f.total ? euros_(f.total) : '') + '</td>' +
          '<td style="padding:8px 10px">' + esc(f.que) + '</td>' +
          '<td style="padding:8px 10px;color:#999;font-size:12px;white-space:nowrap">' + dia(f.fecha) +
          (f.enlace ? ' · <a href="' + esc(f.enlace) + '" style="color:#2471a3">PDF</a>' : '') + '</td></tr>';
      }).join('') + '</table>');
    texto.push(s[1].toUpperCase() + ' (' + cuenta + ')\n' + deEsta.map(function (f) {
      return '• ' + f.proveedor + (f.total ? ' — ' + euros_(f.total) : '') + ': ' + f.que;
    }).join('\n'));
  });
  return {
    cuantos: filas.length,
    asunto: 'Facturas para revisar ' + periodo + (filas.length ? ' (' + filas.length + ')' : ''),
    html: '<div style="font-family:Arial,Helvetica,sans-serif;max-width:640px;color:#222">' +
      '<h1 style="font-size:20px;margin:0 0 4px">Facturas para revisar</h1>' +
      '<p style="margin:0;color:#888;font-size:13px">Semana ' + periodo + '. Solo lo que necesita que lo mires (lo pendiente, como está hoy).</p>' +
      (html.length ? html.join('') : '<p style="margin-top:24px">No hay nada que revisar esta semana. 👍</p>') +
      '<p style="margin-top:32px;color:#aaa;font-size:11px">Llega los lunes a las 8:00 si hay algo.</p></div>',
    texto: texto.join('\n\n') || 'No hay nada que revisar esta semana.'
  };
}

// Cómo está ALBARANES hoy: repetidas y sin pagar (este año y, hasta marzo,
// también el anterior), y los albaranes escaneados sin su factura (de los
// últimos 3 meses).
function estadoAlbaranes_(hoy) {
  var filas = [];
  var dia = function (iso) { return iso ? Utilities.formatDate(fechaComoDate_(iso), 'Europe/Madrid', 'dd/MM') : '?'; };
  var hace = function (iso) {
    var d = Math.floor((hoy - fechaComoDate_(iso)) / 86400000);
    return d >= 2 ? ' · hace ' + d + ' días' : '';
  };
  var anios = [hoy.getFullYear()];
  if (hoy.getMonth() < 3) anios.unshift(hoy.getFullYear() - 1);
  anios.forEach(function (anio) {
    var r = revisarAlbaranes_(anio);
    if (!r.ok) return;
    r.repetidos.forEach(function (g) {
      filas.push({
        tipo: 'repetido', fecha: fechaComoDate_(g[0].fecha || anio + '-01-01'), proveedor: g[0].proveedor, total: g[0].importe,
        que: 'Está ' + g.length + ' veces: ' + g.map(function (x) {
          return (x.caja ? 'caja del ' + dia(x.caja) : 'a mano') + (x.forma ? ' (' + x.forma.toLowerCase() + ')' : '');
        }).join(', ')
      });
    });
    r.sinPagar.forEach(function (x) {
      filas.push({ tipo: 'sinpagar', fecha: fechaComoDate_(x.fecha || anio + '-01-01'), proveedor: x.proveedor, total: x.importe, que: 'Del ' + dia(x.fecha) + hace(x.fecha) });
    });
  });
  albaranesSinFactura_(2).forEach(function (a) {
    var iso = /^\d{4}-\d{2}-\d{2}$/.test(a.fecha) ? a.fecha : '';
    filas.push({ tipo: 'albaran', fecha: iso ? fechaComoDate_(iso) : hoy, proveedor: a.proveedor, total: '', que: 'Albarán del ' + dia(iso) + (iso ? hace(iso) : ''), enlace: a.enlace });
  });
  return filas;
}
