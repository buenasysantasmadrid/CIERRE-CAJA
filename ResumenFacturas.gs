// ============================================================================
// RESUMEN SEMANAL DE FACTURAS — los lunes a las 8:00 llega un email (a la
// cuenta que publica la app) solo con lo que hay que mirar, no con lo que
// salió bien:
//   · Fallos: facturas de las carpetas de Drive (Bandejas.gs) que no se
//     pudieron cargar o cuyo albarán no apareció.
//   · Diferencias: la factura dice otro importe que el albarán.
//   · Sin proveedor: facturas puestas en las carpetas de Drive que no son de
//     ningún proveedor de la lista (lo cargado en Varios en la caja, no).
//   · Y cómo está todo hoy (RevisionAlbaranes.gs): errores de la semana,
//     facturas repetidas en ALBARANES, lo que falta pagar y los albaranes
//     que todavía no tienen su factura.
//   · La semana en números (ver extrasDeLaSemana_): ventas, diferencias de
//     caja, proveedores que gastaron más que de costumbre, subidas de
//     precio y facturas del mes sin PDF.
// De cada uno, solo lo necesario (nombre, importe y qué pasa). Los fallos,
// diferencias y "sin proveedor" se apuntan en la pestaña "Movimientos
// facturas" del Índice; lo demás se calcula al armar el email.
// ============================================================================

var MOV_FACTURAS_PESTANA_ = 'Movimientos facturas';
var MOV_FACTURAS_ENCABEZADOS_ = ['Fecha', 'Tipo', 'Proveedor', 'Total', 'Qué pasa', 'Enlace'];
// Secciones del email: tipo, título, explicación, color.
var MOV_FACTURAS_SECCIONES_ = [
  ['ventas', 'Ventas de la semana', 'Lo vendido (local, delivery y web), comparado.', '#1e8449'],
  ['fallo', 'Fallos', 'No se pudieron hacer solas: hay que mirarlas.', '#c0392b'],
  ['error', 'Errores al guardar', 'Lo que falló esta semana (también llegó en los avisos de las 9 y las 18).', '#c0392b'],
  ['repetido', 'Repetidas en ALBARANES', 'La misma factura está dos veces: una sobra.', '#8e44ad'],
  ['diferencia', 'Diferencias de importe', 'La factura no dice lo mismo que el albarán.', '#d35400'],
  ['sinpagar', 'Falta pagar', 'Cargadas como "No pagado" y sin día de pago.', '#2471a3'],
  ['albaran', 'Albaranes sin factura', 'Escaneados como albarán; su factura todavía no llegó.', '#16a085'],
  ['varios', 'Sin proveedor', 'Facturas de las carpetas que no son de ningún proveedor de la lista: ¿hay que pasarlas a uno?', '#b7950b'],
  ['caja', 'Diferencias de caja', 'Días que la caja no cuadró (más de 5 €).', '#c0392b'],
  ['gasto', 'Proveedores por encima de lo normal', 'Gastado esta semana comparado con su media de las 8 semanas anteriores.', '#d35400'],
  ['precio', 'Subidas de precio', 'Productos más caros que la última vez (de las facturas escaneadas).', '#8e44ad'],
  ['sinpdf', 'Facturas del mes sin PDF', 'Están en ALBARANES pero no en Drive: faltarán en el paquete de los contables del día 5.', '#7f8c8d']
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
  MailApp.sendEmail({ to: Session.getEffectiveUser().getEmail(), subject: r.asunto, htmlBody: r.html, body: r.texto });
  props.setProperty('resumen_facturas_desde', ahora.toISOString());
}

// Una muestra (Claude: GET ?muestraResumenFacturas=1): lo de los últimos 7
// días, sin cambiar desde cuándo cuenta el resumen de verdad.
function mandarMuestraResumenFacturas_() {
  var r = resumenFacturas_(new Date(Date.now() - 7 * 86400000), new Date());
  MailApp.sendEmail({ to: Session.getEffectiveUser().getEmail(), subject: 'MUESTRA — ' + r.asunto, htmlBody: r.html, body: r.texto });
  return { ok: true, cuantos: r.cuantos };
}

function resumenFacturas_(desde, hasta) {
  var hoja = SpreadsheetApp.openById(INDICE_SHEET_ID_).getSheetByName(MOV_FACTURAS_PESTANA_);
  var filas = [];
  if (hoja && hoja.getLastRow() > 1 && hoja.getRange(1, 3).getValue() === 'Proveedor') {
    filas = hoja.getRange(2, 1, hoja.getLastRow() - 1, 6).getValues()
      .filter(function (f) { return f[0] instanceof Date && f[0] >= desde && f[0] <= hasta; })
      .map(function (f) { return { fecha: f[0], tipo: f[1], proveedor: f[2], total: f[3], que: f[4], enlace: f[5] }; });
  }
  try { extrasDeLaSemana_(hasta).forEach(function (f) { filas.push(f); }); } catch (err) { Logger.log('Resumen de facturas, la semana: ' + err); }
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
    if (s[0] !== 'ventas' && s[0] !== 'gasto' && s[0] !== 'precio') deEsta.sort(function (a, b) { return a.fecha - b.fecha; });
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
          '<td style="padding:8px 10px;color:#999;font-size:12px;white-space:nowrap">' + (['ventas', 'gasto', 'precio', 'caja', 'error'].indexOf(f.tipo) > -1 ? '' : dia(f.fecha)) +
          (f.enlace ? ' · <a href="' + esc(f.enlace) + '" style="color:#2471a3">PDF</a>' : '') + '</td></tr>';
      }).join('') + '</table>');
    texto.push(s[1].toUpperCase() + ' (' + cuenta + ')\n' + deEsta.map(function (f) {
      return '• ' + f.proveedor + (f.total ? ' — ' + euros_(f.total) : '') + ': ' + f.que;
    }).join('\n'));
  });
  return {
    cuantos: filas.filter(function (f) { return f.tipo !== 'ventas'; }).length,
    asunto: 'Resumen de la semana ' + periodo,
    html: '<div style="font-family:Arial,Helvetica,sans-serif;max-width:640px;color:#222">' +
      '<h1 style="font-size:20px;margin:0 0 4px">Resumen de la semana</h1>' +
      '<p style="margin:0;color:#888;font-size:13px">Semana ' + periodo + '. Solo lo que necesita que lo mires (lo pendiente, como está hoy).</p>' +
      (html.length ? html.join('') : '<p style="margin-top:24px">No hay nada que revisar esta semana. 👍</p>') +
      '<p style="margin-top:32px;color:#aaa;font-size:11px">Llega los lunes a las 8:00.</p></div>',
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

// La semana (los 7 días hasta ayer) en números.
function extrasDeLaSemana_(hoy) {
  var filas = [];
  var iso = function (d) { return Utilities.formatDate(d, 'Europe/Madrid', 'yyyy-MM-dd'); };
  var ayer = new Date(hoy.getTime() - 86400000);
  var hasta = iso(ayer), desde = iso(new Date(ayer.getTime() - 6 * 86400000));
  var pct = function (a, b) { return b ? (a >= b ? '+' : '') + Math.round((a - b) / b * 100) + ' %' : '—'; };

  // Ventas y diferencias de caja: de Contabilidad (este año y el anterior).
  try {
    var dias = {};
    [ayer.getFullYear() - 1, ayer.getFullYear()].forEach(function (anio) {
      var id = buscarEnIndice_('CONTABILIDAD', String(anio));
      if (!id) return;
      var ss = SpreadsheetApp.openById(id);
      var mesHasta = ayer.getMonth(), mesDesde = Math.max(0, mesHasta - 1);
      for (var m = mesDesde; m <= mesHasta; m++) {
        var d = datosMesContabilidad_(ss, m);
        if (d) d.lista.forEach(function (x) { dias[x.fecha] = x; });
      }
    });
    var ventas = function (ini, fin) {
      var t = 0;
      Object.keys(dias).forEach(function (f) { if (f >= ini && f <= fin) t += (dias[f].sistema || 0) + (dias[f].delivery || 0); });
      return Math.round(t * 100) / 100;
    };
    var menos = function (f, n) { return iso(new Date(fechaComoDate_(f).getTime() - n * 86400000)); };
    var esta = ventas(desde, hasta), anterior = ventas(menos(desde, 7), menos(hasta, 7)), anioPasado = ventas(menos(desde, 364), menos(hasta, 364));
    filas.push({ tipo: 'ventas', fecha: ayer, proveedor: 'Esta semana', total: esta, que: fechaCorta_(desde).slice(0, 5) + ' al ' + fechaCorta_(hasta).slice(0, 5) });
    if (anterior) filas.push({ tipo: 'ventas', fecha: ayer, proveedor: 'Semana anterior', total: anterior, que: 'Esta semana ' + pct(esta, anterior) });
    if (anioPasado) filas.push({ tipo: 'ventas', fecha: ayer, proveedor: 'Hace un año', total: anioPasado, que: 'Esta semana ' + pct(esta, anioPasado) + ' (los mismos días de la semana)' });
    var nombresDia = ['Dom', 'Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb'];
    Object.keys(dias).sort().forEach(function (f) {
      var dif = dias[f].diferencia || 0;
      if (f < desde || f > hasta || Math.abs(dif) <= 5) return;
      filas.push({ tipo: 'caja', fecha: fechaComoDate_(f), proveedor: nombresDia[fechaComoDate_(f).getDay()] + ' ' + fechaCorta_(f).slice(0, 5), total: '', que: 'Diferencia de ' + (dif > 0 ? '+' : '') + euros_(dif) });
    });
  } catch (err) { Logger.log('Resumen, ventas: ' + err); }

  // Gasto por proveedor y precios: lo mismo que la página de informes
  // (InformeSemanalPedidos.gs).
  try {
    var inf = armarInformeSemanal_(hasta).datos;
    inf.proveedores.proveedores.filter(function (p) { return p.cambio !== null && p.cambio >= 30 && p.diferencia >= 50; }).slice(0, 10).forEach(function (p) {
      filas.push({ tipo: 'gasto', fecha: ayer, proveedor: p.proveedor, total: p.estaSemana, que: 'Lo normal: ' + euros_(p.media) + ' (+' + p.cambio + ' %)' });
    });
    inf.cambiosDePrecio.cambios.filter(function (c) { return c.cambio >= 3; }).slice(0, 15).forEach(function (c) {
      filas.push({ tipo: 'precio', fecha: ayer, proveedor: c.producto + ' (' + c.proveedor + ')', total: c.ahora, que: 'Antes ' + euros_(c.antes) + (c.unidad ? ' / ' + c.unidad : '') + ' · +' + String(c.cambio).replace('.', ',') + ' %' });
    });
  } catch (err) { Logger.log('Resumen, compras: ' + err); }

  // Facturas del mes (hasta hace 3 días, para dar tiempo a escanearlas) sin PDF.
  try {
    var anio = hasta.slice(0, 4), mm = hasta.slice(5, 7);
    var limite = iso(new Date(hoy.getTime() - 3 * 86400000));
    facturasConPdf_(anio, mm, pdfsDelMes_(DriveApp.getFolderById(CARPETA_FACTURAS_ID_), anio, mm)).forEach(function (f) {
      if (f.pdf !== 'NO' || f.fecha > limite) return;
      filas.push({ tipo: 'sinpdf', fecha: fechaComoDate_(f.fecha), proveedor: f.proveedor, total: f.total, que: 'Del ' + fechaCorta_(f.fecha).slice(0, 5) });
    });
  } catch (err) { Logger.log('Resumen, sin PDF: ' + err); }
  return filas;
}
