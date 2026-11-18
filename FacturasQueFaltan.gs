// ============================================================================
// FACTURAS QUE FALTAN — el día 3 de cada mes a las 9:00 llega un email con
// las facturas habituales del mes anterior que no están cargadas.
//   · "Habitual": un proveedor (o un servicio de la pestaña SERVICIOS, por
//     su nombre) que tuvo algo cargado en ALBARANES en al menos 2 de los 3
//     meses anteriores a ese mes. Si ese mes no tiene nada, falta.
//   · Pollería, La Jaula, Mercandalucía y Las 3B (factura mensual): falta si
//     hay albaranes del mes pero su factura del mes no está en Drive.
//   · Atlanta, Disbesa, Iberolactal y CBG (primero el albarán, después la
//     factura): falta la factura de cada albarán del mes que no pasó a factura.
//   · Varios, Super, Extras y Glovo no cuentan: no son fijos.
// El activador se pone solo desde los guardados de la caja (Bandejas.gs).
// ============================================================================

// Los que traen albarán y después la factura (nombres de pestaña de ALBARANES).
var FALTAN_ALBARAN_LUEGO_FACTURA_ = ['ATLANTA', 'DISBESA', 'IBEROLACTAL', 'CBG'];

var FALTAN_NOMBRES_MES_ = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];

function asegurarActivadorFacturasQueFaltan_() {
  var hay = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'mandarFacturasQueFaltan'; });
  if (!hay) ScriptApp.newTrigger('mandarFacturasQueFaltan').timeBased().onMonthDay(3).atHour(9).nearMinute(0).inTimezone('Europe/Madrid').create();
}

// Lo del día 3: el mes anterior.
function mandarFacturasQueFaltan() {
  var hoy = new Date();
  var anio = hoy.getFullYear(), mes = hoy.getMonth() - 1;
  if (mes < 0) { mes = 11; anio--; }
  return mandarFacturasQueFaltanDe_(anio, mes, '');
}

// muestra: texto delante del asunto (Claude: GET ?muestraFaltanFacturas=AAAA-MM).
function mandarFacturasQueFaltanDe_(anio, mes, muestra) {
  var faltan = facturasQueFaltan_(anio, mes);
  var nombreMes = FALTAN_NOMBRES_MES_[mes] + ' ' + anio;
  var esc = function (s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); };
  var asunto = (muestra ? muestra + ' — ' : '') + (faltan.length ? '¿Y estas facturas de ' + nombreMes + '? Faltan ' + faltan.length : 'Facturas de ' + nombreMes + ': no falta ninguna');
  var html = '<div style="font-family:Arial,Helvetica,sans-serif;max-width:640px;color:#222">' +
    '<h1 style="font-size:20px;margin:0 0 4px">Facturas de ' + esc(nombreMes) + '</h1>' +
    (faltan.length
      ? '<p style="color:#555;font-size:14px">Estas llegan casi todos los meses y en ' + esc(nombreMes) + ' no están cargadas. ¿Ya las tienes?</p>' +
        '<table style="border-collapse:collapse;width:100%;font-size:14px;border-left:4px solid #d35400">' +
        faltan.map(function (f, i) {
          return '<tr style="background:' + (i % 2 ? '#fff' : '#f7f7f7') + '"><td style="padding:8px 10px;font-weight:bold">' + esc(f.proveedor) + '</td>' +
            '<td style="padding:8px 10px">' + esc(f.que) + '</td></tr>';
        }).join('') + '</table>' +
        '<p style="color:#555;font-size:14px;margin-top:16px">Si la tienes, guárdala en la carpeta <b>📥 Facturas nuevas</b> y se carga sola. ' +
        'Si alguna ya no va a llegar (por ejemplo, un servicio que se dio de baja), díselo a Claude.</p>'
      : '<p style="color:#1e8449;font-size:14px">✔ Están todas las facturas habituales.</p>') +
    '<p style="margin-top:32px;color:#aaa;font-size:11px">Llega el día 3 de cada mes a las 9:00.</p></div>';
  var texto = faltan.length
    ? 'Facturas de ' + nombreMes + ' que faltan:\n' + faltan.map(function (f) { return '• ' + f.proveedor + ': ' + f.que; }).join('\n') +
      '\n\nSi la tienes, guárdala en la carpeta "📥 Facturas nuevas".'
    : 'Están todas las facturas habituales de ' + nombreMes + '.';
  MailApp.sendEmail({ to: Session.getEffectiveUser().getEmail(), subject: asunto, htmlBody: html, body: texto });
  return { ok: true, mes: nombreMes, faltan: faltan };
}

// [{ proveedor, que }]
function facturasQueFaltan_(anio, mes) {
  var cache = {};
  var planilla = function (a) {
    if (!(a in cache)) { var id = buscarEnIndice_('ALBARANES', String(a)); cache[a] = id ? SpreadsheetApp.openById(id) : null; }
    return cache[a];
  };
  // Qué hay cargado un mes: { clave: { nombre, total, n } }. SERVICIOS, por servicio.
  var delMes = function (a, m) {
    var r = {};
    var ss = planilla(a);
    if (!ss) return null;
    ss.getSheets().forEach(function (h) {
      var nombre = h.getName(), clave = normalizarClave_(nombre);
      if (ALBARANES_PESTANAS_EXCLUIDAS_.indexOf(nombre) > -1 || nombre === INFORME_MES_PESTANA_) return;
      if (BANDEJA_PROVEEDORES_GENERICOS_.indexOf(clave) > -1) return;
      var b;
      try { b = filasDelMesAlbaranes_(h, m); } catch (err) { return; }
      if (!b) return;
      b.filas.forEach(function (f) {
        if (!f.importe) return;
        var k = clave, n = nombre;
        if (clave === normalizarClave_(ALBARANES_PESTANA_SERVICIOS_)) {
          if (!f.detalle) return;
          k = clave + ' / ' + normalizarClave_(f.detalle); n = 'Servicios — ' + f.detalle;
        }
        var x = r[k] = r[k] || { nombre: n, total: 0, n: 0 };
        x.total += f.importe; x.n++;
      });
    });
    return r;
  };
  var antes = function (k) { var a = anio, m = mes - k; while (m < 0) { m += 12; a--; } return [a, m]; };

  var actual = delMes(anio, mes) || {};
  var cuenta = {}, ultimo = {};
  [1, 2, 3].forEach(function (k) {
    var am = antes(k);
    var d = delMes(am[0], am[1]);
    if (!d) return;
    Object.keys(d).forEach(function (c) {
      cuenta[c] = (cuenta[c] || 0) + 1;
      if (!ultimo[c]) ultimo[c] = { nombre: d[c].nombre, total: d[c].total, mes: FALTAN_NOMBRES_MES_[am[1]] };
    });
  });

  var faltan = [];
  Object.keys(cuenta).forEach(function (c) {
    if (cuenta[c] < 2 || actual[c]) return;
    var u = ultimo[c];
    faltan.push({ proveedor: u.nombre, que: 'No hay nada cargado. La última, en ' + u.mes + ': ' + euros_(Math.round(u.total * 100) / 100) });
  });

  // Los de factura mensual: albaranes del mes sin la factura del mes en Drive.
  var pdfs = facturasMensualesEnDrive_(anio, mes);
  INFORME_MES_FACTURA_MENSUAL_.forEach(function (clave) {
    var x = actual[clave];
    if (!x || (pdfs[clave] || []).length) return;
    faltan.push({ proveedor: x.nombre, que: 'Falta la factura del mes (' + x.n + ' albaranes, ' + euros_(Math.round(x.total * 100) / 100) + ')' });
  });

  // Atlanta, Disbesa, Iberolactal y CBG: primero traen el albarán y después
  // la factura. Falta la de cada albarán del mes que no pasó a factura (al
  // pasar, en observaciones queda "Ticket <nº del albarán>", ver
  // FacturaTicket.gs). Los "SIN FACTURA" no cuentan.
  var ssMes = planilla(anio);
  if (ssMes) ssMes.getSheets().forEach(function (h) {
    if (FALTAN_ALBARAN_LUEGO_FACTURA_.indexOf(normalizarClave_(h.getName())) === -1) return;
    var b;
    try { b = filasDelMesAlbaranes_(h, mes); } catch (err) { return; }
    if (!b) return;
    var sin = b.filas.filter(function (f) {
      var notas = normalizarClave_(f.obs);
      return f.importe && !/(^|·\s*)TICKET\b/.test(notas) && !/SIN FACTURA/.test(notas);
    });
    if (!sin.length) return;
    var total = Math.round(sin.reduce(function (s, f) { return s + f.importe; }, 0) * 100) / 100;
    faltan.push({
      proveedor: h.getName(),
      que: (sin.length === 1 ? 'Falta la factura de 1 albarán: ' : 'Faltan las facturas de ' + sin.length + ' albaranes (' + euros_(total) + '): ') +
        sin.map(function (f) {
          return (f.numero ? 'nº ' + f.numero + ' ' : '') + 'del ' + (f.fecha ? Utilities.formatDate(f.fecha, 'Europe/Madrid', 'dd/MM') : '?') + ' (' + euros_(f.importe) + ')';
        }).join(', ')
    });
  });
  return faltan.sort(function (a, b) { return a.proveedor.localeCompare(b.proveedor); });
}
