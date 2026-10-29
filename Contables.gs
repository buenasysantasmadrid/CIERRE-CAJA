// ============================================================================
// PAQUETE PARA LOS CONTABLES — el día 5 de cada mes a las 8:00 se prepara
// lo del mes anterior en "Facturas proveedores / Contables / AAAA-MM":
//   · Un ZIP con los PDF del mes ordenados por proveedor (lo de
//     "Facturas proveedores / AAAA / MM / <proveedor>"), sin las subcarpetas
//     "Albaranes" (albaranes y tickets que ya tienen su factura: estarían
//     repetidos). Si pesa demasiado para un solo archivo, va en varias partes.
//   · Un Excel con cada factura del mes según ALBARANES (proveedor, fecha,
//     nº, base, IVA, total, forma de pago) y si tiene su PDF; y en otras
//     dos pestañas, lo pendiente a hoy: albaranes sin su factura y lo que
//     falta pagar (ver RevisionAlbaranes.gs).
// Llega un email a la cuenta que publica la app con el ZIP y el Excel
// adjuntos (si pesan demasiado, solo los enlaces), para revisarlo y
// reenviarlo; arriba, lo que hay que mirar (facturas sin PDF, albaranes sin
// factura, falta pagar).
// Volver a prepararlo reemplaza lo anterior de ese mes.
// ============================================================================

var CONTABLES_CARPETA_ = 'Contables';
var CONTABLES_ZIP_MAX_ = 40 * 1024 * 1024; // por parte (Apps Script no maneja archivos de más de 50 MB)
// Pestañas de ALBARANES que no son facturas de proveedores.
var CONTABLES_PESTANAS_FUERA_ = ['EXTRAS'];
var CONTABLES_ADJUNTOS_MAX_ = 20 * 1024 * 1024; // Gmail no deja más de 25 MB por email

function asegurarActivadorPaqueteContables_() {
  var cache = CacheService.getScriptCache();
  if (cache.get('activador_contables')) return;
  var hay = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'prepararPaqueteContables'; });
  if (!hay) ScriptApp.newTrigger('prepararPaqueteContables').timeBased().onMonthDay(5).atHour(8).nearMinute(0).inTimezone('Europe/Madrid').create();
  cache.put('activador_contables', '1', 21600);
}

// Lo que corre el día 5: el mes anterior.
function prepararPaqueteContables() {
  var hoy = new Date();
  var mes = new Date(hoy.getFullYear(), hoy.getMonth() - 1, 1, 12);
  return paqueteContables_(Utilities.formatDate(mes, 'Europe/Madrid', 'yyyy-MM'), true);
}

// periodo: "AAAA-MM". conEmail: manda el email (Claude puede prepararlo sin él).
function paqueteContables_(periodo, conEmail) {
  if (!/^\d{4}-\d{2}$/.test(String(periodo))) return { ok: false, error: 'Mes no válido' };
  var anio = periodo.slice(0, 4), mm = periodo.slice(5, 7);
  var raiz = DriveApp.getFolderById(CARPETA_FACTURAS_ID_);
  var salida = carpetaHija_(carpetaHija_(raiz, CONTABLES_CARPETA_), periodo);
  var viejos = salida.getFiles();
  while (viejos.hasNext()) viejos.next().setTrashed(true);

  var archivos = pdfsDelMes_(raiz, anio, mm);
  var facturas = facturasConPdf_(anio, mm, archivos);
  var pend = pendientesContables_();

  var nombreMes = MESES_MAYUS_[Number(mm) - 1].charAt(0) + MESES_MAYUS_[Number(mm) - 1].slice(1).toLowerCase();
  var titulo = 'Facturas ' + nombreMes + ' ' + anio;
  var excel = excelFacturas_(salida, titulo, facturas, [
    { nombre: 'Albaranes sin factura', columnas: ['Proveedor', 'Fecha', 'Archivo', 'Enlace'], fechas: [2], euros: [],
      filas: pend.albaranes.map(function (a) { return [a.proveedor, /^\d{4}-\d{2}-\d{2}$/.test(a.fecha) ? fechaComoDate_(a.fecha) : '', a.nombre, a.enlace]; }) },
    { nombre: 'Falta pagar', columnas: ['Proveedor', 'Fecha', 'Nº factura', 'Importe', 'Días'], fechas: [2], euros: [4],
      filas: pend.sinPagar.map(function (x) { return [x.proveedor, x.fecha ? fechaComoDate_(x.fecha) : '', x.numero, x.importe, x.dias]; }) }
  ]);

  // El ZIP (en partes si hace falta), con el Excel dentro de la primera.
  var partes = [[]], peso = 0;
  partes[0].push(excel.blob.setName(titulo + '.xlsx'));
  archivos.forEach(function (a) {
    var tam = a.archivo.getSize();
    if (peso + tam > CONTABLES_ZIP_MAX_ && partes[partes.length - 1].length) { partes.push([]); peso = 0; }
    partes[partes.length - 1].push(a.archivo.getBlob().setName(sinCaracteresRaros_(a.carpeta) + '/' + a.archivo.getName()));
    peso += tam;
  });
  var zips = partes.map(function (blobs, i) {
    var nombre = titulo + (partes.length > 1 ? ' (parte ' + (i + 1) + ' de ' + partes.length + ')' : '') + '.zip';
    return salida.createFile(Utilities.zip(blobs, nombre));
  });

  var sinPdf = facturas.filter(function (f) { return f.pdf === 'NO'; });
  var res = {
    ok: true, periodo: periodo, carpeta: salida.getUrl(), pdfs: archivos.length, facturas: facturas.length,
    sinPdf: sinPdf.map(function (f) { return f.proveedor + ' ' + euros_(f.total) + ' (' + fechaCorta_(f.fecha) + ')'; }),
    zips: zips.map(function (z) { return z.getName(); })
  };
  res.albaranesSinFactura = pend.albaranes.length;
  res.sinPagar = pend.sinPagar.length;
  if (conEmail) mandarEmailContables_(titulo, res, sinPdf, facturas, zips, excel.archivo, pend);
  return res;
}

// Los PDF del mes ("Facturas proveedores / AAAA / MM / <proveedor>"), sin
// las subcarpetas Albaranes.
function pdfsDelMes_(raiz, anio, mm) {
  var archivos = [];
  var itAnio = raiz.getFoldersByName(anio);
  var itMes = itAnio.hasNext() ? itAnio.next().getFoldersByName(mm) : null;
  if (itMes && itMes.hasNext()) {
    var provs = itMes.next().getFolders();
    while (provs.hasNext()) {
      var prov = provs.next();
      var fs = prov.getFiles();
      while (fs.hasNext()) archivos.push({ carpeta: prov.getName(), archivo: fs.next() });
    }
  }
  archivos.sort(function (a, b) { return (a.carpeta + a.archivo.getName()).localeCompare(b.carpeta + b.archivo.getName()); });
  return archivos;
}

// Las facturas del mes según ALBARANES, y si tienen PDF (pdf: Sí / NO / ?
// si no tienen número para buscarlo).
function facturasConPdf_(anio, mm, archivos) {
  var claves = archivos.map(function (a) { return claveNumeroFactura_(a.archivo.getName().replace(/\.[a-z0-9]{2,4}$/i, '').replace(/ \(\d+\)$/, '')); });
  return facturasDelMesAlbaranes_(anio, Number(mm) - 1).map(function (f) {
    var k = claveNumeroFactura_(f.numero);
    f.pdf = !k ? '?' : (claves.some(function (c) { return c.slice(-k.length) === k; }) ? 'Sí' : 'NO');
    return f;
  });
}

// Lo pendiente a hoy: albaranes escaneados sin su factura (últimos 3
// meses) y lo que falta pagar (este año y, hasta marzo, el anterior).
function pendientesContables_() {
  var hoy = new Date();
  var r = { albaranes: [], sinPagar: [] };
  try { r.albaranes = albaranesSinFactura_(3); } catch (err) { Logger.log('Albaranes sin factura: ' + err); }
  var anios = [hoy.getFullYear()];
  if (hoy.getMonth() < 3) anios.unshift(hoy.getFullYear() - 1);
  anios.forEach(function (a) {
    var rev = revisarAlbaranes_(a);
    (rev.sinPagar || []).forEach(function (x) {
      x.dias = x.fecha ? Math.floor((hoy - fechaComoDate_(x.fecha)) / 86400000) : '';
      r.sinPagar.push(x);
    });
  });
  return r;
}

// Cada factura del mes en ALBARANES: una fila por gasto de cada pestaña de
// proveedor (en VARIOS, SUPER y SERVICIOS, el nombre real).
function facturasDelMesAlbaranes_(anio, mesIndex) {
  var id = buscarEnIndice_('ALBARANES', anio);
  if (!id) return [];
  var dia = function (d) { return d ? Utilities.formatDate(d, 'Europe/Madrid', 'yyyy-MM-dd') : ''; };
  var lista = [];
  SpreadsheetApp.openById(id).getSheets().forEach(function (h) {
    var nombre = h.getName();
    if (ALBARANES_PESTANAS_EXCLUIDAS_.indexOf(nombre) > -1 || nombre === INFORME_MES_PESTANA_ || CONTABLES_PESTANAS_FUERA_.indexOf(normalizarClave_(nombre)) > -1) return;
    var b = filasDelMesAlbaranes_(h, mesIndex);
    if (!b) return;
    b.filas.forEach(function (f) {
      if (!f.importe) return;
      var iva = Math.round(f.iva * 100) / 100;
      lista.push({
        proveedor: f.detalle ? nombre + ' — ' + f.detalle : nombre, fecha: dia(f.fecha), numero: f.numero,
        base: Math.round((f.importe - iva) * 100) / 100, iva: iva, total: f.importe, forma: f.forma
      });
    });
  });
  lista.sort(function (a, b) { return a.proveedor.localeCompare(b.proveedor) || (a.fecha < b.fecha ? -1 : a.fecha > b.fecha ? 1 : 0); });
  return lista;
}

// Se arma como hoja de Google y se guarda en .xlsx (la hoja se borra).
// extras: más pestañas, [{nombre, columnas, filas, fechas: [nº de columna], euros: [nº de columna]}].
function excelFacturas_(carpeta, titulo, facturas, extras) {
  var ss = SpreadsheetApp.create(titulo);
  var hoja = ss.getSheets()[0];
  hoja.setName('Facturas');
  var filas = [['Proveedor', 'Fecha', 'Nº factura', 'Base', 'IVA', 'Total', 'Forma de pago', 'PDF']].concat(facturas.map(function (f) {
    return [f.proveedor, f.fecha ? fechaComoDate_(f.fecha) : '', f.numero, f.base, f.iva, f.total, f.forma, f.pdf];
  }));
  var n = filas.length;
  filas.push(['TOTAL', '', '', '=SUM(D2:D' + n + ')', '=SUM(E2:E' + n + ')', '=SUM(F2:F' + n + ')', '', '']);
  hoja.getRange(1, 1, filas.length, 8).setValues(filas);
  hoja.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#e8e8e8');
  hoja.getRange(filas.length, 1, 1, 8).setFontWeight('bold');
  hoja.getRange(2, 2, Math.max(n - 1, 1), 1).setNumberFormat('dd/MM/yyyy');
  hoja.getRange(2, 4, filas.length - 1, 3).setNumberFormat('#,##0.00 €');
  hoja.setFrozenRows(1);
  hoja.autoResizeColumns(1, 8);
  (extras || []).forEach(function (x) {
    var h = ss.insertSheet(x.nombre);
    var datos = [x.columnas].concat(x.filas.length ? x.filas : [x.columnas.map(function (c, i) { return i ? '' : 'Nada pendiente'; })]);
    h.getRange(1, 1, datos.length, x.columnas.length).setValues(datos);
    h.getRange(1, 1, 1, x.columnas.length).setFontWeight('bold').setBackground('#e8e8e8');
    (x.fechas || []).forEach(function (c) { h.getRange(2, c, datos.length - 1, 1).setNumberFormat('dd/MM/yyyy'); });
    (x.euros || []).forEach(function (c) { h.getRange(2, c, datos.length - 1, 1).setNumberFormat('#,##0.00 €'); });
    h.setFrozenRows(1);
    h.autoResizeColumns(1, x.columnas.length);
  });
  SpreadsheetApp.flush();
  var resp = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx', {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
  });
  var blob = resp.getBlob().setName(titulo + '.xlsx');
  var archivo = carpeta.createFile(blob.copyBlob());
  DriveApp.getFileById(ss.getId()).setTrashed(true);
  return { blob: blob, archivo: archivo };
}

function mandarEmailContables_(titulo, res, sinPdf, facturas, zips, excel, pend) {
  var esc = function (s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); };
  var total = facturas.reduce(function (s, f) { return s + f.total; }, 0);
  var revisar = sinPdf.length
    ? '<h2 style="margin:24px 0 4px;font-size:17px;color:#c0392b">Facturas sin PDF <span style="font-weight:normal;color:#888">(' + sinPdf.length + ')</span></h2>' +
      '<p style="margin:0 0 10px;color:#666;font-size:13px">Están en ALBARANES pero su PDF no está en Drive. Conviene buscarlas antes de mandarlo.</p>' +
      '<table style="border-collapse:collapse;width:100%;font-size:14px;border-left:4px solid #c0392b">' + sinPdf.map(function (f, i) {
        return '<tr style="background:' + (i % 2 ? '#fff' : '#f7f7f7') + '"><td style="padding:8px 10px;font-weight:bold">' + esc(f.proveedor) + '</td>' +
          '<td style="padding:8px 10px;text-align:right;white-space:nowrap">' + euros_(f.total) + '</td>' +
          '<td style="padding:8px 10px;color:#999;font-size:12px">' + fechaCorta_(f.fecha) + '</td></tr>';
      }).join('') + '</table>'
    : '<p style="margin-top:20px;color:#1e8449">✔ Todas las facturas de ALBARANES tienen su PDF.</p>';
  var tabla = function (color, titulo2, explica, filas) {
    return '<h2 style="margin:28px 0 4px;font-size:17px;color:' + color + '">' + titulo2 + '</h2>' +
      '<p style="margin:0 0 10px;color:#666;font-size:13px">' + explica + '</p>' +
      '<table style="border-collapse:collapse;width:100%;font-size:14px;border-left:4px solid ' + color + '">' + filas.map(function (f, i) {
        return '<tr style="background:' + (i % 2 ? '#fff' : '#f7f7f7') + '"><td style="padding:8px 10px;font-weight:bold">' + esc(f[0]) + '</td>' +
          '<td style="padding:8px 10px;text-align:right;white-space:nowrap">' + f[1] + '</td>' +
          '<td style="padding:8px 10px;color:#999;font-size:12px">' + f[2] + '</td></tr>';
      }).join('') + '</table>';
  };
  if (pend && pend.albaranes.length) {
    revisar += tabla('#16a085', 'Albaranes sin factura <span style="font-weight:normal;color:#888">(' + pend.albaranes.length + ')</span>',
      'Escaneados como albarán; su factura todavía no llegó (también en el Excel).',
      pend.albaranes.map(function (a) { return [a.proveedor, '<a href="' + a.enlace + '" style="color:#2471a3">PDF</a>', /^\d{4}-\d{2}-\d{2}$/.test(a.fecha) ? fechaCorta_(a.fecha) : '']; }));
  }
  if (pend && pend.sinPagar.length) {
    var debe = pend.sinPagar.reduce(function (t, x) { return t + x.importe; }, 0);
    revisar += tabla('#2471a3', 'Falta pagar <span style="font-weight:normal;color:#888">(' + pend.sinPagar.length + ' · ' + euros_(debe) + ')</span>',
      'Cargadas como "No pagado" y sin día de pago (también en el Excel).',
      pend.sinPagar.map(function (x) { return [x.proveedor, euros_(x.importe), (x.fecha ? fechaCorta_(x.fecha) : '') + (x.dias >= 2 ? ' · hace ' + x.dias + ' días' : '')]; }));
  }
  // Adjuntos: el ZIP y el Excel, si caben en un email.
  var peso = zips.reduce(function (t, z) { return t + z.getSize(); }, excel.getSize());
  var adjuntar = peso <= CONTABLES_ADJUNTOS_MAX_;
  var enlaces = zips.map(function (z) { return '<a href="' + z.getUrl() + '" style="color:#2471a3">' + esc(z.getName()) + '</a>'; }).join('<br>') +
    '<br><a href="' + excel.getUrl() + '" style="color:#2471a3">' + esc(excel.getName()) + '</a>';
  MailApp.sendEmail({
    to: Session.getEffectiveUser().getEmail(),
    subject: titulo + ': paquete para los contables listo',
    attachments: adjuntar ? zips.map(function (z) { return z.getBlob(); }).concat([excel.getBlob()]) : [],
    body: titulo + ': ' + res.pdfs + ' PDF y ' + res.facturas + ' facturas en la lista. Carpeta: ' + res.carpeta +
      (sinPdf.length ? '\nSin PDF: ' + res.sinPdf.join('; ') : ''),
    htmlBody: '<div style="font-family:Arial,Helvetica,sans-serif;max-width:640px;color:#222">' +
      '<h1 style="font-size:20px;margin:0 0 4px">' + esc(titulo) + '</h1>' +
      '<p style="margin:0;color:#888;font-size:13px">Paquete para los contables: ' + res.pdfs + ' PDF · ' + res.facturas + ' facturas · ' + euros_(total) + '</p>' +
      revisar +
      '<h2 style="margin:28px 0 8px;font-size:17px">Para reenviar</h2>' +
      '<p style="margin:0 0 6px;font-size:13px;color:#666">' + (adjuntar ? 'Van adjuntos en este email: el ZIP con los PDF y el Excel. También en Drive:' : 'Pesan demasiado para adjuntarlos (' + Math.round(peso / 1048576) + ' MB): aquí están en Drive.') + '</p>' +
      '<p style="margin:0;font-size:14px;line-height:1.8">' + enlaces + '</p>' +
      '<p style="margin:12px 0 0;font-size:13px"><a href="' + res.carpeta + '" style="color:#2471a3">Abrir la carpeta</a></p>' +
      '<p style="margin-top:32px;color:#aaa;font-size:11px">Se prepara el día 5 de cada mes con lo del mes anterior.</p></div>'
  });
}
