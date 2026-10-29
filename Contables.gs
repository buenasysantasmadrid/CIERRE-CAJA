// ============================================================================
// PAQUETE PARA LOS CONTABLES — el día 5 de cada mes a las 8:00 se prepara
// lo del mes anterior en "Facturas proveedores / Contables / AAAA-MM":
//   · Un ZIP con los PDF del mes ordenados por proveedor (lo de
//     "Facturas proveedores / AAAA / MM / <proveedor>"), sin las subcarpetas
//     "Albaranes" (albaranes y tickets que ya tienen su factura: estarían
//     repetidos). Si pesa demasiado para un solo archivo, va en varias partes.
//   · Un Excel con cada factura del mes según ALBARANES (proveedor, fecha,
//     nº, base, IVA, total, forma de pago) y si tiene su PDF.
// Llega un email a la cuenta que publica la app con el enlace, para
// revisarlo y reenviarlo; arriba, lo que hay que mirar (facturas sin PDF).
// Volver a prepararlo reemplaza lo anterior de ese mes.
// ============================================================================

var CONTABLES_CARPETA_ = 'Contables';
var CONTABLES_ZIP_MAX_ = 40 * 1024 * 1024; // por parte (Apps Script no maneja archivos de más de 50 MB)
// Pestañas de ALBARANES que no son facturas de proveedores.
var CONTABLES_PESTANAS_FUERA_ = ['EXTRAS'];

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

  // Los PDF del mes, por proveedor.
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

  // Las facturas del mes según ALBARANES, y si tienen PDF.
  var claves = archivos.map(function (a) { return claveNumeroFactura_(a.archivo.getName().replace(/\.[a-z0-9]{2,4}$/i, '').replace(/ \(\d+\)$/, '')); });
  var facturas = facturasDelMesAlbaranes_(anio, Number(mm) - 1).map(function (f) {
    var k = claveNumeroFactura_(f.numero);
    f.pdf = !k ? '?' : (claves.some(function (c) { return c.slice(-k.length) === k; }) ? 'Sí' : 'NO');
    return f;
  });

  var nombreMes = MESES_MAYUS_[Number(mm) - 1].charAt(0) + MESES_MAYUS_[Number(mm) - 1].slice(1).toLowerCase();
  var titulo = 'Facturas ' + nombreMes + ' ' + anio;
  var excel = excelFacturas_(salida, titulo, facturas);

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
  if (conEmail) mandarEmailContables_(titulo, res, sinPdf, facturas, zips, excel.archivo);
  return res;
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
function excelFacturas_(carpeta, titulo, facturas) {
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
  SpreadsheetApp.flush();
  var resp = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx', {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
  });
  var blob = resp.getBlob().setName(titulo + '.xlsx');
  var archivo = carpeta.createFile(blob.copyBlob());
  DriveApp.getFileById(ss.getId()).setTrashed(true);
  return { blob: blob, archivo: archivo };
}

function mandarEmailContables_(titulo, res, sinPdf, facturas, zips, excel) {
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
  var enlaces = zips.map(function (z) { return '<a href="' + z.getUrl() + '" style="color:#2471a3">' + esc(z.getName()) + '</a>'; }).join('<br>') +
    '<br><a href="' + excel.getUrl() + '" style="color:#2471a3">' + esc(excel.getName()) + '</a>';
  MailApp.sendEmail({
    to: Session.getEffectiveUser().getEmail(),
    subject: titulo + ': paquete para los contables listo',
    body: titulo + ': ' + res.pdfs + ' PDF y ' + res.facturas + ' facturas en la lista. Carpeta: ' + res.carpeta +
      (sinPdf.length ? '\nSin PDF: ' + res.sinPdf.join('; ') : ''),
    htmlBody: '<div style="font-family:Arial,Helvetica,sans-serif;max-width:640px;color:#222">' +
      '<h1 style="font-size:20px;margin:0 0 4px">' + esc(titulo) + '</h1>' +
      '<p style="margin:0;color:#888;font-size:13px">Paquete para los contables: ' + res.pdfs + ' PDF · ' + res.facturas + ' facturas · ' + euros_(total) + '</p>' +
      revisar +
      '<h2 style="margin:28px 0 8px;font-size:17px">Para reenviar</h2><p style="margin:0;font-size:14px;line-height:1.8">' + enlaces + '</p>' +
      '<p style="margin:12px 0 0;font-size:13px"><a href="' + res.carpeta + '" style="color:#2471a3">Abrir la carpeta</a></p>' +
      '<p style="margin-top:32px;color:#aaa;font-size:11px">Se prepara el día 5 de cada mes con lo del mes anterior.</p></div>'
  });
}
