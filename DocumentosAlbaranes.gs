// ============================================================================
// ALBARANES: columna DOCUMENTO y pestaña "FALTA FACTURA" (24/11/2026, Luciana)
// ============================================================================
// En cada bloque de mes de cada pestaña de proveedor, la columna J
// (DOCUMENTO) dice qué papel hay de cada fila:
//   ALBARÁN  · un albarán que todavía no pasó a factura.
//   FACTURA  · la factura (escaneada, del email, o un albarán que ya pasó a
//              factura: "Ticket <nº>" en observaciones).
//   TICKET   · escaneado como ticket.
//   (vacío)  · no hay nada: "SIN FACTURA", o sin número y sin papel.
// Cómo se decide, en este orden:
//   1. "SIN FACTURA" en observaciones → vacío.
//   2. "Ticket <nº>" en observaciones (ya pasó a factura) → FACTURA.
//   3. Pollería, La Jaula, Mercandalucía, Las 3B (factura mensual) → ALBARÁN.
//   4. Cargada desde "Facturas nuevas" (ID "pdf-…") → FACTURA.
//   5. El PDF en Drive con ese número: escaneado como albarán → ALBARÁN,
//      como ticket → TICKET, como factura o del email → FACTURA.
//   6. "ALBARÁN" en observaciones, o Atlanta/Disbesa/Iberolactal/CBG → ALBARÁN.
//   7. Con número → FACTURA; sin número → vacío.
// La pestaña "FALTA FACTURA" (al principio de la planilla) lista los
// ALBARÁN que esperan su factura y, de los de factura mensual, los meses
// ya terminados sin la factura del mes en Drive.
// Se rehace cada 2 horas (activador actualizarDocumentosAlbaranes). Al
// reordenar un bloque (guardarBloqueOrdenado_), J se mueve con su fila.
// Claude: GET ?documentosAlbaranes=AAAA (&activador=1 para poner el activador).
// ============================================================================

var ALBARANES_COL_DOC_ = 10; // J
var ALBARANES_ENC_DOC_ = 'DOCUMENTO';
var FALTA_FACTURA_PESTANA_ = 'FALTA FACTURA';
ALBARANES_PESTANAS_EXCLUIDAS_.push(FALTA_FACTURA_PESTANA_);

function actualizarDocumentosAlbaranes() {
  var hoy = new Date();
  var r = documentosAlbaranes_(hoy.getFullYear());
  if (hoy.getMonth() === 0) { try { documentosAlbaranes_(hoy.getFullYear() - 1); } catch (e) { Logger.log(e); } }
  Logger.log(JSON.stringify(r));
  return r;
}

function asegurarActivadorDocumentos_() {
  var hay = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'actualizarDocumentosAlbaranes'; });
  if (!hay) ScriptApp.newTrigger('actualizarDocumentosAlbaranes').timeBased().everyHours(2).create();
  return !hay;
}

// Los PDF del año en Drive, con la clave de su número (nombre sin la fecha,
// solo letras y números: "POLLERIA94359").
function indiceDocsDrive_(anio) {
  var lista = [];
  var itA = carpetaFacturas_().getFoldersByName(String(anio));
  if (!itA.hasNext()) return lista;
  var meses = itA.next().getFolders();
  while (meses.hasNext()) {
    var mes = meses.next();
    if (!/^\d{2}$/.test(mes.getName())) continue;
    var provs = mes.getFolders();
    while (provs.hasNext()) {
      var prov = provs.next();
      if (/^[📥⚠🗑]/.test(prov.getName())) continue;
      (function recorrer(carpeta, enAlbaranes, nivel) {
        var fs = carpeta.getFiles();
        while (fs.hasNext()) {
          var f = fs.next(), nombre = f.getName().replace(/\.[a-z0-9]{2,4}$/i, '');
          var m = /^(\d{4}-\d{2}-\d{2})\s*(.*)$/.exec(nombre);
          lista.push({ k: claveNumeroFactura_(m ? m[2] : nombre), fecha: m ? m[1] : '', mes: Number(mes.getName()) - 1, carpeta: normalizarClave_(prov.getName()),
            desc: String(f.getDescription() || '').split('\n')[0].toLowerCase(), url: f.getUrl(), enAlbaranes: enAlbaranes });
        }
        if (nivel > 1) return;
        var subs = carpeta.getFolders();
        while (subs.hasNext()) { var s = subs.next(); recorrer(s, enAlbaranes || s.getName() === 'Albaranes', nivel + 1); }
      })(prov, false, 0);
    }
  }
  return lista;
}

// El PDF de una fila: el número tiene que ser el final del nombre (y no
// pegado a otra cifra). Mejor el de la carpeta del proveedor que el de
// Albaranes, y el de la fecha más cercana.
function docDeFila_(docs, numero, fecha) {
  var k = claveNumeroFactura_(numero);
  if (k.length < 3) return null;
  var t = fecha instanceof Date ? fecha.getTime() : null;
  var cands = docs.filter(function (d) {
    if (!d.k.endsWith(k)) return false;
    var antes = d.k.charAt(d.k.length - k.length - 1);
    if (/\d/.test(k.charAt(0)) && /\d/.test(antes)) return false;
    if (k.length < 4 && t != null && d.fecha !== Utilities.formatDate(fecha, 'Europe/Madrid', 'yyyy-MM-dd')) return false;
    return true;
  });
  if (!cands.length) return null;
  var dist = function (d) { return t == null || !d.fecha ? 0 : Math.abs(fechaComoDate_(d.fecha).getTime() - t); };
  cands.sort(function (a, b) { return (a.enAlbaranes - b.enAlbaranes) || (dist(a) - dist(b)); });
  return cands[0];
}

function tipoDocumentoFila_(f, pestanaClave, docs, conDrive) {
  var notas = normalizarClave_(f.obs + ' ' + f.info);
  if (/SIN FACTURA/.test(notas)) return { tipo: '' };
  if (/(^|·\s*)TICKET\s/.test(normalizarClave_(f.obs))) return { tipo: 'FACTURA' };
  if (INFORME_MES_FACTURA_MENSUAL_.indexOf(pestanaClave) > -1) return { tipo: 'ALBARÁN', mensual: true };
  if (f.id.indexOf('pdf-') === 0) return { tipo: 'FACTURA' };
  var d = docDeFila_(docs, f.numero, f.fecha);
  if (d) return { tipo: d.desc === 'albaran' ? 'ALBARÁN' : d.desc === 'ticket' ? 'TICKET' : 'FACTURA', url: d.url };
  if (/ALBARAN/.test(notas)) return { tipo: 'ALBARÁN' };
  // Antes de que hubiera PDF en Drive no se puede saber: con número, factura.
  if (conDrive && FALTAN_ALBARAN_LUEGO_FACTURA_.indexOf(pestanaClave) > -1) return { tipo: 'ALBARÁN' };
  return { tipo: f.numero ? 'FACTURA' : '' };
}

function documentosAlbaranes_(anio) {
  var id = buscarEnIndice_('ALBARANES', String(anio));
  if (!id) return { ok: false, error: 'No hay ALBARANES ' + anio + ' en el Índice' };
  var ss = SpreadsheetApp.openById(id);
  var docs = indiceDocsDrive_(anio);
  var hoy = new Date(), mesHoy = hoy.getFullYear() === Number(anio) ? hoy.getMonth() : 12;
  // Primer mes con PDF en Drive: antes, lo de Drive no dice nada.
  var mesInicio = docs.length ? Math.min.apply(null, docs.map(function (d) { return d.mes; })) : 12;
  var porMes = {};
  var faltan = [], mensuales = [], saltadas = [], cuenta = { 'ALBARÁN': 0, FACTURA: 0, TICKET: 0, '': 0 };

  ss.getSheets().forEach(function (h) {
    var nombre = h.getName(), clave = normalizarClave_(nombre);
    if (ALBARANES_PESTANAS_EXCLUIDAS_.indexOf(nombre) > -1) return;
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      var ultima = h.getLastRow();
      if (ultima < 2) return;
      var todo = h.getRange(1, 1, ultima, ALBARANES_COL_ID_).getValues();
      var encs = [];
      todo.forEach(function (v, i) { if (normalizarClave_(v[0]) === 'FECHA') encs.push(i); });
      if (!encs.length) return;
      // J libre en todos los encabezados (o ya es DOCUMENTO).
      if (encs.some(function (e) { var j = normalizarClave_(todo[e][ALBARANES_COL_DOC_ - 1]); return j && j !== ALBARANES_ENC_DOC_; })) {
        saltadas.push(nombre + ' (la columna J ya tiene otra cosa)'); return;
      }
      var colJ = todo.map(function (v) { return [v[ALBARANES_COL_DOC_ - 1]]; });
      encs.forEach(function (e, mes) {
        colJ[e][0] = ALBARANES_ENC_DOC_;
        var mapa = mapaColumnasAlbaranes_(todo[e].slice(0, ALBARANES_ANCHO_));
        var albaranesMes = [];
        for (var i = e + 1; i <= e + ALBARANES_FILAS_DATOS_ && i < todo.length; i++) {
          var v = todo[i];
          if (v.slice(0, ALBARANES_ANCHO_).every(function (x) { return x === '' || x == null; })) { colJ[i][0] = ''; continue; }
          var f = {
            fecha: mapa.fecha != null && v[mapa.fecha] instanceof Date ? v[mapa.fecha] : null,
            numero: mapa.factura != null ? String(v[mapa.factura] || '').trim() : '',
            importe: mapa.importe != null ? Number(v[mapa.importe]) || 0 : 0,
            forma: mapa.formaPago != null ? String(v[mapa.formaPago] || '') : '',
            detalle: mapa.detalle != null ? String(v[mapa.detalle] || '') : '',
            obs: mapa.obs != null ? String(v[mapa.obs] || '') : '',
            info: mapa.info != null ? String(v[mapa.info] || '') : '',
            id: String(v[ALBARANES_COL_ID_ - 1] || '')
          };
          var r = tipoDocumentoFila_(f, clave, docs, mes >= mesInicio);
          colJ[i][0] = r.tipo;
          cuenta[r.tipo]++;
          if (r.tipo !== 'ALBARÁN') continue;
          if (r.mensual) { albaranesMes.push(f); continue; }
          porMes[mes + 1] = (porMes[mes + 1] || 0) + 1;
          faltan.push([nombre + (f.detalle ? ' — ' + f.detalle : ''), f.fecha || '', f.numero, f.importe, f.forma,
            f.fecha ? Math.floor((hoy - f.fecha) / 86400000) : '', r.url || (docDeFila_(docs, f.numero, f.fecha) || {}).url || '']);
        }
        // Factura mensual: el mes ya terminado sin la factura del mes en Drive.
        if (albaranesMes.length && mes < mesHoy && mes >= mesInicio) {
          var mm = ('0' + (mes + 1)).slice(-2);
          var hay = docs.some(function (d) { return d.carpeta === clave && d.desc === 'factura' && d.fecha.slice(0, 7) === anio + '-' + mm; });
          if (!hay) {
            var total = Math.round(albaranesMes.reduce(function (s, x) { return s + x.importe; }, 0) * 100) / 100;
            mensuales.push([nombre, FALTAN_NOMBRES_MES_[mes], albaranesMes.length, total]);
          }
        }
      });
      h.getRange(1, ALBARANES_COL_DOC_, ultima, 1).setValues(colJ);
      encs.forEach(function (e) { h.getRange(e + 1, ALBARANES_ANCHO_).copyTo(h.getRange(e + 1, ALBARANES_COL_DOC_), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false); });
    } finally {
      lock.releaseLock();
    }
  });

  escribirFaltaFactura_(ss, faltan, mensuales);
  var pendientes = '';
  if (Number(anio) === new Date().getFullYear()) { try { pendientes = escribirPendientes_(ss, anio, faltan, mensuales); } catch (err) { pendientes = 'error: ' + err; } }
  return { ok: true, anio: anio, pendientes: pendientes, cuenta: cuenta, faltan: faltan.length, mensualesSinFactura: mensuales.length, saltadas: saltadas, pdfsEnDrive: docs.length, desdeMes: mesInicio + 1, faltanPorMes: porMes, mensualesLista: mensuales.map(function (m) { return m[0] + ' ' + m[1]; }) };
}

function escribirFaltaFactura_(ss, faltan, mensuales) {
  var h = ss.getSheetByName(FALTA_FACTURA_PESTANA_) || ss.insertSheet(FALTA_FACTURA_PESTANA_, 0);
  h.clear();
  faltan.sort(function (a, b) { return (a[1] instanceof Date ? a[1].getTime() : 0) - (b[1] instanceof Date ? b[1].getTime() : 0); });
  var verde = '#2f4f3a';
  h.getRange(1, 1).setValue('Albaranes que esperan su factura').setFontWeight('bold').setFontSize(14);
  h.getRange(2, 1).setValue('Actualizado el ' + Utilities.formatDate(new Date(), 'Europe/Madrid', 'dd/MM/yyyy HH:mm') +
    ' (se rehace solo cada 2 horas). Cuando llega la factura, sale sola de la lista. No escribir aquí.').setFontColor('#666');
  var fila = 4;
  var cab = ['Proveedor', 'Fecha albarán', 'Nº albarán', 'Importe', 'Forma de pago', 'Días esperando', 'PDF'];
  h.getRange(fila, 1, 1, cab.length).setValues([cab]).setFontWeight('bold').setBackground(verde).setFontColor('#fff');
  if (faltan.length) {
    var filas = faltan.map(function (f) { return f.slice(0, 6).concat([f[6] ? '=HYPERLINK("' + f[6] + '";"Ver")' : '']); });
    h.getRange(fila + 1, 1, filas.length, 6).setValues(filas.map(function (f) { return f.slice(0, 6); }));
    h.getRange(fila + 1, 7, filas.length, 1).setFormulas(filas.map(function (f) { return [f[6]]; }));
    h.getRange(fila + 1, 2, filas.length, 1).setNumberFormat('dd/MM/yyyy');
    h.getRange(fila + 1, 4, filas.length, 1).setNumberFormat('#,##0.00 €');
    h.getRange(fila + 1, 6, filas.length, 1).setBackgrounds(filas.map(function (f) { return [f[5] !== '' && f[5] > 30 ? '#fde8e4' : null]; }));
    var total = Math.round(faltan.reduce(function (s, f) { return s + f[3]; }, 0) * 100) / 100;
    fila += filas.length + 1;
    h.getRange(fila, 1, 1, 4).setValues([['TOTAL (' + faltan.length + ' albaranes)', '', '', total]]).setFontWeight('bold').setBackground('#f1efe6');
    h.getRange(fila, 4).setNumberFormat('#,##0.00 €');
  } else {
    fila++;
    h.getRange(fila, 1).setValue('No hay ninguno: todos los albaranes tienen su factura.');
  }
  fila += 3;
  h.getRange(fila, 1).setValue('Factura mensual que falta (Pollería, La Jaula, Mercandalucía, Las 3B: meses ya terminados)').setFontWeight('bold').setFontSize(12);
  fila++;
  h.getRange(fila, 1, 1, 4).setValues([['Proveedor', 'Mes', 'Albaranes', 'Total albaranes']]).setFontWeight('bold').setBackground(verde).setFontColor('#fff');
  if (mensuales.length) {
    h.getRange(fila + 1, 1, mensuales.length, 4).setValues(mensuales);
    h.getRange(fila + 1, 4, mensuales.length, 1).setNumberFormat('#,##0.00 €');
  } else h.getRange(fila + 1, 1).setValue('Ninguna: están todas en Drive.');
  h.setFrozenRows(4);
  h.setColumnWidth(1, 240); h.setColumnWidth(3, 140); h.setColumnWidth(5, 170);
}
