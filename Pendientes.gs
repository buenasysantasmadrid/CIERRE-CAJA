// ============================================================================
// ALBARANES: pestaña "PENDIENTES" (24/11/2026, Luciana)
// ============================================================================
// Todo lo que hay que mirar de facturas y albaranes, en un solo sitio y como
// está AHORA (lo mismo que llega por email). Se rehace cada 2 horas, junto
// con la columna DOCUMENTO (actualizarDocumentosAlbaranes). Lo que se
// soluciona sale solo:
//   · Facturas en "⚠️ Revisar" / "⚠️ No se pudo leer" (Drive): salen cuando
//     el archivo ya no está ahí.
//   · Albaranes sin factura y factura del mes que falta: salen al llegar la
//     factura.
//   · Repetidas y falta pagar (revisarAlbaranes_): salen al arreglar la fila.
//   · Facturas del mes sin PDF: salen al escanearla.
//   · PDF que se ven mal: salen cuando el PDF se cambia (va a la papelera).
// Lo que no se puede comprobar solo (diferencias de importe, sin proveedor,
// otros fallos de los últimos 30 días, errores de la última semana) lleva una
// casilla "Ya está": al marcarla, la fila desaparece y no vuelve
// (alEditarAlbaranes → alEditarPendientes_).
// ============================================================================

var PENDIENTES_PESTANA_ = 'PENDIENTES';
ALBARANES_PESTANAS_EXCLUIDAS_.push(PENDIENTES_PESTANA_);
var PENDIENTES_HECHOS_ = 'pendientes_hechos';
var PENDIENTES_CAB_ = ['Qué pasa', 'Proveedor', 'Fecha', 'Nº', 'Importe', 'Detalle', 'PDF', 'Ya está', 'clave'];
var PENDIENTES_COL_HECHO_ = 8, PENDIENTES_COL_CLAVE_ = 9;

// Secciones, en este orden: [código, título, color].
var PENDIENTES_SECCIONES_ = [
  ['bandeja', 'Facturas del email sin cargar', '#c0392b'],
  ['fallo', 'Otros fallos al cargar facturas', '#c0392b'],
  ['diferencia', 'Diferencias de importe', '#d35400'],
  ['repetido', 'Repetidas en ALBARANES (una sobra)', '#8e44ad'],
  ['varios', 'Sin proveedor (¿a qué proveedor va?)', '#b7950b'],
  ['mensual', 'Falta la factura del mes', '#16a085'],
  ['albaran', 'Albaranes que esperan su factura', '#16a085'],
  ['sinpdf', 'Facturas sin PDF en Drive', '#7f8c8d'],
  ['mal', 'PDF que se ven mal (volver a escanear)', '#7f8c8d'],
  ['sinpagar', 'Falta pagar', '#2471a3'],
  ['error', 'Errores de la app (última semana)', '#c0392b']
];

function leerHechosPendientes_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(PENDIENTES_HECHOS_) || '{}'); } catch (e) { return {}; }
}

function idDeUrlDrive_(url) {
  var m = /\/d\/([\w-]{20,})/.exec(String(url || '')) || /[?&]id=([\w-]{20,})/.exec(String(url || ''));
  return m ? m[1] : '';
}

// faltan / mensuales: los de documentosAlbaranes_ (FALTA FACTURA).
function escribirPendientes_(ss, anio, faltan, mensuales) {
  var hoy = new Date();
  var hechos = leerHechosPendientes_();
  var items = [];
  var poner = function (sec, prov, fecha, num, importe, que, url, clave) {
    if (clave && hechos[clave]) return;
    items.push({ sec: sec, prov: prov || '', fecha: fecha instanceof Date ? fecha : (/^\d{4}-\d{2}-\d{2}$/.test(String(fecha)) ? fechaComoDate_(fecha) : ''),
      num: num || '', importe: Number(importe) || '', que: que || '', url: url || '', clave: clave || '' });
  };
  var log = function (q, err) { Logger.log('PENDIENTES, ' + q + ': ' + err); };

  // Lo apuntado en "Movimientos facturas" (Índice) los últimos 30 días.
  var mov = [];
  try {
    var hm = SpreadsheetApp.openById(INDICE_SHEET_ID_).getSheetByName(MOV_FACTURAS_PESTANA_);
    var desde = new Date(hoy.getTime() - 30 * 86400000);
    if (hm && hm.getLastRow() > 1) mov = hm.getRange(2, 1, hm.getLastRow() - 1, 6).getValues().filter(function (f) { return f[0] instanceof Date && f[0] >= desde; });
  } catch (err) { log('movimientos', err); }
  var queDeEnlace = {};
  mov.forEach(function (f) { if (f[5] && f[1] === 'fallo') queDeEnlace[idDeUrlDrive_(f[5])] = f[4]; });

  // 1. Facturas del email en "⚠️ Revisar" y "⚠️ No se pudo leer".
  try {
    var itB = carpetaFacturas_().getFoldersByName(BANDEJA_UNICA_);
    if (itB.hasNext()) {
      var bandeja = itB.next();
      [BANDEJA_REVISAR_, BANDEJA_NO_LEIDAS_].forEach(function (sub) {
        var itS = bandeja.getFoldersByName(sub);
        if (!itS.hasNext()) return;
        var fs = itS.next().getFiles();
        while (fs.hasNext()) {
          var f = fs.next(), esp = leerEsperando_(f);
          if (esp && esp.r) {
            var r = esp.r;
            poner('bandeja', r.proveedor_lista || r.proveedor, r.fecha, r.numero_factura, r.importe_total,
              'No encontré su albarán ni su foto en la caja. La sigo probando cada hora; si en un día no aparece, la cargo sola (NO PAGADO, o pagada si fue con tarjeta).', f.getUrl());
          } else {
            poner('bandeja', f.getName(), f.getDateCreated(), '', '',
              (queDeEnlace[f.getId()] || (sub === BANDEJA_NO_LEIDAS_ ? 'No se pudo leer.' : 'Hay que mirarla.')) + ' Está en "' + sub + '": cargarla a mano o volver a ponerla en "' + BANDEJA_UNICA_ + '".', f.getUrl());
          }
        }
      });
    }
  } catch (err) { log('bandejas', err); }

  // 2. Fallos, diferencias y sin proveedor (con casilla "Ya está"). Los de
  // las carpetas Revisar / No se pudo leer ya salen arriba.
  mov.forEach(function (f) {
    var tipo = f[1], que = String(f[4] || '');
    if (['fallo', 'diferencia', 'varios'].indexOf(tipo) === -1) return;
    if (tipo === 'varios' && que.indexOf('¿') === -1) return; // solo las que preguntan algo
    if (tipo === 'fallo' && (que.indexOf(BANDEJA_REVISAR_) > -1 || que.indexOf(BANDEJA_NO_LEIDAS_) > -1 || que.indexOf('No sé si es de albarán') > -1)) return;
    var clave = tipo + '|' + Utilities.formatDate(f[0], 'Europe/Madrid', 'yyyyMMddHHmmss') + '|' + String(f[2]).slice(0, 40);
    poner(tipo, f[2], f[0], '', f[3], que, f[5], clave);
  });

  // 3. Albaranes sin factura y factura del mes que falta.
  (mensuales || []).forEach(function (m) { poner('mensual', m[0], '', '', m[3], 'Factura de ' + m[1] + ' (' + m[2] + ' albaranes).'); });
  (faltan || []).forEach(function (f) {
    poner('albaran', f[0], f[1], f[2], f[3], (f[4] ? f[4].toLowerCase() + ' · ' : '') + (f[5] !== '' ? 'esperando hace ' + f[5] + ' días' : ''), f[6]);
  });

  // 4. Repetidas y falta pagar.
  try {
    var rev = revisarAlbaranes_(anio);
    (rev.repetidos || []).forEach(function (g) {
      poner('repetido', g[0].proveedor, g[0].fecha, g[0].numero, g[0].importe,
        'Está ' + g.length + ' veces: ' + g.map(function (x) { return 'pestaña ' + x.pestana + ' fila ' + x.fila + (x.caja ? ' (caja del ' + fechaCorta_(x.caja).slice(0, 5) + ')' : ' (a mano)'); }).join(', '));
    });
    (rev.sinPagar || []).forEach(function (x) {
      poner('sinpagar', x.proveedor, x.fecha, x.numero, x.importe, 'NO PAGADO (pestaña ' + x.pestana + ', fila ' + x.fila + ')');
    });
  } catch (err) { log('repetidas / sin pagar', err); }

  // 5. Facturas sin PDF (este mes y el anterior, hasta hace 3 días).
  try {
    var limite = Utilities.formatDate(new Date(hoy.getTime() - 3 * 86400000), 'Europe/Madrid', 'yyyy-MM-dd');
    [1, 0].forEach(function (atras) {
      var d = new Date(hoy.getFullYear(), hoy.getMonth() - atras, 1);
      if (d.getFullYear() !== Number(anio)) return;
      var mm = ('0' + (d.getMonth() + 1)).slice(-2);
      facturasConPdf_(String(anio), mm, pdfsDelMes_(DriveApp.getFolderById(CARPETA_FACTURAS_ID_), String(anio), mm)).forEach(function (f) {
        if (f.pdf !== 'NO' || f.fecha > limite) return;
        poner('sinpdf', f.proveedor, f.fecha, f.numero, f.total, 'Está en ALBARANES pero no en Drive: escanearla (los contables la necesitan).');
      });
    });
  } catch (err) { log('sin PDF', err); }

  // 6. PDF que se ven mal (revisarCalidadPdfs_): mientras sigan en Drive.
  try {
    var props = PropertiesService.getScriptProperties().getProperties();
    Object.keys(props).filter(function (k) { return k.indexOf('calidadPdfs_') === 0; }).forEach(function (k) {
      var malos = [];
      try { malos = JSON.parse(props[k]); } catch (e) { return; }
      malos.forEach(function (m) {
        var id = idDeUrlDrive_(m.url);
        try { if (!id || DriveApp.getFileById(id).isTrashed()) return; } catch (e) { return; }
        var partes = String(m.nombre).split(' / ');
        poner('mal', partes.length > 1 ? partes[0] : '', (/\d{4}-\d{2}-\d{2}/.exec(m.nombre) || [''])[0], '', '', m.problema, m.url, 'mal|' + id);
      });
    });
  } catch (err) { log('PDF que se ven mal', err); }

  // 7. Errores de la app de la última semana.
  try {
    erroresDeLaSemana_(new Date(hoy.getTime() - 7 * 86400000)).forEach(function (e) {
      poner('error', e.asunto, new Date(e.ultima), '', '', (e.veces > 1 ? e.veces + ' veces. ' : '') + String(e.detalle || '').split('\n')[0].slice(0, 200), '',
        'error|' + String(e.asunto).slice(0, 60) + '|' + String(e.ultima));
    });
  } catch (err) { log('errores', err); }

  pintarPendientes_(ss, items);
  try { instalarTriggerInformeMes_(ss.getId()); } catch (err) { log('activador', err); }
  var cuenta = {};
  items.forEach(function (x) { cuenta[x.sec] = (cuenta[x.sec] || 0) + 1; });
  return cuenta;
}

function pintarPendientes_(ss, items) {
  var h = ss.getSheetByName(PENDIENTES_PESTANA_) || ss.insertSheet(PENDIENTES_PESTANA_, 0);
  h.getRange(1, 1, h.getMaxRows(), h.getMaxColumns()).breakApart().clearDataValidations();
  h.clear();
  var ancho = PENDIENTES_CAB_.length;
  var cuenta = {};
  items.forEach(function (x) { cuenta[x.sec] = (cuenta[x.sec] || 0) + 1; });
  h.getRange(1, 1).setValue('Pendientes de facturas y albaranes').setFontWeight('bold').setFontSize(14);
  h.getRange(2, 1).setValue('Actualizado el ' + Utilities.formatDate(new Date(), 'Europe/Madrid', 'dd/MM/yyyy HH:mm') +
    ' (se rehace solo cada 2 horas). Lo que se soluciona sale solo. Lo que tiene casilla "Ya está": marcarla cuando esté arreglado y desaparece.').setFontColor('#666');
  var resumen = PENDIENTES_SECCIONES_.filter(function (s) { return cuenta[s[0]]; }).map(function (s) { return s[1].split(' (')[0] + ': ' + cuenta[s[0]]; });
  h.getRange(3, 1).setValue(resumen.length ? resumen.join(' · ') : 'No hay nada pendiente. 👍').setFontWeight('bold');

  var fila = 5;
  var valores = [], formulas = [], formatos = [];
  PENDIENTES_SECCIONES_.forEach(function (s) {
    var deEsta = items.filter(function (x) { return x.sec === s[0]; });
    if (!deEsta.length) return;
    deEsta.sort(function (a, b) { return (a.fecha ? a.fecha.getTime() : 0) - (b.fecha ? b.fecha.getTime() : 0); });
    var total = deEsta.reduce(function (t, x) { return t + (Number(x.importe) || 0); }, 0);
    h.getRange(fila, 1, 1, ancho - 1).merge().setValue(s[1] + '  (' + deEsta.length + (total ? ' · ' + euros_(Math.round(total * 100) / 100) : '') + ')')
      .setFontWeight('bold').setFontColor('#fff').setBackground(s[2]);
    fila++;
    h.getRange(fila, 1, 1, ancho).setValues([PENDIENTES_CAB_]).setFontWeight('bold').setBackground('#f1efe6');
    fila++;
    var filas = deEsta.map(function (x) { return [s[1].split(' (')[0], x.prov, x.fecha, x.num, x.importe, x.que, '', '', x.clave]; });
    h.getRange(fila, 1, filas.length, ancho).setValues(filas);
    h.getRange(fila, 7, filas.length, 1).setFormulas(deEsta.map(function (x) { return [x.url ? '=HYPERLINK("' + x.url + '";"Ver")' : '']; }));
    h.getRange(fila, 3, filas.length, 1).setNumberFormat('dd/MM/yyyy');
    h.getRange(fila, 5, filas.length, 1).setNumberFormat('#,##0.00 €');
    deEsta.forEach(function (x, i) {
      var c = h.getRange(fila + i, PENDIENTES_COL_HECHO_);
      if (x.clave) c.insertCheckboxes(); else c.setValue('sale sola').setFontColor('#aaa');
    });
    fila += filas.length + 1;
  });
  h.setFrozenRows(0);
  h.setColumnWidth(1, 170); h.setColumnWidth(2, 200); h.setColumnWidth(3, 90); h.setColumnWidth(4, 120);
  h.setColumnWidth(5, 90); h.setColumnWidth(6, 520); h.setColumnWidth(7, 50); h.setColumnWidth(8, 70);
  h.getRange(1, 6, h.getMaxRows(), 1).setWrap(true);
  h.hideColumns(PENDIENTES_COL_CLAVE_);
}

// Desde alEditarAlbaranes: casilla "Ya está" marcada → se guarda y la fila
// se borra.
function alEditarPendientes_(e) {
  var hoja = e.range.getSheet();
  if (e.range.getColumn() !== PENDIENTES_COL_HECHO_ || e.range.getNumRows() !== 1) return;
  if (hoja.getRange(e.range.getRow(), PENDIENTES_COL_HECHO_).getValue() !== true) return;
  var clave = String(hoja.getRange(e.range.getRow(), PENDIENTES_COL_CLAVE_).getValue() || '');
  if (!clave) return;
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var hechos = leerHechosPendientes_();
    hechos[clave] = new Date().toISOString().slice(0, 10);
    // Solo se guardan los de los últimos 60 días (lo más viejo ya no sale).
    var limite = new Date(Date.now() - 60 * 86400000).toISOString().slice(0, 10);
    Object.keys(hechos).forEach(function (k) { if (hechos[k] < limite && k.indexOf('mal|') !== 0) delete hechos[k]; });
    PropertiesService.getScriptProperties().setProperty(PENDIENTES_HECHOS_, JSON.stringify(hechos));
  } finally {
    lock.releaseLock();
  }
  hoja.deleteRow(e.range.getRow());
}
