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
// Una sola tabla con filtro (24/11/2026, Luciana: poder mirar por meses).
var PENDIENTES_CAB_ = ['Mes', 'Qué pasa', 'Proveedor', 'Fecha', 'Nº', 'Importe', 'Detalle', 'PDF', 'Ya está', 'clave'];
var PENDIENTES_COL_HECHO_ = 9, PENDIENTES_COL_CLAVE_ = 10, PENDIENTES_FILA_CAB_ = 5;
var PENDIENTES_MESES_ = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
function textoMesPendientes_(anio, mesIndex) { return anio + '-' + ('0' + (mesIndex + 1)).slice(-2) + ' ' + PENDIENTES_MESES_[mesIndex]; }

// Secciones, en este orden: [código, título, color].
var PENDIENTES_SECCIONES_ = [
  ['bandeja', 'Facturas del email sin cargar', '#c0392b'],
  ['fallo', 'Otros fallos al cargar facturas', '#c0392b'],
  ['diferencia', 'Diferencias de importe', '#d35400'],
  ['repetido', 'Repetidas en ALBARANES (una sobra)', '#8e44ad'],
  ['varios', 'Sin proveedor (¿a qué proveedor va?)', '#b7950b'],
  ['mensual', 'Falta la factura del mes', '#16a085'],
  ['albaran', 'Albaranes que esperan su factura', '#16a085'],
  ['habitual', '¿Falta alguna factura? (suele llegar cada mes)', '#b03a2e'],
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

// Lo que corre cada 2 horas (activador aparte de actualizarDocumentosAlbaranes:
// juntas pasan de los 6 minutos). Los albaranes sin factura los lee de la
// pestaña FALTA FACTURA, que hace documentosAlbaranes_.
function actualizarPendientes() {
  var anio = new Date().getFullYear();
  var id = buscarEnIndice_('ALBARANES', String(anio));
  if (!id) return { ok: false, error: 'No hay ALBARANES ' + anio };
  var ss = SpreadsheetApp.openById(id);
  var ff = leerFaltaFactura_(ss);
  return { ok: true, pendientes: escribirPendientes_(ss, anio, ff.faltan, ff.mensuales) };
}

// Las dos tablas de la pestaña FALTA FACTURA (escribirFaltaFactura_), en el
// mismo formato que documentosAlbaranes_.
function leerFaltaFactura_(ss) {
  var r = { faltan: [], mensuales: [] };
  var h = ss.getSheetByName(FALTA_FACTURA_PESTANA_);
  if (!h || h.getLastRow() < 5) return r;
  var rango = h.getRange(1, 1, h.getLastRow(), 7);
  var v = rango.getValues(), fo = rango.getFormulas();
  var parte = '';
  for (var i = 0; i < v.length; i++) {
    var a = String(v[i][0] || '');
    if (a === 'Proveedor' && v[i][1] === 'Fecha albarán') { parte = 'faltan'; continue; }
    if (a === 'Proveedor' && v[i][1] === 'Mes') { parte = 'mensual'; continue; }
    if (!a || /^TOTAL \(/.test(a) || /^(No hay ninguno|Ninguna)/.test(a)) { if (/^TOTAL/.test(a)) parte = ''; continue; }
    if (parte === 'faltan') {
      var url = (/HYPERLINK\("([^"]+)"/i.exec(fo[i][6] || '') || [])[1] || '';
      r.faltan.push([a, v[i][1], v[i][2], v[i][3], v[i][4], v[i][5], url]);
    } else if (parte === 'mensual') {
      r.mensuales.push([a, v[i][1], v[i][2], v[i][3]]);
    }
  }
  return r;
}

// faltan / mensuales: los de documentosAlbaranes_ (FALTA FACTURA).
function escribirPendientes_(ss, anio, faltan, mensuales) {
  var hoy = new Date();
  var hechos = leerHechosPendientes_();
  var items = [];
  var poner = function (sec, prov, fecha, num, importe, que, url, clave, mes) {
    if (clave && hechos[clave]) return;
    var f = fecha instanceof Date ? fecha : (/^\d{4}-\d{2}-\d{2}$/.test(String(fecha)) ? fechaComoDate_(fecha) : '');
    items.push({ sec: sec, prov: prov || '', fecha: f, mes: mes || (f ? textoMesPendientes_(f.getFullYear(), f.getMonth()) : ''),
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
  (mensuales || []).forEach(function (m) { poner('mensual', m[0], '', '', m[3], 'Factura de ' + m[1] + ' (' + m[2] + ' albaranes).', '', '', PENDIENTES_MESES_.indexOf(m[1]) > -1 ? textoMesPendientes_(Number(anio), PENDIENTES_MESES_.indexOf(m[1])) : ''); });
  (faltan || []).forEach(function (f) {
    poner('albaran', f[0], f[1], f[2], f[3], (f[4] ? f[4].toLowerCase() + ' · ' : '') + (f[5] !== '' ? 'esperando hace ' + f[5] + ' días' : ''), f[6]);
  });

  // 3b. Facturas que suelen llegar cada mes y el mes pasado no llegó
  // (FacturasQueFaltan.gs; desde el día 5, para darles tiempo).
  try {
    if (hoy.getDate() >= 5) {
      var mp = new Date(hoy.getFullYear(), hoy.getMonth() - 1, 1);
      facturasQueFaltan_(mp.getFullYear(), mp.getMonth()).forEach(function (x) {
        if (!/^No hay nada cargado/.test(x.que)) return; // lo demás ya sale en otras secciones
        poner('habitual', x.proveedor, '', '', '', 'Suele llegar cada mes y en ' + PENDIENTES_MESES_[mp.getMonth()] + ' no hay nada. ' + x.que.replace('No hay nada cargado. ', ''), '',
          'habitual|' + mp.getFullYear() + '-' + mp.getMonth() + '|' + normalizarClave_(x.proveedor), textoMesPendientes_(mp.getFullYear(), mp.getMonth()));
      });
    }
  } catch (err) { log('habituales', err); }

  // 4. Repetidas y falta pagar.
  try {
    var rev = revisarAlbaranes_(anio);
    (rev.repetidos || []).forEach(function (g) {
      poner('repetido', g[0].proveedor, g[0].fecha, g[0].numero, g[0].importe,
        'Está ' + g.length + ' veces: ' + g.map(function (x) { return 'pestaña ' + x.pestana + ' fila ' + x.fila + (x.caja ? ' (caja del ' + fechaCorta_(x.caja).slice(0, 5) + ')' : ' (a mano)'); }).join(', '));
    });
    // Los de factura a fin de mes (o cada 3 meses: Iberolactal, Pollería...)
    // se pagan todos juntos con la factura: una sola línea por mes
    // (25/11/2026, Luciana: "es mucha info para una cosa que no es un problema").
    var juntos = {};
    (rev.sinPagar || []).forEach(function (x) {
      var f = x.fecha instanceof Date ? x.fecha : (/^\d{4}-\d{2}-\d{2}$/.test(String(x.fecha)) ? fechaComoDate_(x.fecha) : null);
      if (f && INFORME_MES_FACTURA_MENSUAL_.indexOf(normalizarClave_(x.pestana)) > -1) {
        var k = x.pestana + '|' + f.getFullYear() + '-' + f.getMonth();
        var j = juntos[k] = juntos[k] || { prov: x.proveedor, pestana: x.pestana, fecha: f, n: 0, total: 0 };
        j.n++; j.total += Number(x.importe) || 0;
        return;
      }
      poner('sinpagar', x.proveedor, x.fecha, x.numero, x.importe, 'NO PAGADO (pestaña ' + x.pestana + ', fila ' + x.fila + ')');
    });
    Object.keys(juntos).forEach(function (k) {
      var j = juntos[k];
      poner('sinpagar', j.prov, new Date(j.fecha.getFullYear(), j.fecha.getMonth(), 1), '', Math.round(j.total * 100) / 100,
        'Esperando la factura del ' + (esTrimestral_(normalizarClave_(j.pestana)) ? 'trimestre' : 'mes') + ' para pagar (' + j.n + ' albarán' + (j.n === 1 ? '' : 'es') + ' de ' + PENDIENTES_MESES_[j.fecha.getMonth()] + '). No es un problema.');
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
        var archivo;
        try { archivo = id && DriveApp.getFileById(id); if (!archivo || archivo.isTrashed()) return; } catch (e) { return; }
        var partes = String(m.nombre).split(' / ');
        // Solo importan las facturas que van a los contables (25/11/2026,
        // Luciana): un albarán, o la foto de una factura que después llega por
        // email (Envapro, Frutapro), da igual que se vea mal.
        var desc = String(archivo.getDescription() || '').split('\n')[0];
        if (desc === DOC_ALBARAN_) return;
        var trat = partes.length > 1 ? tratamientoDe_(partes[0]) : null;
        if (trat && trat.c === 'FACTURA' && trat.f === 'DE_FACTURA') return;
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

// 24/11/2026 (Luciana): por secciones, y arriba un desplegable "Ver mes"
// (B3) que deja ver solo ese mes (esconde las demás filas). La columna K
// (escondida) dice qué es cada fila: S|<sección>, C (títulos) o D (dato).
var PENDIENTES_COL_MARCA_ = 11, PENDIENTES_CELDA_MES_ = 'B3', PENDIENTES_TODOS_ = 'Todos los meses';

function pintarPendientes_(ss, items) {
  var h = ss.getSheetByName(PENDIENTES_PESTANA_) || ss.insertSheet(PENDIENTES_PESTANA_, 0);
  var mesAntes = String(h.getRange(PENDIENTES_CELDA_MES_).getValue() || '');
  if (h.getFilter()) h.getFilter().remove();
  h.getRange(1, 1, h.getMaxRows(), h.getMaxColumns()).breakApart().clearDataValidations();
  h.clear();
  h.showColumns(1, h.getMaxColumns());
  h.showRows(1, h.getMaxRows());
  h.setFrozenRows(0);
  var ancho = PENDIENTES_CAB_.length; // A–J
  var orden = {}, titulo = {};
  PENDIENTES_SECCIONES_.forEach(function (s, i) { orden[s[0]] = i; titulo[s[0]] = s[1]; });

  h.getRange(1, 1).setValue('Pendientes de facturas y albaranes').setFontWeight('bold').setFontSize(14);
  h.getRange(2, 1).setValue('Actualizado el ' + Utilities.formatDate(new Date(), 'Europe/Madrid', 'dd/MM/yyyy HH:mm') +
    ' (se rehace solo cada 2 horas). Lo que se soluciona sale solo; lo que tiene casilla "Ya está": marcarla cuando esté arreglado y desaparece.').setFontColor('#666');
  var meses = {};
  items.forEach(function (x) { meses[x.mes || '(sin fecha)'] = true; });
  var lista = [PENDIENTES_TODOS_].concat(Object.keys(meses).sort().reverse());
  h.getRange('A3').setValue('Ver mes:').setFontWeight('bold').setHorizontalAlignment('right');
  h.getRange(PENDIENTES_CELDA_MES_).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(lista, true).setAllowInvalid(false).build())
    .setValue(lista.indexOf(mesAntes) > -1 ? mesAntes : PENDIENTES_TODOS_)
    .setFontWeight('bold').setBackground('#fff2cc').setBorder(true, true, true, true, false, false);
  // Casilla para rehacerla a mano sin esperar las 2 horas (25/11/2026, Luciana).
  h.getRange(PENDIENTES_ETIQUETA_ACTUALIZAR_).setValue('Actualizar ahora:').setFontWeight('bold').setHorizontalAlignment('right');
  h.getRange(PENDIENTES_CELDA_ACTUALIZAR_).insertCheckboxes().setValue(false).setBackground('#fff2cc');

  items.sort(function (a, b) {
    return (orden[a.sec] - orden[b.sec]) || ((a.fecha ? a.fecha.getTime() : 0) - (b.fecha ? b.fecha.getTime() : 0));
  });
  var fila = 6;
  PENDIENTES_SECCIONES_.forEach(function (s) {
    var deEsta = items.filter(function (x) { return x.sec === s[0]; });
    if (!deEsta.length) return;
    h.getRange(fila, 1, 1, PENDIENTES_COL_HECHO_).merge().setFontWeight('bold').setFontColor('#fff').setBackground(s[2]);
    h.getRange(fila, PENDIENTES_COL_MARCA_).setValue('S|' + s[0]);
    fila++;
    h.getRange(fila, 1, 1, ancho).setValues([PENDIENTES_CAB_]).setFontWeight('bold').setBackground('#f1efe6');
    h.getRange(fila, PENDIENTES_COL_MARCA_).setValue('C|' + s[0]);
    fila++;
    var n = deEsta.length;
    h.getRange(fila, 1, n, ancho).setValues(deEsta.map(function (x) {
      return [x.mes || '(sin fecha)', s[1].split(' (')[0], x.prov, x.fecha, x.num, x.importe, x.que, '', x.clave ? false : 'sale sola', x.clave];
    }));
    h.getRange(fila, PENDIENTES_COL_MARCA_, n, 1).setValue('D|' + s[0]);
    h.getRange(fila, 8, n, 1).setFormulas(deEsta.map(function (x) { return [x.url ? '=HYPERLINK("' + x.url + '";"Ver")' : '']; }));
    h.getRange(fila, 4, n, 1).setNumberFormat('dd/MM/yyyy');
    h.getRange(fila, 6, n, 1).setNumberFormat('#,##0.00 €');
    h.getRange(fila, 9, n, 1).setFontColor('#aaa');
    var k = 0;
    while (k < n) {
      if (!deEsta[k].clave) { k++; continue; }
      var desde = k;
      while (k < n && deEsta[k].clave) k++;
      h.getRange(fila + desde, 9, k - desde, 1).insertCheckboxes();
    }
    fila += n + 1;
  });
  h.setColumnWidth(1, 130); h.setColumnWidth(2, 190); h.setColumnWidth(3, 190); h.setColumnWidth(4, 90); h.setColumnWidth(5, 120);
  h.setColumnWidth(6, 90); h.setColumnWidth(7, 480); h.setColumnWidth(8, 50); h.setColumnWidth(9, 75);
  if (fila > 6) h.getRange(6, 7, fila - 6, 1).setWrap(true);
  h.hideColumns(PENDIENTES_COL_CLAVE_, 2);
  h.setFrozenRows(4);
  aplicarMesPendientes_(h);
}

// Deja ver solo el mes elegido en B3 (o todo), con los títulos de cada
// sección y el resumen de arriba contando solo ese mes.
function aplicarMesPendientes_(h) {
  var mes = String(h.getRange(PENDIENTES_CELDA_MES_).getValue() || PENDIENTES_TODOS_);
  var todos = mes === PENDIENTES_TODOS_;
  var ultima = h.getLastRow();
  var titulo = {};
  PENDIENTES_SECCIONES_.forEach(function (s) { titulo[s[0]] = s[1]; });
  if (ultima < 6) { h.getRange('A4').setValue('No hay nada pendiente. 👍').setFontWeight('bold'); return; }
  h.showRows(6, ultima - 5);
  var v = h.getRange(6, 1, ultima - 5, PENDIENTES_COL_MARCA_).getValues();
  var ver = v.map(function () { return true; });
  var cuenta = {}, suma = {};
  v.forEach(function (r, i) {
    var m = String(r[PENDIENTES_COL_MARCA_ - 1] || '').split('|');
    if (m[0] !== 'D') return;
    var esta = todos || r[0] === mes;
    ver[i] = esta;
    if (esta) { cuenta[m[1]] = (cuenta[m[1]] || 0) + 1; suma[m[1]] = (suma[m[1]] || 0) + (Number(r[5]) || 0); }
  });
  v.forEach(function (r, i) {
    var m = String(r[PENDIENTES_COL_MARCA_ - 1] || '').split('|');
    if (m[0] === 'S' || m[0] === 'C') ver[i] = !!cuenta[m[1]];
    if (m[0] === 'S' && cuenta[m[1]]) {
      h.getRange(6 + i, 1).setValue(titulo[m[1]] + '  (' + cuenta[m[1]] + (suma[m[1]] ? ' · ' + euros_(Math.round(suma[m[1]] * 100) / 100) : '') + ')');
    }
    if (!m[0]) ver[i] = todos; // las filas en blanco entre secciones
  });
  var i = 0;
  while (i < ver.length) {
    if (ver[i]) { i++; continue; }
    var desde = i;
    while (i < ver.length && !ver[i]) i++;
    h.hideRows(6 + desde, i - desde);
  }
  var resumen = PENDIENTES_SECCIONES_.filter(function (s) { return cuenta[s[0]]; }).map(function (s) { return s[1].split(' (')[0] + ': ' + cuenta[s[0]]; });
  h.getRange('A4').setValue(resumen.length ? resumen.join(' · ') : 'Nada pendiente ' + (todos ? '' : 'en ' + mes.slice(8) + ' ') + '👍').setFontWeight('bold');
}

// Desde alEditarAlbaranes: casilla "Ya está" marcada → se guarda y la fila
// se borra.
function alEditarPendientes_(e) {
  var hoja = e.range.getSheet();
  if (e.range.getA1Notation() === PENDIENTES_CELDA_MES_) { aplicarMesPendientes_(hoja); return; }
  if (e.range.getA1Notation() === PENDIENTES_CELDA_ACTUALIZAR_) {
    if (hoja.getRange(PENDIENTES_CELDA_ACTUALIZAR_).getValue() !== true) return;
    hoja.getRange('A4').setValue('⏳ Actualizando… tarda unos minutos (la pestaña se rehace sola al terminar).').setFontWeight('bold');
    pedirActualizarPendientes_('actualizarPendientesAMano1');
    return;
  }
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


// ---------- Actualizar a mano (casilla E3) ----------
// Primero FALTA FACTURA (documentosAlbaranes_) y después PENDIENTES, en dos
// ejecuciones seguidas: juntas pueden pasar de los 6 minutos.
var PENDIENTES_ETIQUETA_ACTUALIZAR_ = 'D3', PENDIENTES_CELDA_ACTUALIZAR_ = 'E3';
function pedirActualizarPendientes_(funcion) {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === funcion) ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger(funcion).timeBased().after(1000).create();
}
function borrarActivador_(funcion) {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === funcion) ScriptApp.deleteTrigger(t); });
}
function actualizarPendientesAMano1() {
  borrarActivador_('actualizarPendientesAMano1');
  try { actualizarDocumentosAlbaranes(); } catch (e) { Logger.log(e); }
  pedirActualizarPendientes_('actualizarPendientesAMano2');
}
function actualizarPendientesAMano2() {
  borrarActivador_('actualizarPendientesAMano2');
  actualizarPendientes();
}
