// ============================================================================
// ¿SERVICIOS o VARIOS? Para facturas que no son de ningún proveedor de la lista.
//
// Regla de Luciana (23/11/2026): si no se sabe dónde va, a VARIOS, y se le
// pregunta. Solo van a SERVICIOS los servicios conocidos (esta lista). Cuando
// ella confirme que algo es un servicio, se añade a SERVICIOS_CONOCIDOS_; si
// confirma que es de VARIOS, a VARIOS_CONOCIDOS_ (así no se vuelve a preguntar).
// La misma lista de servicios está en index.html (SERVICIOS_CONOCIDOS): cambiar las dos.
// ============================================================================

// [nombre que queda como detalle (null = el nombre leído), palabras con que se reconoce]
var SERVICIOS_CONOCIDOS_ = [
  ['Carsan', ['carsan']], ['GSC', ['gsc', 'grupo de software', 'software y consulting', 'software consulting']], ['Prosegur', ['prosegur']],
  ['Logos Energía', ['logos energia', 'logos energy', 'logos']], ['Love Energy', ['love energy', 'love energia', 'love']],
  ['Endered', ['endered', 'edenred']], ['O2', ['o2']], ['Pluxee', ['pluxee', 'pluxe']], ['Up', ['up']],
  ['AGEDI-AIE', ['agedi']],
  ['SGAE', ['sgae', 'sociedad general de autores', 'sociedad general de autories', 'autores y editores']],
  [null, ['zhjang', 'zhang', 'zhiang']]
];

// Confirmados por Luciana como VARIOS (no se pregunta).
var VARIOS_CONOCIDOS_ = [
  ['Amazon', ['amazon']],
  ['Birou Gas', ['birou gas', 'birou']],
  ['Realisaprint', ['realisaprint']],
  ['Yiwu Xiaoheng', ['yiwu xiaoheng', 'xiaoheng']]
];

function nombreParaComparar_(s) {
  return ' ' + String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[^a-z0-9]+/g, ' ').trim() + ' ';
}
function buscarEnLista_(lista, nombre) {
  var n = nombreParaComparar_(nombre);
  if (!n.trim()) return null;
  for (var i = 0; i < lista.length; i++) {
    if (lista[i][1].some(function (w) { return n.indexOf(' ' + w + ' ') > -1; })) return lista[i][0] || String(nombre).trim();
  }
  return null;
}
function servicioConocido_(nombre) { return buscarEnLista_(SERVICIOS_CONOCIDOS_, nombre); }
function variosConocido_(nombre) { return buscarEnLista_(VARIOS_CONOCIDOS_, nombre); }

// Solo Claude (GET ?serviciosAVarios=AAAA[&probar=1]): las facturas de
// SERVICIOS que son de VARIOS_CONOCIDOS_ pasan a la pestaña VARIOS (la fila
// tal cual, con el nombre real en PROVEEDOR) y su PDF a la carpeta Varios.
// Las que vienen de una caja no se tocan (se avisa). Con probar=1 solo dice qué haría.
function pasarServiciosAVarios_(anio, probar) {
  var ss = SpreadsheetApp.openById(buscarEnIndice_('ALBARANES', String(anio)));
  var serv = pestanaAlbaranesParaProveedor_(ss, ALBARANES_PESTANA_SERVICIOS_);
  if (!serv) return { ok: true, hechos: ['No hay pestaña SERVICIOS'] };
  var varios = pestanaAlbaranesParaProveedor_(ss, 'VARIOS');
  if (!varios) return { ok: false, error: 'No hay pestaña VARIOS' };
  var raiz = DriveApp.getFolderById(CARPETA_FACTURAS_ID_);
  var hechos = [], meses = [];
  var encServ = filasEncabezadoAlbaranes_(serv), encVar = filasEncabezadoAlbaranes_(varios);
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    encServ.forEach(function (filaEnc, mes) {
      var cab = serv.getRange(filaEnc, 1, 1, ALBARANES_ANCHO_).getValues()[0];
      var mapa = mapaColumnasAlbaranes_(cab);
      if (mapa.detalle == null) return;
      var valores = serv.getRange(filaEnc + 1, 1, ALBARANES_FILAS_DATOS_, ALBARANES_ANCHO_).getValues();
      var ids = serv.getRange(filaEnc + 1, ALBARANES_COL_ID_, ALBARANES_FILAS_DATOS_, 1).getValues();
      var quedan = [], pasan = [];
      valores.forEach(function (v, i) {
        var id = String(ids[i][0] || '');
        if (v.every(function (x) { return x === '' || x == null; }) && !id) return;
        var nombre = variosConocido_(v[mapa.detalle]);
        if (!nombre) { quedan.push({ id: id, fila: v }); return; }
        var p = { id: id, fila: v.slice(), nombre: String(v[mapa.detalle]), numero: mapa.factura != null ? String(v[mapa.factura] || '') : '',
          fecha: mapa.fecha != null && v[mapa.fecha] instanceof Date ? Utilities.formatDate(v[mapa.fecha], 'Europe/Madrid', 'yyyy-MM-dd') : '',
          importe: mapa.importe != null ? Number(v[mapa.importe]) || 0 : 0 };
        var cual = p.nombre + (p.numero ? ' nº ' + p.numero : '') + ' del ' + fechaCorta_(p.fecha) + ' (' + euros_(p.importe) + ')';
        if (id.indexOf('pdf-') !== 0) { // cargada desde una caja: se cambia en la caja
          quedan.push({ id: id, fila: v });
          hechos.push(cual + ': viene de una caja, se queda en SERVICIOS (cambiarla en la caja)');
          return;
        }
        p.cual = cual;
        pasan.push(p);
      });
      if (!pasan.length) return;
      if (!encVar[mes]) throw new Error('VARIOS no tiene el bloque del mes ' + (mes + 1));
      var cabVar = varios.getRange(encVar[mes], 1, 1, ALBARANES_ANCHO_).getValues()[0];
      if (cabVar.join('|') !== cab.join('|')) throw new Error('Las columnas de SERVICIOS y VARIOS no son iguales (mes ' + (mes + 1) + ')');
      if (probar) { pasan.forEach(function (p) { hechos.push(p.cual + ': pasaría de SERVICIOS a VARIOS'); }); return; }
      reescribirBloqueAlbaranes_(varios, encVar[mes], {}, pasan.map(function (p) { return { id: p.id, fila: p.fila }; }));
      guardarBloqueOrdenado_(serv, filaEnc, quedan);
      meses.push(mes);
      pasan.forEach(function (p) {
        var pdf = '';
        if (p.fecha) {
          try { pdf = ' · PDF a la carpeta Varios: ' + moverFacturaACarpetaProveedor_(DriveApp.getFileById(p.id.slice(4)), raiz, p.fecha, 'Varios', p.nombre, p.numero); }
          catch (err) { pdf = ' · el PDF no se pudo mover: ' + err; }
        }
        hechos.push(p.cual + ': de SERVICIOS a VARIOS' + pdf);
      });
    });
  } finally {
    lock.releaseLock();
  }
  if (!probar && meses.length) {
    try { refrescarInformeMesSiToca_(ss, meses); } catch (err) { hechos.push('INFORME MES: ' + err); }
    try { CacheService.getScriptCache().remove('totales_sin_servicios_extras_' + ss.getId()); totalSinServiciosNiExtras_(ss); } catch (err) { hechos.push('TOTAL: ' + err); }
    try { actualizarProveedoresContabilidad_(String(anio)); } catch (err) { hechos.push('Contabilidad: ' + err); }
    invalidarCacheListados_();
  }
  return { ok: true, hechos: hechos };
}
