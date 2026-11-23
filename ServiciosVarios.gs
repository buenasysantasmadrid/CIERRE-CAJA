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
  ['Yiwu Xiaoheng', ['yiwu xiaoheng', 'xiaoheng']],
  ['Controlplaga', ['controlplaga', 'control plaga', 'albores']]
];

// A dónde va una factura de SERVICIOS que estaba mal puesta (o null si se queda).
function destinoDesdeServicios_(detalle) {
  if (/glovo/i.test(String(detalle || ''))) return { pestana: ALBARANES_PESTANA_GLOVO_, carpeta: 'Glovo', detalle: '' };
  if (variosConocido_(detalle)) return { pestana: 'VARIOS', carpeta: 'Varios', detalle: String(detalle) };
  return null;
}

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
// SERVICIOS que estaban mal puestas (destinoDesdeServicios_: VARIOS_CONOCIDOS_
// a VARIOS, Glovo a GLOVO) pasan a su pestaña y su PDF a su carpeta. Las
// columnas se copian por nombre (GLOVO no tiene columna PROVEEDOR). Las que
// vienen de una caja no se tocan (se avisa). Con probar=1 solo dice qué haría.
function pasarServiciosAVarios_(anio, probar) {
  var ss = SpreadsheetApp.openById(buscarEnIndice_('ALBARANES', String(anio)));
  var serv = pestanaAlbaranesParaProveedor_(ss, ALBARANES_PESTANA_SERVICIOS_);
  if (!serv) return { ok: true, hechos: ['No hay pestaña SERVICIOS'] };
  var raiz = DriveApp.getFolderById(CARPETA_FACTURAS_ID_);
  var hechos = [], meses = [], destinos = {};
  var hojaDestino = function (nombre) {
    if (!(nombre in destinos)) {
      var h = nombre === ALBARANES_PESTANA_GLOVO_ ? asegurarPestanaGlovoAlbaranes_(ss) : pestanaAlbaranesParaProveedor_(ss, nombre);
      destinos[nombre] = h ? { hoja: h, enc: filasEncabezadoAlbaranes_(h) } : null;
    }
    if (!destinos[nombre]) throw new Error('No hay pestaña ' + nombre);
    return destinos[nombre];
  };
  var encServ = filasEncabezadoAlbaranes_(serv);
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    encServ.forEach(function (filaEnc, mes) {
      var mapa = mapaColumnasAlbaranes_(serv.getRange(filaEnc, 1, 1, ALBARANES_ANCHO_).getValues()[0]);
      if (mapa.detalle == null) return;
      var valores = serv.getRange(filaEnc + 1, 1, ALBARANES_FILAS_DATOS_, ALBARANES_ANCHO_).getValues();
      var ids = serv.getRange(filaEnc + 1, ALBARANES_COL_ID_, ALBARANES_FILAS_DATOS_, 1).getValues();
      var quedan = [], pasan = [];
      valores.forEach(function (v, i) {
        var id = String(ids[i][0] || '');
        if (v.every(function (x) { return x === '' || x == null; }) && !id) return;
        if (probar === "todas") hechos.push("mes " + (mes + 1) + ": " + v.filter(function (x) { return x !== "" && x != null; }).map(function (x) { return x instanceof Date ? Utilities.formatDate(x, "Europe/Madrid", "dd/MM") : String(x); }).join(" · ") + " [" + id.slice(0, 4) + "]");
        var destino = destinoDesdeServicios_(v[mapa.detalle]);
        if (!destino) { quedan.push({ id: id, fila: v }); return; }
        var p = { id: id, fila: v, destino: destino, nombre: String(v[mapa.detalle]), numero: mapa.factura != null ? String(v[mapa.factura] || '') : '',
          fecha: mapa.fecha != null && v[mapa.fecha] instanceof Date ? Utilities.formatDate(v[mapa.fecha], 'Europe/Madrid', 'yyyy-MM-dd') : '',
          importe: mapa.importe != null ? Number(v[mapa.importe]) || 0 : 0 };
        p.cual = p.nombre + (p.numero ? ' nº ' + p.numero : '') + ' del ' + fechaCorta_(p.fecha) + ' (' + euros_(p.importe) + ')';
        if (id.indexOf('pdf-') !== 0) { // cargada desde una caja: se cambia en la caja
          quedan.push({ id: id, fila: v });
          hechos.push(p.cual + ': viene de una caja, se queda en SERVICIOS (cambiarla en la caja)');
          return;
        }
        pasan.push(p);
      });
      if (!pasan.length) return;
      if (probar) { pasan.forEach(function (p) { hechos.push(p.cual + ': pasaría de SERVICIOS a ' + p.destino.pestana); }); return; }
      // Por pestaña de destino, la fila con las columnas en su sitio.
      var porDestino = {};
      pasan.forEach(function (p) {
        var d = hojaDestino(p.destino.pestana);
        if (!d.enc[mes]) throw new Error(p.destino.pestana + ' no tiene el bloque del mes ' + (mes + 1));
        var mapaD = mapaColumnasAlbaranes_(d.hoja.getRange(d.enc[mes], 1, 1, ALBARANES_ANCHO_).getValues()[0]);
        var fila = []; for (var c = 0; c < ALBARANES_ANCHO_; c++) fila.push('');
        Object.keys(mapa).forEach(function (k) { if (mapaD[k] != null) fila[mapaD[k]] = p.fila[mapa[k]]; });
        if (mapaD.detalle != null) fila[mapaD.detalle] = p.destino.detalle;
        (porDestino[p.destino.pestana] = porDestino[p.destino.pestana] || []).push({ id: p.id, fila: fila });
      });
      Object.keys(porDestino).forEach(function (n) { var d = hojaDestino(n); reescribirBloqueAlbaranes_(d.hoja, d.enc[mes], {}, porDestino[n]); });
      guardarBloqueOrdenado_(serv, filaEnc, quedan);
      meses.push(mes);
      pasan.forEach(function (p) {
        var pdf = '';
        if (p.fecha) {
          try { pdf = ' · PDF a la carpeta ' + p.destino.carpeta + ': ' + moverFacturaACarpetaProveedor_(DriveApp.getFileById(p.id.slice(4)), raiz, p.fecha, p.destino.carpeta, p.destino.detalle, p.numero); }
          catch (err) { pdf = ' · el PDF no se pudo mover: ' + err; }
        }
        hechos.push(p.cual + ': de SERVICIOS a ' + p.destino.pestana + pdf);
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
