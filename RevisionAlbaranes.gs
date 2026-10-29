// ============================================================================
// REVISIÓN DE ALBARANES — lo que va en el resumen de los lunes además de lo
// de las facturas de las carpetas (ResumenFacturas.gs). No mira lo que pasó
// en la semana sino cómo está todo hoy:
//   · Repetidos: el mismo proveedor con el mismo nº (o, sin nº, la misma
//     fecha e importe) dos veces en ALBARANES.
//   · Sin pagar: filas con FORMA DE PAGO "NO PAGADO" (o "SIN PAGAR" en
//     observaciones) y sin DIA DE PAGO.
//   · Albaranes sin factura: PDF escaneados como albarán (la descripción
//     del archivo en Drive dice "albaran", ver guardarFotoFactura_) que
//     siguen en la carpeta del proveedor: cuando llega la factura, el
//     albarán se aparta a <proveedor>/Albaranes y deja de salir.
//   · Errores de la semana: lo apuntado en la pestaña Errores del Índice.
// ============================================================================

// Las filas de ALBARANES de todo el año de una pestaña, leyendo la hoja una
// sola vez. fila: la fila de la hoja (1-based).
function filasDelAnioAlbaranes_(hoja) {
  var ultima = hoja.getLastRow();
  if (ultima < 2) return [];
  var todo = hoja.getRange(1, 1, ultima, ALBARANES_COL_ID_).getValues();
  var filas = [];
  var encabezados = [];
  todo.forEach(function (v, i) { if (normalizarClave_(v[0]) === 'FECHA') encabezados.push(i); });
  encabezados.forEach(function (e, mes) {
    var mapa = mapaColumnasAlbaranes_(todo[e].slice(0, ALBARANES_ANCHO_));
    for (var i = e + 1; i <= e + ALBARANES_FILAS_DATOS_ && i < todo.length; i++) {
      var v = todo[i];
      if (v.slice(0, ALBARANES_ANCHO_).every(function (x) { return x === '' || x == null; })) continue;
      filas.push({
        mes: mes, fila: i + 1,
        fecha: mapa.fecha != null && v[mapa.fecha] instanceof Date ? v[mapa.fecha] : null,
        numero: mapa.factura != null ? String(v[mapa.factura] || '') : '',
        importe: mapa.importe != null ? Number(v[mapa.importe]) || 0 : 0,
        forma: mapa.formaPago != null ? normalizarClave_(v[mapa.formaPago]) : '',
        diaPago: mapa.diaPago != null ? v[mapa.diaPago] : '',
        detalle: mapa.detalle != null ? String(v[mapa.detalle] || '') : '',
        obs: mapa.obs != null ? String(v[mapa.obs] || '') : '',
        id: String(v[ALBARANES_COL_ID_ - 1] || '')
      });
    }
  });
  return filas;
}

function revisarAlbaranes_(anio) {
  var id = buscarEnIndice_('ALBARANES', String(anio));
  if (!id) return { ok: false, error: 'No hay ALBARANES de ' + anio };
  var ss = SpreadsheetApp.openById(id);
  var cajaDe = {};
  getOrCrearHojaIdsAlbaranes_(ss).getDataRange().getValues().slice(1).forEach(function (f) { cajaDe[String(f[0])] = textoFechaIds_(f[1]); });
  var dia = function (d) { return d instanceof Date ? Utilities.formatDate(d, 'Europe/Madrid', 'yyyy-MM-dd') : ''; };
  var repetidos = [], sinPagar = [];

  ss.getSheets().forEach(function (h) {
    var nombre = h.getName();
    if (ALBARANES_PESTANAS_EXCLUIDAS_.indexOf(nombre) > -1 || nombre === INFORME_MES_PESTANA_ || CONTABLES_PESTANAS_FUERA_.indexOf(normalizarClave_(nombre)) > -1) return;
    var conDetalle = ALBARANES_PESTANAS_CON_PROVEEDOR_.indexOf(normalizarClave_(nombre)) > -1;
    var filas = filasDelAnioAlbaranes_(h).filter(function (f) { return f.importe; });
    filas.forEach(function (f) {
      f.pestana = nombre;
      f.proveedor = conDetalle && f.detalle ? f.detalle : nombre;
      f.fechaISO = dia(f.fecha);
      f.caja = cajaDe[f.id] || '';
    });

    // Repetidos: mismo nº (de 3 o más letras/cifras, con alguna cifra) e
    // importe; o misma fecha e importe cuando alguno de los dos no tiene nº.
    var grupos = {};
    filas.forEach(function (f) {
      var k = claveNumeroFactura_(f.numero);
      if (k.length < 3 || !/\d/.test(k)) return;
      var clave = nombre + '|' + (conDetalle ? normalizarClave_(f.detalle) : '') + '|n' + k + '|' + f.importe.toFixed(2);
      (grupos[clave] = grupos[clave] || []).push(f);
    });
    filas.forEach(function (f) {
      if (!f.fechaISO) return;
      var mismos = filas.filter(function (g) {
        return g !== f && g.fechaISO === f.fechaISO && Math.abs(g.importe - f.importe) < 0.01 &&
          (!conDetalle || normalizarClave_(g.detalle) === normalizarClave_(f.detalle)) &&
          (!claveNumeroFactura_(f.numero) || !claveNumeroFactura_(g.numero));
      });
      if (!mismos.length) return;
      var k = nombre + '|fi|' + [f].concat(mismos).map(function (x) { return x.fila; }).sort().join(',');
      grupos[k] = [f].concat(mismos);
    });
    var vistos = {};
    Object.keys(grupos).forEach(function (k) {
      var g = grupos[k];
      if (g.length < 2) return;
      var huella = g.map(function (x) { return x.fila; }).sort().join(',');
      if (vistos[huella]) return;
      vistos[huella] = true;
      repetidos.push(g.map(function (x) {
        return { pestana: x.pestana, fila: x.fila, proveedor: x.proveedor, fecha: x.fechaISO, numero: x.numero, importe: x.importe, forma: x.forma, id: x.id, caja: x.caja };
      }));
    });

    filas.forEach(function (f) {
      var noPagado = f.forma === 'NO PAGADO' || /SIN\s+PAGAR|NO\s+PAGAD[OA]/i.test(f.obs);
      if (noPagado && !(f.diaPago instanceof Date) && !/(^|·\s*)PAGADO/.test(f.obs)) {
        sinPagar.push({ pestana: f.pestana, fila: f.fila, proveedor: f.proveedor, fecha: f.fechaISO, numero: f.numero, importe: f.importe });
      }
    });
  });
  sinPagar.sort(function (a, b) { return a.fecha < b.fecha ? -1 : a.fecha > b.fecha ? 1 : 0; });
  return { ok: true, repetidos: repetidos, sinPagar: sinPagar };
}

// PDF escaneados como albarán que siguen en la carpeta de su proveedor (los
// meses de `meses` atrás hasta hoy).
function albaranesSinFactura_(meses) {
  var raiz = carpetaFacturas_();
  var hoy = new Date();
  var lista = [];
  for (var i = meses; i >= 0; i--) {
    var d = new Date(hoy.getFullYear(), hoy.getMonth() - i, 1);
    var anio = raiz.getFoldersByName(String(d.getFullYear()));
    if (!anio.hasNext()) continue;
    var mes = anio.next().getFoldersByName(('0' + (d.getMonth() + 1)).slice(-2));
    if (!mes.hasNext()) continue;
    var provs = mes.next().getFolders();
    while (provs.hasNext()) {
      var prov = provs.next();
      if (prov.getName().indexOf('📥') === 0 || prov.getName().indexOf('⚠️') === 0) continue;
      var archivos = prov.getFiles();
      while (archivos.hasNext()) {
        var a = archivos.next();
        if (a.getDescription() !== DOC_ALBARAN_) continue;
        var nombre = a.getName().replace(/\.[a-z0-9]{2,4}$/i, '');
        lista.push({ proveedor: prov.getName(), fecha: nombre.slice(0, 10), nombre: nombre, enlace: a.getUrl() });
      }
    }
  }
  lista.sort(function (a, b) { return a.fecha < b.fecha ? -1 : a.fecha > b.fecha ? 1 : 0; });
  return lista;
}
var DOC_ALBARAN_ = 'albaran';

// Errores apuntados en la pestaña Errores desde `desde`, juntos por asunto.
function erroresDeLaSemana_(desde) {
  var r = leerRegistroErrores_(desde.toISOString());
  var porAsunto = {};
  (r.errores || []).forEach(function (e) {
    var x = porAsunto[e.asunto] = porAsunto[e.asunto] || { asunto: e.asunto, veces: 0, detalle: e.detalle, ultima: e.fecha };
    x.veces++;
    if (e.fecha > x.ultima) { x.ultima = e.fecha; x.detalle = e.detalle; }
  });
  return Object.keys(porAsunto).map(function (k) { return porAsunto[k]; });
}

// Claude (GET ?revisarAlbaranes=AAAA): lo mismo que sale en el resumen.
function revisionParaClaude_(anio) {
  var r = revisarAlbaranes_(anio);
  try { r.albaranesSinFactura = albaranesSinFactura_(2); } catch (err) { r.errorDrive = String(err); }
  return r;
}

// Claude (POST accionBorrarFilasAlbaranes): borra filas repetidas de
// ALBARANES que no vienen de una caja (sin ID, o "pdf-"). Cada una se
// comprueba antes (fecha, nº e importe) para no borrar otra.
function borrarFilasAlbaranes_(data) {
  var ss = SpreadsheetApp.openById(buscarEnIndice_('ALBARANES', String(data.anio)));
  var hechos = [];
  var porPestana = {};
  (data.filas || []).forEach(function (f) { (porPestana[f.pestana] = porPestana[f.pestana] || []).push(f); });
  Object.keys(porPestana).forEach(function (p) {
    var hoja = ss.getSheetByName(p);
    var filas = filasDelAnioAlbaranes_(hoja);
    var porMes = {};
    porPestana[p].forEach(function (b) {
      var f = filas.filter(function (x) { return x.fila === b.fila; })[0];
      var fechaF = f && f.fecha ? Utilities.formatDate(f.fecha, 'Europe/Madrid', 'yyyy-MM-dd') : '';
      if (!f || fechaF !== b.fecha || Math.abs(f.importe - b.importe) > 0.001 || String(f.numero) !== String(b.numero || '') || f.id !== String(b.id || '')) {
        hechos.push(p + ' fila ' + b.fila + ': no coincide, no se borra');
        return;
      }
      if (f.id && f.id.indexOf('pdf-') !== 0) {
        hechos.push(p + ' fila ' + b.fila + ': viene de una caja, no se borra aquí');
        return;
      }
      (porMes[f.mes] = porMes[f.mes] || []).push(f);
    });
    var encabezados = filasEncabezadoAlbaranes_(hoja);
    Object.keys(porMes).forEach(function (m) {
      var filaEnc = encabezados[m];
      var quitar = {};
      porMes[m].forEach(function (f) { quitar[f.fila] = true; });
      var rango = hoja.getRange(filaEnc + 1, 1, ALBARANES_FILAS_DATOS_, ALBARANES_ANCHO_).getValues();
      var ids = hoja.getRange(filaEnc + 1, ALBARANES_COL_ID_, ALBARANES_FILAS_DATOS_, 1).getValues();
      var quedan = [];
      rango.forEach(function (v, i) {
        if (quitar[filaEnc + 1 + i]) return;
        if (v.every(function (x) { return x === '' || x == null; }) && !ids[i][0]) return;
        quedan.push({ id: String(ids[i][0] || ''), fila: v });
      });
      guardarBloqueOrdenado_(hoja, filaEnc, quedan);
      porMes[m].forEach(function (f) { hechos.push(p + ' fila ' + f.fila + ' (' + euros_(f.importe) + '): borrada'); });
    });
  });
  invalidarCacheListados_();
  return { ok: true, hechos: hechos };
}

// Claude (POST accionQuitarGastoRepetido): un albarán "No pagado" que, al
// pagarlo, se volvió a cargar en otra caja (con tarjeta o transferencia) en
// vez de marcarlo como pagado. Se quita la segunda carga de su caja (solo si
// no es efectivo: la caja no cambia) y, si se dice `original`, éste queda
// "PAGADO <forma> <fecha de la segunda>", igual que desde la app.
// data: { fecha, idMovimiento, original: { fecha, idMovimiento } }
function quitarGastoRepetido_(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  var quitado = null;
  try {
    var ss = planillaCierreCaja_(data.fecha);
    var dia = obtenerDiaJSON_(ss, data.fecha);
    if (!dia || !dia.encontrado) return { ok: false, error: 'no se encontró la caja del ' + data.fecha };
    var tocados = {};
    ['mediodia', 'noche'].forEach(function (k) {
      var turno = dia[k];
      if (!turno || quitado) return;
      var movs = turno.movimientos || [];
      for (var i = 0; i < movs.length; i++) {
        var m = movs[i];
        if (m.id !== data.idMovimiento || m.tipo !== 'gasto') continue;
        if (['tarjeta', 'transferencia', 'no_pagado', 'no_efectivo'].indexOf(m.subtipo) === -1) return;
        movs.splice(i, 1);
        var c = turno.calc || (turno.calc = {});
        c.gastoNoEfectivo = redondear2_((c.gastoNoEfectivo || 0) - (Number(m.importe) || 0));
        quitado = m;
        tocados[k] = true;
        return;
      }
    });
    if (!quitado) return { ok: false, error: 'no está en la caja del ' + data.fecha + ', o es en efectivo' };
    var r = guardarDiaCambiado_(ss, data.fecha, dia, tocados, lock);
    r.quitado = { proveedor: quitado.proveedor, factura: quitado.factura, importe: quitado.importe, subtipo: quitado.subtipo };
    if (data.original && FORMA_PAGO_LABEL_[quitado.subtipo]) {
      r.original = cambiarFormaPagoAlbaran_({ fecha: data.original.fecha, idMovimiento: data.original.idMovimiento, nuevaFormaPago: quitado.subtipo, fechaPago: data.fecha });
    }
    invalidarCacheListados_();
    return r;
  } finally {
    if (!quitado) lock.releaseLock();
  }
}
