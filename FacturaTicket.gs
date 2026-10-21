// ============================================================================
// FACTURA DE UN TICKET — llega la factura de una compra que ya se anotó con
// el ticket (Dia, Disbesa...). La app encuentra el ticket (o los tickets, si
// la factura junta varios) y manda aquí la factura:
//  - Los gastos de esos tickets pasan a tener el número, la fecha, el IVA y
//    el importe de la factura (manda siempre la factura; con varios tickets
//    el importe y el IVA se reparten en proporción). No se anota un gasto
//    nuevo, para no contarlo dos veces. En Info queda "Ticket <nº viejo>".
//  - Se rehacen la caja de esos días (Registro, Movimientos, pestaña del
//    día), Contabilidad y Albaranes.
//  - En Drive, el PDF de la factura sustituye al del ticket: la factura va a
//    su carpeta (año / mes de la factura / proveedor) y el PDF del ticket se
//    mueve a <proveedor>/Albaranes, en la carpeta de su mes.
// ============================================================================

function redondear2_(x) { return Math.round((Number(x) || 0) * 100) / 100; }

function aplicarFacturaATickets_(data) {
  var tickets = (data.tickets || []).filter(function (t) { return t && t.idMovimiento && t.fecha; });
  if (!tickets.length) return { ok: false, error: 'No llegó ningún ticket' };
  var f = data.factura || {};
  var numero = String(f.numero || '').trim();
  var fechaFactura = /^\d{4}-\d{2}-\d{2}$/.test(String(f.fecha || '')) ? f.fecha : hoyISO_();
  var total = redondear2_(f.importe);
  var iva = redondear2_(f.iva);
  if (!total) return { ok: false, error: 'La factura no tiene importe' };

  // Reparto: cada ticket se lleva su parte de la factura; el último, lo que
  // quede (para que la suma dé justo el total).
  var suma = tickets.reduce(function (s, t) { return s + (Number(t.importe) || 0); }, 0);
  var llevado = 0, llevadoIva = 0;
  tickets.forEach(function (t, i) {
    var parte = suma > 0 ? (Number(t.importe) || 0) / suma : 1 / tickets.length;
    if (i === tickets.length - 1) {
      t.nuevoImporte = redondear2_(total - llevado);
      t.nuevoIva = redondear2_(iva - llevadoIva);
    } else {
      t.nuevoImporte = redondear2_(total * parte);
      t.nuevoIva = redondear2_(iva * parte);
      llevado += t.nuevoImporte; llevadoIva += t.nuevoIva;
    }
  });

  var porDia = {};
  tickets.forEach(function (t) { (porDia[t.fecha] = porDia[t.fecha] || []).push(t); });

  var resultado = { ok: true, dias: [], cambios: [] };
  Object.keys(porDia).sort().forEach(function (fecha) {
    var r = aplicarFacturaEnDia_(fecha, porDia[fecha], numero, fechaFactura, data.quien);
    resultado.dias.push(r);
    if (!r.ok) resultado.ok = false;
    (r.cambios || []).forEach(function (c) { resultado.cambios.push(c); });
  });
  invalidarCacheListados_();

  // La app manda lo de Drive en otra llamada (sinDrive), para no tener la
  // pantalla esperando: mover los PDF tarda.
  if (!data.sinDrive) {
    var d = facturaDeTicketEnDrive_(data);
    if (d.errorFoto) resultado.errorFoto = d.errorFoto;
    if (d.errorTicketsDrive) resultado.errorTicketsDrive = d.errorTicketsDrive;
  }

  var cual = (data.proveedor || '?') + ', factura ' + (numero || '?') + ' del ' + fechaFactura;
  var fallos = resultado.dias.filter(function (d) { return !d.ok; }).map(function (d) { return '- caja del ' + d.fecha + ': ' + d.error; });
  if (resultado.errorFoto) fallos.push('- el PDF de la factura no se guardó en Drive: ' + resultado.errorFoto);
  if (resultado.errorTicketsDrive) fallos.push('- no se pudo mover el PDF del ticket a Albaranes: ' + resultado.errorTicketsDrive);
  if (fallos.length) {
    avisarPorEmail_('algo falló al pasar una factura a su ticket', 'Al aplicar la factura (' + cual + ') a sus tickets:\n' + fallos.join('\n'));
  }
  return resultado;
}

// Drive: primero los tickets a Albaranes y después la factura (así un
// número de factura que contenga el del ticket no se mueve por error).
function facturaDeTicketEnDrive_(data) {
  var f = data.factura || {};
  var numero = String(f.numero || '').trim();
  var fechaFactura = /^\d{4}-\d{2}-\d{2}$/.test(String(f.fecha || '')) ? f.fecha : hoyISO_();
  var proveedor = String(data.proveedor || '').trim();
  var r = { ok: true };
  try {
    (data.tickets || []).forEach(function (t) { moverTicketAAlbaranes_(t, proveedor); });
  } catch (err) { r.ok = false; r.errorTicketsDrive = String(err); }
  try {
    if (data.pdfBase64) {
      r.foto = guardarFotoFactura_({
        pdfBase64: data.pdfBase64, numero_factura: numero, idEscaneo: data.idEscaneo, detalle: data.detalle
      }, proveedor || 'Sin proveedor', fechaFactura);
    }
  } catch (err) { r.ok = false; r.errorFoto = String(err); }
  return r;
}

// La llamada aparte de la app con el PDF (después de pasar la factura).
function facturaDeTicketSoloDrive_(data) {
  var r = facturaDeTicketEnDrive_(data);
  var fallos = [];
  if (r.errorFoto) fallos.push('- el PDF de la factura no se guardó en Drive: ' + r.errorFoto);
  if (r.errorTicketsDrive) fallos.push('- no se pudo mover el PDF del ticket a Albaranes: ' + r.errorTicketsDrive);
  if (fallos.length) {
    var f = data.factura || {};
    avisarPorEmail_('algo falló al guardar en Drive una factura de ticket',
      'Factura ' + (f.numero || '?') + ' de ' + (data.proveedor || '?') + ' del ' + (f.fecha || '?') + ' (ya está pasada a su ticket en la caja):\n' + fallos.join('\n'));
  }
  return r;
}

// Un día: cambia los gastos de esos tickets en su turno, ajusta las cuentas
// del turno (solo cambia lo gastado) y vuelve a escribir todo lo de ese día.
function aplicarFacturaEnDia_(fecha, tickets, numero, fechaFactura, quien) {
  var lock = LockService.getScriptLock(), conCandado = false;
  try {
    var ss = planillaCierreCaja_(fecha);
    // Leer y escribir Registro/Movimientos de a un guardado por vez (igual
    // que doPost), para no pisar un guardado de la app que llegue a la vez.
    lock.waitLock(30000);
    conCandado = true;
    var dia = obtenerDiaJSON_(ss, fecha);
    if (!dia || !dia.encontrado) return { ok: false, fecha: fecha, error: 'no se encontró la caja de ese día' };
    var NO_EFECTIVO = ['no_efectivo', 'tarjeta', 'no_pagado', 'transferencia'];
    var cambios = [], tocados = {};
    tickets.forEach(function (t) {
      ['mediodia', 'noche'].forEach(function (k) {
        var turno = dia[k];
        if (!turno) return;
        (turno.movimientos || []).forEach(function (m) {
          if (m.id !== t.idMovimiento || m.tipo !== 'gasto') return;
          var viejo = redondear2_(m.importe);
          var delta = redondear2_(t.nuevoImporte - viejo);
          // Ya pasado antes (se corrige el importe de la factura): el ticket
          // sigue siendo el de la primera vez.
          // (o, para arreglar uno pasado antes de que se apuntara esto, lo que
          // venga en t.original).
          var previo = m.pasoAFactura || t.original || null;
          var numeroTicket = previo ? String(previo.ticket || '') : String(m.factura || '').trim();
          // Quién y cuándo lo pasó, y cómo era el ticket: se ve en la caja
          // ("Tickets que pasaron a factura").
          m.pasoAFactura = {
            ticket: numeroTicket,
            importeTicket: previo ? previo.importeTicket : viejo,
            ivaTicket: previo ? previo.ivaTicket : (m.iva === null || m.iva === undefined ? '' : m.iva),
            fechaTicket: previo ? previo.fechaTicket : (m.fecha || fecha),
            cuando: (previo && previo.cuando) || new Date().toISOString(),
            quien: String(quien || (previo && previo.quien) || '').slice(0, 60)
          };
          m.importe = t.nuevoImporte;
          m.iva = t.nuevoIva;
          m.factura = numero;
          m.fecha = fechaFactura;
          var marca = 'Ticket ' + (numeroTicket || 'del ' + fecha);
          if (String(m.info || '').indexOf(marca) === -1) m.info = m.info ? m.info + ' · ' + marca : marca;
          if (delta) {
            var c = turno.calc || (turno.calc = {});
            if (NO_EFECTIVO.indexOf(m.subtipo) > -1) {
              c.gastoNoEfectivo = redondear2_((c.gastoNoEfectivo || 0) + delta);
            } else {
              c.gastoEfectivo = redondear2_((c.gastoEfectivo || 0) + delta);
              c.esperado = redondear2_((c.esperado || 0) - delta);
              c.diferencia = redondear2_((c.diferencia || 0) + delta);
            }
          }
          cambios.push({ fecha: fecha, turno: k, numeroTicket: numeroTicket, importeAntes: viejo, importeAhora: t.nuevoImporte });
          tocados[k] = true;
        });
      });
    });
    if (!cambios.length) return { ok: false, fecha: fecha, error: 'no se encontró el ticket en esa caja (¿se borró?)' };

    var registro = getOrCrearRegistro_(ss);
    var mov = getOrCrearMovimientos_(ss);
    var ahora = new Date();
    var datos = { fecha: fecha, negocio: dia.negocio, dispositivo: 'sheet' };
    Object.keys(tocados).forEach(function (k) {
      escribirRegistroYMovimientos_(registro, mov, datos, turnoComoLoMandaLaApp_(dia[k]), k === 'noche' ? 'Noche' : 'Mediodía', ahora);
    });
    SpreadsheetApp.flush();
    lock.releaseLock();
    conCandado = false;

    var hoja = { mediodia: turnoParaHojaExacta_(dia.mediodia), noche: turnoParaHojaExacta_(dia.noche) };
    var r = { ok: true, fecha: fecha, cambios: cambios };
    try { escribirHojaDelDiaComprobada_(ss, { fecha: fecha, negocio: dia.negocio, mediodia: hoja.mediodia, noche: hoja.noche }); }
    catch (errHoja) { r.ok = false; r.error = 'pestaña del día: ' + errHoja; }
    marcarCajaCambiada_(fecha, ahora);
    var cont = escribirContabilidad_({ fecha: fecha, mediodia: dia.mediodia, noche: dia.noche });
    if (cont && cont.ok === false) { r.ok = false; r.error = (r.error ? r.error + '; ' : '') + 'Contabilidad: ' + cont.error; }
    try { sincronizarAlbaranesDelDia_(fecha, gastosDelDia_(obtenerDiaJSON_(ss, fecha)), true); }
    catch (errAlb) { r.ok = false; r.error = (r.error ? r.error + '; ' : '') + 'Albaranes: ' + errAlb; }
    return r;
  } catch (err) {
    return { ok: false, fecha: fecha, error: String(err) };
  } finally {
    if (conCandado) lock.releaseLock();
  }
}

// El turno guardado (claves internas) con la forma en que lo manda la app
// al guardar (ver doPost): movimientos con etiquetas y movimientosRaw.
function turnoComoLoMandaLaApp_(turno) {
  var t = {};
  for (var k in turno) if (k !== 'actualizadoEn') t[k] = turno[k];
  t.estado = 'Sincronizado';
  t.movimientosRaw = turno.movimientos || [];
  t.movimientos = (turno.movimientos || []).map(function (m) {
    return {
      id: m.id, tipo: TIPO_LABEL_MAP_[m.tipo] || m.tipo, subtipo: SUBTIPO_LABEL_MAP_[m.subtipo] || '',
      proveedor: m.proveedor || '', proveedorDetalle: m.proveedorDetalle || '', motivo: m.motivo || '',
      responsable: m.responsable || '', factura: m.factura || '', info: m.info || '',
      iva: m.iva, importe: m.importe, fechaFactura: m.tipo === 'gasto' ? (m.fecha || '') : ''
    };
  });
  return t;
}

// El PDF del ticket (guardado al escanearlo, con su número en el nombre) se
// mueve a <proveedor>/Albaranes dentro de la carpeta de su mes.
function moverTicketAAlbaranes_(t, proveedorFactura) {
  var numeroTicket = sinCaracteresRaros_(String(t.factura || '').trim());
  if (!numeroTicket) return;
  var raiz = carpetaFacturas_();
  var fechas = [t.fechaFactura, t.fecha].filter(function (x, i, a) { return x && a.indexOf(x) === i; });
  var proveedores = [String(t.proveedorTexto || '').split(' —')[0].trim(), proveedorFactura]
    .filter(function (x, i, a) { return x && a.indexOf(x) === i; });
  fechas.forEach(function (fecha) {
    var anio = raiz.getFoldersByName(fecha.slice(0, 4));
    if (!anio.hasNext()) return;
    var mes = anio.next().getFoldersByName(fecha.slice(5, 7));
    if (!mes.hasNext()) return;
    var dirMes = mes.next();
    proveedores.forEach(function (p) {
      var it = dirMes.getFoldersByName(sinCaracteresRaros_(p));
      if (!it.hasNext()) return;
      var dirProv = it.next();
      var archivos = dirProv.getFiles();
      while (archivos.hasNext()) {
        var a = archivos.next();
        if (a.getName().indexOf(numeroTicket) === -1) continue;
        a.moveTo(carpetaHija_(dirProv, 'Albaranes'));
      }
    });
  });
}

// Los tickets que se pasaron a factura en esta fecha (el ticket puede ser de
// la caja de otro día): la app los muestra en "Tickets que pasaron a
// factura" de la caja de ese día. Se buscan en el Registro de las cajas de
// los últimos meses; de cada movimiento, la fila guardada más nueva.
function ticketsPasadosElDia_(fecha) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha)) return { ok: false, error: 'Fecha inválida' };
  try {
    var porId = {};
    planillasCierreCajaRecientes_(3).forEach(function (p) {
      var registro = p.ss.getSheetByName('Registro');
      if (!registro) return;
      var valores = registro.getDataRange().getValues();
      var header = valores[0];
      var idxFecha = header.indexOf('Fecha'), idxTurno = header.indexOf('Turno'), idxUltima = header.indexOf('Última actualización');
      var idxJSON = -1;
      for (var h = 0; h < header.length; h++) if (String(header[h] || '').indexOf('Datos JSON') === 0) { idxJSON = h; break; }
      if (idxJSON === -1) return;
      for (var i = 1; i < valores.length; i++) {
        var texto = String(valores[i][idxJSON] || '');
        if (texto.indexOf('pasoAFactura') === -1) continue;
        var turno;
        try { turno = JSON.parse(texto); } catch (e) { continue; }
        var fechaRaw = valores[i][idxFecha];
        var fechaCaja = fechaRaw instanceof Date ? Utilities.formatDate(fechaRaw, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(fechaRaw || '');
        var ultima = idxUltima > -1 && valores[i][idxUltima] instanceof Date ? valores[i][idxUltima].getTime() : 0;
        (turno.movimientos || []).forEach(function (m) {
          var paso = m.pasoAFactura;
          if (m.tipo !== 'gasto' || !paso || !paso.cuando) return;
          var cuando = new Date(paso.cuando);
          if (isNaN(cuando) || Utilities.formatDate(cuando, 'Europe/Madrid', 'yyyy-MM-dd') !== fecha) return;
          var previo = porId[m.id];
          if (previo && previo.ultima > ultima) return;
          porId[m.id] = {
            ultima: ultima, fechaCaja: fechaCaja, turno: String(valores[i][idxTurno] || ''),
            mov: { id: m.id, tipo: m.tipo, subtipo: m.subtipo, proveedor: m.proveedor || '', proveedorDetalle: m.proveedorDetalle || '',
              factura: m.factura || '', fecha: m.fecha || '', importe: m.importe, iva: m.iva, info: m.info || '', pasoAFactura: paso }
          };
        });
      }
    });
    var lista = Object.keys(porId).map(function (k) { var x = porId[k]; delete x.ultima; return x; });
    lista.sort(function (a, b) { return a.mov.pasoAFactura.cuando < b.mov.pasoAFactura.cuando ? -1 : 1; });
    return { ok: true, tickets: lista };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}
