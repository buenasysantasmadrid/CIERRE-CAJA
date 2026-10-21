// ============================================================================
// INFORME SEMANAL DE COMPRAS — se ve en la página de informes (solo
// lfisbein), sección "Compras de la semana": los 7 días hasta la fecha
// elegida. Ya no se manda por email ni lleva conclusiones de la IA (costaba
// dinero cada semana); se arma solo con los datos, gratis.
//   1. Cambios de precio (facturas leídas esa semana).
//   2. Productos que se pidieron más que de costumbre.
//   3. Gasto por proveedor comparado con su media.
//   4. Pedidos de la app cuya factura no se ha escaneado.
// La "media" es el promedio semanal de las INFORME_SEMANAS_MEDIA_ semanas
// anteriores.
//
// De dónde sale cada cosa:
// - Precios y pedidos: la hoja de la app de Pedidos (APP_Precios,
//   APP_Pedidos), que es de lfisbein y se abre directamente. Los precios de
//   Frutapro (no está en Pedidos), de "Precios proveedores" del Índice.
// - Gasto por proveedor: la planilla de Albaranes (las facturas cargadas en
//   la caja), por la fecha de la factura.
// - Facturas escaneadas: la carpeta "Facturas proveedores" de Drive (ver
//   Precios.gs) y las facturas cuyos precios ya están en APP_Precios.
//
// Claude puede verlo con doPost accionInformeSemanal: 'vista' y su clave.
// ============================================================================

var INFORME_SEMANAS_MEDIA_ = 8;
var PEDIDOS_HOJA_ID_ = '11X7L9CcmklQtHbqdCyi29O0QPzpeEUPAK-GH5c8Fr0I';
// Pestañas de Albaranes que no son compras a proveedores.
var INFORME_SEMANAL_SIN_ALBARANES_ = ['SERVICIOS', 'EXTRAS', 'NOO', 'INFORME MES', 'GLOVO'];

// Antes la llamaba el activador de los viernes para mandar el email: ahora,
// si todavía salta, solo borra ese activador (no manda nada).
function informeSemanalPedidos() {
  quitarActivadorInformeSemanal_();
}
function quitarActivadorInformeSemanal_() {
  var quitados = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'informeSemanalPedidos') { ScriptApp.deleteTrigger(t); quitados++; }
  });
  return quitados;
}

function accionInformeSemanal_(data) {
  try {
    if (data.accionInformeSemanal === 'vista') { var v = armarInformeSemanal_(data.hasta); return { ok: true, asunto: v.asunto, html: v.html, datos: v.datos }; }
    if (data.accionInformeSemanal === 'quitarEmail') return { ok: true, activadoresQuitados: quitarActivadorInformeSemanal_() };
    return { ok: false, error: 'Acción desconocida' };
  } catch (err) {
    return { ok: false, error: String(err && err.stack || err) };
  }
}

// ---------- fechas (siempre como texto AAAA-MM-DD, hora de Madrid) ----------
function semanalISO_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Europe/Madrid', 'yyyy-MM-dd');
  var s = String(v || '');
  return /^\d{4}-\d{2}-\d{2}/.test(s) ? s.slice(0, 10) : '';
}
function semanalSumarDias_(iso, n) {
  return new Date(new Date(iso + 'T12:00:00Z').getTime() + n * 86400000).toISOString().slice(0, 10);
}
function semanalDiasEntre_(desde, hasta) {
  return Math.round((new Date(hasta + 'T12:00:00Z') - new Date(desde + 'T12:00:00Z')) / 86400000);
}
// 0 = esta semana, 1..N = hace N semanas, -1 = fuera.
function semanalSemanaDe_(iso, inicio, hoy) {
  if (!iso || iso > hoy) return -1;
  if (iso >= inicio) return 0;
  var k = Math.ceil(semanalDiasEntre_(iso, inicio) / 7);
  return k <= INFORME_SEMANAS_MEDIA_ ? k : -1;
}
function semanalFechaCorta_(iso) {
  var m = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
  return Number(iso.slice(8, 10)) + ' ' + m[Number(iso.slice(5, 7)) - 1];
}
function semanalNum_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  var s = String(v || '').replace(/[^\d,.\-]/g, '');
  if (s.indexOf(',') >= 0) s = s.replace(/\./g, '').replace(',', '.');
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}
function semanalR2_(n) { return Math.round(n * 100) / 100; }

// ---------- datos ----------
// hasta: último día de la semana (AAAA-MM-DD); si no, hoy.
function armarInformeSemanal_(hasta) {
  var hoy = Utilities.formatDate(new Date(), 'Europe/Madrid', 'yyyy-MM-dd');
  if (/^\d{4}-\d{2}-\d{2}$/.test(String(hasta || '')) && hasta < hoy) hoy = hasta;
  var inicio = semanalSumarDias_(hoy, -6);
  var ss = SpreadsheetApp.openById(PEDIDOS_HOJA_ID_);
  // Y los de proveedores que no están en Pedidos (Frutapro, ver Precios.gs).
  var precios = semanalLeerPestana_(ss, 'APP_Precios').map(function (f) { return f.slice(0, 13); });
  try { precios = precios.concat(filasPreciosPropios_()); } catch (err) { Logger.log('Precios proveedores: ' + err); }
  var pedidos = semanalLeerPestana_(ss, 'APP_Pedidos');
  var precioActual = {};
  semanalLeerPestana_(ss, 'APP_Productos').forEach(function (f) {
    if (f[0]) precioActual[f[0]] = semanalNum_(f[7]);
  });

  var datos = {
    semana: { desde: inicio, hasta: hoy, semanasParaLaMedia: INFORME_SEMANAS_MEDIA_ },
    cambiosDePrecio: semanalCambiosDePrecio_(precios, inicio, hoy),
    productos: semanalProductos_(pedidos, precioActual, inicio, hoy),
    proveedores: semanalGastoProveedores_(inicio, hoy),
    facturasQueFaltan: semanalFacturasQueFaltan_(pedidos, precios, hoy)
  };
  var asunto = 'Compras de la semana: ' + semanalFechaCorta_(inicio) + ' – ' + semanalFechaCorta_(hoy);
  return { asunto: asunto, html: semanalHtml_(datos), datos: datos };
}

function semanalLeerPestana_(ss, nombre) {
  var h = ss.getSheetByName(nombre);
  if (!h || h.getLastRow() < 2) return [];
  return h.getRange(2, 1, h.getLastRow() - 1, h.getLastColumn()).getValues();
}

// APP_Precios: Fecha · Proveedor · Factura · En la factura · ID producto ·
// Producto · Cantidad · Unidad · Precio · Precio anterior · Cambio % · Foto ·
// Registrado. Cuenta la fila si la factura es de esta semana o se leyó esta
// semana. Del mismo producto, la más nueva.
function semanalCambiosDePrecio_(filas, inicio, hoy) {
  var porProducto = {}, facturas = {};
  filas.forEach(function (f) {
    var fecha = semanalISO_(f[0]), reg = semanalISO_(f[12]);
    if (!((fecha >= inicio && fecha <= hoy) || (reg >= inicio && reg <= hoy))) return;
    facturas[f[1] + '|' + (f[2] || fecha)] = 1;
    var precio = semanalNum_(f[8]), antes = semanalNum_(f[9]);
    var clave = f[4] || (f[1] + '|' + String(f[3]).toLowerCase());
    var previo = porProducto[clave];
    if (previo && previo.fecha > fecha) return;
    porProducto[clave] = {
      proveedor: f[1], producto: f[5] || f[3], unidad: f[7], fecha: fecha,
      antes: antes, ahora: precio, cambio: antes ? Math.round((precio - antes) / antes * 1000) / 10 : 0
    };
  });
  var cambios = Object.keys(porProducto).map(function (k) { return porProducto[k]; })
    .filter(function (x) { return x.antes > 0 && x.ahora > 0 && Math.abs(x.ahora - x.antes) >= 0.005; })
    .sort(function (a, b) { return Math.abs(b.cambio) - Math.abs(a.cambio); });
  return { facturasLeidas: Object.keys(facturas).length, productosConPrecio: Object.keys(porProducto).length, cambios: cambios };
}

// APP_Pedidos: Fecha · Hora · Entrega · Proveedor · Producto · Cantidad ·
// Unidad · ID producto · Pedido · Dispositivo · Precio aprox. · Importe aprox.
function semanalProductos_(filas, precioActual, inicio, hoy) {
  var primera = '', prods = {};
  filas.forEach(function (f) {
    var iso = semanalISO_(f[0]);
    if (!iso) return;
    if (!primera || iso < primera) primera = iso;
    var k = semanalSemanaDe_(iso, inicio, hoy);
    if (k < 0) return;
    var clave = f[3] + '|' + (f[7] || String(f[4]).toLowerCase().trim());
    var p = prods[clave] = prods[clave] || { proveedor: f[3], producto: f[4], unidad: f[6], id: f[7], semanas: {} };
    p.semanas[k] = (p.semanas[k] || 0) + semanalNum_(f[5]);
  });
  // Semanas anteriores que cuentan: solo desde que se empezó a usar la app.
  var previas = primera && primera < inicio ? Math.min(INFORME_SEMANAS_MEDIA_, Math.ceil(semanalDiasEntre_(primera, inicio) / 7)) : 0;
  var lista = Object.keys(prods).map(function (c) {
    var p = prods[c], suma = 0;
    for (var k = 1; k <= previas; k++) suma += p.semanas[k] || 0;
    var media = previas ? suma / previas : 0, esta = p.semanas[0] || 0;
    var precio = precioActual[p.id] || 0;
    return {
      proveedor: p.proveedor, producto: p.producto, unidad: p.unidad, estaSemana: semanalR2_(esta), media: semanalR2_(media),
      cambio: media ? Math.round((esta - media) / media * 100) : null,
      eurosDeMas: precio ? semanalR2_((esta - media) * precio) : null
    };
  });
  var mas = lista.filter(function (x) { return x.estaSemana > 0 && (x.media === 0 ? previas >= 2 : x.estaSemana >= x.media * 1.3 && x.estaSemana - x.media >= 1); })
    .sort(function (a, b) {
      return (b.eurosDeMas || 0) - (a.eurosDeMas || 0) || (b.cambio === null ? 1e9 : b.cambio) - (a.cambio === null ? 1e9 : a.cambio);
    });
  var menos = lista.filter(function (x) { return x.media >= 1 && x.estaSemana <= x.media * 0.5; })
    .sort(function (a, b) { return (a.cambio || 0) - (b.cambio || 0); });
  var pedidosSemana = {};
  filas.forEach(function (f) { if (semanalSemanaDe_(semanalISO_(f[0]), inicio, hoy) === 0) pedidosSemana[f[8]] = 1; });
  return {
    primerPedidoEnLaApp: primera, semanasAnteriores: previas, pedidosEstaSemana: Object.keys(pedidosSemana).length,
    pedidoMasQueDeCostumbre: mas.slice(0, 20), pedidoMuchoMenos: menos.slice(0, 10)
  };
}

// Albaranes: cada pestaña de proveedor tiene 12 bloques (uno por mes) con
// encabezados FECHA … IMPORTE (ver Albaranes.gs).
function semanalGastoProveedores_(inicio, hoy) {
  var desde = semanalSumarDias_(inicio, -7 * INFORME_SEMANAS_MEDIA_);
  var anios = {};
  anios[desde.slice(0, 4)] = 1; anios[hoy.slice(0, 4)] = 1;
  var provs = {}, faltaAnio = [];
  Object.keys(anios).forEach(function (anio) {
    var id = buscarEnIndice_('ALBARANES', anio);
    if (!id) { faltaAnio.push(anio); return; }
    SpreadsheetApp.openById(id).getSheets().forEach(function (hoja) {
      var nombre = hoja.getName();
      if (ALBARANES_PESTANAS_EXCLUIDAS_.indexOf(nombre) > -1 || INFORME_SEMANAL_SIN_ALBARANES_.indexOf(normalizarClave_(nombre)) > -1) return;
      if (hoja.getLastRow() < 2) return;
      var v = hoja.getRange(1, 1, hoja.getLastRow(), ALBARANES_ANCHO_).getValues(), mapa = null;
      v.forEach(function (fila) {
        if (normalizarClave_(fila[0]) === 'FECHA') { mapa = mapaColumnasAlbaranes_(fila); return; }
        if (!mapa || mapa.fecha === undefined || mapa.importe === undefined) return;
        var iso = semanalISO_(fila[mapa.fecha]), imp = semanalNum_(fila[mapa.importe]);
        var k = semanalSemanaDe_(iso, inicio, hoy);
        if (k < 0 || !imp) return;
        var p = provs[normalizarClave_(nombre)] = provs[normalizarClave_(nombre)] || { semanas: {}, facturas: 0 };
        p.semanas[k] = (p.semanas[k] || 0) + imp;
        if (k === 0) p.facturas++;
      });
    });
  });
  var lista = Object.keys(provs).map(function (n) {
    var p = provs[n], suma = 0;
    for (var k = 1; k <= INFORME_SEMANAS_MEDIA_; k++) suma += p.semanas[k] || 0;
    var media = suma / INFORME_SEMANAS_MEDIA_, esta = p.semanas[0] || 0;
    return {
      proveedor: n, estaSemana: semanalR2_(esta), media: semanalR2_(media), diferencia: semanalR2_(esta - media),
      cambio: media ? Math.round((esta - media) / media * 100) : null, facturasEstaSemana: p.facturas
    };
  }).sort(function (a, b) { return b.diferencia - a.diferencia; });
  var total = lista.reduce(function (t, x) { t.estaSemana += x.estaSemana; t.media += x.media; return t; }, { estaSemana: 0, media: 0 });
  return {
    proveedores: lista, totalEstaSemana: semanalR2_(total.estaSemana), totalMedia: semanalR2_(total.media),
    sinPlanilla: faltaAnio
  };
}

// Pedidos enviados de hace 1 a 14 días: cada uno busca una factura del mismo
// proveedor (en Drive o con precios en APP_Precios) de 3 días antes a 10 días
// después. Cada factura vale para un solo pedido.
function semanalFacturasQueFaltan_(pedidos, precios, hoy) {
  var desde = semanalSumarDias_(hoy, -14), hasta = semanalSumarDias_(hoy, -1);
  var porPedido = {};
  pedidos.forEach(function (f) {
    var iso = semanalISO_(f[0]);
    if (!iso || iso < desde || iso > hasta || !f[8]) return;
    var p = porPedido[f[8]] = porPedido[f[8]] || { proveedor: f[3], fecha: iso, entrega: f[2], lineas: 0, importe: 0 };
    p.lineas++;
    p.importe += semanalNum_(f[11]);
  });
  var lista = Object.keys(porPedido).map(function (k) { return porPedido[k]; })
    .sort(function (a, b) { return a.fecha.localeCompare(b.fecha); });
  if (!lista.length) return { revisados: 0, faltan: [], desde: desde, hasta: hasta };

  var enDrive = [];
  try { enDrive = (facturasParaPedidos_({ proveedor: '' }).facturas || []); } catch (err) { /* sin Drive: solo APP_Precios */ }
  var enPrecios = {};
  precios.forEach(function (f) {
    var iso = semanalISO_(f[0]);
    if (iso) enPrecios[f[1] + '|' + (f[2] || iso)] = { proveedor: f[1], fecha: iso, fotoId: String(f[11] || '').replace(/^.*\/d\/([^\/]+).*$/, '$1') };
  });
  // Todas las parejas pedido-factura posibles; primero las más cercanas en
  // fecha (una factura de antes del pedido cuenta el doble de lejos).
  var parejas = [];
  lista.forEach(function (ped, i) {
    var palabras = sinAcentos_(ped.proveedor).split(' ').filter(function (w) { return w.length >= 3; });
    var candidatas = [];
    enDrive.forEach(function (x) {
      var texto = ' ' + sinAcentos_(x.carpeta + ' ' + x.nombre) + ' ';
      if (palabras.some(function (w) { return texto.indexOf(' ' + w + ' ') >= 0; })) candidatas.push({ clave: x.id, fecha: x.fecha });
    });
    Object.keys(enPrecios).forEach(function (k) {
      var x = enPrecios[k];
      if (x.proveedor === ped.proveedor) candidatas.push({ clave: x.fotoId || k, fecha: x.fecha });
    });
    candidatas.forEach(function (c) {
      var dias = semanalDiasEntre_(ped.fecha, c.fecha);
      if (dias >= -3 && dias <= 10) parejas.push({ pedido: i, factura: c.clave, coste: dias < 0 ? -2 * dias : dias });
    });
  });
  parejas.sort(function (a, b) { return a.coste - b.coste; });
  var conFactura = {}, usadas = {};
  parejas.forEach(function (p) {
    if (conFactura[p.pedido] || usadas[p.factura]) return;
    conFactura[p.pedido] = 1; usadas[p.factura] = 1;
  });
  var faltan = lista.filter(function (ped, i) { return !conFactura[i]; })
    .map(function (p) { p.importe = semanalR2_(p.importe); return p; });
  return { revisados: lista.length, faltan: faltan, desde: desde, hasta: hasta };
}

// ---------- HTML (se ve en la página de informes) ----------
function semanalEsc_(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
function semanalEur_(n) {
  var s = Math.abs(n).toFixed(2).split('.');
  return (n < 0 ? '−' : '') + s[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.') + ',' + s[1] + ' €';
}
function semanalCant_(n) { return String(semanalR2_(n)).replace('.', ','); }
function semanalPct_(n) {
  if (n === null || n === undefined) return '';
  var c = n > 0 ? '#b3261e' : '#2e7d32';
  return '<span style="color:' + c + ';font-weight:600">' + (n > 0 ? '+' : '−') + String(Math.abs(n)).replace('.', ',') + ' %</span>';
}

function semanalHtml_(d) {
  var OLIVO = '#556b2f', GRIS = '#6b6f63';
  var th = 'style="text-align:left;padding:6px 8px;border-bottom:2px solid #dfe3d3;font-size:12px;color:' + GRIS + ';font-weight:600"';
  var thd = 'style="text-align:right;padding:6px 8px;border-bottom:2px solid #dfe3d3;font-size:12px;color:' + GRIS + ';font-weight:600"';
  var td = 'style="padding:6px 8px;border-bottom:1px solid #eceee5;font-size:14px"';
  var tdd = 'style="padding:6px 8px;border-bottom:1px solid #eceee5;font-size:14px;text-align:right;white-space:nowrap"';
  var nota = function (t) { return '<p style="margin:6px 0 0;font-size:13px;color:' + GRIS + '">' + t + '</p>'; };
  var nada = function (t) { return '<p style="margin:4px 0;font-size:14px;color:' + GRIS + '">' + t + '</p>'; };
  var seccion = function (n, titulo, cuerpo) {
    return '<div style="background:#fff;border:1px solid #dfe3d3;border-radius:14px;padding:16px 18px;margin:0 0 14px">' +
      '<h2 style="margin:0 0 10px;font-family:Georgia,serif;font-style:italic;font-weight:600;font-size:20px;color:' + OLIVO + '">' + n + '. ' + titulo + '</h2>' + cuerpo + '</div>';
  };
  var tabla = function (cab, filas) { return '<table style="width:100%;border-collapse:collapse">' + '<tr>' + cab + '</tr>' + filas.join('') + '</table>'; };

  // 1. Precios
  var cp = d.cambiosDePrecio, s1;
  if (cp.cambios.length) {
    s1 = tabla('<th ' + th + '>Producto</th><th ' + thd + '>Antes</th><th ' + thd + '>Ahora</th><th ' + thd + '>Cambio</th>',
      cp.cambios.map(function (x) {
        return '<tr><td ' + td + '>' + semanalEsc_(x.producto) + '<br><span style="font-size:12px;color:' + GRIS + '">' + semanalEsc_(x.proveedor) +
          (x.unidad ? ' · ' + semanalEsc_(x.unidad) : '') + '</span></td><td ' + tdd + '>' + semanalEur_(x.antes) + '</td><td ' + tdd + '>' +
          semanalEur_(x.ahora) + '</td><td ' + tdd + '>' + semanalPct_(x.cambio) + '</td></tr>';
      })) + nota('Precios sin IVA, de ' + cp.facturasLeidas + (cp.facturasLeidas === 1 ? ' factura leída' : ' facturas leídas') + ' esta semana.');
  } else {
    s1 = nada(cp.facturasLeidas ? 'Ningún precio ha cambiado en las ' + cp.facturasLeidas + ' facturas leídas esta semana (' + cp.productosConPrecio + ' productos con precio).'
      : 'Esta semana no se ha leído ninguna factura con productos de la app de pedidos.');
  }

  // 2. Productos
  var pr = d.productos, s2;
  if (pr.semanasAnteriores < 2) {
    s2 = nada('Todavía no hay bastantes semanas de pedidos en la app para comparar' +
      (pr.primerPedidoEnLaApp ? ' (el primero es del ' + semanalFechaCorta_(pr.primerPedidoEnLaApp) + ')' : ' (aún no hay pedidos enviados desde la app)') +
      '. Esta comparación saldrá cuando haya al menos 2 semanas anteriores.' + (pr.pedidosEstaSemana ? ' Esta semana se enviaron ' + pr.pedidosEstaSemana + ' pedidos.' : ''));
  } else if (!pr.pedidoMasQueDeCostumbre.length) {
    s2 = nada('Nada se ha pedido bastante más que de costumbre (' + pr.pedidosEstaSemana + ' pedidos esta semana).');
  } else {
    s2 = tabla('<th ' + th + '>Producto</th><th ' + thd + '>Esta semana</th><th ' + thd + '>Media</th><th ' + thd + '>Cambio</th>',
      pr.pedidoMasQueDeCostumbre.map(function (x) {
        return '<tr><td ' + td + '>' + semanalEsc_(x.producto) + '<br><span style="font-size:12px;color:' + GRIS + '">' + semanalEsc_(x.proveedor) +
          (x.eurosDeMas ? ' · ≈ ' + semanalEur_(x.eurosDeMas) + ' de más' : '') + '</span></td><td ' + tdd + '>' + semanalCant_(x.estaSemana) + ' ' + semanalEsc_(x.unidad || '') +
          '</td><td ' + tdd + '>' + (x.media ? semanalCant_(x.media) : 'no se pedía') + '</td><td ' + tdd + '>' + (x.cambio === null ? '' : semanalPct_(x.cambio)) + '</td></tr>';
      })) + nota('Cantidades pedidas en la app. Media = promedio semanal de las ' + pr.semanasAnteriores + ' semanas anteriores.');
  }

  // 3. Proveedores
  var pv = d.proveedores, s3;
  var conGasto = pv.proveedores.filter(function (x) { return x.estaSemana || x.media >= 1; });
  if (!conGasto.length) {
    s3 = nada('No hay facturas cargadas en Albaranes en estas semanas.');
  } else {
    s3 = tabla('<th ' + th + '>Proveedor</th><th ' + thd + '>Esta semana</th><th ' + thd + '>Media</th><th ' + thd + '>Diferencia</th>',
      conGasto.map(function (x) {
        var alto = x.media ? x.estaSemana > x.media * 1.25 && x.diferencia >= 30 : x.estaSemana >= 30;
        return '<tr' + (alto ? ' style="background:#fdf3e7"' : '') + '><td ' + td + '>' + (alto ? '<b>' : '') + semanalEsc_(x.proveedor) + (alto ? '</b>' : '') +
          '</td><td ' + tdd + '>' + semanalEur_(x.estaSemana) + '</td><td ' + tdd + '>' + semanalEur_(x.media) + '</td><td ' + tdd + '>' +
          (x.diferencia > 0 ? '+' : '') + semanalEur_(x.diferencia) + (x.cambio !== null && x.estaSemana ? '<br>' + semanalPct_(x.cambio) : '') + '</td></tr>';
      }).concat(['<tr><td ' + td + '><b>Total</b></td><td ' + tdd + '><b>' + semanalEur_(pv.totalEstaSemana) + '</b></td><td ' + tdd + '><b>' +
        semanalEur_(pv.totalMedia) + '</b></td><td ' + tdd + '><b>' + (pv.totalEstaSemana > pv.totalMedia ? '+' : '') + semanalEur_(pv.totalEstaSemana - pv.totalMedia) + '</b></td></tr>'])) +
      nota('Facturas cargadas en la caja (Albaranes, con IVA), por fecha de factura. Media = promedio semanal de las ' + INFORME_SEMANAS_MEDIA_ +
        ' semanas anteriores. En naranja, los que están bastante por encima de su media. Las facturas de esta semana que aún no se han cargado en la caja no cuentan.');
  }

  // 4. Facturas
  var fq = d.facturasQueFaltan, s4;
  if (!fq.revisados) {
    s4 = nada('No hay pedidos enviados desde la app entre el ' + semanalFechaCorta_(fq.desde) + ' y el ' + semanalFechaCorta_(fq.hasta) + ', así que no hay nada que comprobar.');
  } else if (!fq.faltan.length) {
    s4 = nada('✓ Los ' + fq.revisados + ' pedidos enviados entre el ' + semanalFechaCorta_(fq.desde) + ' y el ' + semanalFechaCorta_(fq.hasta) + ' tienen su factura escaneada.');
  } else {
    s4 = tabla('<th ' + th + '>Proveedor</th><th ' + th + '>Pedido del</th><th ' + thd + '>Líneas</th><th ' + thd + '>Importe aprox.</th>',
      fq.faltan.map(function (x) {
        return '<tr><td ' + td + '>' + semanalEsc_(x.proveedor) + '</td><td ' + td + '>' + semanalFechaCorta_(x.fecha) + (x.entrega ? ' (' + semanalEsc_(x.entrega) + ')' : '') +
          '</td><td ' + tdd + '>' + x.lineas + '</td><td ' + tdd + '>' + (x.importe ? semanalEur_(x.importe) : '—') + '</td></tr>';
      })) + nota(fq.faltan.length + ' de ' + fq.revisados + ' pedidos de las últimas dos semanas no tienen factura escaneada en el Cierre de caja. ' +
        'Si un proveedor factura una vez al mes, o una factura cubre varios pedidos, puede salir aquí sin que falte nada.');
  }

  return '<div style="background:#f4f5ef;padding:18px 12px;font-family:Helvetica,Arial,sans-serif;color:#23261f">' +
    '<div style="max-width:640px;margin:0 auto">' +
    '<div style="background:' + OLIVO + ';color:#fff;border-radius:14px;padding:16px 18px;margin:0 0 14px">' +
    '<div style="font-family:Georgia,serif;font-style:italic;font-size:24px">Compras de la semana</div>' +
    '<div style="font-size:14px;opacity:.9">Buenas y Santas · del ' + semanalFechaCorta_(d.semana.desde) + ' al ' + semanalFechaCorta_(d.semana.hasta) + '</div></div>' +
    seccion(1, 'Cambios de precio', s1) +
    seccion(2, 'Lo que se ha pedido más que de costumbre', s2) +
    seccion(3, 'Gasto por proveedor frente a su media', s3) +
    seccion(4, 'Facturas que faltan por escanear', s4) +
    '</div></div>';
}
