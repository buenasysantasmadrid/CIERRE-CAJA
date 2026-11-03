// ============================================================================
// ANÁLISIS DE VENTAS — sin IA (no cuesta nada), todo calculado con los datos:
//   · La semana: va arriba del email de los lunes (ResumenFacturas.gs).
//   · El mes: un email aparte el día 1 a las 9:00 con el mes anterior.
//
// De dónde sale cada cosa:
// - Día por día: las pestañas de cada mes de la Contabilidad (hay desde 2026).
// - Meses desde 2009: pestaña "años"; promedio por día de la semana desde
//   2016: pestaña "comparativa"; compras por proveedor: albaranes. Todo eso
//   lo lee Informes.gs (datosInformes_).
// - Semanas de años sin día por día: APROXIMADAS. Cada día = el promedio de
//   ese día de la semana en ese mes y año ("comparativa") + el delivery del
//   mes repartido entre los días trabajados ("años"). Cuando haya
//   Contabilidad día por día de ese año (2026 en adelante), se usa la real.
//
// "Facturado" = local (TOTAL SIST) + delivery + web. La venta de empanadas a
// otros negocios va aparte (en algunos años estaba dentro de TOTAL SIST).
// ============================================================================

var AN_DIAS_ = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo'];
var AN_DIAS_CORTOS_ = ['Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb', 'Dom'];
var AN_MESES_ = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
var AN_VERDE_ = '#1e8449', AN_ROJO_ = '#c0392b', AN_GRIS_ = '#888';

// ---------- utilidades ----------
function anR_(n) { return Math.round((Number(n) || 0) * 100) / 100; }
function anDow_(iso) { return (new Date(iso + 'T12:00:00Z').getUTCDay() + 6) % 7; } // lunes = 0
function anFechas_(desde, hasta) { var r = []; for (var f = desde; f <= hasta; f = semanalSumarDias_(f, 1)) r.push(f); return r; }
function anPct_(a, b) { return b ? (a - b) / b * 100 : null; }
function anDosDig_(n) { return (n < 10 ? '0' : '') + n; }
function anUltimoDia_(anio, mes) { return new Date(Date.UTC(anio, mes, 0)).getUTCDate(); } // mes 1-12
function anDDMM_(iso) { return iso.slice(8, 10) + '/' + iso.slice(5, 7); }
function anDiaCorto_(iso) { return AN_DIAS_CORTOS_[anDow_(iso)] + ' ' + anDDMM_(iso); }
function anCap_(s) { return s.charAt(0).toUpperCase() + s.slice(1); }
function anEsc_(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

// 12.345 € (sin céntimos, con punto de miles)
function anEur_(n) {
  if (n == null || !isFinite(n)) return '—';
  var s = String(Math.round(Math.abs(n))).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  return (Math.round(n) < 0 ? '−' : '') + s + ' €';
}
function anPctTxt_(p) {
  if (p == null || !isFinite(p)) return '—';
  var r = Math.round(p);
  return (r > 0 ? '+' : r < 0 ? '−' : '') + Math.abs(r) + ' %';
}
// subeEsBueno = false para compras y descuadres (que suban es malo).
function anPctHtml_(p, subeEsBueno) {
  if (p == null || !isFinite(p)) return '<span style="color:' + AN_GRIS_ + '">—</span>';
  var r = Math.round(p), bueno = subeEsBueno === false ? r < 0 : r > 0;
  var color = r === 0 ? '#666' : bueno ? AN_VERDE_ : AN_ROJO_;
  return '<span style="color:' + color + ';font-weight:bold;white-space:nowrap">' + anPctTxt_(p) + '</span>';
}
// 'un +4 %' o 'igual' (para frases: '... que la semana anterior').
function anUn_(p) { return Math.round(p) === 0 ? 'igual' : 'un ' + anPctTxt_(p); }
function anDiasTxt_(n) { return String(n).replace('.', ','); }
function anSuma_(lista) { return lista.reduce(function (t, x) { return t + (Number(x) || 0); }, 0); }

// ---------- HTML del email ----------
function anH2_(titulo, explicacion) {
  return '<h2 style="margin:28px 0 4px;font-size:17px;color:#1a5276">' + titulo + '</h2>' +
    (explicacion ? '<p style="margin:0 0 8px;color:#777;font-size:12px">' + explicacion + '</p>' : '');
}
// filas: arrays de celdas (HTML). La primera columna a la izquierda, las demás a la derecha.
function anTabla_(cabecera, filas, resaltar) {
  var td = 'padding:6px 8px;border-bottom:1px solid #eee;';
  return '<table style="border-collapse:collapse;width:100%;font-size:13px">' +
    '<tr>' + cabecera.map(function (c, i) {
      return '<th style="' + td + 'color:#888;font-weight:normal;font-size:11px;text-align:' + (i ? 'right' : 'left') + '">' + c + '</th>';
    }).join('') + '</tr>' +
    filas.map(function (f, j) {
      var fondo = resaltar && resaltar(j) ? 'background:#fef9e7;' : '';
      return '<tr>' + f.map(function (c, i) {
        return '<td style="' + td + fondo + 'text-align:' + (i ? 'right' : 'left') + (i ? ';white-space:nowrap' : ';font-weight:bold') + '">' + c + '</td>';
      }).join('') + '</tr>';
    }).join('') + '</table>';
}
function anGrande_(importe, detalle) {
  return '<div style="margin:16px 0 4px;padding:14px 16px;background:#eaf2f8;border-radius:8px">' +
    '<div style="font-size:28px;font-weight:bold;color:#1a5276">' + importe + '</div>' +
    '<div style="font-size:13px;color:#555;margin-top:2px">' + detalle + '</div></div>';
}
function anDestacados_(lista) {
  if (!lista.length) return '';
  return '<div style="margin:12px 0;padding:10px 14px;border-left:4px solid #f39c12;background:#fef9e7">' +
    '<div style="font-weight:bold;margin-bottom:4px">Lo destacado</div>' +
    '<ul style="margin:0;padding-left:18px;font-size:14px;line-height:1.5">' + lista.map(function (x) { return '<li>' + x + '</li>'; }).join('') + '</ul></div>';
}
function anSinHtml_(s) { return String(s).replace(/<[^>]+>/g, ''); }

// ---------- datos ----------
// Día por día de la Contabilidad entre dos fechas (AAAA-MM-DD), de los años que
// tengan Contabilidad en el Índice. fecha -> { sistema, mediodia, noche,
// delivery (con web), empanadas, tarjetas, diferencia }.
function anDias_(desde, hasta) {
  var res = {}, libros = {};
  var y = Number(desde.slice(0, 4)), m = Number(desde.slice(5, 7)) - 1;
  var yh = Number(hasta.slice(0, 4)), mh = Number(hasta.slice(5, 7)) - 1;
  while (y * 12 + m <= yh * 12 + mh) {
    if (!(y in libros)) {
      var id = buscarEnIndice_('CONTABILIDAD', String(y));
      libros[y] = id ? SpreadsheetApp.openById(id) : null;
    }
    if (libros[y]) {
      var d = datosMesContabilidad_(libros[y], m);
      if (d) d.lista.forEach(function (x) { if (x.fecha >= desde && x.fecha <= hasta) res[x.fecha] = x; });
    }
    if (++m > 11) { m = 0; y++; }
  }
  return res;
}
function anFact_(x) { return (x.sistema || 0) + (x.delivery || 0); }
function anSumar_(dias, fechas, fn) {
  var t = 0, n = 0;
  fechas.forEach(function (f) { var x = dias[f]; if (x) { t += fn(x); n++; } });
  return { total: anR_(t), dias: n };
}

// Un mes de la pestaña "años" (mes 1-12), o null.
function anMes_(datos, anio, mes) {
  var fila = (datos.ventas[String(anio)] || []).filter(function (r) { return r.mes === mes; })[0];
  if (!fila) return null;
  return {
    fact: anR_((fila.sistema || 0) + (fila.delivery || 0) + (fila.web || 0)),
    local: fila.sistema || 0, md: fila.mediodia, nc: fila.noche, deliv: anR_((fila.delivery || 0) + (fila.web || 0)),
    emp: fila.empanadas || 0, dias: fila.dias || 0, prov: fila.proveedores, nominas: fila.nominas, extras: fila.extras,
    tarjetas: fila.tarjetas, diferencia: fila.diferencia
  };
}

// Lo facturado en esas fechas de un año pasado: real si hay día por día, si no
// aproximado (ver arriba). null si no hay datos de ese año.
function anOtroAnio_(datos, fechas) {
  var real = anSumar_(anDias_(fechas[0], fechas[fechas.length - 1]), fechas, anFact_);
  if (real.dias) return { total: real.total, dias: real.dias, aprox: false };
  var t = 0, n = 0;
  for (var i = 0; i < fechas.length; i++) {
    var iso = fechas[i], anio = iso.slice(0, 4), mes = Number(iso.slice(5, 7));
    var sem = datos.semana[anio] && datos.semana[anio][String(mes)];
    var mm = anMes_(datos, anio, mes);
    if (!sem || !mm || !mm.dias) return null;
    var local = sem.total[anDow_(iso)];
    if (!local) continue; // ese día de la semana no se abría
    t += local + mm.deliv / mm.dias;
    n++;
  }
  return { total: anR_(t), dias: n, aprox: true };
}
function anMoverAnios_(fechas, anios) { return fechas.map(function (f) { return semanalSumarDias_(f, -364 * anios); }); } // mismo día de la semana

// Compras de un mes (mes 1-12) según los albaranes; si no, PROVEEDORES de "años".
function anCompras_(datos, anio, mes) {
  var alb = datos.albaranes[String(anio)];
  if (alb) {
    var t = 0, hay = false;
    Object.keys(alb).forEach(function (p) { var v = alb[p][mes - 1] || 0; if (v) hay = true; t += v; });
    if (hay) return anR_(t);
  }
  var m = anMes_(datos, anio, mes);
  return m && m.prov ? m.prov : null;
}

// ============================================================================
// LA SEMANA (para el email de los lunes). hasta = último día (AAAA-MM-DD).
// compras = { total, media } de la semana (de extrasDeLaSemana_), o null.
// Devuelve { html, texto, facturado, cambio } o null si no hubo ventas.
// ============================================================================
function analisisSemana_(hasta, compras) {
  var datos = datosInformes_(false);
  var desde = semanalSumarDias_(hasta, -6);
  var dias = anDias_(semanalSumarDias_(desde, -56), hasta);
  var semana = function (k) { return anFechas_(semanalSumarDias_(desde, -7 * k), semanalSumarDias_(hasta, -7 * k)); };
  var fechas = semana(0);
  var esta = anSumar_(dias, fechas, anFact_);
  if (!esta.dias) return null;
  var anio = Number(hasta.slice(0, 4));
  var destacados = [], html = [];

  // --- comparaciones ---
  var ant = anSumar_(dias, semana(1), anFact_);
  var previas = [];
  for (var k = 1; k <= 4; k++) { var s = anSumar_(dias, semana(k), anFact_); if (s.dias) previas.push(s); }
  var media4 = previas.length ? anR_(anSuma_(previas.map(function (x) { return x.total; })) / previas.length) : null;
  var pasados = [1, 2, 3].map(function (a) { var r = anOtroAnio_(datos, anMoverAnios_(fechas, a)); if (r) r.anio = anio - a; return r; }).filter(Boolean);
  var porDia = function (s) { return s && s.dias ? s.total / s.dias : null; };

  var filas = [];
  if (ant.dias) filas.push(['Semana anterior', anEur_(ant.total), ant.dias, anEur_(porDia(ant)), anPctHtml_(anPct_(esta.total, ant.total))]);
  if (media4) filas.push(['Media de las ' + previas.length + ' semanas anteriores', anEur_(media4), '', '', anPctHtml_(anPct_(esta.total, media4))]);
  pasados.forEach(function (p) {
    filas.push(['Misma semana de ' + p.anio + (p.aprox ? ' <span style="color:#999;font-weight:normal">(aprox.)</span>' : ''), anEur_(p.total), p.dias, anEur_(porDia(p)), anPctHtml_(anPct_(esta.total, p.total))]);
  });

  var cambio = ant.dias ? anPct_(esta.total, ant.total) : null;
  destacados.push('Facturaste <b>' + anEur_(esta.total) + '</b>' + (cambio != null ? ', <b>' + anUn_(cambio) + '</b> que la semana anterior' : '') +
    (pasados[0] ? ' y <b>' + anUn_(anPct_(esta.total, pasados[0].total)) + '</b> que la misma semana de ' + pasados[0].anio : '') + '.');
  if (pasados.length >= 2) {
    var mejorAntes = pasados.reduce(function (a, b) { return a.total > b.total ? a : b; });
    if (esta.total > mejorAntes.total) destacados.push('Es mejor que esa misma semana en los ' + pasados.length + ' años anteriores.');
    else {
      var peores = pasados.filter(function (p) { return esta.total < p.total; }).map(function (p) { return p.anio; });
      if (peores.length === pasados.length) destacados.push('Está por debajo de esa misma semana en los ' + pasados.length + ' años anteriores (' + peores.join(', ') + ').');
    }
  }

  html.push(anH2_('Ventas de la semana', 'Del ' + anDDMM_(desde) + ' al ' + anDDMM_(hasta) + '. Facturado = local + delivery + web (las empanadas a otros negocios, aparte).'));
  html.push(anGrande_(anEur_(esta.total), esta.dias + ' días abiertos · ' + anEur_(porDia(esta)) + ' por día'));
  html.push(anTabla_(['Comparado con', 'Facturado', 'Días', 'Por día', 'Esta semana'], filas));
  if (pasados.some(function (p) { return p.aprox; })) {
    html.push('<p style="margin:4px 0 0;color:#999;font-size:11px">Aprox.: de esos años solo hay promedios por día de la semana de cada mes, no el día a día.</p>');
  }

  // --- desglose ---
  var partes = [
    ['Mediodía', function (x) { return x.mediodia || 0; }],
    ['Noche', function (x) { return x.noche || 0; }],
    ['Delivery y web', function (x) { return x.delivery || 0; }],
    ['Empanadas (aparte)', function (x) { return x.empanadas || 0; }]
  ];
  var filasDes = partes.map(function (p) {
    var a = anSumar_(dias, fechas, p[1]).total, b = anSumar_(dias, semana(1), p[1]).total;
    return [p[0], anEur_(a), p[0].indexOf('Empanadas') === 0 ? '' : Math.round(esta.total ? a / esta.total * 100 : 0) + ' %', ant.dias ? anEur_(b) : '—', ant.dias ? anPctHtml_(anPct_(a, b)) : '—'];
  }).filter(function (f, i) { return i < 3 || f[1] !== '0 €' || f[3] !== '0 €'; });
  html.push(anH2_('De dónde salió'));
  html.push(anTabla_(['', 'Esta semana', 'Del total', 'Semana anterior', 'Cambio'], filasDes));
  var tj = anSumar_(dias, fechas, function (x) { return x.tarjetas || 0; }).total, loc = anSumar_(dias, fechas, function (x) { return x.sistema || 0; }).total;
  var tjA = anSumar_(dias, semana(1), function (x) { return x.tarjetas || 0; }).total, locA = anSumar_(dias, semana(1), function (x) { return x.sistema || 0; }).total;
  if (loc && tj) html.push('<p style="margin:6px 0 0;font-size:13px;color:#555">Pagado con tarjeta: <b>' + Math.round(tj / loc * 100) + ' %</b> de lo del local' + (locA && tjA ? ' (la semana anterior, ' + Math.round(tjA / locA * 100) + ' %)' : '') + '.</p>');

  // --- día por día ---
  var mejor = null, peor = null, filasDia = [];
  fechas.forEach(function (f) {
    var x = dias[f];
    var antes = [];
    for (var k = 1; k <= 8; k++) { var y = dias[semanalSumarDias_(f, -7 * k)]; if (y) antes.push(anFact_(y)); }
    var suMedia = antes.length >= 2 ? anSuma_(antes) / antes.length : null;
    if (!x) { filasDia.push([anDiaCorto_(f), '<span style="color:#999">cerrado</span>', '', '', '', '']); return; }
    var v = anFact_(x), prev = dias[semanalSumarDias_(f, -7)];
    var vsMedia = suMedia ? anPct_(v, suMedia) : null;
    if (!mejor || v > mejor.v) mejor = { f: f, v: v };
    if (!peor || v < peor.v) peor = { f: f, v: v };
    if (vsMedia != null && vsMedia <= -25) destacados.push('El ' + AN_DIAS_[anDow_(f)].toLowerCase() + ' ' + anDDMM_(f) + ' fue flojo: ' + anEur_(v) + ', un ' + anPctTxt_(vsMedia) + ' que un ' + AN_DIAS_[anDow_(f)].toLowerCase() + ' normal (' + anEur_(suMedia) + ').');
    if (vsMedia != null && vsMedia >= 25) destacados.push('El ' + AN_DIAS_[anDow_(f)].toLowerCase() + ' ' + anDDMM_(f) + ' fue muy bueno: ' + anEur_(v) + ', un ' + anPctTxt_(vsMedia) + ' que un ' + AN_DIAS_[anDow_(f)].toLowerCase() + ' normal (' + anEur_(suMedia) + ').');
    filasDia.push([anDiaCorto_(f), anEur_(v), anEur_(x.mediodia), anEur_(x.noche), prev ? anPctHtml_(anPct_(v, anFact_(prev))) : '—', vsMedia != null ? anPctHtml_(vsMedia) : '—']);
  });
  html.push(anH2_('Día por día', '"Vs. su media": comparado con ese mismo día de la semana en las 8 semanas anteriores.'));
  html.push(anTabla_(['', 'Facturado', 'Mediodía', 'Noche', 'Vs. semana ant.', 'Vs. su media'], filasDia,
    function (j) { return mejor && fechas[j] === mejor.f; }));
  if (mejor && esta.dias > 1) destacados.push('El mejor día fue el <b>' + AN_DIAS_[anDow_(mejor.f)].toLowerCase() + ' ' + anDDMM_(mejor.f) + '</b> (' + anEur_(mejor.v) + ') y el más flojo el ' + AN_DIAS_[anDow_(peor.f)].toLowerCase() + ' ' + anDDMM_(peor.f) + ' (' + anEur_(peor.v) + ').');

  // --- el mes y el año en curso ---
  var mes = Number(hasta.slice(5, 7)), inicioMes = hasta.slice(0, 8) + '01';
  var fechasMes = anFechas_(inicioMes, hasta);
  var mesAct = anSumar_(dias, fechasMes, anFact_);
  var filasMes = [];
  if (mesAct.dias >= 5) { // con pocos días del mes no tiene sentido
    // Proyección: lo que queda del mes con la media de cada día de la semana
    // (8 semanas); un día de la semana que casi no se abrió cuenta como cerrado.
    var resto = anFechas_(semanalSumarDias_(hasta, 1), hasta.slice(0, 8) + anDosDig_(anUltimoDia_(anio, mes)));
    var proy = mesAct.total;
    resto.forEach(function (f) {
      var vals = [], vistas = 0;
      for (var k = 1; vistas < 8 && k <= 14; k++) {
        var g = semanalSumarDias_(f, -7 * k);
        if (g > hasta) continue;
        vistas++;
        if (dias[g]) vals.push(anFact_(dias[g]));
      }
      if (vals.length >= 4) proy += anSuma_(vals) / vals.length;
    });
    var mesPas = anMes_(datos, anio - 1, mes);
    [1, 2, 3].forEach(function (a) {
      var r = anOtroAnio_(datos, anMoverAnios_(fechasMes, a));
      if (r) filasMes.push([anCap_(AN_MESES_[mes - 1]) + ' ' + (anio - a) + ', mismos días' + (r.aprox ? ' <span style="color:#999;font-weight:normal">(aprox.)</span>' : ''), anEur_(r.total), anPctHtml_(anPct_(mesAct.total, r.total))]);
    });
    html.push(anH2_(anCap_(AN_MESES_[mes - 1]) + ' hasta el ' + anDDMM_(hasta),
      'Llevas ' + anEur_(mesAct.total) + ' en ' + mesAct.dias + ' días.' + (resto.length ? ' Si sigue como las últimas semanas, cerraría en unos <b>' + anEur_(proy) + '</b>' + (mesPas ? ' (' + AN_MESES_[mes - 1] + ' ' + (anio - 1) + ' cerró en ' + anEur_(mesPas.fact) + ')' : '') + '.' : '')));
    if (filasMes.length) html.push(anTabla_(['Comparado con', 'Facturado', anCap_(AN_MESES_[mes - 1]) + ' ' + anio], filasMes));
  }

  // Año: meses completos de "años" + lo que va del mes.
  var anioAct = mesAct.total;
  for (var m = 1; m < mes; m++) { var mm = anMes_(datos, anio, m); if (mm) anioAct += mm.fact; }
  var filasAnio = [], anioAnt1 = 0;
  [1, 2, 3].forEach(function (a) {
    var t = 0, ok = true;
    for (var m = 1; m < mes; m++) { var mm = anMes_(datos, anio - a, m); if (!mm) { ok = false; break; } t += mm.fact; }
    var parcial = mesAct.dias ? anOtroAnio_(datos, anMoverAnios_(fechasMes, a)) : { total: 0 };
    if (!ok || !parcial) return;
    t += parcial.total;
    if (a === 1) anioAnt1 = t;
    filasAnio.push([String(anio - a) + ', mismo periodo', anEur_(t), anPctHtml_(anPct_(anioAct, t))]);
  });
  if (filasAnio.length) {
    html.push(anH2_('El año hasta el ' + anDDMM_(hasta), 'Llevas ' + anEur_(anioAct) + ' en ' + anio + '.'));
    html.push(anTabla_(['Comparado con', 'Facturado', String(anio)], filasAnio));
    var a1 = filasAnio[0];
    if (a1 && anioAnt1) destacados.push('En el año llevas ' + anEur_(anioAct) + ', ' + anUn_(anPct_(anioAct, anioAnt1)) + ' que ' + (anio - 1) + ' a estas alturas.');
  }

  // --- compras ---
  if (compras && compras.total) {
    var ventasMedia = previas.length ? anSuma_(previas.map(function (x) { return x.total; })) / previas.length : 0;
    var pc = esta.total ? compras.total / esta.total * 100 : null, pcMedia = ventasMedia && compras.media ? compras.media / ventasMedia * 100 : null;
    html.push(anH2_('Compras de la semana', 'Facturas de proveedores cargadas con fecha de esta semana.'));
    html.push(anTabla_(['', 'Compras', 'Sobre lo facturado'], [
      ['Esta semana', anEur_(compras.total), pc != null ? Math.round(pc) + ' %' : '—'],
      ['Lo normal (media de 8 semanas)', anEur_(compras.media), pcMedia != null ? Math.round(pcMedia) + ' %' : '—']
    ]));
    if (pc != null && pcMedia != null && pc - pcMedia >= 5) destacados.push('Las compras fueron el <b>' + Math.round(pc) + ' %</b> de lo facturado (lo normal, ' + Math.round(pcMedia) + ' %).');
  }

  // --- caja ---
  var difs = fechas.filter(function (f) { return dias[f]; }).map(function (f) { return dias[f].diferencia || 0; });
  var difTotal = anSuma_(difs);
  if (Math.abs(difTotal) > 5) destacados.push('La caja de la semana descuadró ' + (difTotal > 0 ? '+' : '') + euros_(difTotal) + ' en total (los días, abajo).');

  var cuerpo = anDestacados_(destacados) + html.join('');
  return {
    html: cuerpo,
    texto: 'VENTAS DE LA SEMANA\n' + destacados.map(function (d) { return '• ' + anSinHtml_(d); }).join('\n'),
    facturado: esta.total,
    cambio: cambio
  };
}

// ============================================================================
// EL MES — email el día 1 con el mes anterior. mes 1-12.
// ============================================================================
function analisisMes_(anio, mes) {
  var datos = datosInformes_(true);
  var ini = anio + '-' + anDosDig_(mes) + '-01', fin = anio + '-' + anDosDig_(mes) + '-' + anDosDig_(anUltimoDia_(anio, mes));
  var dias = anDias_(ini, fin), fechas = anFechas_(ini, fin);
  var nombreMes = AN_MESES_[mes - 1], Mes = anCap_(nombreMes);
  var hoja = anMes_(datos, anio, mes);
  var abiertos = fechas.filter(function (f) { return dias[f]; });

  // Lo del mes: día por día si hay; si no, la fila de "años".
  var cur;
  if (abiertos.length) {
    var sum = function (fn) { return anSumar_(dias, fechas, fn).total; };
    cur = {
      fact: sum(anFact_), local: sum(function (x) { return x.sistema || 0; }), md: sum(function (x) { return x.mediodia || 0; }),
      nc: sum(function (x) { return x.noche || 0; }), deliv: sum(function (x) { return x.delivery || 0; }),
      emp: sum(function (x) { return x.empanadas || 0; }), tarjetas: sum(function (x) { return x.tarjetas || 0; }),
      diferencia: sum(function (x) { return x.diferencia || 0; }), dias: hoja && hoja.dias ? hoja.dias : abiertos.length
    };
    if (hoja) { cur.nominas = hoja.nominas; cur.extras = hoja.extras; }
  } else if (hoja) {
    cur = hoja;
  } else {
    return { asunto: 'Análisis de ' + nombreMes + ' ' + anio + ': sin datos', html: '<p>No hay ventas de ' + nombreMes + ' ' + anio + ' en la Contabilidad.</p>', texto: 'Sin datos.' };
  }
  var porDia = function (m) { return m && m.dias ? m.fact / m.dias : null; };
  var destacados = [], html = [];

  // --- comparaciones del total ---
  var mAnt = mes === 1 ? anMes_(datos, anio - 1, 12) : anMes_(datos, anio, mes - 1);
  var pasados = [1, 2, 3].map(function (a) { var m = anMes_(datos, anio - a, mes); if (m) m.anio = anio - a; return m; }).filter(Boolean);
  var filas = [];
  if (mAnt) filas.push(['Mes anterior (' + AN_MESES_[(mes + 10) % 12] + ')', anEur_(mAnt.fact), anPctHtml_(anPct_(cur.fact, mAnt.fact)), anEur_(porDia(mAnt)), anPctHtml_(anPct_(porDia(cur), porDia(mAnt)))]);
  pasados.forEach(function (p) { filas.push([Mes + ' ' + p.anio, anEur_(p.fact), anPctHtml_(anPct_(cur.fact, p.fact)), anEur_(porDia(p)), anPctHtml_(anPct_(porDia(cur), porDia(p)))]); });
  if (pasados.length >= 2) {
    var med = anSuma_(pasados.map(function (p) { return p.fact; })) / pasados.length;
    var medDia = anSuma_(pasados.map(function (p) { return porDia(p) || 0; })) / pasados.length;
    filas.push(['Media de esos ' + pasados.length + ' años', anEur_(med), anPctHtml_(anPct_(cur.fact, med)), anEur_(medDia), anPctHtml_(anPct_(porDia(cur), medDia))]);
  }
  // Ranking de ese mes en todos los años, y de todos los meses de la historia.
  var historia = [], todos = [];
  Object.keys(datos.ventas).forEach(function (a) {
    datos.ventas[a].forEach(function (r) {
      var f = (Number(a) === anio && r.mes === mes) ? cur.fact : (r.sistema || 0) + (r.delivery || 0) + (r.web || 0);
      todos.push({ anio: Number(a), mes: r.mes, fact: f });
      if (r.mes === mes) historia.push({ anio: Number(a), fact: f });
    });
  });
  if (!historia.some(function (h) { return h.anio === anio; })) { historia.push({ anio: anio, fact: cur.fact }); todos.push({ anio: anio, mes: mes, fact: cur.fact }); }
  historia.sort(function (a, b) { return b.fact - a.fact; });
  var puesto = historia.map(function (h) { return h.anio; }).indexOf(anio) + 1;

  var cambioAnio = pasados[0] ? anPct_(cur.fact, pasados[0].fact) : null;
  destacados.push('Facturaste <b>' + anEur_(cur.fact) + '</b> en ' + anDiasTxt_(cur.dias) + ' días' +
    (cambioAnio != null ? ', <b>' + anUn_(cambioAnio) + '</b> que ' + nombreMes + ' de ' + pasados[0].anio : '') +
    (pasados[0] && porDia(pasados[0]) ? ' (por día trabajado, ' + anPctTxt_(anPct_(porDia(cur), porDia(pasados[0]))) + ')' : '') + '.');
  if (historia.length > 2) {
    destacados.push(puesto === 1 ? '<b>Es el mejor ' + nombreMes + ' de la historia</b> (desde ' + historia.reduce(function (m, h) { return Math.min(m, h.anio); }, 9999) + ').'
      : 'Es el ' + puesto + 'º mejor ' + nombreMes + ' de ' + historia.length + ' (el mejor fue ' + historia[0].anio + ', con ' + anEur_(historia[0].fact) + ').');
  }
  todos.sort(function (a, b) { return b.fact - a.fact; });
  if (todos[0] && todos[0].anio === anio && todos[0].mes === mes) destacados.push('<b>¡Es el mes que más se facturó en toda la historia!</b>');

  html.push(anH2_('Facturado en ' + nombreMes, 'Local + delivery + web. Las empanadas a otros negocios, aparte' + (cur.emp ? ' (' + anEur_(cur.emp) + ' este mes)' : '') + '.'));
  html.push(anGrande_(anEur_(cur.fact), anDiasTxt_(cur.dias) + ' días trabajados · ' + anEur_(porDia(cur)) + ' por día' + (historia.length > 2 ? ' · ' + puesto + 'º de ' + historia.length + ' ' + nombreMes + 's' : '')));
  html.push(anTabla_(['Comparado con', 'Facturado', Mes + ' ' + anio, 'Por día', 'Por día ' + anio], filas));

  // --- desglose vs mismo mes del año pasado ---
  var p1 = pasados[0];
  var partes = [['Mediodía', 'md'], ['Noche', 'nc'], ['Delivery y web', 'deliv'], ['Empanadas (aparte)', 'emp']];
  var filasDes = partes.map(function (p) {
    var a = cur[p[1]] || 0, b = p1 ? p1[p[1]] : null;
    var cuota = p[1] === 'emp' ? '' : (cur.fact ? Math.round(a / cur.fact * 100) + ' %' : '');
    var cuotaB = p[1] === 'emp' || !p1 || b == null ? '' : (p1.fact ? Math.round(b / p1.fact * 100) + ' %' : '');
    return [p[0], anEur_(a), cuota, b != null ? anEur_(b) : '—', cuotaB, b ? anPctHtml_(anPct_(a, b)) : '—'];
  }).filter(function (f) { return f[1] !== '0 €' || (f[3] !== '—' && f[3] !== '0 €'); });
  html.push(anH2_('De dónde salió'));
  html.push(anTabla_(['', String(anio), 'Del total', p1 ? String(p1.anio) : '', p1 ? 'Del total' : '', 'Cambio'], filasDes));
  if (p1 && p1.deliv && cur.fact && p1.fact) {
    var cd = cur.deliv / cur.fact * 100, cdA = p1.deliv / p1.fact * 100;
    if (Math.abs(cd - cdA) >= 3) destacados.push('El delivery pesa el ' + Math.round(cd) + ' % de lo facturado (en ' + p1.anio + ', el ' + Math.round(cdA) + ' %).');
  }
  if (p1 && p1.md && p1.nc && cur.md && cur.nc) {
    var cmd = anPct_(cur.md, p1.md), cnc = anPct_(cur.nc, p1.nc);
    if (Math.abs(cmd - cnc) >= 10) destacados.push('El mediodía ' + (cmd > cnc ? 'fue mejor' : 'fue peor') + ' que la noche: mediodía ' + anPctTxt_(cmd) + ', noche ' + anPctTxt_(cnc) + ' que en ' + p1.anio + '.');
  }

  // --- por día de la semana (solo el local, como la pestaña "comparativa") ---
  if (abiertos.length) {
    var suma = [0, 0, 0, 0, 0, 0, 0], cuenta = [0, 0, 0, 0, 0, 0, 0];
    abiertos.forEach(function (f) { var w = anDow_(f); suma[w] += dias[f].sistema || 0; cuenta[w]++; });
    var anios = [anio - 1, anio - 2].filter(function (a) { return datos.semana[String(a)] && datos.semana[String(a)][String(mes)]; });
    var mejorDia = -1, mayorCambio = null;
    var filasSem = AN_DIAS_.map(function (n, w) {
      var v = cuenta[w] ? suma[w] / cuenta[w] : null;
      if (v != null && (mejorDia < 0 || v > suma[mejorDia] / cuenta[mejorDia])) mejorDia = w;
      var fila = [n + (cuenta[w] ? ' <span style="color:#999;font-weight:normal">(' + cuenta[w] + ')</span>' : ''), v != null ? anEur_(v) : '<span style="color:#999">cerrado</span>'];
      anios.forEach(function (a, i) {
        var b = datos.semana[String(a)][String(mes)].total[w];
        fila.push(b ? anEur_(b) : '—');
        if (i === 0) {
          fila.push(v != null && b ? anPctHtml_(anPct_(v, b)) : '—');
          if (v != null && b) { var c = anPct_(v, b); if (!mayorCambio || Math.abs(c) > Math.abs(mayorCambio.c)) mayorCambio = { w: w, c: c }; }
        }
      });
      return fila;
    });
    var cab = ['', anio + ' (promedio)'];
    anios.forEach(function (a, i) { cab.push(String(a)); if (i === 0) cab.push('Cambio'); });
    html.push(anH2_('Por día de la semana', 'Promedio de lo facturado en el local cada día (sin delivery). Entre paréntesis, cuántos hubo en el mes.'));
    html.push(anTabla_(cab, filasSem));
    if (mayorCambio && Math.abs(mayorCambio.c) >= 15) destacados.push('Los ' + AN_DIAS_[mayorCambio.w].toLowerCase() + (/s$/.test(AN_DIAS_[mayorCambio.w]) ? '' : 's') + ' son los que más cambiaron: ' + anPctTxt_(mayorCambio.c) + ' que en ' + nombreMes + ' de ' + anios[0] + '.');

    // Mejores y peores días.
    var orden = abiertos.slice().sort(function (a, b) { return anFact_(dias[b]) - anFact_(dias[a]); });
    var top = orden.slice(0, 3), bottom = orden.length > 6 ? orden.slice(-3).reverse() : [];
    html.push(anH2_('Mejores y peores días'));
    html.push(anTabla_(['', 'Facturado', 'Mediodía', 'Noche', 'Delivery'],
      top.map(function (f) { return ['▲ ' + anDiaCorto_(f), anEur_(anFact_(dias[f])), anEur_(dias[f].mediodia), anEur_(dias[f].noche), anEur_(dias[f].delivery)]; })
        .concat(bottom.map(function (f) { return ['▼ ' + anDiaCorto_(f), anEur_(anFact_(dias[f])), anEur_(dias[f].mediodia), anEur_(dias[f].noche), anEur_(dias[f].delivery)]; }))));
  }

  // --- el año hasta este mes, y los últimos 12 meses ---
  var acumula = function (a, hastaMes) {
    var t = 0;
    for (var m = 1; m <= hastaMes; m++) {
      if (a === anio && m === mes) { t += cur.fact; continue; }
      var mm = anMes_(datos, a, m);
      if (!mm) return null;
      t += mm.fact;
    }
    return t;
  };
  var anioAct = acumula(anio, mes);
  if (anioAct != null) {
    var filasA = [];
    [1, 2, 3].forEach(function (a) { var t = acumula(anio - a, mes); if (t != null) filasA.push([String(anio - a), anEur_(t), anPctHtml_(anPct_(anioAct, t))]); });
    // Últimos 12 meses contra los 12 anteriores.
    var doce = function (corrimiento) {
      var t = 0;
      for (var i = 0; i < 12; i++) {
        var n = anio * 12 + (mes - 1) - i - corrimiento, a = Math.floor(n / 12), m = n % 12 + 1;
        if (a === anio && m === mes) { t += cur.fact; continue; }
        var mm = anMes_(datos, a, m);
        if (!mm) return null;
        t += mm.fact;
      }
      return t;
    };
    var u12 = doce(0), p12 = doce(12);
    html.push(anH2_('El año', 'De enero a ' + nombreMes + '.'));
    html.push(anGrande_(anEur_(anioAct), 'facturado en ' + anio + ' hasta ' + nombreMes));
    if (filasA.length) html.push(anTabla_(['Mismo periodo de', 'Facturado', String(anio)], filasA));
    if (u12 != null && p12 != null) {
      html.push('<p style="margin:8px 0 0;font-size:13px;color:#555">Últimos 12 meses: <b>' + anEur_(u12) + '</b>, ' + anPctHtml_(anPct_(u12, p12)) + ' que los 12 anteriores (' + anEur_(p12) + ').</p>');
    }
    if (filasA[0]) destacados.push('En el año llevas ' + anEur_(anioAct) + ', ' + anUn_(anPct_(anioAct, acumula(anio - 1, mes))) + ' que en ' + (anio - 1) + ' a estas alturas.');
  }

  // --- compras ---
  var comp = anCompras_(datos, anio, mes);
  if (comp) {
    var filasC = [[Mes + ' ' + anio, anEur_(comp), Math.round(comp / cur.fact * 100) + ' %', '']];
    pasados.forEach(function (p) {
      var c = anCompras_(datos, p.anio, mes);
      if (c) filasC.push([Mes + ' ' + p.anio, anEur_(c), Math.round(c / p.fact * 100) + ' %', anPctHtml_(anPct_(comp, c), false)]);
    });
    html.push(anH2_('Compras a proveedores', 'Lo de las facturas de los albaranes, y qué parte de lo facturado se fue en compras.'));
    html.push(anTabla_(['', 'Compras', 'Sobre lo facturado', 'Cambio de ' + anio], filasC));
    if (filasC[1]) {
      var pcAct = comp / cur.fact * 100, pcAnt = anCompras_(datos, pasados[0].anio, mes) / pasados[0].fact * 100;
      if (Math.abs(pcAct - pcAnt) >= 3) destacados.push('Las compras fueron el <b>' + Math.round(pcAct) + ' %</b> de lo facturado (en ' + nombreMes + ' de ' + pasados[0].anio + ', el ' + Math.round(pcAnt) + ' %).');
    }

    // Por proveedor: los que más se llevaron y los que más cambiaron contra el año pasado.
    var alb = datos.albaranes[String(anio)] || {}, albA = datos.albaranes[String(anio - 1)] || {};
    var provs = Object.keys(alb).map(function (n) {
      var v = alb[n][mes - 1] || 0, va = albA[n] ? (albA[n][mes - 1] || 0) : null;
      return { n: n, v: v, va: va };
    }).filter(function (p) { return p.v > 0; }).sort(function (a, b) { return b.v - a.v; });
    if (provs.length) {
      html.push(anH2_('Proveedores del mes', 'Los que más se llevaron, comparados con ' + nombreMes + ' de ' + (anio - 1) + '.'));
      html.push(anTabla_(['', String(anio), String(anio - 1), 'Cambio'], provs.slice(0, 10).map(function (p) {
        return [anEsc_(p.n), anEur_(p.v), p.va ? anEur_(p.va) : '—', p.va ? anPctHtml_(anPct_(p.v, p.va), false) : '—'];
      })));
      var cambios = Object.keys(alb).concat(Object.keys(albA).filter(function (n) { return !alb[n]; })).map(function (n) {
        var v = alb[n] ? (alb[n][mes - 1] || 0) : 0, va = albA[n] ? (albA[n][mes - 1] || 0) : 0;
        return { n: n, v: v, va: va, d: v - va };
      }).filter(function (p) { return p.va > 0 && p.v > 0 && Math.abs(p.d) >= 150 && Math.abs(p.d) / p.va >= 0.2; });
      var suben = cambios.filter(function (p) { return p.d > 0; }).sort(function (a, b) { return b.d - a.d; }).slice(0, 5);
      var bajan = cambios.filter(function (p) { return p.d < 0; }).sort(function (a, b) { return a.d - b.d; }).slice(0, 5);
      if (suben.length || bajan.length) {
        html.push(anH2_('Los que más cambiaron', 'Contra ' + nombreMes + ' de ' + (anio - 1) + ' (más de 150 € y más de un 20 %).'));
        html.push(anTabla_(['', String(anio), String(anio - 1), 'Diferencia'],
          suben.map(function (p) { return ['▲ ' + anEsc_(p.n), anEur_(p.v), anEur_(p.va), '<span style="color:' + AN_ROJO_ + ';font-weight:bold">+' + anEur_(p.d) + '</span>']; })
            .concat(bajan.map(function (p) { return ['▼ ' + anEsc_(p.n), anEur_(p.v), anEur_(p.va), '<span style="color:' + AN_VERDE_ + ';font-weight:bold">' + anEur_(p.d) + '</span>']; }))));
        if (suben[0]) destacados.push('El proveedor que más subió: <b>' + anEsc_(suben[0].n) + '</b>, ' + anEur_(suben[0].v) + ' (' + anEur_(suben[0].va) + ' en ' + nombreMes + ' de ' + (anio - 1) + ').');
      }
    }
  }

  // --- lo que queda después de compras y nóminas ---
  if (comp && cur.nominas) {
    var queda = function (fact, emp, c, nom, ext) { return fact + (emp || 0) - c - nom - (ext || 0); };
    var q = queda(cur.fact, cur.emp, comp, cur.nominas, cur.extras);
    var filasQ = [['Facturado (con empanadas)', anEur_(cur.fact + (cur.emp || 0))], ['− Compras', anEur_(comp)], ['− Nóminas', anEur_(cur.nominas)]];
    if (cur.extras) filasQ.push(['− Extras', anEur_(cur.extras)]);
    filasQ.push(['<span style="font-size:15px">Queda</span>', '<b style="font-size:15px">' + anEur_(q) + '</b>']);
    var cA = p1 ? anCompras_(datos, p1.anio, mes) : null, extra = '';
    if (p1 && cA && p1.nominas) {
      var qA = queda(p1.fact, p1.emp, cA, p1.nominas, p1.extras);
      extra = ' En ' + nombreMes + ' de ' + p1.anio + ' quedaron ' + anEur_(qA) + ' (' + anPctTxt_(anPct_(q, qA)) + ').';
    }
    html.push(anH2_('Después de compras y nóminas', 'Sin alquiler, luz, impuestos ni otros gastos: es solo una idea.' + extra));
    html.push(anTabla_(['', ''], filasQ));
  }

  // --- caja ---
  if (abiertos.length) {
    var descuadres = abiertos.filter(function (f) { return Math.abs(dias[f].diferencia || 0) > 5; });
    var peorDif = descuadres.slice().sort(function (a, b) { return Math.abs(dias[b].diferencia) - Math.abs(dias[a].diferencia); })[0];
    var tjPc = cur.local && cur.tarjetas ? cur.tarjetas / cur.local * 100 : null;
    var tjPcA = p1 && p1.local && p1.tarjetas ? p1.tarjetas / p1.local * 100 : null;
    var lineas = [];
    lineas.push('Descuadre del mes: <b>' + (cur.diferencia > 0 ? '+' : '') + euros_(cur.diferencia) + '</b>' + (descuadres.length ? ' · ' + descuadres.length + ' días descuadraron más de 5 €' + (peorDif ? ' (el peor, el ' + anDDMM_(peorDif) + ': ' + euros_(dias[peorDif].diferencia) + ')' : '') : ' · ningún día descuadró más de 5 €') + '.');
    if (tjPc != null) lineas.push('Pagado con tarjeta: <b>' + Math.round(tjPc) + ' %</b> de lo del local' + (tjPcA != null ? ' (en ' + p1.anio + ', ' + Math.round(tjPcA) + ' %)' : '') + '.');
    html.push(anH2_('Caja'));
    html.push('<ul style="margin:0;padding-left:18px;font-size:13px;line-height:1.6">' + lineas.map(function (l) { return '<li>' + l + '</li>'; }).join('') + '</ul>');
    if (Math.abs(cur.diferencia) > 50) destacados.push('La caja descuadró ' + (cur.diferencia > 0 ? '+' : '') + euros_(cur.diferencia) + ' en el mes.');
  }

  var asunto = 'Análisis de ' + nombreMes + ' ' + anio + ': ' + anEur_(cur.fact) + (cambioAnio != null ? ' (' + anPctTxt_(cambioAnio) + ' que ' + pasados[0].anio + ')' : '');
  return {
    asunto: asunto,
    html: '<div style="font-family:Arial,Helvetica,sans-serif;max-width:640px;color:#222">' +
      '<h1 style="font-size:20px;margin:0 0 4px">' + Mes + ' ' + anio + '</h1>' +
      '<p style="margin:0;color:#888;font-size:13px">Análisis del mes, comparado con el mes anterior y con los años anteriores.</p>' +
      anDestacados_(destacados) + html.join('') +
      '<p style="margin-top:32px;color:#aaa;font-size:11px">Llega el día 1 de cada mes. Para preguntas concretas, la página de informes ("Pregúntale a tus datos").</p></div>',
    texto: Mes.toUpperCase() + ' ' + anio + '\n' + destacados.map(function (d) { return '• ' + anSinHtml_(d); }).join('\n')
  };
}

// Lo que corre el día 1 a las 9:00: el mes anterior.
function mandarAnalisisMes() {
  var hoy = new Date();
  var anio = Number(Utilities.formatDate(hoy, 'Europe/Madrid', 'yyyy')), mes = Number(Utilities.formatDate(hoy, 'Europe/Madrid', 'M')) - 1;
  if (mes === 0) { mes = 12; anio--; }
  var r = analisisMes_(anio, mes);
  MailApp.sendEmail({ to: Session.getEffectiveUser().getEmail(), subject: r.asunto, htmlBody: r.html, body: r.texto });
}

// Claude: ?muestraAnalisisMes=AAAA-MM manda una muestra (y deja puesto el activador).
function mandarMuestraAnalisisMes_(aaaamm) {
  asegurarActivadorAnalisisMes_();
  var r = analisisMes_(Number(aaaamm.slice(0, 4)), Number(aaaamm.slice(5, 7)));
  MailApp.sendEmail({ to: Session.getEffectiveUser().getEmail(), subject: 'MUESTRA — ' + r.asunto, htmlBody: r.html, body: r.texto });
  return { ok: true, asunto: r.asunto };
}

function asegurarActivadorAnalisisMes_() {
  var hay = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'mandarAnalisisMes'; });
  if (!hay) ScriptApp.newTrigger('mandarAnalisisMes').timeBased().onMonthDay(1).atHour(9).nearMinute(0).inTimezone('Europe/Madrid').create();
}
