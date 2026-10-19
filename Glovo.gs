// ============================================================================
// GLOVO — los tickets de Glovo que se cargan en el paso 3 del cierre de cada
// turno se escriben también en la planilla anual de Glovo (tipo GLOVO del
// Índice, ver Indice.gs). Esa planilla NO se crea sola: tiene que estar
// cargada a mano en la pestaña "Archivos" del Índice (GLOVO | 2026 | <ID>).
//
// Forma de la planilla (la de siempre): una pestaña por mes (ENERO, FEBRERO,
// ... DICIEMBRE). En cada una, fila 4 = encabezados y desde la fila 5 los
// pedidos, hasta la fila 206 (la 207 tiene la suma):
//   B  FECHA
//   C  nº PEDIDO
//   D  IMPORTE
//   E  comisión — fórmula, no se toca
//   F  TOTAL — fórmula, no se toca
// La app solo escribe B, C y D. Los ticket nuevos van debajo del último
// pedido cargado (a mano o por la app).
//
// La columna Z (oculta) guarda el ID del ticket de la app, para poder
// actualizarlo o borrarlo sin duplicar: la app manda el día entero en cada
// guardado. Las filas cargadas a mano (sin ID en la columna Z) no se tocan.
// Un ticket borrado en la app deja su fila vacía (B, C, D y Z en blanco).
// ============================================================================

var GLOVO_FILA_DESDE_ = 5;
var GLOVO_FILA_HASTA_ = 206;
var GLOVO_COL_ID_ = 26; // Z

// Tickets de Glovo del día (Mediodía + Noche), tal como los manda la app o
// como quedan en "Datos JSON" de Registro: [{id, pedido, importe}].
function ticketsGlovoDelDia_(dia) {
  var tickets = [];
  ['mediodia', 'noche'].forEach(function (k) {
    var t = dia && dia[k];
    if (!t) return;
    (t.glovo || []).forEach(function (g) {
      // Una línea con nº de pedido pero sin importe todavía no es un ticket.
      if (g && g.id && Number(g.importe) > 0) tickets.push(g);
    });
  });
  return tickets;
}

function totalGlovoDelDia_(dia) {
  var total = 0;
  ticketsGlovoDelDia_(dia).forEach(function (g) { total += Number(g.importe) || 0; });
  return Math.round(total * 100) / 100;
}

// Deja la pestaña del mes de la planilla de Glovo igual a lo que tiene la
// app para la caja de `fechaISO`. `tickets` vacío = borrar lo que la app
// había escrito de ese día.
function sincronizarGlovoDelDia_(fechaISO, tickets) {
  var anio = periodoAnual_(fechaISO);
  var idPlanilla = buscarEnIndice_('GLOVO', anio);
  if (!idPlanilla) {
    return tickets.length
      ? { ok: false, error: 'No hay planilla de Glovo ' + anio + ' en la pestaña "Archivos" del Índice.' }
      : { ok: true, nada: 'No hay planilla de Glovo ' + anio + '.' };
  }

  var mes = parseInt(String(fechaISO).split('-')[1], 10);
  var ss = SpreadsheetApp.openById(idPlanilla);
  var hoja = ss.getSheetByName(MESES_MAYUS_[mes - 1]);
  if (!hoja) return { ok: false, error: 'La planilla de Glovo no tiene la pestaña ' + MESES_MAYUS_[mes - 1] + '.' };

  var cantidad = GLOVO_FILA_HASTA_ - GLOVO_FILA_DESDE_ + 1;
  var datos = hoja.getRange(GLOVO_FILA_DESDE_, 2, cantidad, 3).getValues(); // B:D
  var formulas = hoja.getRange(GLOVO_FILA_DESDE_, 2, cantidad, 3).getFormulas();
  var ids = hoja.getRange(GLOVO_FILA_DESDE_, GLOVO_COL_ID_, cantidad, 1).getValues();

  // Los pedidos van solo hasta antes de la primera fila con una fórmula en
  // B:D (ej. la suma "=SUM(C5:C204)" de la fila 205): esa fila y las de
  // abajo no son para pedidos.
  var libres = cantidad;
  for (var f = 0; f < cantidad; f++) {
    if (formulas[f][0] || formulas[f][1] || formulas[f][2]) { libres = f; break; }
  }

  // Qué fila tiene cada ticket que la app ya había escrito de este día.
  var filaDeId = {};
  for (var i = 0; i < cantidad; i++) {
    var id = String(ids[i][0] || '');
    if (id && textoFechaIds_(datos[i][0]) === fechaISO) filaDeId[id] = i;
  }

  var actuales = {};
  tickets.forEach(function (g) { actuales[g.id] = true; });

  // Solo se escriben las filas que cambian: las cargadas a mano quedan tal cual.
  var cambiadas = {};

  // Los que ya no están en la app, o que quedaron fuera de la zona de
  // pedidos (debajo de la suma): se vacía su fila (y los segundos se vuelven
  // a escribir en su sitio).
  Object.keys(filaDeId).forEach(function (id) {
    if (actuales[id] && filaDeId[id] < libres) return;
    var i = filaDeId[id];
    if (actuales[id]) delete filaDeId[id]; // se vuelve a escribir más abajo, como nuevo
    datos[i] = ['', '', ''];
    ids[i] = [''];
    cambiadas[i] = true;
  });

  // La primera fila libre es la de debajo del último pedido cargado.
  var ultima = -1;
  for (var j = 0; j < libres; j++) {
    if (datos[j][0] !== '' || datos[j][1] !== '' || datos[j][2] !== '' || ids[j][0] !== '') ultima = j;
  }

  var fecha = fechaComoDate_(fechaISO);
  var sinLugar = 0;
  tickets.forEach(function (g) {
    var fila = [fecha, String(g.pedido || ''), Number(g.importe) || 0];
    var i = filaDeId[g.id];
    if (i == null) {
      if (ultima + 1 >= libres) { sinLugar++; return; }
      i = ++ultima;
    }
    datos[i] = fila;
    ids[i] = [g.id];
    cambiadas[i] = true;
  });

  Object.keys(cambiadas).forEach(function (k) {
    var r = GLOVO_FILA_DESDE_ + Number(k);
    hoja.getRange(r, 3).setNumberFormat('@'); // nº PEDIDO como texto (ej. "0DB3B")
    hoja.getRange(r, 2, 1, 3).setValues([datos[k]]);
    hoja.getRange(r, GLOVO_COL_ID_).setValue(ids[k][0]);
  });
  if (Object.keys(cambiadas).length && !hoja.isColumnHiddenByUser(GLOVO_COL_ID_)) hoja.hideColumns(GLOVO_COL_ID_);

  var salida = { ok: sinLugar === 0, escritos: tickets.length - sinLugar };
  if (sinLugar) salida.error = 'La pestaña ' + hoja.getName() + ' de Glovo está llena: no entraron ' + sinLugar + ' tickets.';
  return salida;
}

// ---------- Totales por día (columnas I y L, filas 14 a 44) ----------
// H = el día, I = suma de TOTAL (F) de ese día, L = suma de IMPORTE (D).
// Las fórmulas de antes (SUMIF con "=1/10/2026") solo sumaban las fechas
// escritas a mano: las que escribe la app llevan hora y no eran "iguales".
// Ahora se suma lo que cae entre ese día y el siguiente. Desde octubre de
// 2026 (las pestañas de antes se dejan como estaban), una vez por pestaña.
// También se borran las filas escritas a mano que repiten un pedido que ya
// escribió la app (mismo día, nº de pedido e importe): se copiaron a mano
// porque la fórmula vieja no contaba las de la app, y ahora sumarían doble.
var GLOVO_TOTALES_FILA_ = 14;
function arreglarTotalesGlovo_(ss, anio) {
  var props = PropertiesService.getScriptProperties();
  var hechos = [];
  for (var mes = 10; mes <= 12; mes++) {
    if (anio < 2026) break;
    var hoja = ss.getSheetByName(MESES_MAYUS_[mes - 1]);
    var marca = 'glovo_totales_v3_' + ss.getId() + '_' + mes;
    if (!hoja || props.getProperty(marca)) continue;
    if (normalizarClave_(hoja.getRange('H' + (GLOVO_TOTALES_FILA_ - 1)).getValue()) !== 'TOTALES') continue; // otra forma: no se toca
    var diasDelMes = new Date(anio, mes, 0).getDate();
    var filasI = [], filasL = [];
    for (var d = 1; d <= 31; d++) {
      if (d > diasDelMes) { filasI.push(['']); filasL.push(['']); continue; }
      var desde = 'DATE(' + anio + ',' + mes + ',' + d + ')';
      var rango = '$B$' + GLOVO_FILA_DESDE_ + ':$B$252,">="&' + desde + ',$B$' + GLOVO_FILA_DESDE_ + ':$B$252,"<"&(' + desde + '+1)';
      filasI.push(['=SUMIFS($F$' + GLOVO_FILA_DESDE_ + ':$F$252,' + rango + ')']);
      filasL.push(['=SUMIFS($D$' + GLOVO_FILA_DESDE_ + ':$D$252,' + rango + ')']);
    }
    var f0 = GLOVO_TOTALES_FILA_, f1 = GLOVO_TOTALES_FILA_ + 30;
    hoja.getRange('I' + f0 + ':I' + f1).setValues(filasLocales_(hoja, filasI));
    hoja.getRange('L' + f0 + ':L' + f1).setValues(filasLocales_(hoja, filasL));
    // J y M: del 1 al 15 (en la fila del día 15) y del 16 a fin de mes (en
    // la del día 31).
    var f15 = f0 + 14;
    hoja.getRange('J' + f15).setFormula('=SUM(I' + f0 + ':I' + f15 + ')');
    hoja.getRange('M' + f15).setFormula('=SUM(L' + f0 + ':L' + f15 + ')');
    hoja.getRange('J' + f1).setFormula('=SUM(I' + (f15 + 1) + ':I' + f1 + ')');
    hoja.getRange('M' + f1).setFormula('=SUM(L' + (f15 + 1) + ':L' + f1 + ')');
    var quitados = quitarRepetidosAManoGlovo_(hoja);
    props.setProperty(marca, '1');
    hechos.push(hoja.getName() + (quitados.length ? ' (quitadas filas repetidas a mano: ' + quitados.join(', ') + ')' : ''));
  }
  return hechos;
}
function quitarRepetidosAManoGlovo_(hoja) {
  var n = GLOVO_FILA_HASTA_ - GLOVO_FILA_DESDE_ + 1;
  var datos = hoja.getRange(GLOVO_FILA_DESDE_, 2, n, 3).getValues();
  var formulas = hoja.getRange(GLOVO_FILA_DESDE_, 2, n, 3).getFormulas();
  var ids = hoja.getRange(GLOVO_FILA_DESDE_, GLOVO_COL_ID_, n, 1).getValues();
  var zona = hoja.getParent().getSpreadsheetTimeZone();
  function clave(i) {
    var f = datos[i][0];
    var dia = f instanceof Date ? Utilities.formatDate(f, zona, 'yyyy-MM-dd') : String(f || '');
    return dia + '|' + String(datos[i][1] || '').trim() + '|' + (Math.round((Number(datos[i][2]) || 0) * 100) / 100);
  }
  var deLaApp = {};
  for (var i = 0; i < n; i++) if (ids[i][0] && datos[i][1] !== '') deLaApp[clave(i)] = true;
  var quitados = [];
  for (var j = 0; j < n; j++) {
    if (ids[j][0] || formulas[j].some(function (x) { return x; })) continue;
    if (datos[j][0] === '' || datos[j][1] === '' || !deLaApp[clave(j)]) continue;
    hoja.getRange(GLOVO_FILA_DESDE_ + j, 2, 1, 3).clearContent();
    quitados.push(String(GLOVO_FILA_DESDE_ + j));
  }
  return quitados;
}

// Desde el menú del Índice, para aplicarlo ya sin esperar a un guardado.
function arreglarTotalesGlovoDesdeMenu() {
  var ui = SpreadsheetApp.getUi();
  var anio = periodoAnual_(hoyISO_());
  var id = buscarEnIndice_('GLOVO', anio);
  if (!id) { ui.alert('No hay planilla de Glovo ' + anio + ' en la pestaña "Archivos" del Índice.'); return; }
  var hechos = arreglarTotalesGlovo_(SpreadsheetApp.openById(id), Number(anio));
  ui.alert('Glovo', hechos.length ? 'Arreglado: ' + hechos.join('; ') + '.' : 'Ya estaba arreglado (no hacía falta cambiar nada).', ui.ButtonSet.OK);
}

// Lo que corre en cada guardado de la app (ver doPost).
function escribirGlovo_(data) {
  try {
    try {
      var anio = periodoAnual_(data.fecha);
      var idGlovo = buscarEnIndice_('GLOVO', anio);
      if (idGlovo) {
        var hechos = arreglarTotalesGlovo_(SpreadsheetApp.openById(idGlovo), Number(anio));
        if (hechos.length) avisarPorEmail_('Glovo: totales por día arreglados', 'Se cambiaron las fórmulas de las columnas I y L (totales por día) en: ' + hechos.join('; ') + '.');
      }
    } catch (errTot) {
      avisarPorEmail_('Glovo: no se pudieron arreglar los totales', String(errTot));
    }
    return sincronizarGlovoDelDia_(data.fecha, ticketsGlovoDelDia_(data));
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

// ---------- Leer un ticket de Glovo desde una foto (con Claude) ----------
var ESQUEMA_TICKET_GLOVO = {
  type: 'object',
  properties: {
    tickets: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          pedido: { type: 'string' },
          importe: { type: 'string' }
        },
        required: ['pedido', 'importe'],
        additionalProperties: false
      }
    }
  },
  required: ['tickets'],
  additionalProperties: false
};

var INSTRUCCIONES_TICKET_GLOVO =
  'Esta es la foto de uno o varios tickets de pedidos de Glovo impresos en un restaurante ' +
  '(tiras de papel térmico, pueden estar torcidas, arrugadas, de lado o una al lado de la otra). ' +
  'Revisa toda la foto y devuelve UN elemento por cada ticket que se vea, en orden de izquierda a ' +
  'derecha y de arriba abajo. Cada ticket es así:\n' +
  '- Arriba del todo, en la primera línea: "glovo #359". El número que va después del "#" es el ' +
  'número de pedido (en este ejemplo, 359). Suele tener 2 a 4 cifras.\n' +
  '- Debajo vienen el nombre del restaurante, la fecha y hora, y un número largo de unas 12 cifras ' +
  '(ej. 101805097658): ese número largo NO es el pedido, no lo uses. Tampoco el PIN ni la hora de recogida.\n' +
  '- Abajo del todo, a la derecha de la palabra "Total", está el importe total (ej. "Total 37,80 €"). ' +
  'Usa siempre la línea "Total", no "Subtotal" ni los precios de cada plato.\n\n' +
  'Por cada ticket devuelve:\n' +
  '- pedido: solo las cifras que van después del "#" (sin el "#").\n' +
  '- importe: el Total, con punto como separador decimal y sin símbolo de euro (ejemplo: 37.80).\n' +
  'Si de un ticket se ve solo una parte (ej. está cortado por el borde de la foto o tapado), ' +
  'devuélvelo igual con el dato que no se lee vacío (""). No inventes ni adivines cifras.';

// Deja el pedido solo con lo que va después del "#" y el importe con
// punto decimal. El número largo del ticket (unas 12 cifras) no es el
// pedido: si la IA lo trae por error, se deja vacío para escribirlo a mano.
function limpiarTicketGlovo_(t) {
  var pedido = String(t.pedido || '').replace(/^.*#/, '').replace(/\s+/g, '');
  if (/^\d{7,}$/.test(pedido)) pedido = '';
  var importe = String(t.importe || '').replace(/[^\d,.]/g, '');
  if (importe.indexOf(',') > -1) importe = importe.replace(/\./g, '').replace(',', '.');
  if (!(Number(importe) > 0)) importe = '';
  return { pedido: pedido, importe: importe };
}

// Primero se lee rápido; si no aparece ningún ticket o alguno queda sin
// pedido o sin total, se vuelve a leer con más cuidado y se queda la
// lectura más completa.
function escanearTicketGlovo_(data) {
  if (!data.imageBase64) return { ok: false, error: 'No llegó la foto' };
  var contenido = [
    { type: 'image', source: { type: 'base64', media_type: data.mediaType || 'image/jpeg', data: data.imageBase64 } },
    { type: 'text', text: INSTRUCCIONES_TICKET_GLOVO }
  ];
  function leer(esfuerzo) {
    var r = pedirJsonAClaude_(contenido, ESQUEMA_TICKET_GLOVO, esfuerzo);
    if (r.ok) r.tickets = (r.tickets || []).map(limpiarTicketGlovo_).filter(function (t) { return t.pedido || t.importe; });
    return r;
  }
  function completos(r) {
    return r.ok ? r.tickets.filter(function (t) { return t.pedido && t.importe; }).length : -1;
  }
  var r = leer('low');
  if (r.ok && r.tickets.length && completos(r) === r.tickets.length) return r;
  var r2 = leer('medium');
  if (!r.ok) return r2;
  if (!r2.ok) return r;
  return completos(r2) > completos(r) || (completos(r2) === completos(r) && r2.tickets.length > r.tickets.length) ? r2 : r;
}
