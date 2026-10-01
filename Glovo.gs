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
      if (g && g.id) tickets.push(g);
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
  var ids = hoja.getRange(GLOVO_FILA_DESDE_, GLOVO_COL_ID_, cantidad, 1).getValues();

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

  // Los que ya no están en la app: se vacía su fila.
  Object.keys(filaDeId).forEach(function (id) {
    if (actuales[id]) return;
    var i = filaDeId[id];
    datos[i] = ['', '', ''];
    ids[i] = [''];
    cambiadas[i] = true;
  });

  // La primera fila libre es la de debajo del último pedido cargado.
  var ultima = -1;
  for (var j = 0; j < cantidad; j++) {
    if (datos[j][0] !== '' || datos[j][1] !== '' || datos[j][2] !== '' || ids[j][0] !== '') ultima = j;
  }

  var fecha = fechaComoDate_(fechaISO);
  var sinLugar = 0;
  tickets.forEach(function (g) {
    var fila = [fecha, String(g.pedido || ''), Number(g.importe) || 0];
    var i = filaDeId[g.id];
    if (i == null) {
      if (ultima + 1 >= cantidad) { sinLugar++; return; }
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

// Lo que corre en cada guardado de la app (ver doPost).
function escribirGlovo_(data) {
  try {
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
  'Esta es la foto de uno o varios tickets de pedidos de Glovo impresos en un restaurante. ' +
  'Por cada ticket que se vea, extrae:\n' +
  '- pedido: el número o código del pedido de Glovo, tal como aparece (sin el "#").\n' +
  '- importe: el total del pedido.\n' +
  'Los importes con punto como separador decimal y sin símbolo de euro (ejemplo: 24.30). ' +
  'Si un dato no se lee con seguridad, déjalo vacío ("") en vez de inventarlo.';

function escanearTicketGlovo_(data) {
  if (!data.imageBase64) return { ok: false, error: 'No llegó la foto' };
  return pedirJsonAClaude_([
    { type: 'image', source: { type: 'base64', media_type: data.mediaType || 'image/jpeg', data: data.imageBase64 } },
    { type: 'text', text: INSTRUCCIONES_TICKET_GLOVO }
  ], ESQUEMA_TICKET_GLOVO);
}
