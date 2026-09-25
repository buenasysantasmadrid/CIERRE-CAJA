// ============================================================================
// ÍNDICE — registro central de qué planilla corresponde a cada archivo del
// negocio: Cierre de Caja (una planilla por MES), Contabilidad (una planilla
// por AÑO), y los que se vayan sumando más adelante, todos anuales:
// Albaranes, Comparativa Carne, Glovo, Stock Producción, Cambio de Tinta,
// Arreglo Electrodomésticos.
//
// Esta es la ÚNICA planilla cuyo ID hace falta pegar a mano en el código
// (ahí abajo, en INDICE_SHEET_ID_). Todo lo demás — "cuál es la planilla de
// Cierre de Caja de este mes", "cuál es la de Contabilidad de este año" — el
// script lo resuelve solo consultando esta, y si hace falta, crea el
// archivo nuevo automáticamente (por eso funciona solo al cambiar de mes o
// de año, sin tocar nada a mano).
//
// ---- Cómo armar la planilla Índice (una sola vez) --------------------------
// 1) Crear una planilla nueva en Drive, llamarla por ejemplo "Índice
//    CIERRE-CAJA", y pegar su ID acá abajo en INDICE_SHEET_ID_.
// 2) Adentro, crear dos pestañas:
//
//    Pestaña "Tipos" (un renglón por tipo de archivo; se completa a mano):
//      Tipo | Granularidad | Plantilla ID | Carpeta Drive ID | Nombre base
//      CIERRE_CAJA | MENSUAL | <id de la planilla plantilla> |  | Cierre de Caja
//      CONTABILIDAD | ANUAL | <id de la planilla plantilla> |  | Contabilidad
//      ALBARANES | ANUAL |  |  | Albaranes
//      COMPARATIVA_CARNE | ANUAL |  |  | Comparativa Carne
//      GLOVO | ANUAL |  |  | Glovo
//      STOCK_PRODUCCION | ANUAL |  |  | Stock Producción
//      CAMBIO_TINTA | ANUAL |  |  | Cambio de Tinta
//      ARREGLO_ELECTRODOMESTICOS | ANUAL |  |  | Arreglo Electrodomésticos
//
//      Los tipos con "Plantilla ID" vacío son el "hueco" para cuando se
//      construyan: el script explica con un error claro que falta eso, en
//      vez de crear cualquier cosa a ciegas. "Carpeta Drive ID" es opcional
//      (si se deja vacía, el archivo nuevo se crea suelto en Mi unidad).
//
//    Pestaña "Archivos" (con encabezados; el script la completa solo):
//      Tipo | Periodo | ID Planilla | Nombre | Creado
//
//    "Periodo" es "YYYY-MM" para los tipos MENSUAL (Cierre de Caja) y
//    "YYYY" para los ANUAL (Contabilidad y el resto).
//
// 3) Preparar la planilla PLANTILLA de Cierre de Caja: una copia de la
//    planilla de Cierre de Caja actual, dejando solo la pestaña "1" (la
//    plantilla del día bonito) — Registro y Movimientos las crea el script
//    solas la primera vez que hace falta. Pegar su ID en "Tipos".
// 4) Preparar la planilla PLANTILLA de Contabilidad: una copia de la
//    planilla de Contabilidad actual, dejando solo la pestaña "MASTER"
//    (el mes en blanco). Pegar su ID en "Tipos".
// 5) Pegar este archivo (Indice.gs) y CierreCaja.gs en el proyecto de Apps
//    Script de la PROPIA planilla Índice (Extensiones > Apps Script desde
//    el Índice) — NO en la plantilla ni en las planillas de cada mes: lo que
//    tenga la plantilla se copia a cada mes nuevo, y el onEdit correría dos
//    veces. Implementar como aplicación web y poner esa URL en index.html
//    (es un solo Web App para todos los meses/años).
// 6) Correr una vez el menú "Cierre de Caja — Pruebas" > "Conectar los
//    meses cargados en el Índice", para conectar las planillas de mes que
//    se cargaron a mano en "Archivos".
// ============================================================================

var INDICE_SHEET_ID_ = '1ZqE4UZqXIYmnmwu-uNLgHDhRpl62WbbRNKY9i2ePyEY';

// Cuántos meses hacia atrás (incluyendo el actual) se buscan por defecto en
// los listados que cruzan meses: calendario de días con datos, fondo fijo
// sugerido del turno anterior y, más adelante, Albaranes.
var MESES_HISTORIAL_ = 6;

function hoyISO_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function periodoMensual_(fechaISO) {
  var partes = String(fechaISO || '').split('-');
  if (partes.length !== 3) throw new Error('Fecha inválida para calcular período mensual: "' + fechaISO + '".');
  return partes[0] + '-' + partes[1];
}

function periodoAnual_(fechaISOoAnio) {
  var s = String(fechaISOoAnio || '');
  var anio = s.indexOf('-') > -1 ? s.split('-')[0] : s;
  if (!/^\d{4}$/.test(anio)) throw new Error('Fecha/año inválido para calcular período anual: "' + fechaISOoAnio + '".');
  return anio;
}

// Resta `n` meses a un período "YYYY-MM".
function periodoMensualMenos_(periodo, n) {
  var partes = periodo.split('-');
  var anio = parseInt(partes[0], 10), mes = parseInt(partes[1], 10);
  var d = new Date(anio, mes - 1 - n, 1);
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
}

function hojaIndice_(nombre) {
  var indice = SpreadsheetApp.openById(INDICE_SHEET_ID_);
  var hoja = indice.getSheetByName(nombre);
  if (!hoja) throw new Error('La planilla Índice no tiene la pestaña "' + nombre + '". Revisá INDICE_SHEET_ID_ y la estructura de esa planilla.');
  return hoja;
}

// Limpia un valor de celda para compararlo como texto: saca espacios de
// más y, si quedó un BOM (﻿) pegado al principio de la primera celda
// (pasa seguido al importar un CSV), también lo saca.
function normalizarTexto_(v) {
  if (v == null) return '';
  return String(v).replace(/^﻿/, '').trim();
}

// Busca el nombre de columna `nombre` en la fila de encabezados, sin
// importar espacios de más.
function indiceColumna_(header, nombre) {
  for (var i = 0; i < header.length; i++) {
    if (normalizarTexto_(header[i]) === nombre) return i;
  }
  return -1;
}

// Compara la celda de "Periodo" contra el período buscado. Google Sheets
// suele autoconvertir textos como "2026-09" o "2026" en una fecha real —
// acá se contempla ese caso además de la comparación de texto normal.
function coincidePeriodo_(valorCelda, periodo) {
  if (valorCelda instanceof Date) {
    var comoMes = Utilities.formatDate(valorCelda, Session.getScriptTimeZone(), 'yyyy-MM');
    var comoAnio = Utilities.formatDate(valorCelda, Session.getScriptTimeZone(), 'yyyy');
    return periodo === comoMes || periodo === comoAnio;
  }
  return normalizarTexto_(valorCelda) === periodo;
}

function buscarEnIndice_(tipo, periodo) {
  var hoja = hojaIndice_('Archivos');
  var valores = hoja.getDataRange().getValues();
  for (var i = 1; i < valores.length; i++) {
    if (normalizarTexto_(valores[i][0]) === tipo && coincidePeriodo_(valores[i][1], periodo)) {
      return valores[i][2] || null; // ID Planilla
    }
  }
  return null;
}

// Todas las planillas de `tipo` que figuran en "Archivos": [{id, nombre}].
function planillasDelTipo_(tipo) {
  var valores = hojaIndice_('Archivos').getDataRange().getValues();
  var resultado = [];
  for (var i = 1; i < valores.length; i++) {
    if (normalizarTexto_(valores[i][0]) !== tipo) continue;
    var id = normalizarTexto_(valores[i][2]);
    if (id) resultado.push({ id: id, nombre: normalizarTexto_(valores[i][3]) || id });
  }
  return resultado;
}

function registrarEnIndice_(tipo, periodo, id, nombre) {
  var hoja = hojaIndice_('Archivos');
  hoja.appendRow([tipo, periodo, id, nombre, new Date()]);
}

function configTipo_(tipo) {
  var hoja = hojaIndice_('Tipos');
  var valores = hoja.getDataRange().getValues();
  var header = valores[0];
  var idxTipo = indiceColumna_(header, 'Tipo');
  var idxPlantilla = indiceColumna_(header, 'Plantilla ID');
  var idxCarpeta = indiceColumna_(header, 'Carpeta Drive ID');
  var idxNombreBase = indiceColumna_(header, 'Nombre base');
  if (idxTipo === -1) {
    throw new Error('La pestaña "Tipos" del Índice no tiene una columna llamada exactamente "Tipo" en la fila 1 — revisá el encabezado (fila 1) de esa pestaña.');
  }
  var tiposEncontrados = [];
  for (var i = 1; i < valores.length; i++) {
    var tipoFila = normalizarTexto_(valores[i][idxTipo]);
    if (!tipoFila) continue;
    tiposEncontrados.push(tipoFila);
    if (tipoFila === tipo) {
      return {
        plantillaId: valores[i][idxPlantilla],
        carpetaId: valores[i][idxCarpeta],
        nombreBase: valores[i][idxNombreBase] || tipo
      };
    }
  }
  throw new Error('El tipo "' + tipo + '" no está dado de alta en la pestaña "Tipos" del Índice. Tipos que SÍ encontró ahí: [' + tiposEncontrados.join(', ') + '].');
}

// Devuelve el ID de la planilla de `tipo` para `periodo`, creándola desde su
// plantilla si todavía no existe (por ejemplo, al entrar a un mes o año
// nuevo). Si el tipo no está dado de alta en "Tipos", o está dado de alta
// pero sin plantilla todavía (el "hueco" para un archivo que falta
// construir), tira un error explicando exactamente eso.
function obtenerOCrearPlanilla_(tipo, periodo) {
  var idExistente = buscarEnIndice_(tipo, periodo);
  if (idExistente) return idExistente;

  var cfg = configTipo_(tipo);
  if (!cfg.plantillaId) {
    throw new Error('El tipo "' + tipo + '" todavía no tiene una planilla plantilla configurada en el Índice (columna "Plantilla ID" vacía) — falta terminar de construir ese archivo.');
  }

  var plantilla = DriveApp.getFileById(cfg.plantillaId);
  var nombre = cfg.nombreBase + ' ' + periodo;
  var archivoNuevo = cfg.carpetaId
    ? plantilla.makeCopy(nombre, DriveApp.getFolderById(cfg.carpetaId))
    : plantilla.makeCopy(nombre);
  var id = archivoNuevo.getId();

  if (tipo === 'CIERRE_CAJA') {
    try { instalarTriggerOnEditSiFalta_(id); }
    catch (errTrigger) { Logger.log('No se pudo instalar el trigger onEdit en "' + nombre + '": ' + errTrigger); }
  }

  registrarEnIndice_(tipo, periodo, id, nombre);
  return id;
}

// Conecta el onEdit de este proyecto a una planilla mensual de Cierre de
// Caja (para que las ediciones a mano en "Registro" o en la pestaña del día
// se reflejen). No duplica: si esa planilla ya tiene el trigger, no hace
// nada. Devuelve true si lo instaló.
function instalarTriggerOnEditSiFalta_(idPlanilla) {
  var yaEsta = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === 'onEdit' && t.getTriggerSourceId() === idPlanilla;
  });
  if (yaEsta) return false;
  ScriptApp.newTrigger('onEdit').forSpreadsheet(idPlanilla).onEdit().create();
  return true;
}

// Correr a mano (menú "Cierre de Caja — Pruebas" > "Conectar los meses
// cargados en el Índice", o desde el editor): conecta el onEdit a TODAS las
// planillas CIERRE_CAJA de la pestaña "Archivos" que todavía no lo tengan —
// hace falta para las que se cargaron a mano en el Índice (las que crea el
// script solo ya quedan conectadas). Se puede correr las veces que haga
// falta: no duplica.
function conectarMesesCierreCaja() {
  var resumen = [];
  planillasDelTipo_('CIERRE_CAJA').forEach(function (p) {
    try {
      resumen.push(p.nombre + ': ' + (instalarTriggerOnEditSiFalta_(p.id) ? 'conectado ahora' : 'ya estaba conectado'));
    } catch (err) {
      resumen.push(p.nombre + ': ERROR — ' + err);
    }
  });
  Logger.log(resumen.join('\n'));
  return resumen;
}

function planillaCierreCaja_(fechaISO) {
  return SpreadsheetApp.openById(obtenerOCrearPlanilla_('CIERRE_CAJA', periodoMensual_(fechaISO)));
}

function planillaContabilidad_(fechaISOoAnio) {
  return SpreadsheetApp.openById(obtenerOCrearPlanilla_('CONTABILIDAD', periodoAnual_(fechaISOoAnio)));
}

// Últimas `cantidadMeses` (incluyendo el actual) planillas de Cierre de Caja
// que YA EXISTEN — no crea meses pasados que nunca se usaron. Se usa para
// listados que necesitan mirar hacia atrás cruzando meses: calendario de
// días con datos, fondo fijo sugerido del turno anterior y, más adelante,
// Albaranes. Devuelve [{periodo, ss}], del más nuevo al más viejo.
function planillasCierreCajaRecientes_(cantidadMeses) {
  var periodoActual = periodoMensual_(hoyISO_());
  var resultado = [];
  for (var i = 0; i < cantidadMeses; i++) {
    var periodo = i === 0 ? periodoActual : periodoMensualMenos_(periodoActual, i);
    var id = buscarEnIndice_('CIERRE_CAJA', periodo);
    if (id) resultado.push({ periodo: periodo, ss: SpreadsheetApp.openById(id) });
  }
  return resultado;
}
