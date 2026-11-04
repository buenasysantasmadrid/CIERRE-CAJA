// ============================================================================
// SALUD — ¿la app puede guardar? Comprueba que la cuenta con la que corre es
// la dueña de las planillas (EDITOR_UNICO_) y que puede abrir y escribir lo
// que usa al guardar una caja: el Índice, Cierre de Caja del mes (y del mes
// que viene, si ya existe), Contabilidad y Albaranes del año, y la carpeta de
// facturas en Drive.
//
// Se usa de dos formas:
//  · Al publicar (.claude/publicar.ps1 llama a ?salud=1): si algo falla, la
//    publicación lo dice en rojo. Así se ve en el momento si se publicó con
//    otra cuenta (pasó el 3/11: un día entero sin guardar sin que nadie se
//    enterara).
//  · Cada hora (activador revisarSalud): si algo falla, manda un email en el
//    momento (como mucho uno cada 3 horas por el mismo problema), sin esperar
//    al resumen de las 9 y las 18.
// ============================================================================

var SALUD_SILENCIO_SEG_ = 3 * 60 * 60;

function comprobarSalud_() {
  var problemas = [];
  var bien = [];
  var cuenta = '';
  try { cuenta = String(Session.getEffectiveUser().getEmail() || '').toLowerCase(); } catch (e) {}
  if (cuenta !== EDITOR_UNICO_) {
    problemas.push('La app corre con la cuenta "' + (cuenta || 'desconocida') + '" y tiene que ser ' + EDITOR_UNICO_ +
      ' (la dueña de las planillas). Hay que volver a publicar con ' + EDITOR_UNICO_ + ' (.claude/publicar.ps1).');
  }

  var probar = function (nombre, fn) {
    try { var r = fn(); bien.push(nombre + (r ? ': ' + r : '')); }
    catch (err) { problemas.push(nombre + ': ' + err); }
  };

  // El Índice: abrir y escribir (una celda en la pestaña "Salud").
  probar('Índice (abrir y escribir)', function () {
    var ss = SpreadsheetApp.openById(INDICE_SHEET_ID_);
    var hoja = ss.getSheetByName('Salud') || ss.insertSheet('Salud');
    hoja.getRange('A1:B1').setValues([['Última comprobación', new Date()]]);
    return '';
  });

  var hoy = hoyISO_();
  var periodo = periodoMensual_(hoy);
  var siguiente = periodoMensualMenos_(periodo, -1);
  [periodo, siguiente].forEach(function (p, i) {
    var id = buscarEnIndice_('CIERRE_CAJA', p);
    if (!id) { if (i === 0) bien.push('Cierre de Caja ' + p + ': todavía no existe (se crea sola al guardar)'); return; }
    probar('Cierre de Caja ' + p + ' (abrir y escribir)', function () {
      var ss = SpreadsheetApp.openById(id);
      var reg = ss.getSheetByName('Registro');
      if (!reg) return 'sin Registro todavía';
      // Escribir lo mismo que ya hay: comprueba el permiso sin cambiar nada.
      var celda = reg.getRange(1, 1);
      celda.setValue(celda.getValue());
      return ss.getName();
    });
  });

  var anio = hoy.slice(0, 4);
  ['CONTABILIDAD', 'ALBARANES'].forEach(function (tipo) {
    var id = buscarEnIndice_(tipo, anio);
    if (!id) { problemas.push(tipo + ' ' + anio + ': no está en el Índice.'); return; }
    probar(tipo + ' ' + anio + ' (abrir)', function () { return SpreadsheetApp.openById(id).getName(); });
  });

  probar('Drive (carpeta de facturas)', function () {
    return typeof CARPETA_FACTURAS_ID_ !== 'undefined' ? DriveApp.getFolderById(CARPETA_FACTURAS_ID_).getName() : DriveApp.getRootFolder().getName();
  });

  return { ok: problemas.length === 0, cuenta: cuenta, problemas: problemas, bien: bien, cuando: new Date().toISOString() };
}

// Activador de cada hora (lo pone asegurarActivadorSalud_).
function revisarSalud() {
  var r = comprobarSalud_();
  if (r.ok) return r;
  var detalle = 'La app de la caja NO puede guardar bien:\n\n- ' + r.problemas.join('\n- ') +
    '\n\nLas cajas siguen en los aparatos y se vuelven a mandar solas cuando se arregle.';
  registrarError_('la app no puede guardar (comprobación de cada hora)', detalle);
  avisoUrgente_('URGENTE: la caja no se está guardando', detalle);
  return r;
}

// Email en el momento (no espera al resumen), como mucho uno cada 3 horas
// por el mismo asunto.
function avisoUrgente_(asunto, detalle) {
  try {
    var cache = CacheService.getScriptCache();
    var clave = 'urgente_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, asunto));
    if (cache.get(clave)) return false;
    var para = EDITOR_UNICO_;
    MailApp.sendEmail(para, 'Cierre de caja: ' + asunto,
      detalle + '\n\n(' + Utilities.formatDate(new Date(), 'Europe/Madrid', 'dd/MM/yyyy HH:mm') + ')');
    cache.put(clave, '1', SALUD_SILENCIO_SEG_);
    return true;
  } catch (err) {
    Logger.log('No se pudo mandar el aviso urgente: ' + err);
    return false;
  }
}

function asegurarActivadorSalud_() {
  var hay = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'revisarSalud'; });
  if (!hay) ScriptApp.newTrigger('revisarSalud').timeBased().everyHours(1).create();
  return !hay;
}
