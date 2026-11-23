// ============================================================================
// ¿Se leen bien los PDF de facturas guardados en Drive? (23/11/2026)
// ============================================================================
// Claude mira cada PDF de "Facturas proveedores" / AAAA / MM (con todas sus
// subcarpetas) y dice si los datos importantes se leen. Al terminar el mes
// manda un correo con la lista de los que se ven mal, para volver a
// escanearlos si todavía está el papel.
//
// Solo Claude, por tandas (Google corta a los 6 minutos):
//   GET ?revisarPdfs=AAAA-MM&desde=0&cuantos=15  → devuelve "siguiente"
//   ... seguir con &desde=<siguiente> hasta "terminado": true.
//   Con &correo=1 la última tanda manda el correo.
// ============================================================================

var ESQUEMA_CALIDAD_ = {
  type: 'object',
  properties: { se_lee_bien: { type: 'boolean' }, problema_imagen: { type: 'string' } },
  required: ['se_lee_bien', 'problema_imagen'],
  additionalProperties: false
};

function pdfsParaRevisar_(periodo) {
  var p = String(periodo).split('-');
  var itA = carpetaFacturas_().getFoldersByName(p[0]); if (!itA.hasNext()) return null;
  var itM = itA.next().getFoldersByName(p[1]); if (!itM.hasNext()) return null;
  var lista = [];
  (function recorrer(carpeta, ruta) {
    var fs = carpeta.getFiles();
    while (fs.hasNext()) {
      var f = fs.next(), tipo = f.getMimeType();
      if (tipo === 'application/pdf' || /^image\/(jpeg|png)$/.test(tipo)) lista.push({ archivo: f, ruta: ruta + f.getName(), tipo: tipo });
    }
    var subs = carpeta.getFolders();
    while (subs.hasNext()) { var s = subs.next(); recorrer(s, ruta + s.getName() + ' / '); }
  })(itM.next(), '');
  lista.sort(function (a, b) { return a.ruta < b.ruta ? -1 : 1; });
  return lista;
}

function revisarCalidadPdfs_(periodo, desde, cuantos, correo) {
  var lista = pdfsParaRevisar_(periodo);
  if (!lista) return { ok: false, error: 'No hay carpeta ' + periodo + ' en Facturas proveedores' };
  desde = Math.max(0, Number(desde) || 0);
  cuantos = Math.max(1, Math.min(40, Number(cuantos) || 15));
  var props = PropertiesService.getScriptProperties(), clave = 'calidadPdfs_' + periodo;
  var malos = [];
  if (desde > 0) { try { malos = JSON.parse(props.getProperty(clave) || '[]'); } catch (e) { malos = []; } }
  var hechos = [], inicio = Date.now(), i;
  for (i = desde; i < lista.length && i < desde + cuantos && Date.now() - inicio < 4.5 * 60 * 1000; i++) {
    var x = lista[i], r;
    try {
      var b64 = Utilities.base64Encode(x.archivo.getBlob().getBytes());
      var hoja = x.tipo === 'application/pdf'
        ? { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: b64 } }
        : { type: 'image', source: { type: 'base64', media_type: x.tipo, data: b64 } };
      r = pedirJsonAClaude_([hoja, { type: 'text', text:
        'Es el escaneo de una factura, albarán o ticket de un restaurante. Di si se lee bien para la contabilidad:\n' + CALIDAD_IMAGEN_TEXTO_ }],
        ESQUEMA_CALIDAD_, 'low', MODELO_ESCANEO);
    } catch (err) { r = { ok: false, error: String(err) }; }
    var fila = { nombre: x.ruta, url: x.archivo.getUrl(),
      bien: r.ok ? r.se_lee_bien : null, problema: r.ok ? r.problema_imagen : 'no se pudo revisar: ' + r.error };
    hechos.push(fila);
    if (fila.bien !== true) malos.push({ nombre: fila.nombre, url: fila.url, problema: fila.problema });
  }
  props.setProperty(clave, JSON.stringify(malos));
  var terminado = i >= lista.length;
  if (terminado && correo) mandarCorreoCalidad_(periodo, lista.length, malos);
  return { ok: true, total: lista.length, siguiente: i, terminado: terminado, hechos: hechos, malosHastaAhora: malos.length };
}

function mandarCorreoCalidad_(periodo, total, malos) {
  var para = Session.getEffectiveUser().getEmail();
  if (!para) return;
  var p = periodo.split('-');
  var mes = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'][+p[1] - 1] + ' ' + p[0];
  var texto = 'Revisados los ' + total + ' PDF de facturas de ' + mes + ' guardados en Drive.\n\n';
  if (!malos.length) texto += 'Todos se leen bien.';
  else {
    texto += 'Estos ' + malos.length + ' se ven mal. Si todavía tienes el papel, conviene volver a escanearlos:\n\n' +
      malos.map(function (m) { return '• ' + m.nombre + '\n   ' + (m.problema || '') + '\n   ' + m.url; }).join('\n\n');
  }
  MailApp.sendEmail(para, 'Facturas de ' + mes + ' que se ven mal', texto);
}
