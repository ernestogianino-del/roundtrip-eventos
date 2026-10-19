/**
 * Cuadre Sisifuz - receptor en Google Drive y Google Sheets
 *
 * QUE HACE
 *  Recibe lo que Nery llena en la pagina /cuadre y lo guarda en tu Drive:
 *   - Una carpeta por fecha dentro de tu carpeta de cuadres:  "Cuadre AAAA-MM-DD"
 *   - Dentro, una hoja "Datos Nery DD.MM" con una pestana por seccion
 *     (Cajas, Taquilla, Totem, Planilla, Extras, Gastos local, Resumen)
 *   - Dentro, una subcarpeta "Fotos" con las fotos que subio, nombradas por tipo
 *  Si Nery vuelve a enviar la misma fecha, la hoja se reescribe y las fotos se agregan.
 *  No borra ni modifica ninguna otra cosa de tu Drive.
 *
 * INSTALACION (una sola vez, con tu cuenta de Google)
 * 1. Entra a https://script.google.com > Proyecto nuevo. Borra lo que haya y pega todo este archivo.
 * 2. Configuracion del proyecto (engranaje) > Propiedades de la secuencia de comandos > Agregar:
 *      CODIGO      = el codigo de acceso que va a escribir Nery
 *      CARPETA_ID  = el ID de la carpeta de Drive donde quieres las carpetas por fecha
 *                    (es el texto largo al final de la URL de esa carpeta)
 * 3. Implementar > Nueva implementacion > tipo "Aplicacion web":
 *        Ejecutar como: Yo      Quien tiene acceso: Cualquier persona
 *    Autoriza los permisos (Drive y Hojas de calculo) y copia la URL que termina en /exec.
 * 4. Pasale esa URL a Claude (o pegala en cuadre/index.html, constante WEB_APP_URL).
 *
 * El codigo y el ID de carpeta NO estan en este archivo ni en el repositorio: viven solo en las
 * propiedades del script. Sin el codigo correcto el script no lee ni escribe nada.
 *
 * Si cambias este archivo despues, hay que crear una NUEVA VERSION de la implementacion
 * (Implementar > Administrar implementaciones > editar > Nueva version).
 */

function props_() { return PropertiesService.getScriptProperties(); }

function salida_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function codigoOk_(c) {
  var k = props_().getProperty('CODIGO');
  return !!k && String(c || '') === k;
}

function denegado_() {
  Utilities.sleep(1000);
  return salida_({ ok: false, error: 'codigo' });
}

function num_(v) {
  var n = parseFloat(String(v === undefined || v === null ? '' : v).replace(',', '.'));
  return isFinite(n) ? n : 0;
}

function txt_(v) { return String(v === undefined || v === null ? '' : v).slice(0, 500); }

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (!codigoOk_(p.code)) return denegado_();
  return salida_({ ok: true, ping: true });
}

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (x) { return salida_({ ok: false, error: 'formato' }); }
  if (!codigoOk_(body.code)) return denegado_();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(body.fecha || ''))) return salida_({ ok: false, error: 'fecha' });
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    if (body.action === 'save') return salida_(guardar_(body));
    if (body.action === 'photo') return salida_(foto_(body));
    return salida_({ ok: false, error: 'accion' });
  } catch (err) {
    return salida_({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

// ---------------------------------------------------------------- carpetas

function carpetaRaiz_() {
  var id = props_().getProperty('CARPETA_ID');
  if (!id) throw new Error('Falta la propiedad CARPETA_ID');
  return DriveApp.getFolderById(id);
}

function subcarpeta_(padre, nombre) {
  var it = padre.getFoldersByName(nombre);
  return it.hasNext() ? it.next() : padre.createFolder(nombre);
}

function carpetaFecha_(fecha) { return subcarpeta_(carpetaRaiz_(), 'Cuadre ' + fecha); }

function ddmm_(fecha) { return fecha.slice(8, 10) + '.' + fecha.slice(5, 7); }

// ---------------------------------------------------------------- datos

function guardar_(b) {
  var d = b.data || {};
  var carpeta = carpetaFecha_(b.fecha);
  var nombre = 'Datos Nery ' + ddmm_(b.fecha);
  var it = carpeta.getFilesByName(nombre);
  var ss;
  if (it.hasNext()) {
    ss = SpreadsheetApp.openById(it.next().getId());
  } else {
    ss = SpreadsheetApp.create(nombre);
    DriveApp.getFileById(ss.getId()).moveTo(carpeta);
  }

  var cajas = (d.cajas || []).map(function (r) {
    var e = num_(r.efectivo), y = num_(r.yape), p = num_(r.pos);
    return [txt_(r.tipo), txt_(r.resp), e, y, p, e + y + p, num_(r.cont)];
  });
  tabla_(ss, 'Cajas', ['Tipo', 'Responsable', 'Efectivo', 'Yape Luis', 'Culqui (POS)', 'Total', 'Contómetro'], cajas, [3, 4, 5, 6, 7]);

  var cover = (d.cover || []).map(function (r) {
    var e = num_(r.efectivo), y = num_(r.yape);
    return [txt_(r.item), num_(r.cant), num_(r.precio), e, y, e + y];
  });
  tabla_(ss, 'Taquilla', ['Concepto', 'Cant', 'Precio', 'Efectivo', 'Yape', 'Total'], cover, [4, 5, 6],
    [['Barra vendida en taquilla (botellas)', '', '', '', '', num_(d.barraTaq)]]);

  var totem = (d.totem || []).map(function (r) {
    var t = num_(r.totem), e = num_(r.efectivo);
    return [txt_(r.item), num_(r.cant), num_(r.precio), t, e, t + e];
  });
  tabla_(ss, 'Totem', ['Concepto', 'Cant', 'Precio', 'Totem', 'Efectivo', 'Total'], totem, [4, 5, 6]);

  var plan = (d.planilla || []).map(function (r) {
    return [txt_(r.cargo), txt_(r.nombre), num_(r.monto), txt_(r.paga)];
  });
  tabla_(ss, 'Planilla', ['Cargo', 'Nombre', 'Monto S/', 'Paga'], plan, [3]);
  var shP = ss.getSheetByName('Planilla');
  var fin = plan.length + 1;
  shP.getRange(fin + 2, 1, 2, 3).setFormulas([
    ['Total Sisifuz', '', '=SUMIF(D2:D' + fin + ',"Sisifuz",C2:C' + fin + ')'],
    ['Total local', '', '=SUMIF(D2:D' + fin + ',"Local",C2:C' + fin + ')']
  ]).setFontWeight('bold');

  var ext = (d.extras || []).map(function (r) { return [txt_(r.concepto), txt_(r.prov), num_(r.monto)]; });
  tabla_(ss, 'Extras', ['Concepto', 'Proveedor', 'Monto S/'], ext, [3]);

  var gastos = [['Hielo', num_(d.hielo)], ['Frutas para cócteles', num_(d.frutas)]];
  (d.otros || []).forEach(function (r) { gastos.push([txt_(r.concepto), num_(r.monto)]); });
  gastos.push(['Cortesías promotores y DJs', num_(d.cprom)]);
  gastos.push(['Cortesías Sr. Luis', num_(d.cluis)]);
  gastos.push(['Cortesías Srta. Nery', num_(d.cnery)]);
  tabla_(ss, 'Gastos local', ['Concepto', 'Monto S/'], gastos, []);

  var ahora = Utilities.formatDate(new Date(), 'America/Lima', 'yyyy-MM-dd HH:mm');
  var res = [['Fecha', b.fecha], ['Llenó', txt_(d.resp)], ['Enviado', ahora], ['Observaciones', String(d.obs || '').slice(0, 2000)], ['', '']];
  String(b.resumen || '').split('\n').slice(0, 200).forEach(function (l) { res.push([l.slice(0, 300), '']); });
  var shR = ss.getSheetByName('Resumen') || ss.insertSheet('Resumen');
  shR.clear();
  shR.getRange(1, 1, res.length, 2).setValues(res);
  shR.getRange(1, 1, 4, 1).setFontWeight('bold');
  shR.setColumnWidth(1, 420);

  ['Hoja 1', 'Sheet1'].forEach(function (n) {
    var s = ss.getSheetByName(n);
    if (s && ss.getSheets().length > 1) ss.deleteSheet(s);
  });
  return { ok: true, carpeta: carpeta.getUrl(), hoja: ss.getUrl() };
}

// Escribe una tabla con encabezado en negrita y, si pide, fila de totales con SUM por columna.
function tabla_(ss, nombre, encabezado, filas, colsSuma, extra) {
  var sh = ss.getSheetByName(nombre) || ss.insertSheet(nombre);
  sh.clear();
  var todo = [encabezado].concat(filas);
  sh.getRange(1, 1, todo.length, encabezado.length).setValues(todo);
  sh.getRange(1, 1, 1, encabezado.length).setFontWeight('bold');
  sh.setFrozenRows(1);
  var fin = todo.length;
  if (colsSuma && colsSuma.length && filas.length) {
    var tot = [];
    for (var c = 1; c <= encabezado.length; c++) {
      if (colsSuma.indexOf(c) >= 0) {
        var L = letra_(c);
        tot.push('=SUM(' + L + '2:' + L + fin + ')');
      } else { tot.push(c === 1 ? 'TOTAL' : ''); }
    }
    sh.getRange(fin + 1, 1, 1, encabezado.length).setFormulas([tot]).setFontWeight('bold');
    fin += 1;
  }
  if (extra && extra.length) {
    sh.getRange(fin + 2, 1, extra.length, extra[0].length).setValues(extra);
  }
  for (var k = 1; k <= encabezado.length; k++) sh.autoResizeColumn(k);
}

function letra_(n) {
  var s = '';
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// ---------------------------------------------------------------- fotos

function foto_(b) {
  var mime = String(b.mime || 'image/jpeg');
  if (!/^image\//.test(mime)) throw new Error('Solo se aceptan imagenes');
  if (!b.b64 || String(b.b64).length > 14000000) throw new Error('Foto demasiado grande');
  var fotos = subcarpeta_(carpetaFecha_(b.fecha), 'Fotos');
  var tipo = String(b.tipo || 'Foto').replace(/[^\w\sáéíóúñÁÉÍÓÚÑ-]/g, '').slice(0, 40);
  var base = String(b.nombre || 'foto').replace(/[^\w.\- ]/g, '').slice(0, 60);
  var n = 1;
  var it = fotos.getFiles();
  while (it.hasNext()) { it.next(); n++; }
  var nombre = ('0' + n).slice(-2) + ' - ' + tipo + ' - ' + base;
  if (!/\.(jpe?g|png|webp|heic)$/i.test(nombre)) nombre += '.jpg';
  var blob = Utilities.newBlob(Utilities.base64Decode(String(b.b64)), mime, nombre);
  var f = fotos.createFile(blob);
  return { ok: true, id: f.getId(), nombre: nombre };
}
