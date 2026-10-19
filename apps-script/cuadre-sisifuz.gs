/**
 * Cuadre web por marca - receptor y calculadora en Google Drive / Hojas
 *
 * QUE HACE
 *  Cada marca de eventos tiene su propio link (cuadre/?m=<marca>) y su propio codigo. El operador de
 *  la noche llena el formulario al final de la fiesta. Este script:
 *   1) Guarda todo en el Drive de la marca: carpeta "Cuadre AAAA-MM-DD" con una hoja
 *      "Cuadre <Marca> DD.MM" y una subcarpeta "Fotos".
 *   2) Calcula el cuadre con las reglas de la marca: ingresos, gastos del local (insumos, alquiler por
 *      tramo, personal, hielo, frutas, IGV, comision POS), cuanto deposita Luis al local, cobros
 *      pendientes y utilidad parcial. La hoja "Cuadre" queda lista cuando Luis despierta.
 *  Las reglas y los codigos NO estan en este archivo: viven en la hoja privada "Cuadre web - Config marcas".
 *  Agregar una marca = agregar una fila en esa hoja (y su nombre y dia en cuadre/index.html).
 *
 * INSTALACION (una sola vez, con tu cuenta de Google)
 * 1. https://script.google.com > Proyecto nuevo. Borra lo que haya y pega todo este archivo.
 * 2. Configuracion del proyecto (engranaje) > Propiedades de la secuencia de comandos > Agregar:
 *      CONFIG_ID = el ID de la hoja "Cuadre web - Config marcas"
 *                  (texto largo de su URL, entre /d/ y /edit)
 * 3. Implementar > Nueva implementacion > tipo "Aplicacion web":
 *        Ejecutar como: Yo      Quien tiene acceso: Cualquier persona
 *    Autoriza los permisos (Drive y Hojas de calculo) y copia la URL que termina en /exec.
 * 4. Pega esa URL en cuadre/index.html, constante WEB_APP_URL.
 *
 * Si cambias este archivo despues, crea una NUEVA VERSION de la implementacion
 * (Implementar > Administrar implementaciones > editar > Nueva version).
 */

// ---------------------------------------------------------------- utilidades

function salida_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function denegado_() { Utilities.sleep(1000); return salida_({ ok: false, error: 'codigo' }); }
function num_(v) {
  var n = parseFloat(String(v === undefined || v === null ? '' : v).replace(',', '.'));
  return isFinite(n) ? n : 0;
}
function r2_(n) { return Math.round(n * 100) / 100; }
function txt_(v) { return String(v === undefined || v === null ? '' : v).slice(0, 500); }
function ddmm_(f) { return f.slice(8, 10) + '.' + f.slice(5, 7); }
function letra_(n) {
  var s = '';
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// Lee la hoja de configuracion y devuelve la marca pedida, o null.
function marca_(slug) {
  var id = PropertiesService.getScriptProperties().getProperty('CONFIG_ID');
  if (!id) throw new Error('Falta la propiedad CONFIG_ID');
  var v = SpreadsheetApp.openById(id).getSheets()[0].getDataRange().getValues();
  var h = v[0].map(function (x) { return String(x).trim(); });
  for (var i = 1; i < v.length; i++) {
    var o = {};
    h.forEach(function (k, j) { o[k] = v[i][j]; });
    if (String(o.slug).trim() === String(slug || '').trim() && o.slug !== '') return o;
  }
  return null;
}

function validar_(slug, code) {
  var m = marca_(slug);
  if (!m || !m.codigo || String(code || '') !== String(m.codigo)) return null;
  return m;
}

// ---------------------------------------------------------------- entradas web

function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    var m = validar_(p.m, p.code);
    if (!m) return denegado_();
    return salida_({ ok: true, nombre: m.nombre, local: m.local });
  } catch (err) { return salida_({ ok: false, error: String(err && err.message ? err.message : err) }); }
}

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (x) { return salida_({ ok: false, error: 'formato' }); }
  var lock = LockService.getScriptLock();
  try {
    var m = validar_(body.m, body.code);
    if (!m) return denegado_();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(body.fecha || ''))) return salida_({ ok: false, error: 'fecha' });
    lock.waitLock(20000);
    if (body.action === 'save') return salida_(guardar_(body, m));
    if (body.action === 'photo') return salida_(foto_(body, m));
    return salida_({ ok: false, error: 'accion' });
  } catch (err) {
    return salida_({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

// ---------------------------------------------------------------- carpetas

function subcarpeta_(padre, nombre) {
  var it = padre.getFoldersByName(nombre);
  return it.hasNext() ? it.next() : padre.createFolder(nombre);
}
function carpetaFecha_(fecha, m) {
  return subcarpeta_(DriveApp.getFolderById(String(m.carpeta_id).trim()), 'Cuadre ' + fecha);
}

// ---------------------------------------------------------------- calculo del cuadre

function tramoAlquiler_(base, tramos) {
  // tramos = "0:1500;8000:2000;13000:3000;16000:4000" -> el monto del mayor umbral que la base supere.
  var monto = 0, mejor = -1;
  String(tramos || '').split(';').forEach(function (t) {
    var p = t.split(':');
    var u = num_(p[0]), a = num_(p[1]);
    if (String(p[0]).trim() === '') return;
    if ((u === 0 && mejor < 0) || base > u) { if (u >= mejor) { mejor = u; monto = a; } }
  });
  return monto;
}

function calcular_(d, m) {
  var pPos = num_(m.pct_pos), pTot = num_(m.pct_totem), igv = num_(m.igv), pLic = num_(m.pct_costo_licor);
  var pUtil = num_(m.pct_utilidad_local);
  var cajaLocal = String(m.caja_local || 'CAJA').trim();
  var cajas = d.cajas || [];
  var tot = function (r) { return num_(r.efectivo) + num_(r.yape) + num_(r.pos); };
  var sumT = function (f) { return cajas.filter(f).reduce(function (a, r) { return a + tot(r); }, 0); };
  var sumP = function (f) { return cajas.filter(f).reduce(function (a, r) { return a + num_(r.pos); }, 0); };
  var sumEY = function (f) { return cajas.filter(f).reduce(function (a, r) { return a + num_(r.efectivo) + num_(r.yape); }, 0); };

  var cover = (d.cover || []).reduce(function (a, r) { return a + num_(r.efectivo) + num_(r.yape); }, 0);
  var totemBruto = (d.totem || []).reduce(function (a, r) { return a + num_(r.totem) + num_(r.efectivo); }, 0);
  var totemPass = (d.totem || []).reduce(function (a, r) { return a + num_(r.totem); }, 0);
  var barraMozos = sumT(function (r) { return r.tipo === 'MOZO'; });
  var barraLocal = sumT(function (r) { return r.tipo === cajaLocal; });
  var barraTaq = num_(d.barraTaq);
  var ingresos = cover + totemBruto + barraMozos + barraLocal;

  var ventaSistema = barraMozos + barraLocal;
  var ventaLicor = ventaSistema + barraTaq;
  var cortesias = num_(d.cprom) + num_(d.cluis) + num_(d.cnery);
  var costoLicor = r2_(ventaLicor * pLic);
  var costoCort = r2_(cortesias * pLic);
  var costoInsumos = r2_(costoLicor + costoCort);
  var alquiler = tramoAlquiler_(ventaLicor, m.tramos);
  var personalLocal = (d.planilla || []).filter(function (r) { return r.paga === 'Local'; }).reduce(function (a, r) { return a + num_(r.monto); }, 0);
  var hielo = num_(d.hielo), frutas = num_(d.frutas);
  var otros = (d.otros || []).reduce(function (a, r) { return a + num_(r.monto); }, 0);
  var posLocal = sumP(function (r) { return r.tipo === cajaLocal; });
  var igvMonto = r2_(posLocal * igv / (1 + igv));
  var culquiLocal = r2_(posLocal * pPos);
  var gastosLocal = r2_(costoInsumos + alquiler + personalLocal + hielo + frutas + otros + igvMonto + culquiLocal);

  var planSis = (d.planilla || []).filter(function (r) { return r.paga !== 'Local'; }).reduce(function (a, r) { return a + num_(r.monto); }, 0);
  var extras = (d.extras || []).reduce(function (a, r) { return a + num_(r.monto); }, 0);
  var posMozos = sumP(function (r) { return r.tipo === 'MOZO'; });
  var culquiMozos = r2_(posMozos * pPos);
  var comTotem = r2_(totemPass * pTot);
  var gastosMarca = r2_(planSis + extras + culquiMozos + comTotem);

  var utilidadParcial = r2_(ingresos - gastosLocal - gastosMarca);
  var utilLocal = r2_(utilidadParcial * pUtil);
  var deposita = r2_(gastosLocal - barraLocal);

  // Efectivo + Yape de la noche: mozos y cover (sin totem ni caja del local).
  var efectivoNoche = r2_(sumEY(function (r) { return r.tipo !== 'TOTEM' && r.tipo !== cajaLocal; }));

  return {
    cover: cover, totemBruto: totemBruto, barraMozos: barraMozos, barraLocal: barraLocal, ingresos: ingresos,
    ventaSistema: ventaSistema, barraTaq: barraTaq, ventaLicor: ventaLicor, cortesias: cortesias,
    costoLicor: costoLicor, costoCort: costoCort, costoInsumos: costoInsumos, alquiler: alquiler,
    personalLocal: personalLocal, hielo: hielo, frutas: frutas, otros: otros, igvMonto: igvMonto,
    culquiLocal: culquiLocal, gastosLocal: gastosLocal, planSis: planSis, extras: extras,
    culquiMozos: culquiMozos, comTotem: comTotem, gastosMarca: gastosMarca,
    utilidadParcial: utilidadParcial, utilLocal: utilLocal, deposita: deposita,
    efectivoNoche: efectivoNoche, posMozosNeto: r2_(posMozos - culquiMozos), totemNeto: r2_(totemPass - comTotem),
    pPos: pPos, pTot: pTot, igv: igv, pLic: pLic, pUtil: pUtil
  };
}

// ---------------------------------------------------------------- guardar

function guardar_(b, m) {
  var d = b.data || {};
  var c = calcular_(d, m);
  var carpeta = carpetaFecha_(b.fecha, m);
  var nombre = 'Cuadre ' + m.nombre + ' ' + ddmm_(b.fecha);
  var it = carpeta.getFilesByName(nombre);
  var ss;
  if (it.hasNext()) { ss = SpreadsheetApp.openById(it.next().getId()); }
  else { ss = SpreadsheetApp.create(nombre); DriveApp.getFileById(ss.getId()).moveTo(carpeta); }

  escribirCuadre_(ss, b, m, d, c);

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
  var plan = (d.planilla || []).map(function (r) { return [txt_(r.cargo), txt_(r.nombre), num_(r.monto), txt_(r.paga)]; });
  tabla_(ss, 'Planilla', ['Cargo', 'Nombre', 'Monto S/', 'Paga'], plan, [3]);
  var shP = ss.getSheetByName('Planilla'), fin = plan.length + 1;
  shP.getRange(fin + 3, 1, 2, 3).setFormulas([
    ['Total Sisifuz / marca', '', '=SUMIF(D2:D' + fin + ',"<>Local",C2:C' + fin + ')'],
    ['Total local', '', '=SUMIF(D2:D' + fin + ',"Local",C2:C' + fin + ')']
  ]).setFontWeight('bold');
  var ext = (d.extras || []).map(function (r) { return [txt_(r.concepto), txt_(r.prov), num_(r.monto)]; });
  tabla_(ss, 'Extras', ['Concepto', 'Proveedor', 'Monto S/'], ext, [3]);
  var gastos = [['Hielo', num_(d.hielo)], ['Frutas para cócteles', num_(d.frutas)]];
  (d.otros || []).forEach(function (r) { gastos.push([txt_(r.concepto), num_(r.monto)]); });
  gastos.push(['Cortesías promotores y DJs', num_(d.cprom)], ['Cortesías Sr. Luis', num_(d.cluis)], ['Cortesías Srta. Nery', num_(d.cnery)]);
  tabla_(ss, 'Gastos local', ['Concepto', 'Monto S/'], gastos, []);

  var shR = ss.getSheetByName('Datos enviados') || ss.insertSheet('Datos enviados');
  shR.clear();
  var res = [['Fecha', b.fecha], ['Llenó', txt_(d.resp)], ['Enviado', Utilities.formatDate(new Date(), 'America/Lima', 'yyyy-MM-dd HH:mm')],
    ['Observaciones', String(d.obs || '').slice(0, 2000)], ['', '']];
  String(b.resumen || '').split('\n').slice(0, 200).forEach(function (l) { res.push([l.slice(0, 300), '']); });
  shR.getRange(1, 1, res.length, 2).setValues(res);
  shR.getRange(1, 1, 4, 1).setFontWeight('bold');
  shR.setColumnWidth(1, 420);

  ['Hoja 1', 'Sheet1'].forEach(function (n) {
    var s = ss.getSheetByName(n);
    if (s && ss.getSheets().length > 1) ss.deleteSheet(s);
  });
  ss.setActiveSheet(ss.getSheetByName('Cuadre'));
  ss.moveActiveSheet(1);
  return { ok: true, carpeta: carpeta.getUrl(), hoja: ss.getUrl(), deposita: c.deposita };
}

function escribirCuadre_(ss, b, m, d, c) {
  var sh = ss.getSheetByName('Cuadre') || ss.insertSheet('Cuadre');
  sh.clear();
  var pc = function (x) { return (r2_(x * 1000) / 10) + '%'; };
  var cont = (d.cajas || []).reduce(function (a, r) { return a + num_(r.cont); }, 0);
  var totCajas = (d.cajas || []).reduce(function (a, r) { return a + num_(r.efectivo) + num_(r.yape) + num_(r.pos); }, 0);
  var f = [
    ['Cuadre ' + m.nombre + ' - ' + b.fecha + ' (' + m.local + ')', '', ''],
    ['Calculado automáticamente con lo que envió el operador. Falta completar lo marcado PENDIENTE.', '', ''],
    ['', '', ''],
    ['INGRESOS', 'S/', ''],
    ['Taquilla cover', c.cover, ''],
    ['Totems Passline (bruto)', c.totemBruto, ''],
    ['Pre venta web Passline', 0, 'PENDIENTE: revisar el panel de Passline'],
    ['Barra mozos', c.barraMozos, ''],
    ['Barra caja del local', c.barraLocal, 'Queda en la caja del local'],
    ['Total ingresos', c.ingresos, ''],
    ['', '', ''],
    ['GASTOS DEL LOCAL (se recuperan)', 'S/', ''],
    ['Costo insumos (' + pc(c.pLic) + ' vendido + cortesías)', c.costoInsumos, 'Venta por sistema ' + r2_(c.ventaSistema) + ' + barra en taquilla ' + r2_(c.barraTaq) + '; cortesías ' + r2_(c.cortesias)],
    ['Alquiler (tramo)', c.alquiler, 'Base: venta de licor ' + r2_(c.ventaLicor)],
    ['Personal local', c.personalLocal, ''],
    ['Hielo', c.hielo, ''],
    ['Frutas para cócteles', c.frutas, ''],
    ['Otros gastos del local', c.otros, ''],
    ['IGV (POS de la caja del local)', c.igvMonto, 'Solo POS; el efectivo no lleva IGV'],
    ['Comisión POS caja del local ' + pc(c.pPos), c.culquiLocal, ''],
    ['Total gastos del local', c.gastosLocal, ''],
    ['', '', ''],
    ['GASTOS DE LA MARCA', 'S/', ''],
    ['Planilla de la marca', c.planSis, ''],
    ['Técnica y extras', c.extras, ''],
    ['Comisión POS mozos ' + pc(c.pPos), c.culquiMozos, ''],
    ['Comisión totems ' + pc(c.pTot), c.comTotem, ''],
    ['DJs, marketing y promotores', '', 'PENDIENTE: los completa Luis'],
    ['Total gastos de la marca (parcial)', c.gastosMarca, 'Sin DJs, marketing ni promotores'],
    ['', '', ''],
    ['UTILIDAD PARCIAL (antes de DJs, marketing y promotores)', c.utilidadParcial, ''],
    ['Utilidad para el local ' + pc(c.pUtil), c.utilLocal, ''],
    ['', '', ''],
    ['CUADRE DE PAGOS CON EL LOCAL', 'S/', ''],
    ['Costos del local a recuperar', c.gastosLocal, ''],
    ['Ingreso que queda en caja del local', c.barraLocal, ''],
    ['Luis deposita al local (costos - caja)', c.deposita, 'Si es negativo, el local le debe a Luis'],
    ['', '', ''],
    ['COBROS PENDIENTES (Luis)', 'S/', ''],
    ['Efectivo + Yape de la noche (mozos y cover)', c.efectivoNoche, ''],
    ['POS mozos neto', c.posMozosNeto, 'Se cobra al día siguiente'],
    ['Passline totems neto + pre venta', c.totemNeto, 'Se cobra el miércoles siguiente'],
    ['', '', ''],
    ['CONTROLES', 'S/', ''],
    ['Total cajas vs contómetro', r2_(totCajas - cont), cont > 0 ? (r2_(totCajas - cont) === 0 ? 'Cuadra' : 'REVISAR diferencia') : 'No ingresaron contómetro']
  ];
  sh.getRange(1, 1, f.length, 3).setValues(f);
  sh.getRange(2, 2, f.length - 1, 1).setNumberFormat('#,##0.00');
  sh.setColumnWidth(1, 380); sh.setColumnWidth(2, 110); sh.setColumnWidth(3, 380);
  f.forEach(function (row, i) {
    if (row[1] === 'S/' || row[0] === 'Total ingresos' || row[0] === 'Total gastos del local' ||
        String(row[0]).indexOf('UTILIDAD PARCIAL') === 0 || String(row[0]).indexOf('Luis deposita') === 0) {
      sh.getRange(i + 1, 1, 1, 3).setFontWeight('bold');
    }
  });
  sh.getRange(1, 1).setFontWeight('bold').setFontSize(13);
}

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
      if (colsSuma.indexOf(c) >= 0) { var L = letra_(c); tot.push('=SUM(' + L + '2:' + L + fin + ')'); }
      else { tot.push(c === 1 ? 'TOTAL' : ''); }
    }
    sh.getRange(fin + 1, 1, 1, encabezado.length).setFormulas([tot]).setFontWeight('bold');
    fin += 1;
  }
  if (extra && extra.length) sh.getRange(fin + 2, 1, extra.length, extra[0].length).setValues(extra);
  for (var k = 1; k <= encabezado.length; k++) sh.autoResizeColumn(k);
}

// ---------------------------------------------------------------- fotos

function foto_(b, m) {
  var mime = String(b.mime || 'image/jpeg');
  if (!/^image\//.test(mime)) throw new Error('Solo se aceptan imagenes');
  if (!b.b64 || String(b.b64).length > 14000000) throw new Error('Foto demasiado grande');
  var fotos = subcarpeta_(carpetaFecha_(b.fecha, m), 'Fotos');
  var tipo = String(b.tipo || 'Foto').replace(/[^\w\sáéíóúñÁÉÍÓÚÑ-]/g, '').slice(0, 40);
  var base = String(b.nombre || 'foto').replace(/[^\w.\- ]/g, '').slice(0, 60);
  var n = 1, it = fotos.getFiles();
  while (it.hasNext()) { it.next(); n++; }
  var nombre = ('0' + n).slice(-2) + ' - ' + tipo + ' - ' + base;
  if (!/\.(jpe?g|png|webp|heic)$/i.test(nombre)) nombre += '.jpg';
  var f = fotos.createFile(Utilities.newBlob(Utilities.base64Decode(String(b.b64)), mime, nombre));
  return { ok: true, id: f.getId(), nombre: nombre };
}

// ---------------------------------------------------------------- prueba (opcional, desde el editor)

function probarCalculo_() {
  // Ejemplo del 14/10 (Sisifuz): debe dar deposita = 1350.61
  var d = {
    cajas: [{ tipo: 'TOTEM', efectivo: 0, yape: 1502, pos: 0 }, { tipo: 'MOZO', efectivo: 64, yape: 0, pos: 1620 },
      { tipo: 'MOZO', efectivo: 12, yape: 0, pos: 974 }, { tipo: 'CAJA', efectivo: 538, yape: 0, pos: 3195 },
      { tipo: 'COVER', efectivo: 1302, yape: 895, pos: 0 }],
    cover: [{ efectivo: 1302, yape: 895 }], totem: [{ totem: 1502, efectivo: 0 }],
    barraTaq: 2344, cprom: 995, cluis: 496, cnery: 168, hielo: 64.6, frutas: 85,
    planilla: [{ monto: 260, paga: 'Local' }, { monto: 900, paga: 'Sisifuz' }], extras: []
  };
  Logger.log(JSON.stringify(calcular_(d, marca_('sisifuz'))));
}
