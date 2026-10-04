/**
 * Cuadre Wapaz - receptor en Google Sheets
 *
 * HOJAS QUE USA (nunca borra ni modifica ninguna otra pestana de tu archivo):
 *
 *  1) "cuadre web": una fila por noche, cada ingreso y gasto en su columna.
 *     La fila 2 es de TOTALES (suma automatica de cada columna) y los cuadres empiezan en la fila 3.
 *     El detalle de la planilla de cada noche (persona por persona) queda como NOTA en la celda
 *     "Planilla staff" de esa fila (pasa el mouse por la celda para verlo).
 *
 *  2) "Negociacion AAAA-MM-DD": una hoja por cada version de la negociacion. La fecha del nombre es
 *     la fecha desde la que rige. Cuando la negociacion cambia, la pagina crea una hoja nueva y las
 *     anteriores quedan intactas como historial. Un cuadre usa la negociacion vigente a su fecha
 *     (la hoja mas reciente cuya fecha sea igual o anterior).
 *
 * INSTALACION (una sola vez)
 * 1. Abre tu hoja "Proyeccion Viernes - Wapaz" en Google Sheets.
 * 2. Menu Extensiones > Apps Script. Borra lo que haya y pega todo este archivo.
 * 3. Configuracion del proyecto (engranaje) > Propiedades de la secuencia de comandos >
 *    Agregar propiedad:  nombre: CODIGO   valor: el codigo de acceso que van a usar
 * 4. Implementar > Nueva implementacion > tipo "Aplicacion web":
 *        Ejecutar como: Yo      Quien tiene acceso: Cualquier persona
 *    Autoriza los permisos que pida Google y copia la URL que termina en /exec.
 * 5. Pega esa URL en cuadre-wapaz.html, en la constante WEB_APP_URL.
 *
 * El codigo de acceso NO esta en este archivo ni en el repositorio: vive solo en las propiedades
 * del script. Sin el codigo correcto el script no lee ni escribe nada.
 *
 * Si cambias este archivo despues, hay que crear una NUEVA VERSION de la implementacion
 * (Implementar > Administrar implementaciones > editar > Nueva version).
 */

var HOJA = 'cuadre web';
var PREFIJO_NEG = 'Negociación ';
var ENCABEZADOS = ['Fecha', 'Sistema', 'Efectivo Titanium', 'Taquilla efectivo/Yape', 'Taquilla POS',
  'Passline', 'Ingresos', 'Planilla staff', 'Hielo', 'Frutas', 'Cortesias', 'IGV', 'Alquiler',
  'Servicios', 'Costo licor', 'Visa', 'Comision Passline', 'Total gastos', 'Utilidad', 'Wapaz',
  'Titanium', 'Asistentes', 'Ticket promedio', 'Negociacion', 'Guardado'];
var COL_FECHA = 1;
var COL_ASIS = 22;
var COL_TICKET = 23;
var COL_NEG = 24;
var COL_GUARDADO = 25;
var COL_PLANILLA = 8;
var FILA_TOTALES = 2;
var FILA_DATOS = 3;

function doGet(e) {
  var p = (e && e.parameter) || {};
  var codigo = PropertiesService.getScriptProperties().getProperty('CODIGO');
  if (!codigo || p.code !== codigo) {
    Utilities.sleep(1000);
    return salida_({ ok: false, error: 'codigo' });
  }
  var lock = LockService.getScriptLock();
  try {
    switch (p.action) {
      case 'ping':
        return salida_({ ok: true });
      case 'config':
        return salida_(leerConfig_(String(p.fecha || '')));
      case 'saveConfig':
        lock.waitLock(10000);
        return salida_(guardarNegociacion_(JSON.parse(p.d)));
      case 'save':
        lock.waitLock(10000);
        return salida_(guardarCuadre_(JSON.parse(p.d)));
      case 'list':
        return salida_({ ok: true, items: listarCuadres_(15) });
      case 'get':
        return salida_(obtenerCuadre_(String(p.fecha || '')));
      default:
        return salida_({ ok: false, error: 'accion' });
    }
  } catch (err) {
    return salida_({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function salida_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function fechaValida_(f) { return /^\d{4}-\d{2}-\d{2}$/.test(f); }

function num_(v) {
  var n = Number(v);
  return isFinite(n) ? n : 0;
}

// ---------------------------------------------------------------- hoja "cuadre web"

// Devuelve la pestana "cuadre web". La crea si no existe. Si existe con otro formato, avisa y no la toca.
function hoja_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(HOJA);
  if (!sh) sh = ss.insertSheet(HOJA);
  if (sh.getRange(1, 1).getValue() === '') {
    sh.getRange(1, 1, 1, ENCABEZADOS.length).setValues([ENCABEZADOS]).setFontWeight('bold');
    var totales = ['TOTAL'];
    for (var c = 2; c <= COL_ASIS; c++) {
      var L = letra_(c);
      totales.push('=SUM(' + L + FILA_DATOS + ':' + L + ')');
    }
    totales.push('=IFERROR(G' + FILA_TOTALES + '/' + letra_(COL_ASIS) + FILA_TOTALES + ',0)');
    sh.getRange(FILA_TOTALES, 1, 1, totales.length).setFormulas([totales]).setFontWeight('bold');
    sh.setFrozenRows(FILA_TOTALES);
    sh.getRange(1, COL_FECHA, sh.getMaxRows(), 1).setNumberFormat('@');
    sh.getRange(FILA_TOTALES, 2, sh.getMaxRows() - 1, 21).setNumberFormat('#,##0.00');
    sh.getRange(FILA_TOTALES, COL_ASIS, sh.getMaxRows() - 1, 1).setNumberFormat('0');
    sh.getRange(FILA_TOTALES, COL_TICKET, sh.getMaxRows() - 1, 1).setNumberFormat('#,##0.00');
  } else if (sh.getRange(1, COL_GUARDADO).getValue() !== 'Guardado' || sh.getRange(1, COL_GUARDADO + 1).getValue() !== '') {
    throw new Error('hoja_formato');
  }
  return sh;
}

function letra_(n) {
  var s = '';
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// Ultima fila con contenido en UNA columna.
function ultimaFila_(sh, col) {
  var v = sh.getRange(1, col, sh.getMaxRows(), 1).getValues();
  for (var i = v.length - 1; i >= 0; i--) {
    if (v[i][0] !== '') return i + 1;
  }
  return 0;
}

function cantidadCuadres_(sh) {
  return Math.max(ultimaFila_(sh, COL_FECHA) - FILA_TOTALES, 0);
}

function buscarFila_(sh, fecha) {
  var n = cantidadCuadres_(sh);
  if (n < 1) return -1;
  var fechas = sh.getRange(FILA_DATOS, COL_FECHA, n, 1).getValues();
  for (var i = 0; i < fechas.length; i++) {
    if (String(fechas[i][0]) === fecha) return i + FILA_DATOS;
  }
  return -1;
}

function guardarCuadre_(d) {
  var fecha = String(d.fecha || '');
  if (!fechaValida_(fecha)) return { ok: false, error: 'fecha' };
  var neg = String(d.neg || '');
  if (!fechaValida_(neg)) return { ok: false, error: 'sin_negociacion' };
  var s = d.s || {};
  var staff = (d.detalle && d.detalle.staff) || [];
  if (staff.length > 80) return { ok: false, error: 'detalle_grande' };
  var nota = 'Planilla del ' + fecha;
  staff.forEach(function (x) {
    var nombre = String(x[0] || 'Sin nombre').replace(/[\r\n|]+/g, ' ').replace(/ +/g, ' ').trim();
    nota += '\n' + nombre + ' | ' + num_(x[1]).toFixed(2);
  });
  var v = (d.detalle && d.detalle.v) || {};
  var g = (d.detalle && d.detalle.g) || {};
  var fila = [fecha, num_(v.sistema), num_(v.efectivo), num_(v.taqEf), num_(v.taqPos), num_(v.passline),
    num_(s.venta), num_(s.staff), num_(g.hielo), num_(g.frutas), num_(g.cort),
    num_(s.igv), num_(s.alquiler), num_(s.servicio), num_(s.licor), num_(s.visa), num_(s.tpass),
    num_(s.gastos), num_(s.util), num_(s.pago), num_(s.casa), num_(s.asis), num_(s.ticket),
    neg, new Date()];
  var sh = hoja_();
  var r = buscarFila_(sh, fecha);
  var nueva = r === -1;
  if (nueva) r = Math.max(ultimaFila_(sh, COL_FECHA), FILA_TOTALES) + 1;
  sh.getRange(r, COL_FECHA).setNumberFormat('@');
  sh.getRange(r, COL_NEG).setNumberFormat('@');
  sh.getRange(r, 1, 1, fila.length).setValues([fila]);
  sh.getRange(r, COL_PLANILLA).setNote(nota);
  sh.getRange(r, 2, 1, 21).setNumberFormat('#,##0.00');
  sh.getRange(r, COL_ASIS).setNumberFormat('0');
  sh.getRange(r, COL_TICKET).setNumberFormat('#,##0.00');
  sh.getRange(r, COL_GUARDADO).setNumberFormat('dd/mm/yyyy hh:mm');
  return { ok: true, nueva: nueva, fila: r };
}

function listarCuadres_(max) {
  var sh = hoja_();
  var n = cantidadCuadres_(sh);
  if (n < 1) return [];
  var vals = sh.getRange(FILA_DATOS, 1, n, COL_GUARDADO).getValues();
  var items = vals.map(function (v) {
    var g = v[COL_GUARDADO - 1];
    return {
      fecha: String(v[0]), venta: num_(v[6]), util: num_(v[18]), pago: num_(v[19]),
      guardado: g instanceof Date ? g.toISOString() : String(g)
    };
  });
  items.sort(function (a, b) { return a.fecha < b.fecha ? 1 : a.fecha > b.fecha ? -1 : 0; });
  return items.slice(0, max);
}

function obtenerCuadre_(fecha) {
  var sh = hoja_();
  var r = buscarFila_(sh, fecha);
  if (r === -1) return { ok: false, error: 'no_existe' };
  var f = sh.getRange(r, 1, 1, COL_GUARDADO).getValues()[0];
  var neg = String(f[COL_NEG - 1] || '');
  var staff = [];
  String(sh.getRange(r, COL_PLANILLA).getNote() || '').split('\n').slice(1).forEach(function (linea) {
    var k = linea.lastIndexOf(' | ');
    if (k > -1) staff.push([linea.substring(0, k), num_(linea.substring(k + 3))]);
  });
  var cfg = neg ? leerConfig_(neg) : { terms: {} };
  var terms = {};
  Object.keys(cfg.terms || {}).forEach(function (k) { if (/^t_/.test(k)) terms[k] = cfg.terms[k]; });
  return {
    ok: true, neg: neg,
    detalle: {
      v: { sistema: num_(f[1]), efectivo: num_(f[2]), taqEf: num_(f[3]), taqPos: num_(f[4]), passline: num_(f[5]) },
      g: { hielo: num_(f[8]), frutas: num_(f[9]), cort: num_(f[10]) },
      staff: staff, asis: num_(f[COL_ASIS - 1]), terms: terms
    }
  };
}

// ---------------------------------------------------------------- hojas "Negociación AAAA-MM-DD"

// Fechas de inicio de todas las versiones, ordenadas de la mas antigua a la mas reciente.
function versiones_() {
  var out = [];
  SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(function (sh) {
    var n = sh.getName();
    if (n.indexOf(PREFIJO_NEG) === 0 && fechaValida_(n.substring(PREFIJO_NEG.length))) {
      out.push(n.substring(PREFIJO_NEG.length));
    }
  });
  out.sort();
  return out;
}

// Version vigente para una fecha: la mas reciente con inicio <= fecha. Si la fecha es anterior a
// todas las versiones, se usa la primera.
function vigente_(fecha) {
  var vs = versiones_();
  if (!vs.length) return '';
  var elegida = vs[0];
  for (var i = 0; i < vs.length; i++) {
    if (vs[i] <= fecha) elegida = vs[i];
  }
  return elegida;
}

function leerConfig_(fecha) {
  var vs = versiones_();
  if (fecha !== '' && !fechaValida_(fecha)) return { ok: false, error: 'fecha' };
  var desde = vigente_(fecha || '9999-12-31');
  if (!desde) return { ok: true, terms: {}, desde: '', versiones: [] };
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PREFIJO_NEG + desde);
  var n = Math.max(ultimaFila_(sh, 1) - 1, 0);
  var out = {};
  if (n > 0) {
    var vals = sh.getRange(2, 1, n, 2).getValues();
    for (var i = 0; i < vals.length; i++) {
      if (vals[i][0] !== '') out[String(vals[i][0])] = vals[i][1];
    }
  }
  return { ok: true, terms: out, desde: desde, versiones: vs };
}

// Crea (o actualiza, si ya existe esa fecha) la hoja de una version de la negociacion.
function guardarNegociacion_(obj) {
  var desde = String(obj.desde || '');
  if (!fechaValida_(desde)) return { ok: false, error: 'fecha' };
  var filas = [];
  Object.keys(obj).forEach(function (k) {
    if (/^t_[a-z_]+$/.test(k)) filas.push([k, num_(obj[k])]);
  });
  if (!filas.length) return { ok: false, error: 'sin_terminos' };
  if (typeof obj.planilla === 'string' && obj.planilla.length <= 4000) {
    filas.push(['planilla', obj.planilla]);
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var nombre = PREFIJO_NEG + desde;
  var sh = ss.getSheetByName(nombre);
  var nueva = !sh;
  if (nueva) sh = ss.insertSheet(nombre);
  var previa = ultimaFila_(sh, 1);
  if (previa > 0) sh.getRange(1, 1, previa, 2).clearContent();
  sh.getRange(1, 1, 1, 2).setValues([['Clave', 'Valor']]).setFontWeight('bold');
  sh.getRange(2, 1, filas.length, 2).setValues(filas);
  sh.getRange(1, 4).setValue('Rige desde');
  sh.getRange(1, 5).setNumberFormat('@').setValue(desde);
  return { ok: true, nueva: nueva, desde: desde };
}
