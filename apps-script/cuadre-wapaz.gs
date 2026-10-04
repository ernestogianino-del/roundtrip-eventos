/**
 * Cuadre Wapaz - receptor en Google Sheets
 *
 * Escribe SOLO en la pestana "cuadre web" de la hoja donde se instala. Si la pestana no
 * existe, la crea. Nunca borra ni modifica las otras pestanas de la hoja.
 *
 * INSTALACION (una sola vez)
 * 1. Abre tu hoja "Proyeccion Viernes - Wapaz" en Google Sheets.
 * 2. Menu Extensiones > Apps Script. Borra lo que haya y pega todo este archivo.
 * 3. En Apps Script: Configuracion del proyecto (icono de engranaje) > Propiedades de
 *    la secuencia de comandos > Agregar propiedad:
 *        nombre: CODIGO     valor: el codigo de acceso que van a usar (ej. una clave de 6+ caracteres)
 * 4. Implementar > Nueva implementacion > tipo "Aplicacion web":
 *        Ejecutar como: Yo
 *        Quien tiene acceso: Cualquier persona
 *    Autoriza los permisos que pida Google y copia la URL que termina en /exec.
 * 5. Pega esa URL en cuadre-wapaz.html, en la constante WEB_APP_URL.
 *
 * El codigo de acceso NO esta en este archivo ni en el repositorio: vive solo en las
 * propiedades del script. Sin el codigo correcto el script no lee ni escribe nada.
 *
 * Si cambias este archivo despues, hay que crear una NUEVA VERSION de la implementacion
 * (Implementar > Administrar implementaciones > editar > Nueva version) para que se aplique.
 *
 * DISENO DE LA PESTANA "cuadre web"
 *   Columnas A a L: un cuadre por fila (la fecha es la clave; guardar de nuevo la misma
 *   fecha actualiza su fila).
 *   Columnas N y O: terminos del acuerdo y planilla base (los usa la pagina).
 */

var HOJA = 'cuadre web';
var ENCABEZADOS = ['Fecha', 'Ingresos', 'Planilla y gastos', 'Varios', 'Total gastos',
  'Utilidad', 'Wapaz', 'Titanium', 'Asistentes', 'Ticket promedio', 'Guardado', 'Detalle'];
var COL_FECHA = 1;
var COL_DETALLE = 12;
var COL_CLAVE = 14;   // N
var COL_VALOR = 15;   // O

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
        return salida_({ ok: true, terms: leerTerminos_() });
      case 'saveConfig':
        lock.waitLock(10000);
        guardarTerminos_(JSON.parse(p.d));
        return salida_({ ok: true });
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

// Devuelve la pestana "cuadre web". La crea si no existe y le pone encabezados si esta vacia.
function hoja_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(HOJA);
  if (!sh) sh = ss.insertSheet(HOJA);
  if (sh.getRange(1, 1).getValue() === '') {
    sh.getRange(1, 1, 1, ENCABEZADOS.length).setValues([ENCABEZADOS]).setFontWeight('bold');
    sh.getRange(1, COL_CLAVE, 1, 2).setValues([['Clave', 'Valor']]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(1, COL_FECHA, sh.getMaxRows(), 1).setNumberFormat('@');
    sh.getRange(1, COL_DETALLE, sh.getMaxRows(), 1).setNumberFormat('@');
    sh.getRange(2, 2, sh.getMaxRows() - 1, 7).setNumberFormat('#,##0.00');
  }
  return sh;
}

// Ultima fila con contenido en UNA columna (getLastRow() cuenta tambien las columnas N y O).
function ultimaFila_(sh, col) {
  var v = sh.getRange(1, col, sh.getMaxRows(), 1).getValues();
  for (var i = v.length - 1; i >= 0; i--) {
    if (v[i][0] !== '') return i + 1;
  }
  return 0;
}

function num_(v) {
  var n = Number(v);
  return isFinite(n) ? n : 0;
}

function leerTerminos_() {
  var sh = hoja_();
  var n = ultimaFila_(sh, COL_CLAVE) - 1;
  var out = {};
  if (n < 1) return out;
  var vals = sh.getRange(2, COL_CLAVE, n, 2).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (vals[i][0] !== '') out[String(vals[i][0])] = vals[i][1];
  }
  return out;
}

function guardarTerminos_(obj) {
  var sh = hoja_();
  var filas = [];
  Object.keys(obj).forEach(function (k) {
    if (/^t_[a-z_]+$/.test(k)) filas.push([k, num_(obj[k])]);
  });
  // Planilla base (lista de personas y pagos) como texto JSON
  if (typeof obj.planilla === 'string' && obj.planilla.length <= 4000) {
    filas.push(['planilla', obj.planilla]);
  }
  var ultima = ultimaFila_(sh, COL_CLAVE);
  if (ultima > 1) sh.getRange(2, COL_CLAVE, ultima - 1, 2).clearContent();
  if (filas.length) sh.getRange(2, COL_CLAVE, filas.length, 2).setValues(filas);
}

function buscarFila_(sh, fecha) {
  var n = ultimaFila_(sh, COL_FECHA) - 1;
  if (n < 1) return -1;
  var fechas = sh.getRange(2, COL_FECHA, n, 1).getValues();
  for (var i = 0; i < fechas.length; i++) {
    if (String(fechas[i][0]) === fecha) return i + 2;
  }
  return -1;
}

function guardarCuadre_(d) {
  var fecha = String(d.fecha || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha)) return { ok: false, error: 'fecha' };
  var s = d.s || {};
  var detalle = JSON.stringify(d.detalle || {});
  if (detalle.length > 40000) return { ok: false, error: 'detalle_grande' };
  var fila = [fecha, num_(s.venta), num_(s.planilla), num_(s.varios), num_(s.gastos),
    num_(s.util), num_(s.pago), num_(s.casa), num_(s.asis), num_(s.ticket),
    new Date(), detalle];
  var sh = hoja_();
  var r = buscarFila_(sh, fecha);
  var nueva = r === -1;
  if (nueva) r = Math.max(ultimaFila_(sh, COL_FECHA), 1) + 1;
  sh.getRange(r, COL_FECHA).setNumberFormat('@');
  sh.getRange(r, COL_DETALLE).setNumberFormat('@');
  sh.getRange(r, 1, 1, fila.length).setValues([fila]);
  sh.getRange(r, 2, 1, 7).setNumberFormat('#,##0.00');
  sh.getRange(r, 11).setNumberFormat('dd/mm/yyyy hh:mm');
  return { ok: true, nueva: nueva, fila: r };
}

function listarCuadres_(max) {
  var sh = hoja_();
  var n = ultimaFila_(sh, COL_FECHA) - 1;
  if (n < 1) return [];
  var vals = sh.getRange(2, 1, n, 11).getValues();
  var items = vals.map(function (v) {
    var g = v[10];
    return {
      fecha: String(v[0]), venta: num_(v[1]), util: num_(v[5]), pago: num_(v[6]),
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
  var txt = sh.getRange(r, COL_DETALLE).getValue();
  return { ok: true, detalle: JSON.parse(String(txt || '{}')) };
}
