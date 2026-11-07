/**
 * NT Resultados — motor en Google Drive (NT Consultores & Asociados)
 * Lee los mayores de Galac de cada carpeta de cliente, arma el tablero y lo sirve con clave RIF + código.
 */
var CFG = {
  ROOT_ID: '11a7LeS8oytqPzHZ0szZ68CQFGpXR0k0X',     // carpeta NT Resultados
  TASAS_ID: '1peQeE8b0kmZnAa3X5Nuu6q57ukjBMpkv',    // carpeta _TASAS
  SHEET_ID: '1V-kFHv1aBVUvBcMp_Ggd6ciZ0ytRLVrEM1sX-tY2K9g', // hoja Clientes
  LINK_BASE: 'https://navisapp.github.io/resultados/',
  DATA_NAME: '_datos (no tocar).json',
  PKG_NAME: '_tablero (no tocar).json',
  MAX_MS: 270000,          // 4,5 minutos por ejecución
  MAX_FALLOS: 8            // intentos fallidos de clave antes de bloquear 15 minutos
};
var COL = { clave: 1, cliente: 2, rif: 3, carpeta: 4, tablero: 5, codigo: 6, inicio: 7, metodo: 8, activo: 9, ultimo: 10, estado: 11, actualizado: 12 };

/* ===================== MENÚ ===================== */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('NT Resultados')
    .addItem('Actualizar todos los clientes', 'actualizarTodos')
    .addItem('Actualizar solo la fila seleccionada', 'actualizarSeleccionado')
    .addSeparator()
    .addItem('Reprocesar todo desde cero (fila seleccionada)', 'reprocesarSeleccionado')
    .addItem('Activar actualización automática cada noche', 'activarNocturno')
    .addToUi();
}
function actualizarTodos() { procesar_(null, false); }
function actualizarSeleccionado() { var r = SpreadsheetApp.getActiveSheet().getActiveRange().getRow(); if (r > 1) procesar_(r, false); }
function reprocesarSeleccionado() { var r = SpreadsheetApp.getActiveSheet().getActiveRange().getRow(); if (r > 1) procesar_(r, true); }
function activarNocturno() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'actualizarTodos') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('actualizarTodos').timeBased().everyDays(1).atHour(2).create();
  try { SpreadsheetApp.getUi().alert('Listo: se actualizará todos los días a las 2 a. m.'); } catch (e) {}
}

/* ===================== PROCESO ===================== */
function hoja_() { return SpreadsheetApp.openById(CFG.SHEET_ID).getSheets()[0]; }

function procesar_(soloFila, desdeCero) {
  var t0 = Date.now(), sh = hoja_(), last = sh.getLastRow();
  if (last < 2) return;
  var tasas = cargarTasas_();
  var rows = sh.getRange(2, 1, last - 1, 12).getValues();
  for (var i = 0; i < rows.length; i++) {
    var fila = i + 2;
    if (soloFila && fila !== soloFila) continue;
    if (Date.now() - t0 > CFG.MAX_MS) { sh.getRange(fila, COL.estado).setValue('Pendiente: se completa en la próxima actualización'); continue; }
    var r = rows[i];
    if (!String(r[COL.cliente - 1]).trim()) continue;
    if (String(r[COL.activo - 1]).toUpperCase().trim() === 'NO') continue;
    try {
      completarFila_(sh, fila, r);
      r = sh.getRange(fila, 1, 1, 12).getValues()[0];
      var res = procesarCliente_(r, tasas, desdeCero);
      sh.getRange(fila, COL.ultimo, 1, 3).setValues([[res.ultimo, res.estado, new Date()]]);
    } catch (e) {
      sh.getRange(fila, COL.estado, 1, 2).setValues([['ERROR: ' + e.message, new Date()]]);
    }
  }
  escribirTasas_(tasas);
}

/** Completa clave, código, carpeta y link para clientes nuevos */
function completarFila_(sh, fila, r) {
  var cambios = false;
  if (!String(r[COL.clave - 1]).trim()) { r[COL.clave - 1] = slug_(r[COL.cliente - 1]); cambios = true; }
  if (!String(r[COL.codigo - 1]).trim()) { r[COL.codigo - 1] = codigo_(); cambios = true; }
  if (!String(r[COL.carpeta - 1]).trim()) {
    var f = DriveApp.getFolderById(CFG.ROOT_ID).createFolder(String(r[COL.cliente - 1]).trim());
    r[COL.carpeta - 1] = f.getUrl(); cambios = true;
  }
  if (!String(r[COL.tablero - 1]).trim()) { r[COL.tablero - 1] = CFG.LINK_BASE + r[COL.clave - 1] + '/'; cambios = true; }
  if (!String(r[COL.inicio - 1]).trim()) { r[COL.inicio - 1] = '01'; cambios = true; }
  if (!String(r[COL.metodo - 1]).trim()) { r[COL.metodo - 1] = 'cierre'; cambios = true; }
  if (!String(r[COL.activo - 1]).trim()) { r[COL.activo - 1] = 'SI'; cambios = true; }
  if (cambios) sh.getRange(fila, 1, 1, 9).setValues([r.slice(0, 9)]);
}
function slug_(n) {
  return String(n).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/\b(c\.?\s?a\.?|s\.?\s?a\.?|srl)\b/g, '').replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 30);
}
function codigo_() {
  var A = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789', c = '';
  for (var i = 0; i < 10; i++) c += A.charAt(Math.floor(Math.random() * A.length));
  return c.slice(0, 5) + '-' + c.slice(5);
}
function folderId_(url) { var m = /folders\/([A-Za-z0-9_-]+)/.exec(String(url)) || /id=([A-Za-z0-9_-]+)/.exec(String(url)); return m ? m[1] : String(url).trim(); }

function procesarCliente_(r, tasas, desdeCero) {
  var folder = DriveApp.getFolderById(folderId_(r[COL.carpeta - 1]));
  var datos = desdeCero ? { files: {}, cat: null } : (leerJson_(folder, CFG.DATA_NAME) || { files: {}, cat: null });
  var vistos = {}, it = folder.getFiles(), hayCambios = desdeCero;
  while (it.hasNext()) {
    var f = it.next(), name = f.getName(), id = f.getId(), mod = f.getLastUpdated().getTime();
    if (name.charAt(0) === '_') continue;
    if (!/\.(xls|xlsx)$/i.test(name) && f.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;
    var esCat = /cat[aá]logo/i.test(name);
    if (esCat) {
      if (!datos.cat || datos.cat.id !== id || datos.cat.mod !== mod) {
        var c = NT.parseCatalog(leerCeldas_(f));
        datos.cat = { id: id, mod: mod, nombres: c.nombres, empresa: c.empresa, rif: c.rif }; hayCambios = true;
      }
      continue;
    }
    vistos[id] = true;
    var prev = datos.files[id];
    if (prev && prev.mod === mod) continue;
    var L = NT.parseLedger(leerCeldas_(f));
    if (!L.periodo) throw new Error('El archivo "' + name + '" no parece un Libro Mayor de Galac (no encontré el período).');
    datos.files[id] = { name: name, mod: mod, periodo: L.periodo, descuadres: L.descuadres, empresa: L.empresa, rif: L.rif,
      movs: L.movs.map(function (m) { return [m.a, m.an, m.f, m.c, m.t, m.n, m.p, m.r, m.d, m.db, m.hb]; }) };
    hayCambios = true;
  }
  Object.keys(datos.files).forEach(function (id) { if (!vistos[id]) { delete datos.files[id]; hayCambios = true; } });
  var ids = Object.keys(datos.files);
  if (!ids.length) return { ultimo: '', estado: 'Sin mayor cargado: sube el Libro Mayor de Galac a la carpeta' };

  // cada mes lo aporta el archivo más reciente que lo cubre
  ids.sort(function (a, b) { return datos.files[a].mod - datos.files[b].mod; });
  var dueño = {};
  ids.forEach(function (id) { var p = datos.files[id].periodo; mesesEntre_(p.desde, p.hasta).forEach(function (m) { dueño[m] = id; }); });
  var movs = [];
  ids.forEach(function (id) {
    datos.files[id].movs.forEach(function (x) {
      if (dueño[x[2].slice(0, 7)] === id) movs.push({ a: x[0], an: x[1], f: x[2], c: x[3], t: x[4], n: x[5], p: x[6], r: x[7], d: x[8], db: x[9], hb: x[10] });
    });
  });
  var desc = [];
  ids.forEach(function (id) { if (datos.files[id].descuadres.length) desc = desc.concat(datos.files[id].descuadres); });

  var cat = datos.cat || { nombres: {}, empresa: '', rif: '' };
  var meta = { empresa: cat.empresa || datos.files[ids[0]].empresa || r[COL.cliente - 1], rif: String(r[COL.rif - 1]).replace(/[^A-Za-z0-9]/g, '').toUpperCase(),
    inicioMes: +String(r[COL.inicio - 1]).replace(/\D/g, '') || 1, hoy: Utilities.formatDate(new Date(), 'America/Caracas', 'yyyy-MM-dd') };
  var pkg = NT.buildPackage(movs, cat.nombres, tasas.tasas, meta);
  pkg.metodo = String(r[COL.metodo - 1] || 'cierre').toLowerCase().indexOf('prom') === 0 ? 'p' : (String(r[COL.metodo - 1]).toLowerCase().indexOf('dia') === 0 ? 'd' : 'c');
  guardarJson_(folder, CFG.PKG_NAME, pkg);

  // Publicar en GitHub solo si cambió algo (datos, tasas, código o método)
  var pub = '';
  var huella = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5,
    JSON.stringify([pkg.cols, pkg.mr, pkg.alerts, pkg.metodo, pkg.names]) + '|' + r[COL.codigo - 1] + '|' + r[COL.rif - 1]));
  if (datos.pubHash !== huella) {
    try { publicarTablero_(String(r[COL.clave - 1]), r[COL.rif - 1], r[COL.codigo - 1], pkg); datos.pubHash = huella; pub = ' · publicado'; }
    catch (e) { pub = ' · NO PUBLICADO: ' + e.message; }
  } else pub = ' · sin cambios';
  guardarJson_(folder, CFG.DATA_NAME, datos);

  var faltan = pkg.alerts.filter(function (a) { return a.s === 1; }).length;
  var estado = (desc.length ? 'REVISAR: no cuadra con Galac en ' + desc.join(', ') : 'OK · cuadra con Galac') +
    ' · ' + pkg.alerts.length + ' alertas (' + faltan + ' faltantes)' + (datos.cat ? '' : ' · falta el Catálogo de Cuentas') + pub;
  return { ultimo: pkg.months[pkg.lastMi], estado: estado };
}
function mesesEntre_(d, h) { var y = +d.slice(0, 4), m = +d.slice(5, 7), out = []; while ((y + '-' + (m < 10 ? '0' : '') + m) <= h.slice(0, 7)) { out.push(y + '-' + (m < 10 ? '0' : '') + m); m++; if (m === 13) { y++; m = 1; } } return out; }

/** Convierte .xls/.xlsx a Hoja de Google temporal y devuelve sus celdas */
function leerCeldas_(file) {
  if (file.getMimeType() === MimeType.GOOGLE_SHEETS) return SpreadsheetApp.openById(file.getId()).getSheets()[0].getDataRange().getValues();
  var tmp = Drive.Files.create({ name: '_tmp ' + file.getName(), mimeType: MimeType.GOOGLE_SHEETS, parents: [CFG.ROOT_ID] }, file.getBlob());
  try { return SpreadsheetApp.openById(tmp.id).getSheets()[0].getDataRange().getValues(); }
  finally { Drive.Files.remove(tmp.id); }
}
function leerJson_(folder, name) { var it = folder.getFilesByName(name); return it.hasNext() ? JSON.parse(it.next().getBlob().getDataAsString()) : null; }
function guardarJson_(folder, name, obj) {
  var txt = JSON.stringify(obj), it = folder.getFilesByName(name);
  if (it.hasNext()) it.next().setContent(txt); else folder.createFile(name, txt, 'application/json');
}

/* ===================== TASAS ===================== */
function cargarTasas_() {
  var folder = DriveApp.getFolderById(CFG.TASAS_ID), it = folder.getFiles(), base = [], corr = [];
  while (it.hasNext()) {
    var f = it.next(), name = f.getName(), t = {};
    if (/\.(xls|xlsx)$/i.test(name) || f.getMimeType() === MimeType.GOOGLE_SHEETS) t = NT.parseRates(leerCeldas_(f), 'dmy');
    else if (/\.(txt|csv)$/i.test(name) || /^text\//.test(f.getMimeType())) {
      var rows = f.getBlob().getDataAsString('UTF-8').split(/\r?\n/).map(function (l) { return l.split(/\t|;/); });
      t = NT.parseRates(rows, 'dmy');
    } else continue;
    (/correcc/i.test(name) ? corr : base).push({ mod: f.getLastUpdated().getTime(), t: t });
  }
  base.sort(function (a, b) { return a.mod - b.mod; });
  var tasas = {};
  base.concat(corr).forEach(function (x) { for (var k in x.t) tasas[k] = x.t[k]; });
  return { tasas: tasas, avisos: NT.rateChecks(tasas) };
}
function escribirTasas_(tasas) {
  var ss = SpreadsheetApp.openById(CFG.SHEET_ID), sh = ss.getSheetByName('Tasas') || ss.insertSheet('Tasas');
  var ks = Object.keys(tasas.tasas).sort(), ult = ks[ks.length - 1];
  sh.clear();
  sh.getRange(1, 1, 4, 2).setValues([
    ['Última tasa cargada', ult ? ult.split('-').reverse().join('/') + ' · Bs ' + tasas.tasas[ult] : 'Sin tasas'],
    ['Días cargados', ks.length],
    ['Saltos mayores a 5% (revisar)', tasas.avisos.length ? tasas.avisos.map(function (k) { return k.split('-').reverse().join('/'); }).join(', ') : 'Ninguno'],
    ['Revisado', new Date()]]);
  sh.setColumnWidth(1, 220); sh.setColumnWidth(2, 520);
}

function include(n) { return HtmlService.createHtmlOutputFromFile(n).getContent(); }
