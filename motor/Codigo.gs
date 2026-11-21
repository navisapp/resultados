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
var COL = { clave: 1, cliente: 2, rif: 3, carpeta: 4, tablero: 5, codigo: 6, inicio: 7, metodo: 8, activo: 9, ultimo: 10, estado: 11, actualizado: 12,
  inv: 13, invIni: 14, invFecha: 15, invPct: 16, invBase: 17, invReal: 18, prueba: 19 };
var NCOL = 18;
var INV_HEAD = ['¿Lleva inventario?', 'Inventario inicial (Bs)', 'Fecha inventario inicial', '% utilidad bruta', 'Utilidad sobre', 'Inventario real (opcional)'];
var INV_NOTA = [
  'SI: el cliente vende mercancía con inventario periódico y ve la proyección. PRUEBA: se calcula pero el cliente NO la ve (tú la ves con ?simular al final del link). NO o vacío: servicios / sin inventario.',
  'Monto del inventario al inicio del ejercicio, en bolívares (sale del balance de cierre anterior).',
  'Fecha de ese inventario inicial. Vacío = primer día del ejercicio.',
  'Porcentaje de utilidad bruta para proyectar el costo de ventas. Vacío = 45.',
  'costo = el % se aplica sobre el costo (ej. 45% sobre lo que cuesta). ventas = el % es el margen sobre las ventas.',
  'Conteos físicos reales, si los hay. Formato: MM/AAAA = monto; separados por punto y coma. Ej.: 06/2026 = 28.116.287,65'];

function asegurarColumnas_(sh) {
  protegerColumnas_(sh);
  var n = Math.max(sh.getMaxRows() - 1, 1);
  if (!String(sh.getRange(1, COL.prueba).getValue()).trim()) {
    sh.getRange(1, COL.prueba).setValue('Probar % (solo NT)').setNote('Link para abrir el tablero con la casilla de prueba del %. Solo para uso interno: no se lo envíes al cliente. Aparece cuando la columna M dice SI o PRUEBA.');
    sh.getRange(1, 1).copyFormatToRange(sh, COL.prueba, COL.prueba, 1, 1);
    sh.setColumnWidth(COL.prueba, 260);
  }
  if (sh.getRange(1, COL.inv).getNote().indexOf('PRUEBA') < 0) {
    sh.getRange(1, COL.inv).setNote(INV_NOTA[0]);
    sh.getRange(2, COL.inv, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['SI', 'PRUEBA', 'NO'], true).setAllowInvalid(true).build());
  }
  if (String(sh.getRange(1, COL.inv).getValue()).trim()) return;
  var h = sh.getRange(1, COL.inv, 1, INV_HEAD.length);
  h.setValues([INV_HEAD]).setNotes([INV_NOTA]);
  sh.getRange(1, 1).copyFormatToRange(sh, COL.inv, COL.invReal, 1, 1);
  sh.getRange(2, COL.inv, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['SI', 'PRUEBA', 'NO'], true).setAllowInvalid(true).build());
  sh.getRange(2, COL.invBase, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['costo', 'ventas'], true).setAllowInvalid(true).build());
  sh.getRange(2, COL.invIni, n, 1).setNumberFormat('#,##0.00');
  sh.getRange(2, COL.invFecha, n, 1).setNumberFormat('dd/mm/yyyy');
  sh.setColumnWidths(COL.inv, INV_HEAD.length, 150);
}

/** Aviso (sin clave) al editar las columnas que llena el sistema. El programa sí puede escribirlas. */
function protegerColumnas_(sh) {
  var ya = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).some(function (p) { return p.getDescription().indexOf('NT:') === 0; });
  if (ya) return;
  var cols = [['A', 'clave del link'], ['D', 'link de la carpeta'], ['E', 'link del tablero'], ['F', 'código de acceso'],
    ['J', 'último mes'], ['K', 'estado'], ['L', 'última actualización'], ['S', 'link de prueba']];
  cols.forEach(function (c) { sh.getRange(c[0] + ':' + c[0]).protect().setDescription('NT: ' + c[1] + ' (lo llena el sistema)').setWarningOnly(true); });
  sh.getRange('1:1').protect().setDescription('NT: títulos de columnas').setWarningOnly(true);
}

/** Lee la configuración de inventario de una fila de la hoja */
function invDeFila_(r) {
  var mod = String(r[COL.inv - 1]).toUpperCase().trim();
  if (mod.indexOf('S') !== 0 && mod.indexOf('P') !== 0) return null;
  var numVE = function (v) { if (typeof v === 'number') return v; var s = String(v).replace(/[^\d,.\-]/g, ''); if (!s) return 0; return +(s.replace(/\./g, '').replace(',', '.')) || 0; };
  var f = r[COL.invFecha - 1], fecha = '';
  if (f instanceof Date) fecha = Utilities.formatDate(f, 'America/Caracas', 'yyyy-MM-dd');
  else { var m = /(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(String(f)); if (m) fecha = m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2); }
  var real = {};
  String(r[COL.invReal - 1] || '').split(';').forEach(function (p) {
    var m = /(\d{1,2})\s*[\/\-]\s*(\d{4})\s*[=:]\s*(.+)/.exec(p.trim());
    if (m) real[m[2] + '-' + ('0' + m[1]).slice(-2)] = numVE(m[3]);
  });
  var pct = String(r[COL.invPct - 1]).trim() === '' ? 45 : numVE(r[COL.invPct - 1]);
  if (pct > 0 && pct < 1 && typeof r[COL.invPct - 1] === 'number') pct = pct * 100;  // celda con formato %
  return { lleva: true, ii: numVE(r[COL.invIni - 1]), iiFecha: fecha, pct: pct,
    base: String(r[COL.invBase - 1]).toLowerCase().indexOf('vent') === 0 ? 'ventas' : 'costo', real: real, sim: false, oculto: mod.indexOf('P') === 0 };
}

/* ===================== MENÚ ===================== */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('NT Resultados')
    .addItem('Actualizar todos los clientes', 'actualizarTodos')
    .addItem('Actualizar solo la fila seleccionada', 'actualizarSeleccionado')
    .addSeparator()
    .addItem('Reprocesar todo desde cero (fila seleccionada)', 'reprocesarSeleccionado')
    .addItem('Activar actualización automática cada noche', 'activarNocturno')
    .addSeparator()
    .addItem('Eliminar cliente (fila seleccionada)', 'eliminarSeleccionado')
    .addToUi();
}
function actualizarTodos() { procesar_(null, false); }
function actualizarSeleccionado() { var r = SpreadsheetApp.getActiveSheet().getActiveRange().getRow(); if (r > 1) procesar_(r, false); }
function reprocesarSeleccionado() { var r = SpreadsheetApp.getActiveSheet().getActiveRange().getRow(); if (r > 1) procesar_(r, true); }
function activarNocturno() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'actualizarTodos') ScriptApp.deleteTrigger(t); });
  [1, 2, 3, 4].forEach(function (h) { ScriptApp.newTrigger('actualizarTodos').timeBased().everyDays(1).atHour(h).create(); });
  try { SpreadsheetApp.getUi().alert('Listo: se actualizará cada madrugada (1 a 4 a. m.). Los clientes sin cambios no se vuelven a publicar.'); } catch (e) {}
}

/** Elimina el cliente de la fila seleccionada: tablero publicado, carpeta (a la papelera) y fila. Pide confirmación. */
function eliminarSeleccionado() {
  var ui = SpreadsheetApp.getUi(), sh = SpreadsheetApp.getActiveSheet();
  if (sh.getSheetId() !== hoja_().getSheetId()) { ui.alert('Ve a la pestaña de clientes y selecciona la fila del cliente.'); return; }
  var fila = sh.getActiveRange().getRow();
  if (fila < 2) { ui.alert('Selecciona la fila del cliente que quieres eliminar.'); return; }
  var r = sh.getRange(fila, 1, 1, NCOL).getValues()[0];
  var nombre = String(r[COL.cliente - 1]).trim() || '(sin nombre)', clave = String(r[COL.clave - 1]).trim();
  if (clave && !/^[a-z0-9][a-z0-9-]*$/.test(clave) || clave === 'motor') { ui.alert('La clave "' + clave + '" no es válida; no elimino nada.'); return; }
  var resp = ui.alert('Eliminar cliente',
    'Vas a eliminar a ' + nombre + ':\n\n' +
    (clave ? '• Se borra su tablero (' + CFG.LINK_BASE + clave + '/) y el link deja de funcionar.\n' : '') +
    '• Su carpeta de Drive va a la papelera (se puede recuperar durante 30 días).\n' +
    '• Se borra la fila ' + fila + ' de esta hoja.\n\n¿Continuar?', ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;
  var errs = [];
  if (clave) { try { borrarGithub_(clave + '/index.html', 'Elimina tablero ' + clave); } catch (e) { errs.push('tablero: ' + e.message); } }
  if (String(r[COL.carpeta - 1]).trim()) {
    try { var fid = folderId_(r[COL.carpeta - 1]); if (fid !== CFG.ROOT_ID && fid !== CFG.TASAS_ID) DriveApp.getFolderById(fid).setTrashed(true); }
    catch (e) { errs.push('carpeta: ' + e.message); }
  }
  if (errs.length) { ui.alert('No se completó (' + errs.join(' | ') + '). La fila NO se borró; avísame.'); return; }
  sh.deleteRow(fila);
  ui.alert(nombre + ' fue eliminado.');
}

/* ===================== PROCESO ===================== */
function hoja_() { return SpreadsheetApp.openById(CFG.SHEET_ID).getSheets()[0]; }

/** Si una actualización no alcanzó a terminar, sigue sola un minuto después desde donde quedó */
function continuarActualizacion() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'continuarActualizacion') ScriptApp.deleteTrigger(t); });
  var desde = +PropertiesService.getScriptProperties().getProperty('CONTINUAR_DESDE') || 2;
  procesar_(null, false, desde);
}
function programarContinuacion_(fila) {
  PropertiesService.getScriptProperties().setProperty('CONTINUAR_DESDE', String(fila));
  var ya = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'continuarActualizacion'; });
  if (!ya) ScriptApp.newTrigger('continuarActualizacion').timeBased().after(60 * 1000).create();
}

function procesar_(soloFila, desdeCero, desdeFila) {
  var t0 = Date.now(), sh = hoja_(), last = sh.getLastRow(), pendiente = 0;
  if (last < 2) return;
  var tasas = cargarTasas_();
  asegurarColumnas_(sh);
  var rows = sh.getRange(2, 1, last - 1, NCOL).getValues();
  for (var i = 0; i < rows.length; i++) {
    var fila = i + 2;
    if (soloFila && fila !== soloFila) continue;
    if (desdeFila && fila < desdeFila) continue;
    if (Date.now() - t0 > CFG.MAX_MS) { if (!pendiente) pendiente = fila; sh.getRange(fila, COL.estado).setValue('Pendiente: sigue sola en 1 minuto'); continue; }
    var r = rows[i];
    if (!String(r[COL.cliente - 1]).trim()) continue;
    if (String(r[COL.activo - 1]).toUpperCase().trim() === 'NO') continue;
    try {
      completarFila_(sh, fila, r);
      r = sh.getRange(fila, 1, 1, NCOL).getValues()[0];
      var res = procesarCliente_(r, tasas, desdeCero);
      sh.getRange(fila, COL.ultimo, 1, 3).setValues([[res.ultimo, res.estado, new Date()]]);
      sh.getRange(fila, COL.prueba).setValue(invDeFila_(r) && String(r[COL.tablero - 1]).trim() ? String(r[COL.tablero - 1]).trim() + '?simular' : '');
    } catch (e) {
      sh.getRange(fila, COL.estado, 1, 2).setValues([['ERROR: ' + e.message, new Date()]]);
    }
  }
  escribirTasas_(tasas);
  if (pendiente && !soloFila) { try { programarContinuacion_(pendiente); } catch (e) {} }
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
  else {  // la carpeta lleva el mismo nombre que la columna Cliente
    try {
      var fid = folderId_(r[COL.carpeta - 1]), nb = String(r[COL.cliente - 1]).trim();
      if (nb && fid !== CFG.ROOT_ID && fid !== CFG.TASAS_ID) { var fo = DriveApp.getFolderById(fid); if (fo.getName() !== nb) fo.setName(nb); }
    } catch (e) {}
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
    var esCat = /cat[aá]logo|plan\s*de\s*cuentas?/i.test(name), celdas = null;
    if (!esCat) {  // por si el catálogo viene con otro nombre: se reconoce por su contenido
      var prevF = datos.files[id];
      if (!(prevF && prevF.mod === mod) && !(datos.cat && datos.cat.id === id && datos.cat.mod === mod)) {
        celdas = leerCeldas_(f);
        if (!NT.parseMayor(celdas).periodo && Object.keys(NT.parseCatalog(celdas).nombres).length >= 10) esCat = true;
      } else if (datos.cat && datos.cat.id === id) esCat = true;
    }
    if (esCat) {
      if (!datos.cat || datos.cat.id !== id || datos.cat.mod !== mod) {
        var c = NT.parseCatalog(celdas || leerCeldas_(f));
        datos.cat = { id: id, mod: mod, nombres: c.nombres, empresa: c.empresa, rif: c.rif }; hayCambios = true;
      }
      continue;
    }
    vistos[id] = true;
    var prev = datos.files[id];
    if (prev && prev.mod === mod) continue;
    var L = NT.parseMayor(celdas || leerCeldas_(f));
    if (!L.periodo) throw new Error('El archivo "' + name + '" no parece un Libro Mayor de Galac ni un Mayor Analítico de Profit (no encontré el período).');
    datos.files[id] = { name: name, mod: mod, periodo: L.periodo, descuadres: L.descuadres, empresa: L.empresa, rif: L.rif, sistema: L.sistema || 'Galac', nombres: L.nombres || null,
      movs: L.movs.map(function (m) { return [m.a, m.an, m.f, m.c, m.t, m.n, m.p, m.r, m.d, m.db, m.hb]; }) };
    hayCambios = true;
  }
  Object.keys(datos.files).forEach(function (id) { if (!vistos[id]) { delete datos.files[id]; hayCambios = true; } });
  var ids = Object.keys(datos.files);
  if (!ids.length) return { ultimo: '', estado: 'Sin mayor cargado: sube el Libro Mayor de Galac a la carpeta' };

  // cada mes y grupo de cuentas (4, 5, 6, 7) lo aporta el archivo más reciente que lo cubre
  // (así sirven un mayor completo o varios archivos por grupo, como los de Profit)
  ids.sort(function (a, b) { return datos.files[a].mod - datos.files[b].mod; });
  var dueño = {};
  ids.forEach(function (id) {
    var p = datos.files[id].periodo, gs = {};
    datos.files[id].movs.forEach(function (x) { gs[String(x[0]).charAt(0)] = 1; });
    mesesEntre_(p.desde, p.hasta).forEach(function (m) { Object.keys(gs).forEach(function (g) { dueño[m + '|' + g] = id; }); });
  });
  var movs = [];
  ids.forEach(function (id) {
    datos.files[id].movs.forEach(function (x) {
      if (dueño[x[2].slice(0, 7) + '|' + String(x[0]).charAt(0)] === id) movs.push({ a: x[0], an: x[1], f: x[2], c: x[3], t: x[4], n: x[5], p: x[6], r: x[7], d: x[8], db: x[9], hb: x[10] });
    });
  });
  var desc = [];
  ids.forEach(function (id) { if (datos.files[id].descuadres.length) desc = desc.concat(datos.files[id].descuadres); });

  var cat = datos.cat || { nombres: {}, empresa: '', rif: '' };
  var sistema = ids.some(function (id) { return datos.files[id].sistema === 'Profit'; }) ? 'Profit' : 'Galac';
  if (!datos.cat) { var nb = {}; ids.forEach(function (id) { var n = datos.files[id].nombres || {}; Object.keys(n).forEach(function (k) { nb[k] = n[k]; }); }); cat = { nombres: nb, empresa: '', rif: '' }; }
  var meta = { empresa: String(r[COL.cliente - 1]).trim() || cat.empresa || datos.files[ids[0]].empresa, rif: String(r[COL.rif - 1]).replace(/[^A-Za-z0-9]/g, '').toUpperCase(),
    inicioMes: +String(r[COL.inicio - 1]).replace(/\D/g, '') || 1, hoy: Utilities.formatDate(new Date(), 'America/Caracas', 'yyyy-MM-dd'),
    inv: invDeFila_(r) };
  var pkg = NT.buildPackage(movs, cat.nombres, tasas.tasas, meta);
  pkg.metodo = String(r[COL.metodo - 1] || 'cierre').toLowerCase().indexOf('prom') === 0 ? 'p' : (String(r[COL.metodo - 1]).toLowerCase().indexOf('dia') === 0 ? 'd' : 'c');
  guardarJson_(folder, CFG.PKG_NAME, pkg);

  // Publicar en GitHub solo si cambió algo (datos, tasas, código o método)
  var pub = '';
  var huella = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5,
    JSON.stringify([pkg.cols, pkg.mr, pkg.alerts, pkg.metodo, pkg.names, pkg.inv, pkg.empresa]) + '|' + r[COL.codigo - 1] + '|' + r[COL.rif - 1] + '|plantilla-5'));
  if (datos.pubHash !== huella) {
    try { publicarTablero_(String(r[COL.clave - 1]), r[COL.rif - 1], r[COL.codigo - 1], pkg); datos.pubHash = huella; pub = ' · publicado'; }
    catch (e) { pub = ' · NO PUBLICADO: ' + e.message; }
  } else pub = ' · sin cambios';
  guardarJson_(folder, CFG.DATA_NAME, datos);

  var faltan = pkg.alerts.filter(function (a) { return a.s === 1; }).length;
  var estado = (desc.length ? 'REVISAR: no cuadra con ' + sistema + ' en ' + desc.join(', ') : 'OK · cuadra con ' + sistema) +
    ' · ' + pkg.alerts.length + ' alertas (' + faltan + ' faltantes)' + (datos.cat || sistema === 'Profit' ? '' : ' · falta el Catálogo de Cuentas') +
    (pkg.inv ? ' · inventario ' + (pkg.inv.oculto ? 'EN PRUEBA (el cliente no lo ve) ' : '') + pkg.inv.pct + '% s/' + pkg.inv.base + (pkg.inv.ii ? '' : ' (FALTA inventario inicial)') + (pkg.inv.c.length ? '' : ' (no encontré cuentas de compras)') : '') + pub;
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
