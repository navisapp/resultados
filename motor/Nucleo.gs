/**
 * NT Resultados — núcleo de cálculo (sin dependencias de Google, se puede probar fuera).
 * Lee el Libro Mayor y el Catálogo de Galac (como matriz de celdas) y arma el paquete del tablero.
 */
var NT = (function () {
  var DATE_RE = /^(\d{2})\/(\d{2})\/(\d{2,4})$/;
  var DOC_RE = /^(Compra|Venta|Gasto)?\s*(FAC|N\/C|ND|N\/D|NC|FACT)\.?\s*(?:N[º°o]\.?)?\s*([A-Z0-9][A-Z0-9\-]*)?\s*(.*)$/i;
  var PLACEHOLDER = { J12345678: 1, J00000000: 1, J000000000: 1, J123456789: 1 };
  var CO_MARKERS = /\b(C\.?\s?A\.?|S\.?\s?A\.?|S\.?R\.?L\.?|SRL|LLC|LTD|LIMITED|INC|CORP\w*|INVERSIONES|COMERCIAL\w*|GRUPO|DISTRIBUIDORA|SERVICIOS?|FUNDACI\w+|ASOCIACI\w+|INSTITUTO|BANCO|SUMINISTROS|INDUSTRIA\w*|TALLER\w*|FARMACIA|SUPERMERCADO|AUTOMERCADO|MULTISERVICIOS|EMPRESA\w*|COOPERATIVA|ALCALD\w+|GOBERNACI\w+|MINISTERIO|SERVICE\w*|SOLUTIONS?|CENTRO|CLINICA|HOTEL|RESTAURANT\w*|PANADERIA|LICORERIA|FERRETERIA|CA)\b/i;

  function s(v) {
    if (v === null || v === undefined) return '';
    if (Object.prototype.toString.call(v) === '[object Date]') {
      return pad(v.getDate()) + '/' + pad(v.getMonth() + 1) + '/' + String(v.getFullYear()).slice(-2);
    }
    if (typeof v === 'number') return (v === Math.floor(v)) ? String(v) : String(v);
    return String(v).trim();
  }
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function num(v) { if (typeof v === 'number') return v; var t = s(v); if (!t) return 0; var x = Number(t.replace(/\./g, '').replace(',', '.')); return isNaN(x) ? 0 : x; }
  function codeStr(v) {
    if (typeof v === 'number') { var t = String(Math.floor(v)); return t.length <= 2 ? t.split('').join('.') : t; }
    var c = s(v); if (/^\d{1,2}(\.0)?$/.test(c)) { c = c.replace('.0', ''); return c.split('').join('.'); }
    return c;
  }
  function isDateCell(v) { return Object.prototype.toString.call(v) === '[object Date]' || DATE_RE.test(s(v)); }
  function isoDate(v) {
    if (Object.prototype.toString.call(v) === '[object Date]') return v.getFullYear() + '-' + pad(v.getMonth() + 1) + '-' + pad(v.getDate());
    var m = DATE_RE.exec(s(v)); var y = m[3].length === 2 ? '20' + m[3] : m[3];
    return y + '-' + m[2] + '-' + m[1];
  }

  /** Catálogo de cuentas: devuelve {nombres:{codigo:nombre}, empresa, rif} */
  function parseCatalog(rows) {
    var cat = {}, empresa = '', rif = '';
    for (var r = 0; r < rows.length; r++) {
      var v = rows[r];
      var t = s(v[2]);
      if (!empresa && t.indexOf(' - RIF ') > 0) { var p = t.split(' - RIF '); empresa = p[0].trim(); rif = p[1].trim(); }
      var c = v[1]; var cs = s(c);
      if (!cs || cs === 'CODIGO') continue;
      var code = codeStr(c);
      if ('4567'.indexOf(code.charAt(0)) >= 0 && /^[\d.]+$/.test(code)) cat[code] = t;
    }
    return { nombres: cat, empresa: empresa, rif: rif };
  }

  /** Libro Mayor: devuelve {periodo:{desde,hasta}, empresa, rif, movs:[...], totales:{} } */
  function parseLedger(rows) {
    var periodo = null, empresa = '', rif = '';
    for (var r0 = 0; r0 < Math.min(8, rows.length); r0++) {
      var t6 = s(rows[r0][6]);
      var pm = /DEL\s+(\d{2})\/(\d{2})\/(\d{4})\s+AL\s+(\d{2})\/(\d{2})\/(\d{4})/i.exec(t6);
      if (pm) periodo = { desde: pm[3] + '-' + pm[2] + '-' + pm[1], hasta: pm[6] + '-' + pm[5] + '-' + pm[4] };
      if (/^RIF /.test(t6)) rif = t6.slice(4).trim();
      if (r0 === 0 && t6) empresa = t6;
    }
    var movs = [], totales = {}, acct = null, acctName = '', cur = null;
    function flush() { if (cur) movs.push(cur); cur = null; }
    for (var r = 0; r < rows.length; r++) {
      var v = rows[r]; var c0 = s(v[0]);
      if (c0 === 'Código') { flush(); acct = codeStr(v[2]); continue; }
      if (c0 === 'Descripción') { acctName = s(v[2]); continue; }
      if (s(v[5]).indexOf('TOTALES') >= 0) { flush(); totales[acct] = { debe: num(v[7]), haber: num(v[9]) }; continue; }
      if (isDateCell(v[0])) {
        flush();
        if (!acct || '4567'.indexOf(acct.charAt(0)) < 0) continue;
        var fi = isoDate(v[0]), cm = /^(\d{2})-/.exec(s(v[1]));
        // Si Google interpretó la fecha al revés (mes/día), se corrige con el mes del comprobante
        if (cm && fi.slice(5, 7) !== cm[1] && +fi.slice(8, 10) <= 12 && fi.slice(8, 10) === cm[1]) fi = fi.slice(0, 5) + fi.slice(8, 10) + '-' + fi.slice(5, 7);
        cur = { a: acct, an: acctName, f: fi, c: s(v[1]), ref: [s(v[3])], desc: [s(v[4])], db: num(v[7]), hb: num(v[9]) };
        continue;
      }
      if (cur && c0 === '' && s(v[5]) === '' && s(v[6]) === '') {
        var d4 = s(v[4]);
        if (d4.indexOf('Van') === 0) continue;
        if (d4 || s(v[3])) { if (s(v[3])) cur.ref.push(s(v[3])); if (d4) cur.desc.push(d4); continue; }
      }
    }
    flush();
    // cuadre contra TOTALES de Galac
    var sums = {}, bad = [];
    movs.forEach(function (m) { if (!sums[m.a]) sums[m.a] = [0, 0]; sums[m.a][0] += m.db; sums[m.a][1] += m.hb; });
    Object.keys(totales).forEach(function (a) {
      if ('4567'.indexOf(a.charAt(0)) < 0) return;
      var sm = sums[a] || [0, 0];
      if (Math.abs(sm[0] - totales[a].debe) > 0.05 || Math.abs(sm[1] - totales[a].haber) > 0.05) bad.push(a);
    });
    return { periodo: periodo, empresa: empresa, rif: rif, movs: movs.map(enrich), totales: totales, descuadres: bad };
  }

  function enrich(m) {
    var desc = m.desc.filter(Boolean).join(' ').replace(/\s+/g, ' ').trim();
    var ref = m.ref.join('');
    var tipo = 'ASIENTO', doc = '', tercero = '';
    var mm = DOC_RE.exec(desc);
    if (mm && mm[2]) {
      tipo = mm[2].toUpperCase().replace(/^NC$/, 'N/C').replace('N/D', 'ND').replace('FACT', 'FAC');
      doc = (mm[3] || '').trim(); tercero = (mm[4] || '').trim();
      if (!doc && tercero) { var i = tercero.indexOf(' '); doc = i > 0 ? tercero.slice(0, i) : tercero; tercero = i > 0 ? tercero.slice(i + 1) : ''; }
      if (doc && !/\d/.test(doc)) { tercero = (doc + ' ' + tercero).trim(); doc = ''; }
    }
    var rif = '';
    if (ref) {
      var tail = ref;
      if (doc && ref.replace(/-/g, '').indexOf(doc.replace(/-/g, '').slice(0, 6)) === 0) {
        var idx = ref.indexOf('-', Math.max(0, doc.length - 2)); tail = idx >= 0 ? ref.slice(idx + 1) : ref;
      }
      var rm = /([VJGEPC])(\d{6,10})$/.exec(tail.replace(/-/g, ''));
      if (rm) rif = rm[1] + rm[2];
    }
    if (!doc && /^[A-Z]?\d+$/.test(ref || '')) doc = ref;
    return { a: m.a, an: m.an, f: m.f, c: m.c, t: tipo, n: doc, p: tercero, r: rif, d: desc, db: m.db, hb: m.hb };
  }

  /* ---------------- tasas ---------------- */
  /** filas [[fecha(Date|texto), tasa]], devuelve {tasas:{iso:valor}} ; formatoFecha 'dmy' o 'mdy' para textos */
  function parseRates(rows, fmt) {
    var out = {};
    rows.forEach(function (r) {
      var f = r[0], v = r[1], iso = null;
      if (Object.prototype.toString.call(f) === '[object Date]') iso = f.getFullYear() + '-' + pad(f.getMonth() + 1) + '-' + pad(f.getDate());
      else {
        var m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(s(f)); if (!m) return;
        var d = fmt === 'mdy' ? m[2] : m[1], mo = fmt === 'mdy' ? m[1] : m[2];
        iso = m[3] + '-' + pad(+mo) + '-' + pad(+d);
      }
      var x = typeof v === 'number' ? v : Number(s(v).replace(/\./g, '').replace(',', '.'));
      if (x > 0) out[iso] = x;
    });
    return out;
  }
  function rateChecks(t) {
    var ks = Object.keys(t).sort(), prev = null, warn = [];
    ks.forEach(function (k) { if (prev && Math.abs(t[k] / prev - 1) > 0.05) warn.push(k); prev = t[k]; });
    return warn;
  }

  /* ---------------- paquete del tablero ---------------- */
  function monthsOf(desde, n) { var y = +desde.slice(0, 4), m = +desde.slice(5, 7), out = []; for (var i = 0; i < n; i++) { out.push(y + '-' + pad(m)); m++; if (m === 13) { y++; m = 1; } } return out; }
  function lastDay(mo) { var y = +mo.slice(0, 4), m = +mo.slice(5, 7); return new Date(y, m, 0).getDate(); }

  /**
   * movs: movimientos enriquecidos del ejercicio; cat: nombres de cuentas; meta: {empresa,rif,inicioMes(1-12)}
   * tasas: {iso:valor}
   */
  function buildPackage(movsIn, cat, tasas, meta) {
    // ejercicio: 12 meses desde el mes de inicio que contiene el último dato
    var maxF = movsIn.reduce(function (a, m) { return m.f > a ? m.f : a; }, '0000');
    var ini = +meta.inicioMes || 1, y = +maxF.slice(0, 4), mo = +maxF.slice(5, 7);
    var sy = mo >= ini ? y : y - 1;
    var months = monthsOf(sy + '-' + pad(ini), 12);
    var rk = Object.keys(tasas).sort();
    function rate(iso) { var lo = 0, hi = rk.length - 1, best = -1; while (lo <= hi) { var mid = (lo + hi) >> 1; if (rk[mid] <= iso) { best = mid; lo = mid + 1; } else hi = mid - 1; } return best >= 0 ? tasas[rk[best]] : 0; }
    var MR = {};
    months.forEach(function (m) {
      if (!rk.length || m > rk[rk.length - 1].slice(0, 7)) return;
      var ld = lastDay(m), daily = [];
      for (var d = 1; d <= ld; d++) daily.push(rate(m + '-' + pad(d)));
      var inm = rk.filter(function (k) { return k.indexOf(m) === 0; }).map(function (k) { return tasas[k]; });
      MR[m] = { d: daily, c: daily[daily.length - 1], p: inm.length ? Math.round(inm.reduce(function (a, b) { return a + b; }, 0) / inm.length * 10000) / 10000 : daily[daily.length - 1] };
    });
    function nm(a) { return cat[a] || (movsIn.filter(function (x) { return x.a === a; })[0] || {}).an || a; }
    function isInc(a) { return a.charAt(0) === '4' || (a.charAt(0) === '7' && /INGRES/i.test(nm(a))); }

    var excl = { cero: 0, cierre: 0 };
    var movs = movsIn.filter(function (x) {
      if (months.indexOf(x.f.slice(0, 7)) < 0) return false;
      if (/cierre de cuentas/i.test(x.d)) { excl.cierre++; return false; }
      if (Math.abs(x.db) < 0.005 && Math.abs(x.hb) < 0.005) { excl.cero++; return false; }
      return true;
    });
    // ventas masivas
    var cnt = {}; movs.forEach(function (x) { cnt[x.a] = (cnt[x.a] || 0) + 1; });
    var agg = Object.keys(cnt).filter(function (a) { return isInc(a) && cnt[a] > 50; }).sort();
    if (agg.length) {
      var keep = [], grp = {};
      movs.forEach(function (x) {
        if (agg.indexOf(x.a) >= 0) {
          var t = (x.t === 'FAC' || x.t === 'N/C' || x.t === 'ND') ? x.t : 'FAC';
          var k = x.a + '|' + x.f + '|' + t;
          if (!grp[k]) grp[k] = { a: x.a, an: x.an, f: x.f, c: '', t: t, n: '', p: '', r: '', d: 'VENTAS DEL DIA', db: 0, hb: 0, k: 0 };
          grp[k].db += x.db; grp[k].hb += x.hb; grp[k].k++;
        } else keep.push(x);
      });
      Object.keys(grp).forEach(function (k) { if (Math.abs(grp[k].db - grp[k].hb) > 0.004) keep.push(grp[k]); });
      movs = keep;
    }
    var expFac = movs.filter(function (x) { return !isInc(x.a) && x.t !== 'ASIENTO'; });
    var hasRif = expFac.length ? (expFac.filter(function (x) { return x.r; }).length / expFac.length) > 0.5 : false;

    function concept(desc) {
      var toks = desc.split(/\s+/).filter(function (t) { return t && !/\d/.test(t) && ['N°', 'Nº', 'N', 'No', 'Nro', '-', 'OP'].indexOf(t) < 0; });
      var r = toks.slice(0, 6).join(' ').replace(/^[ .,\-]+|[ .,\-]+$/g, '');
      return r.toUpperCase() || 'ASIENTO DE DIARIO';
    }
    function pkey(x) {
      if (x.k) return 'C|VENTAS DEL PERIODO';
      if (x.t === 'ASIENTO' || !x.p) return 'C|' + concept(x.d);
      if (x.r && !PLACEHOLDER[x.r] && x.r.length >= 8) return 'R|' + x.r;
      return 'N|' + x.p.toUpperCase().replace(/[^A-Z0-9]/g, '');
    }
    var keys = movs.map(pkey), pnames = {}, prifs = {};
    movs.forEach(function (x, i) {
      var k = keys[i], n = k.indexOf('C|') === 0 ? (x.k ? 'VENTAS DEL PERIODO' : concept(x.d)) : x.p;
      pnames[k] = pnames[k] || {}; pnames[k][n] = (pnames[k][n] || 0) + 1;
      if (x.r) { prifs[k] = prifs[k] || {}; prifs[k][x.r] = (prifs[k][x.r] || 0) + 1; }
    });
    function top(o) { var b = null, bv = -1; for (var k in o) if (o[k] > bv) { b = k; bv = o[k]; } return b; }
    var provs = [], pidx = {};
    Object.keys(pnames).forEach(function (k) {
      var name = top(pnames[k]), rif = prifs[k] ? top(prifs[k]) : '', kind = k.indexOf('C|') === 0 ? 'A' : 'J';
      if (kind !== 'A') {
        if (PLACEHOLDER[rif] || /\b(LIMITED|LTD|CO\.,|GMBH|INC)\b/i.test(name)) kind = 'X';
        else if (rif.charAt(0) === 'V' || rif.charAt(0) === 'E') kind = 'N';
        else if (!rif && !CO_MARKERS.test(name) && name.split(' ').length >= 2 && name.split(' ').length <= 6) kind = 'n';
        else if (rif.charAt(0) === 'G') kind = 'G';
      }
      pidx[k] = provs.length; provs.push([name, PLACEHOLDER[rif] ? '' : rif, kind]);
    });
    var accts = Object.keys(cnt).filter(function (a) { return movs.some(function (x) { return x.a === a; }); }).sort();
    var aidx = {}; accts.forEach(function (a, i) { aidx[a] = i; });
    var names = {};
    accts.forEach(function (a) { names[a] = nm(a); var p = a.split('.'); for (var i = 1; i < p.length; i++) { var c = p.slice(0, i).join('.'); names[c] = cat[c] || c; } });
    var GN = { '4': 'INGRESOS', '5': 'COSTOS', '6': 'GASTOS', '7': 'OTROS INGRESOS Y EGRESOS' };
    Object.keys(GN).forEach(function (g) { if (!names[g] || names[g] === g) names[g] = GN[g]; });

    var dup = {};
    movs.forEach(function (x, i) { if (x.t !== 'ASIENTO' && x.n) { var k = x.a + '|' + keys[i] + '|' + x.n + '|' + Math.round(Math.abs(x.db - x.hb) * 100); dup[k] = (dup[k] || 0) + 1; } });
    var ADMIN = /^(6\.1\.06|5\.1\.05)/, NOFAC = /IMPUESTO|TASAS|BANCARI|NO DEDUCIBLE|DEPRECIA|AMORTIZ/i;
    var cols = { a: [], mi: [], f: [], t: [], n: [], p: [], bs: [], usd: [], c: [], fl: [], di: [], k: [] }, descs = [], didx = {};
    var TC = { 'FAC': 0, 'N/C': 1, 'ND': 2, 'ASIENTO': 3 };
    movs.forEach(function (x, i) {
      var a = x.a, sign = isInc(a) ? -1 : 1, bs = Math.round((x.db - x.hb) * sign * 100) / 100;
      var mo = x.f.slice(0, 7), r = MR[mo] ? MR[mo].c : 0, pi = pidx[keys[i]], kind = provs[pi][2], fl = 0;
      if (!isInc(a) && a.indexOf('5.2.01') !== 0 && a.indexOf('5.2.03') !== 0) {
        if (kind === 'N' || kind === 'n') fl |= 1;
        if (kind === 'X' || /EXTERIOR/.test(names[a])) fl |= 2;
        if (x.t === 'ASIENTO' && ADMIN.test(a) && !NOFAC.test(names[a])) fl |= 4;
        if (hasRif && x.t !== 'ASIENTO' && !provs[pi][1] && kind !== 'X') fl |= 8;
        if (x.t !== 'ASIENTO' && x.n && dup[a + '|' + keys[i] + '|' + x.n + '|' + Math.round(Math.abs(x.db - x.hb) * 100)] > 1) fl |= 16;
        if (/NO DEDUCIBLE/i.test(names[a])) fl |= 32;
      }
      var d = (x.t === 'ASIENTO' && !x.k) ? x.d : '', di = -1;
      if (d) { if (didx[d] === undefined) { didx[d] = descs.length; descs.push(d); } di = didx[d]; }
      cols.a.push(aidx[a]); cols.mi.push(months.indexOf(mo)); cols.f.push(+x.f.slice(8, 10)); cols.t.push(TC[x.t] === undefined ? 3 : TC[x.t]);
      cols.n.push(x.n); cols.p.push(pi); cols.bs.push(bs); cols.usd.push(r ? Math.round(bs / r * 100) / 100 : 0); cols.c.push(x.c);
      cols.fl.push(fl); cols.di.push(di); cols.k.push(x.k || 1);
    });

    // alertas (USD a cierre)
    var lastMi = Math.max.apply(null, cols.mi);
    var SKIP = /IGTF|INVENTARIO|IMPUESTO SOBRE LA RENTA/i, accM = {}, prvM = {};
    for (var i = 0; i < cols.a.length; i++) {
      var a = accts[cols.a[i]]; if (isInc(a) || SKIP.test(names[a])) continue;
      var mi = cols.mi[i];
      (accM[a] = accM[a] || zeros(months.length))[mi] += cols.usd[i];
      var pk = a + '|' + cols.p[i]; (prvM[pk] = prvM[pk] || zeros(months.length))[mi] += cols.usd[i];
    }
    function zeros(n) { var z = []; for (var j = 0; j < n; j++) z.push(0); return z; }
    function present(v) { var o = []; for (var j = 0; j <= lastMi; j++) if (Math.abs(v[j]) > 0.005) o.push(j); return o; }
    function span(from) { var o = []; for (var j = from; j <= lastMi; j++) o.push(j); return o; }
    var alerts = [];
    Object.keys(accM).forEach(function (a) {
      var v = accM[a], pres = present(v);
      if (pres.length >= 3) {
        var sp = span(pres[0]), miss = sp.filter(function (j) { return pres.indexOf(j) < 0; });
        if (pres.length / sp.length >= 0.6 && miss.length) alerts.push({ s: 3, a: aidx[a], p: -1, m: miss, msg: 'Cuenta sin movimiento en ' + miss.length + ' mes(es) en que normalmente registra gastos' });
      }
      var nz = pres.filter(function (j) { return v[j] > 0; }).map(function (j) { return v[j]; }).sort(function (x, y) { return x - y; });
      if (nz.length >= 4) {
        var med = nz.length % 2 ? nz[(nz.length - 1) / 2] : (nz[nz.length / 2 - 1] + nz[nz.length / 2]) / 2;
        pres.forEach(function (j) {
          if (v[j] > 0 && v[j] < 0.3 * med) alerts.push({ s: 2, a: aidx[a], p: -1, m: [j], msg: 'Mes muy por debajo de lo habitual (' + Math.round((1 - v[j] / med) * 100) + '% menos que la mediana en USD)' });
          else if (v[j] > 4 * med && a.indexOf('5.2') !== 0) alerts.push({ s: 2, a: aidx[a], p: -1, m: [j], msg: 'Mes muy por encima de lo habitual (' + (Math.round(v[j] / med * 10) / 10) + ' veces la mediana en USD)' });
        });
      }
    });
    Object.keys(prvM).forEach(function (pk) {
      var a = pk.split('|')[0], pi = +pk.split('|')[1], v = prvM[pk], pres = present(v);
      if (pres.length >= 3 && a.indexOf('5.2') !== 0) {
        var sp = span(pres[0]), miss = sp.filter(function (j) { return pres.indexOf(j) < 0; });
        if (pres.length / sp.length >= 0.7 && miss.length) alerts.push({ s: 1, a: aidx[a], p: pi, m: miss, msg: 'Gasto recurrente que no aparece en ' + miss.length + ' mes(es)' });
      }
    });
    alerts.sort(function (x, y) { return x.s - y.s || (accts[x.a] < accts[y.a] ? -1 : accts[x.a] > accts[y.a] ? 1 : 0); });
    return {
      empresa: meta.empresa, rif: meta.rif, desde: months[0], hasta: months[11], months: months, lastMi: lastMi,
      generado: meta.hoy, accts: accts, names: names, provs: provs, descs: descs, cols: cols, alerts: alerts,
      hasRif: hasRif, mr: MR, agg: agg, excl: excl
    };
  }

  return { parseCatalog: parseCatalog, parseLedger: parseLedger, parseRates: parseRates, rateChecks: rateChecks, buildPackage: buildPackage };
})();
if (typeof module !== 'undefined') module.exports = NT;
