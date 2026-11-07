/**
 * NT Resultados — cifrado y publicación del tablero en GitHub Pages.
 * Cifrado: PBKDF2-HMAC-SHA256 (RIF+código) -> AES-256-CTR + HMAC-SHA256 (encrypt-then-MAC).
 * El navegador del cliente lo descifra con WebCrypto. En GitHub solo queda texto cifrado.
 */
var PUB = {
  REPO: 'navisapp/resultados',
  ITER: 60000
};

var Cifra = (function () {
  // ---------- AES (solo cifrado de bloque; CTR no necesita descifrar) ----------
  var SBOX = [], RCON = [0x01, 0x02, 0x04, 0x08, 0x10, 0x20, 0x40, 0x80, 0x1b, 0x36];
  (function () {
    var p = 1, q = 1, i;
    SBOX[0] = 0x63;
    do {
      p = p ^ ((p << 1) & 0xff) ^ (p & 0x80 ? 0x1b : 0);
      q ^= q << 1; q ^= q << 2; q ^= q << 4; q &= 0xff; if (q & 0x80) q ^= 0x09;
      var x = q ^ ((q << 1) | (q >> 7)) ^ ((q << 2) | (q >> 6)) ^ ((q << 3) | (q >> 5)) ^ ((q << 4) | (q >> 4));
      SBOX[p] = (x ^ 0x63) & 0xff;
    } while (p !== 1);
    for (i = 0; i < 256; i++) SBOX[i] &= 0xff;
  })();
  function xt(a) { return ((a << 1) ^ (a & 0x80 ? 0x1b : 0)) & 0xff; }
  function expandKey(key) { // 32 bytes -> 240 bytes
    var w = key.slice(0), i = 32, r = 0, t;
    while (w.length < 240) {
      t = w.slice(w.length - 4);
      if (i % 32 === 0) { t = [SBOX[t[1]] ^ RCON[r++], SBOX[t[2]], SBOX[t[3]], SBOX[t[0]]]; }
      else if (i % 32 === 16) { t = [SBOX[t[0]], SBOX[t[1]], SBOX[t[2]], SBOX[t[3]]]; }
      for (var j = 0; j < 4; j++) w.push(w[w.length - 32] ^ t[j]);
      i += 4;
    }
    return w;
  }
  function encBlock(rk, inp) {
    var s = inp.slice(0), r, c, i, a0, a1, a2, a3;
    for (i = 0; i < 16; i++) s[i] ^= rk[i];
    for (r = 1; r <= 14; r++) {
      for (i = 0; i < 16; i++) s[i] = SBOX[s[i]];
      var t = s.slice(0);
      for (c = 0; c < 4; c++) for (i = 0; i < 4; i++) s[c * 4 + i] = t[((c + i) % 4) * 4 + i];
      if (r !== 14) {
        for (c = 0; c < 4; c++) {
          a0 = s[c * 4]; a1 = s[c * 4 + 1]; a2 = s[c * 4 + 2]; a3 = s[c * 4 + 3];
          var all = a0 ^ a1 ^ a2 ^ a3;
          s[c * 4] ^= all ^ xt(a0 ^ a1); s[c * 4 + 1] ^= all ^ xt(a1 ^ a2);
          s[c * 4 + 2] ^= all ^ xt(a2 ^ a3); s[c * 4 + 3] ^= all ^ xt(a3 ^ a0);
        }
      }
      for (i = 0; i < 16; i++) s[i] ^= rk[r * 16 + i];
    }
    return s;
  }
  function ctr(key, iv, data) {
    var rk = expandKey(key), out = new Array(data.length), cnt = iv.slice(0);
    for (var off = 0; off < data.length; off += 16) {
      var ks = encBlock(rk, cnt);
      for (var j = 0; j < 16 && off + j < data.length; j++) out[off + j] = (data[off + j] ^ ks[j]) & 0xff;
      for (var k = 15; k >= 0; k--) { cnt[k] = (cnt[k] + 1) & 0xff; if (cnt[k] !== 0) break; }
    }
    return out;
  }
  // ---------- HMAC / PBKDF2 (usa el HMAC nativo de Google) ----------
  function u8(a) { return a.map(function (b) { return b & 0xff; }); }
  function s8(a) { return a.map(function (b) { return b > 127 ? b - 256 : b; }); }
  function hmac(key, msg) { return u8(Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_256, s8(msg), s8(key))); }
  function pbkdf2(pw, salt, iter, nbytes) {
    var out = [];
    for (var b = 1; out.length < nbytes; b++) {
      var u = hmac(pw, salt.concat([(b >>> 24) & 0xff, (b >>> 16) & 0xff, (b >>> 8) & 0xff, b & 0xff])), t = u.slice(0);
      for (var i = 1; i < iter; i++) { u = hmac(pw, u); for (var j = 0; j < 32; j++) t[j] ^= u[j]; }
      out = out.concat(t);
    }
    return out.slice(0, nbytes);
  }
  function rnd(n) { var a = []; for (var i = 0; i < n; i++) a.push(Math.floor(Math.random() * 256)); return a; }
  function b64(a) { return Utilities.base64Encode(s8(a)); }
  function utf8(str) { return u8(Utilities.newBlob(str).getBytes()); }

  /** Devuelve el JSON del paquete cifrado {v:2,...} */
  function cifrar(texto, clave, iter) {
    var gz = u8(Utilities.gzip(Utilities.newBlob(texto, 'application/json')).getBytes());
    var salt = rnd(16), iv = rnd(16);
    var k = pbkdf2(utf8(clave), salt, iter, 64);
    var ct = ctr(k.slice(0, 32), iv, gz);
    var mac = hmac(k.slice(32), iv.concat(ct));
    return JSON.stringify({ v: 2, it: iter, s: b64(salt), iv: b64(iv), ct: b64(ct), mac: b64(mac) });
  }
  return { cifrar: cifrar, _ctr: ctr, _pbkdf2: pbkdf2, _hmac: hmac };
})();

/** Arma la página del cliente y la sube a GitHub */
function publicarTablero_(slug, rif, codigo, pkg) {
  var token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) throw new Error('Falta la llave de GitHub (GITHUB_TOKEN) en la configuración del programa.');
  var clave = (String(rif) + String(codigo)).toUpperCase().replace(/[^A-Z0-9]/g, '');
  var enc = Cifra.cifrar(JSON.stringify(pkg), clave, PUB.ITER);
  var esc = function (x) { return String(x).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); };
  var body = HtmlService.createHtmlOutputFromFile('TableroEstatico').getContent()
    .replace('__TITLE__', 'Resultados ' + esc(pkg.empresa)).replace('__EMPRESA__', esc(pkg.empresa)).replace('__PKG__', enc);
  var html = '<!doctype html>\n<html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover"><style>body{margin:0}</style></head><body>\n' + body + '\n</body></html>';
  subirGithub_(token, slug + '/index.html', html, 'Actualiza tablero ' + slug);
}

function subirGithub_(token, path, contenido, mensaje) {
  var url = 'https://api.github.com/repos/' + PUB.REPO + '/contents/' + path;
  var h = { Authorization: 'Bearer ' + token, Accept: 'application/vnd.github+json', 'X-GitHub-Api-Version': '2022-11-28' };
  var g = UrlFetchApp.fetch(url, { headers: h, muteHttpExceptions: true });
  var sha = g.getResponseCode() === 200 ? JSON.parse(g.getContentText()).sha : null;
  var payload = { message: mensaje, content: Utilities.base64Encode(Utilities.newBlob(contenido).getBytes()), committer: { name: 'NT Resultados', email: 'infonaviscontable@gmail.com' } };
  if (sha) payload.sha = sha;
  var r = UrlFetchApp.fetch(url, { method: 'put', headers: h, contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true });
  if (r.getResponseCode() >= 300) throw new Error('GitHub respondió ' + r.getResponseCode() + ': ' + r.getContentText().slice(0, 200));
}

/** Prueba de la llave: ejecútala desde el editor después de guardar GITHUB_TOKEN */
function probarGithub() {
  var token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) throw new Error('No encuentro GITHUB_TOKEN en las propiedades del script.');
  var r = UrlFetchApp.fetch('https://api.github.com/repos/' + PUB.REPO, { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
  var j = JSON.parse(r.getContentText());
  Logger.log(r.getResponseCode() === 200 ? 'Llave OK. Permiso de escritura: ' + (j.permissions && j.permissions.push) : 'Llave rechazada: ' + r.getContentText().slice(0, 200));
}
