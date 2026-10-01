/* =========================================================================
 * EmailTags HITSS — MOTOR DE SUGERENCIA Y APRENDIZAJE
 * Sin servidores ni bases externas: el "modelo" vive en los
 * roamingSettings del add-in, que Exchange guarda DENTRO de tu buzón.
 * Funciones puras (se pueden probar con Node).
 * ========================================================================= */
(function (root) {
  var MAX_BYTES = 28000; // roamingSettings admite 32 KB en total
  var STOP = ("the and for with from this that are was your you our una uno unos unas los las del con por para que como sus mas pero sin sobre entre este esta estos estas ese esa eso hay fue son ser muy todo toda todos cada otro otra desde hasta cuando donde quien cual favor saludos gracias buen buenos buenas dias tardes hola adjunto adjunta envio envia correo mail email reunion hitss mx com www http https").split(" ");
  var STOPSET = {};
  STOP.forEach(function (w) { STOPSET[w] = 1; });

  function now() { return Date.now(); }

  function emptyStore() { return { v: 1, t: {}, a: {}, d: {}, w: {} }; }

  function hash(s) {
    s = String(s || "");
    var h = 5381;
    for (var i = 0; i < s.length; i++) { h = ((h << 5) + h + s.charCodeAt(i)) | 0; }
    return (h >>> 0).toString(36) + s.length.toString(36);
  }

  function strip(s) {
    return String(s || "").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "");
  }

  function tokens(subject) {
    var s = strip(subject).replace(/^((re|rv|fw|fwd|reenviar|resp|aw)\s*:\s*)+/g, "");
    var seen = {}, out = [];
    s.split(/[^a-z0-9&]+/).forEach(function (w) {
      if (w.length < 3 || STOPSET[w] || /^\d+$/.test(w) || seen[w]) return;
      seen[w] = 1; out.push(w);
    });
    return out.slice(0, 8);
  }

  function domainOf(addr) {
    var m = String(addr || "").toLowerCase().match(/@([^>\s]+)$/);
    return m ? m[1] : "";
  }

  function bump(map, key, w) { map[key] = (map[key] || 0) + w; }

  function add(target, src, mult) {
    for (var k in src) if (Object.prototype.hasOwnProperty.call(src, k)) bump(target, k, src[k] * mult);
  }

  function aliasOf(cfg, cat) {
    var al = cfg.aliases || {}, c = String(cat).toLowerCase();
    for (var k in al) if (Object.prototype.hasOwnProperty.call(al, k) && k.toLowerCase() === c) return al[k];
    return cat;
  }

  function groupOf(cfg, cat) {
    var c = String(cat).toLowerCase();
    for (var i = 0; i < cfg.groups.length; i++) {
      var g = cfg.groups[i];
      for (var j = 0; j < g.items.length; j++) if (g.items[j].toLowerCase() === c) return g;
    }
    return null;
  }

  /**
   * ctx = { conv, subject, text, contacts:[{a, w}], currentCats:[] }
   * return { folder, folderScore, confidence, alternatives:[{key,score}], cats:[], reasons:[], direct, fromThread }
   */
  function suggest(store, ctx, cfg) {
    store = store || emptyStore();
    var F = {}, C = {}, reasons = [], evidence = 0;
    var own = {}; (cfg.ownDomains || []).forEach(function (d) { own[d.toLowerCase()] = 1; });
    var marker = cfg.marker;

    // 1) Hilo ya clasificado => herencia exacta
    var th = ctx.conv ? store.t[hash(ctx.conv)] : null;
    if (th) {
      return {
        folder: th.f, confidence: 99, alternatives: [], cats: (th.c || []).map(function (c) { return aliasOf(cfg, c); }),
        reasons: ["Hilo ya clasificado: se heredan carpeta y categorías"], direct: true, fromThread: true
      };
    }

    // 2) Contactos (remitente / destinatarios)
    (ctx.contacts || []).forEach(function (c) {
      var a = String(c.a || "").toLowerCase(); if (!a) return;
      var e = store.a[a];
      if (e) {
        add(F, e.f || {}, 4 * (c.w || 1));
        add(C, e.c || {}, 3 * (c.w || 1));
        var n = 0; for (var k in e.f) n += e.f[k];
        evidence += n * (c.w || 1);
        if (n) reasons.push(a + ": " + n + " decisiones previas");
      }
      var d = domainOf(a);
      if (d && !own[d] && store.d[d]) {
        add(F, store.d[d].f || {}, 1.5 * (c.w || 1));
        add(C, store.d[d].c || {}, 1.5 * (c.w || 1));
      }
    });

    // 3) Palabras del asunto
    var toks = tokens(ctx.subject);
    var tokHits = 0;
    toks.forEach(function (t) {
      var e = store.w[t];
      if (!e) return;
      tokHits++;
      add(F, e.f || {}, 1);
      add(C, e.c || {}, 0.7);
    });
    if (tokHits) reasons.push(tokHits + " palabra(s) del asunto ya vistas");

    // 4) Reglas semilla
    var hay = strip((ctx.subject || "") + " " + String(ctx.text || "").slice(0, 1500) + " " +
      (ctx.contacts || []).map(function (c) { return c.a; }).join(" "));
    var ruleHits = [];
    (cfg.rules || []).forEach(function (r) {
      var re; try { re = new RegExp(r.re, "i"); } catch (e) { return; }
      if (!re.test(hay)) return;
      if (r.folder) bump(F, r.folder, 6);
      (r.cats || []).forEach(function (c) { bump(C, c, 6); });
      ruleHits.push(r.re.split("|")[0].replace(/\\b/g, ""));
    });
    if (ruleHits.length) reasons.push("Palabras clave: " + ruleHits.slice(0, 4).join(", "));

    // 5) Categorías que ya tiene el correo
    (ctx.currentCats || []).forEach(function (c) { if (c !== marker) bump(C, c, 10); });

    // Carpeta
    var fk = Object.keys(F).filter(function (k) { return F[k] > 0; }).sort(function (a, b) { return F[b] - F[a]; });
    var sum = fk.reduce(function (s, k) { return s + F[k]; }, 0);
    var top = fk[0] || null;
    var share = top ? F[top] / sum : 0;
    var confidence = top ? Math.round(100 * share * (1 - Math.exp(-sum / 12))) : 0;
    // >5 veces el mismo patrón sube la confianza; >10 lo sugiere directo
    var topCount = 0;
    (ctx.contacts || []).forEach(function (c) {
      var e = store.a[String(c.a || "").toLowerCase()];
      if (e && e.f && top && e.f[top]) topCount = Math.max(topCount, e.f[top]);
    });
    if (topCount > 5) confidence = Math.min(97, confidence + 10);
    var direct = topCount > 10 && share >= 0.75;
    if (direct) confidence = Math.max(confidence, 90);

    // Nombres viejos -> nombre correcto (suma su historial)
    var C2 = {};
    Object.keys(C).forEach(function (k) { bump(C2, aliasOf(cfg, k), C[k]); });
    C = C2;

    // Categorías (respetando grupos exclusivos)
    var cats = [];
    var byGroup = {};
    Object.keys(C).forEach(function (c) {
      if (c === marker || C[c] <= 0) return;
      var g = groupOf(cfg, c);
      var gname = g ? g.name : "_otras";
      (byGroup[gname] = byGroup[gname] || { g: g, list: [] }).list.push(c);
    });
    Object.keys(byGroup).forEach(function (gname) {
      var entry = byGroup[gname];
      var list = entry.list.sort(function (a, b) { return C[b] - C[a]; });
      if (entry.g && entry.g.exclusive) {
        if (C[list[0]] >= 3) cats.push(list[0]);
      } else {
        var mx = C[list[0]];
        list.forEach(function (c) { if (C[c] >= Math.max(3, mx * 0.35)) cats.push(c); });
      }
    });

    return {
      folder: top, confidence: confidence,
      alternatives: fk.slice(1, 4).map(function (k) { return { key: k, score: F[k] }; }),
      cats: cats, reasons: reasons, direct: direct, fromThread: false, evidence: evidence
    };
  }

  /** decision = { f: folderKey|null, c: [cats] }, modified => más peso */
  function learn(store, ctx, decision, modified) {
    store = store || emptyStore();
    var w = modified ? 2 : 1, t = now();
    var cats = (decision.c || []).slice();
    if (ctx.conv) store.t[hash(ctx.conv)] = { f: decision.f || null, c: cats, ts: t };

    function feed(bucket, key, mult) {
      var e = bucket[key] || (bucket[key] = { f: {}, c: {}, ts: t });
      e.ts = t;
      if (decision.f) bump(e.f, decision.f, w * mult);
      cats.forEach(function (c) { bump(e.c, c, w * mult); });
    }
    (ctx.contacts || []).forEach(function (c) {
      var a = String(c.a || "").toLowerCase(); if (!a) return;
      feed(store.a, a, c.w >= 1 ? 1 : 0.5);
      var d = domainOf(a); if (d) feed(store.d, d, 1);
    });
    tokens(ctx.subject).forEach(function (tk) { feed(store.w, tk, 1); });
    return prune(store);
  }

  function prune(store, maxBytes) {
    maxBytes = maxBytes || MAX_BYTES;
    var guard = 0;
    while (JSON.stringify(store).length > maxBytes && guard++ < 50) {
      ["t", "w", "a", "d"].forEach(function (b) {
        var keys = Object.keys(store[b]).sort(function (x, y) { return (store[b][x].ts || 0) - (store[b][y].ts || 0); });
        var drop = Math.max(1, Math.floor(keys.length * 0.15));
        keys.slice(0, drop).forEach(function (k) { delete store[b][k]; });
      });
    }
    return store;
  }

  var Engine = { emptyStore: emptyStore, hash: hash, tokens: tokens, domainOf: domainOf, suggest: suggest, learn: learn, prune: prune, groupOf: groupOf, aliasOf: aliasOf };
  root.ET_ENGINE = Engine;
  if (typeof module !== "undefined") module.exports = Engine;
})(typeof window !== "undefined" ? window : globalThis);
