/* =========================================================================
 * EmailTags HITSS — CONFIGURACIÓN
 * Único archivo que normalmente vas a editar.
 * Las carpetas y sus IDs se leyeron de tu buzón padillaal@hitss.com
 * el 29-sep-2026. Si creas/renombras carpetas, agrega aquí la nueva
 * (el ID lo obtienes con el flujo "Listar carpetas" de la guía).
 * ========================================================================= */
(function (root) {
  var P = "AAMkAGYwODM0MDA3LWI1MjUtNDc4ZS05MThlLWYzMzM3NWU1NmVlNwAuAAAAAAB3VnRzaEwBQYKqdBldNTSOAQDO-xdreMjfQ6cEllJkstX3";

  var CONFIG = {
    // ID del add-in (debe coincidir con <Id> del manifest y con los flujos)
    addinId: "7c1e6f0a-4d2b-4b8e-9a3f-2e5d8c1b6a47",

    // Categoría técnica que avisa al flujo "Archivar" que hay algo que mover.
    // Se quita sola en ≤ 2 min. No la uses para otra cosa.
    marker: "ET-Pendiente",

    // Dominios propios: se ignoran como señal (todos te escriben desde hitss.com)
    ownDomains: ["hitss.com"],

    // Al enviar: "unknown" = pregunta solo si el hilo no está clasificado
    //            "always"  = pregunta siempre
    //            "never"   = nunca pregunta (el flujo hereda del hilo si puede)
    askOnSend: "unknown",

    // ---------------------------------------------------------------------
    // CARPETAS  (key corta: la usa el motor de aprendizaje; no la cambies
    // una vez que empieces a usarlo)
    // ---------------------------------------------------------------------
    folders: [
      { key: "A1", path: "1. Acción / 1. Hoy",          id: P + "AAD7-HeiAAA=" },
      { key: "A2", path: "1. Acción / 2. Mañana",       id: P + "AAD7-HejAAA=" },
      { key: "A3", path: "1. Acción / 3. Esta semana",  id: P + "AAD7-HekAAA=" },
      { key: "A4", path: "1. Acción / 4. Waiting",      id: P + "AAD7-HelAAA=" },

      { key: "P1", path: "2. Proyectos / AI",           id: P + "AAD7-HenAAA=" },
      { key: "P2", path: "2. Proyectos / AMS Testing",  id: P + "AAD0ivvNAAA=" },
      { key: "P3", path: "2. Proyectos / Incidentes N3",id: P + "AAD7-HeoAAA=" },

      { key: "G1", path: "3. Gestión / Finanzas",       id: P + "AAD7-HeqAAA=" },
      { key: "G2", path: "3. Gestión / Personal",       id: P + "AAD7-HetAAA=" },
      { key: "G3", path: "3. Gestión / Procesos",       id: P + "AAD7-HerAAA=" },
      { key: "G4", path: "3. Gestión / Proveedores",    id: P + "AAD7-HesAAA=" },

      { key: "C1", path: "4. Comercial / Contratos",    id: P + "AAD7-HexAAA=" },
      { key: "C2", path: "4. Comercial / Licitaciones", id: P + "AAD7-HewAAA=" },
      { key: "C3", path: "4. Comercial / Propuestas",   id: P + "AAD7-HevAAA=" }
    ],

    // ---------------------------------------------------------------------
    // CATEGORÍAS por grupo. Se muestran TODAS las de tu lista maestra de
    // Outlook; estas solo definen el orden/agrupación. Las que falten en
    // Outlook se crean solas la primera vez que las uses.
    // exclusive: true  => solo una a la vez (Estado, Prioridad)
    // ---------------------------------------------------------------------
    groups: [
      { name: "Cliente",   exclusive: false, items: ["ClaroVideo", "Telcel"] },
      { name: "Proyecto",  exclusive: false, items: ["AMS Pruebas", "AI", "SWF", "Incidentes N3", "Observabilidad", "Insumos", "ReleaseManagement", "Launcher", "Dash", "Recomendador"] },
      { name: "Tipo",      exclusive: false, items: ["Proyecto", "operación", "Incidente", "Presupuesto", "Contrato", "Cambio", "Requerimiento", "agreements", "corpo", "support", "training"] },
      { name: "Estado",    exclusive: true,  items: ["to-do", "in progress", "Waiting", "Blocked", "Review", "Closed"] },
      { name: "Prioridad", exclusive: true,  items: ["Urgente", "Alta", "Media", "Baja"] }
    ],

    // ---------------------------------------------------------------------
    // NOMBRES VIEJOS -> NOMBRE CORRECTO. El panel nunca los muestra, el
    // historial aprendido con el nombre viejo se suma al correcto, y al
    // clasificar un correo que traiga el viejo se le cambia por el correcto.
    // ---------------------------------------------------------------------
    aliases: {
      "AMSTesting":   "AMS Pruebas",
      "IncidentesN3": "Incidentes N3",
      "Operacion":    "operación",
      "ToDo":         "to-do",
      "InProgress":   "in progress"
    },

    // ---------------------------------------------------------------------
    // REGLAS SEMILLA: arrancan el motor mientras aprende de ti.
    // Se evalúan contra asunto + primeras líneas del cuerpo + dominios.
    // Pesan menos que lo que tú ya decidiste antes (historial).
    // ---------------------------------------------------------------------
    rules: [
      { re: "telcel",                                   cats: ["Telcel"] },
      { re: "claro ?video|amco|\\bcv\\b",                cats: ["ClaroVideo"] },
      { re: "@amco\\.mx|@clarovideo",                    cats: ["ClaroVideo"] },
      { re: "insumo|sisap|folio",                        cats: ["Insumos", "Telcel"] },
      { re: "\\bswf\\b",                                 cats: ["SWF", "Telcel"] },
      { re: "\\bams\\b|pruebas|testing|\\bqa\\b|defecto", folder: "P2", cats: ["AMS Pruebas"] },
      { re: "incidente|\\bn3\\b|\\bsev ?[12]\\b|ca[ií]da", folder: "P3", cats: ["Incidentes N3", "Incidente"] },
      { re: "\\bia\\b|\\bai\\b|copilot|claude|\\bllm|agente|genai", folder: "P1", cats: ["AI"] },
      { re: "propuesta|cotizaci",                        folder: "C3", cats: ["Proyecto"] },
      { re: "contrato|anexo|addendum|firma",             folder: "C1", cats: ["Contrato"] },
      { re: "licitaci|\\brfp\\b|\\brfi\\b",              folder: "C2" },
      { re: "factura|presupuesto|costo|\\bp&l\\b|forecast", folder: "G1", cats: ["Presupuesto"] },
      { re: "proveedor|tata|nubiral|neoris|epam|qualtop|testlio", folder: "G4" },
      { re: "recomendador",                              cats: ["Recomendador"] },
      { re: "release|despliegue|deploy|ventana",         cats: ["ReleaseManagement", "Cambio"] },
      { re: "observabilidad|datadog|grafana|monitoreo",  cats: ["Observabilidad"] },
      { re: "launcher",                                  cats: ["Launcher"] },
      { re: "dashboard|\\bdash\\b|power ?bi",            cats: ["Dash"] },
      { re: "requerimiento|solicitud|\\bcr\\b",          cats: ["Requerimiento"] },
      // Estado / prioridad
      { re: "aprobaci|autoriza|vo\\.? ?bo|visto bueno|valida", cats: ["Review"] },
      { re: "bloquead|bloqueo|impedimento",              cats: ["Blocked"] },
      { re: "urgente|asap|cr[ií]tico|inmediat",          cats: ["Urgente"] },
      { re: "favor de|podr[ií]as|necesito|requiero|me apoyas|\\?", cats: ["to-do"] }
    ]
  };

  root.ET_CONFIG = CONFIG;
  if (typeof module !== "undefined") module.exports = CONFIG;
})(typeof window !== "undefined" ? window : globalThis);
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
/* =========================================================================
 * EmailTags HITSS — capa Office.js (compartida por panel y evento de envío)
 * ========================================================================= */
(function (root) {
  var KEY = "et";

  function p(fn) {
    return new Promise(function (resolve, reject) {
      fn(function (r) {
        if (!r || r.status === Office.AsyncResultStatus.Succeeded) resolve(r ? r.value : undefined);
        else reject(r.error || new Error("Office error"));
      });
    });
  }

  function item() { return Office.context.mailbox.item; }
  function isCompose(it) { it = it || item(); return !!(it && it.subject && typeof it.subject.getAsync === "function"); }
  function me() { try { return Office.context.mailbox.userProfile.emailAddress.toLowerCase(); } catch (e) { return ""; } }

  // ---------- aprendizaje (roamingSettings = dentro de tu buzón) ----------
  function loadStore() {
    try {
      var raw = Office.context.roamingSettings.get(KEY);
      if (raw && typeof raw === "object" && raw.t) return raw;
    } catch (e) {}
    return ET_ENGINE.emptyStore();
  }
  function saveStore(store) {
    Office.context.roamingSettings.set(KEY, store);
    return p(function (cb) { Office.context.roamingSettings.saveAsync(cb); });
  }

  // ---------- contexto del correo ----------
  function addrList(arr) { return (arr || []).map(function (x) { return (x.emailAddress || "").toLowerCase(); }).filter(Boolean); }

  async function getContext() {
    var it = item(), mine = me(), ctx = { conv: it.conversationId || "", subject: "", text: "", contacts: [], currentCats: [] };
    if (isCompose(it)) {
      var res = await Promise.all([
        p(function (cb) { it.subject.getAsync(cb); }),
        p(function (cb) { it.to.getAsync(cb); }),
        p(function (cb) { it.cc.getAsync(cb); }),
        p(function (cb) { it.body.getAsync(Office.CoercionType.Text, cb); }).catch(function () { return ""; })
      ]);
      ctx.subject = res[0] || "";
      addrList(res[1]).forEach(function (a) { if (a !== mine) ctx.contacts.push({ a: a, w: 1 }); });
      addrList(res[2]).forEach(function (a) { if (a !== mine) ctx.contacts.push({ a: a, w: 0.5 }); });
      ctx.text = String(res[3] || "").slice(0, 1500);
      ctx.mode = "compose";
    } else {
      ctx.subject = it.subject || "";
      if (it.from && it.from.emailAddress) ctx.contacts.push({ a: it.from.emailAddress.toLowerCase(), w: 1 });
      addrList(it.to).concat(addrList(it.cc)).forEach(function (a) { if (a !== mine) ctx.contacts.push({ a: a, w: 0.3 }); });
      ctx.text = await p(function (cb) { it.body.getAsync(Office.CoercionType.Text, cb); }).then(function (t) { return String(t || "").slice(0, 1500); }).catch(function () { return ""; });
      ctx.currentCats = await p(function (cb) { it.categories.getAsync(cb); }).then(function (v) { return (v || []).map(function (c) { return c.displayName; }); }).catch(function () { return []; });
      ctx.mode = "read";
    }
    return ctx;
  }

  // ---------- categorías maestras ----------
  function getMaster() {
    return p(function (cb) { Office.context.mailbox.masterCategories.getAsync(cb); })
      .then(function (v) { return v || []; })
      .catch(function () { return null; }); // null = sin permiso / no soportado
  }
  var PRESETS = ["Preset0","Preset1","Preset2","Preset3","Preset4","Preset5","Preset7","Preset8","Preset9","Preset10","Preset11","Preset12","Preset13","Preset14","Preset15","Preset16","Preset17","Preset18","Preset19"];
  async function ensureMaster(names) {
    var master = await getMaster();
    if (!master) return;
    var have = {}; master.forEach(function (m) { have[m.displayName.toLowerCase()] = 1; });
    var missing = names.filter(function (n) { return n && !have[n.toLowerCase()]; });
    if (!missing.length) return;
    var add = missing.map(function (n, i) { return { displayName: n, color: n === ET_CONFIG.marker ? "Preset24" : PRESETS[(master.length + i) % PRESETS.length] }; });
    await p(function (cb) { Office.context.mailbox.masterCategories.addAsync(add, cb); });
  }

  // ---------- propiedad que leen los flujos ----------
  function folderById(key) {
    var f = ET_CONFIG.folders.filter(function (x) { return x.key === key; })[0];
    return f ? f.id : "";
  }
  async function writeDecision(decision) {
    var it = item();
    var props = await p(function (cb) { it.loadCustomPropertiesAsync(cb); });
    props.set("d", { f: folderById(decision.f), k: decision.f || "", c: decision.c, fl: decision.flag ? 1 : 0, th: decision.thread ? 1 : 0, ts: Date.now() });
    await p(function (cb) { props.saveAsync(cb); });
  }
  async function readDecision() {
    var it = item();
    var props = await p(function (cb) { it.loadCustomPropertiesAsync(cb); });
    return props.get("d") || null;
  }

  // ---------- aplicar ----------
  async function applyRead(decision, currentCats) {
    var it = item(), marker = ET_CONFIG.marker;
    var needsFlow = !!decision.f || decision.thread || decision.flag;
    await ensureMaster(decision.c.concat(needsFlow ? [marker] : []));
    // 1) primero la decisión (para que el flujo la encuentre)
    if (needsFlow) await writeDecision(decision);
    // 2) categorías visibles al instante
    var want = {}; decision.c.forEach(function (c) { want[c.toLowerCase()] = 1; });
    var remove = (currentCats || []).filter(function (c) { return !want[c.toLowerCase()]; });
    if (remove.length) await p(function (cb) { it.categories.removeAsync(remove, cb); });
    var add = decision.c.filter(function (c) { return (currentCats || []).map(function (x) { return x.toLowerCase(); }).indexOf(c.toLowerCase()) < 0; });
    if (needsFlow) add.push(marker);
    if (add.length) await p(function (cb) { it.categories.addAsync(add, cb); });
  }

  async function applyCompose(decision) {
    await ensureMaster(decision.c).catch(function () {});
    await writeDecision(decision);
    // en Outlook clásico hay que guardar el borrador para persistir la propiedad
    await p(function (cb) { item().saveAsync(cb); }).catch(function () {});
  }

  root.ET_OUTLOOK = {
    p: p, item: item, isCompose: isCompose, loadStore: loadStore, saveStore: saveStore,
    getContext: getContext, getMaster: getMaster, applyRead: applyRead, applyCompose: applyCompose,
    readDecision: readDecision, writeDecision: writeDecision, folderById: folderById
  };
})(typeof window !== "undefined" ? window : globalThis);
/* =========================================================================
 * EmailTags HITSS — EVENTO AL ENVIAR (Smart Alerts / OnMessageSend)
 *  - Si ya elegiste carpeta/categorías en el panel -> deja enviar.
 *  - Si el hilo ya estaba clasificado            -> hereda en silencio y envía.
 *  - Si no                                         -> te detiene con el botón
 *    "🧠 Clasificar" que abre el panel (o puedes "Enviar de todos modos").
 * ========================================================================= */
Office.onReady();

async function onMessageSendHandler(event) {
  var CFG = ET_CONFIG, ENG = ET_ENGINE, OL = ET_OUTLOOK;
  try {
    if (CFG.askOnSend === "never") {
      await inheritIfKnown(CFG, ENG, OL);
      return event.completed({ allowEvent: true });
    }

    var already = await OL.readDecision().catch(function () { return null; });
    if (already) return event.completed({ allowEvent: true });

    if (CFG.askOnSend !== "always") {
      var inherited = await inheritIfKnown(CFG, ENG, OL);
      if (inherited) return event.completed({ allowEvent: true });
    }

    event.completed({
      allowEvent: false,
      errorMessage: "¿Dónde archivo este correo? Elige carpeta y categorías antes de enviarlo (o envíalo así y lo archivo según el hilo).",
      cancelLabel: "🧠 Clasificar",
      commandId: "msgComposeOpenPaneButton",
      sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser
    });
  } catch (e) {
    // Nunca bloquear el envío por un error del add-in
    event.completed({ allowEvent: true });
  }
}

// Hilo conocido => escribe la misma decisión en el correo que sale (con th=1
// para que el correo original también quede con los mismos atributos).
async function inheritIfKnown(CFG, ENG, OL) {
  var it = OL.item();
  var conv = it.conversationId;
  if (!conv) return false;
  var store = OL.loadStore();
  var th = store.t[ENG.hash(conv)];
  if (!th) return false;
  await OL.writeDecision({ f: th.f, c: (th.c || []).map(function (c) { return ENG.aliasOf(CFG, c); }), thread: true, flag: false });
  await OL.p(function (cb) { it.saveAsync(cb); }).catch(function () {});
  return true;
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
