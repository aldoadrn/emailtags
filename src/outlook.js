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
