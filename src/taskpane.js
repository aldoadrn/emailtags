/* =========================================================================
 * EmailTags HITSS — PANEL
 * ========================================================================= */
(function () {
  var CFG = window.ET_CONFIG, ENG = window.ET_ENGINE, OL = window.ET_OUTLOOK;
  var COLORS = { Preset0:"#e74856",Preset1:"#f7630c",Preset2:"#a0522d",Preset3:"#fce100",Preset4:"#16c60c",Preset5:"#00b7c3",Preset6:"#8a8a3c",Preset7:"#0078d4",Preset8:"#8764b8",Preset9:"#c239b3",Preset10:"#6b8bae",Preset11:"#4a5a70",Preset12:"#9a9a9a",Preset13:"#5d5d5d",Preset14:"#1b1b1b",Preset15:"#a4262c",Preset16:"#ca5010",Preset17:"#6b3a1e",Preset18:"#c19c00",Preset19:"#0b6a0b",Preset20:"#005b70",Preset21:"#5c5c26",Preset22:"#003966",Preset23:"#4b2d7f",Preset24:"#6d1f5e" };

  var S = { ctx: null, sug: null, store: null, master: [], folder: null, cats: {}, touched: false, busy: false };
  var $ = function (id) { return document.getElementById(id); };

  function folderPath(key) { var f = CFG.folders.filter(function (x) { return x.key === key; })[0]; return f ? f.path : "— No mover —"; }
  function norm(s) { return String(s || "").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, ""); }
  function status(msg, cls) { var el = $("status"); el.textContent = msg || ""; el.className = "status " + (cls || ""); }

  // Nombre canónico según la lista maestra (respeta mayúsculas de Outlook)
  function isAlias(name) {
    var n = String(name).toLowerCase();
    return Object.keys(CFG.aliases || {}).some(function (k) { return k.toLowerCase() === n; });
  }
  function canon(name) {
    name = ENG.aliasOf(CFG, name);
    var n = String(name).toLowerCase();
    for (var i = 0; i < S.master.length; i++) if (S.master[i].displayName.toLowerCase() === n) return S.master[i].displayName;
    return name;
  }
  function colorOf(name) {
    var n = String(name).toLowerCase();
    var m = S.master.filter(function (x) { return x.displayName.toLowerCase() === n; })[0];
    return m ? (COLORS[m.color] || "#999") : null;
  }

  // ------------------------------------------------------------------ carga
  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Outlook) return;
    try {
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, function () { load(); });
    } catch (e) {}
    bindUI();
    load();
  });

  async function load() {
    $("loading").classList.remove("hidden"); $("app").classList.add("hidden"); $("actions").classList.add("hidden");
    status("");
    var it = OL.item();
    if (!it) { $("loading").textContent = "Selecciona un correo."; return; }
    try {
      S.store = OL.loadStore();
      var res = await Promise.all([OL.getContext(), OL.getMaster()]);
      S.ctx = res[0];
      S.master = (res[1] || []).filter(function (m) { return m.displayName !== CFG.marker && !isAlias(m.displayName) && !(CFG.moveTagPrefix && m.displayName.indexOf(CFG.moveTagPrefix) === 0); });
      S.sug = ENG.suggest(S.store, S.ctx, CFG);
      // estado inicial = sugerencia
      S.folder = S.sug.folder;
      S.cats = {};
      S.sug.cats.forEach(function (c) { S.cats[canon(c)] = true; });
      S.touched = false;
      // en modo lectura, si el correo ya tiene la decisión aplicada, respétala
      if (S.ctx.mode === "compose") {
        var prev = await OL.readDecision().catch(function () { return null; });
        if (prev) { S.folder = prev.k || S.folder; S.cats = {}; (prev.c || []).forEach(function (c) { S.cats[canon(c)] = true; }); }
      }
      $("mode").textContent = S.ctx.mode === "compose" ? "✉️ Enviando" : "📥 Recibido";
      $("optThread").checked = true;
      $("optFlag").checked = S.ctx.mode === "read" && Object.keys(S.cats).some(function (k) { return S.cats[k] && k.toLowerCase() === "to-do"; });
      render();
      $("loading").classList.add("hidden"); $("app").classList.remove("hidden"); $("actions").classList.remove("hidden");
      $("btnApply").focus();
    } catch (e) {
      $("loading").textContent = "No pude leer el correo: " + (e.message || e);
    }
  }

  // ---------------------------------------------------------------- render
  function render() { renderSug(); renderQuick(); renderFolders(); renderCats(); renderApply(); renderLearn(); }

  function renderSug() {
    var s = S.sug;
    $("confTxt").textContent = s.folder || s.cats.length ? s.confidence + "%" : "sin datos";
    $("confBar").style.width = (s.confidence || 0) + "%";
    $("directBadge").classList.toggle("hidden", !s.direct);
    $("sugFolder").textContent = s.folder ? "📁 " + folderPath(s.folder) : (s.cats.length ? "📁 (sin carpeta sugerida)" : "Aún no tengo historial para este correo: elige abajo y aprenderé.");
    var box = $("sugCats"); box.innerHTML = "";
    s.cats.forEach(function (c) { box.appendChild(chip(canon(c), true, null, true)); });
    var ul = $("reasons"); ul.innerHTML = "";
    (s.reasons || []).slice(0, 4).forEach(function (r) { var li = document.createElement("li"); li.textContent = r; ul.appendChild(li); });
  }

  function renderQuick() {
    var box = $("quickFolders"); box.innerHTML = "";
    var keys = [];
    if (S.sug.folder) keys.push(S.sug.folder);
    S.sug.alternatives.forEach(function (a) { keys.push(a.key); });
    ["A1", "A4"].forEach(function (k) { if (keys.indexOf(k) < 0) keys.push(k); }); // atajos útiles
    keys.slice(0, 5).forEach(function (k) {
      var b = document.createElement("button");
      b.className = "chip"; b.type = "button";
      b.setAttribute("aria-pressed", String(S.folder === k));
      b.textContent = folderPath(k).split(" / ").slice(-1)[0];
      b.title = folderPath(k);
      b.onclick = function () { pickFolder(k); };
      box.appendChild(b);
    });
  }

  function renderFolders() {
    var q = norm($("fSearch").value), list = $("fList"); list.innerHTML = "";
    var rows = [{ key: null, path: "— No mover (solo categorías) —" }].concat(CFG.folders);
    rows.forEach(function (f) {
      if (q && norm(f.path).indexOf(q) < 0) return;
      var d = document.createElement("div");
      d.className = "frow"; d.setAttribute("role", "option"); d.tabIndex = 0;
      d.setAttribute("aria-selected", String(S.folder === f.key));
      var parts = f.path.split(" / ");
      var leaf = parts.pop();
      d.innerHTML = (parts.length ? '<span class="parent">' + esc(parts.join(" / ")) + " /</span> " : "") + esc(leaf);
      d.onclick = function () { pickFolder(f.key); };
      d.onkeydown = function (e) { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); pickFolder(f.key); } };
      list.appendChild(d);
    });
  }

  function renderCats() {
    var q = norm($("cSearch").value), box = $("catGroups"); box.innerHTML = "";
    var shown = {};
    CFG.groups.forEach(function (g) {
      var items = g.items.map(canon);
      renderGroup(box, g.name + (g.exclusive ? " · una" : ""), items, g, q, shown);
    });
    // Todas las demás de la lista maestra de Outlook
    var others = S.master.map(function (m) { return m.displayName; }).filter(function (n) { return !shown[n.toLowerCase()]; });
    Object.keys(S.cats).forEach(function (c) { if (S.cats[c] && !shown[c.toLowerCase()] && others.indexOf(c) < 0) others.push(c); });
    if (others.length) renderGroup(box, "Otras de Outlook", others.sort(), null, q, shown);
  }

  function renderGroup(box, title, items, g, q, shown) {
    var vis = items.filter(function (n) { return !q || norm(n).indexOf(q) >= 0; });
    items.forEach(function (n) { shown[n.toLowerCase()] = 1; });
    if (!vis.length) return;
    var wrap = document.createElement("div"); wrap.className = "group";
    var l = document.createElement("span"); l.className = "label"; l.textContent = title; wrap.appendChild(l);
    var chips = document.createElement("div"); chips.className = "chips";
    vis.forEach(function (n) { chips.appendChild(chip(n, !!S.cats[n], g)); });
    wrap.appendChild(chips); box.appendChild(wrap);
  }

  function chip(name, on, group, readonly) {
    var b = document.createElement("button");
    b.type = "button"; b.className = "chip"; b.setAttribute("aria-pressed", String(!!on));
    var col = colorOf(name);
    if (!col && S.master.length) b.classList.add("missing"), b.title = "Se creará en Outlook al aplicar";
    b.innerHTML = '<span class="dot" style="background:' + (col || "transparent") + '"></span>' + esc(name);
    if (readonly) { b.tabIndex = -1; b.style.cursor = "default"; return b; }
    b.onclick = function () { toggleCat(name, group); };
    return b;
  }

  function renderApply() {
    var changed = S.touched;
    $("btnApply").textContent = S.ctx.mode === "compose"
      ? (changed ? "Guardar mi selección" : "Usar sugerencia") + " ↵"
      : (changed ? "Aplicar mi selección" : "Aplicar sugerencia") + " ↵";
    var nothing = !S.folder && !Object.keys(S.cats).some(function (k) { return S.cats[k]; });
    $("btnApply").disabled = nothing || S.busy;
  }

  function renderLearn() {
    var st = S.store, n = function (o) { return Object.keys(o || {}).length; };
    $("learnInfo").textContent = n(st.t) + " hilos · " + n(st.a) + " contactos · " + n(st.d) + " dominios · " + n(st.w) + " palabras · " +
      Math.round(JSON.stringify(st).length / 1024) + " KB de 28 KB (guardado dentro de tu buzón).";
  }

  // ------------------------------------------------------------- acciones
  function pickFolder(k) { S.folder = k; S.touched = true; renderQuick(); renderFolders(); renderApply(); }

  function toggleCat(name, group) {
    var on = !S.cats[name];
    if (on && group && group.exclusive) group.items.forEach(function (x) { delete S.cats[canon(x)]; });
    if (on) S.cats[name] = true; else delete S.cats[name];
    S.touched = true; renderCats(); renderApply();
  }

  async function apply() {
    if (S.busy || $("btnApply").disabled) return;
    S.busy = true; renderApply(); status("Aplicando…");
    var decision = {
      f: S.folder || null,
      c: Object.keys(S.cats).filter(function (k) { return S.cats[k]; }),
      thread: $("optThread").checked,
      flag: $("optFlag").checked
    };
    try {
      if (S.ctx.mode === "compose") {
        await OL.applyCompose(decision);
      } else {
        await OL.applyRead(decision, S.ctx.currentCats);
      }
      S.store = ENG.learn(S.store, S.ctx, { f: decision.f, c: decision.c }, S.touched);
      await OL.saveStore(S.store).catch(function () {});
      renderLearn();
      if (S.ctx.mode === "compose") {
        status("✓ Listo. Presiona Enviar" + (decision.f ? ": al salir se archivará en " + folderPath(decision.f) + (decision.thread ? " junto con el hilo." : ".") : "."), "ok");
      } else {
        status("✓ Categorías aplicadas." + (decision.f ? " Se moverá a " + folderPath(decision.f) + " en ≤ 2 min." : ""), "ok");
      }
    } catch (e) {
      status("Error: " + (e.message || e), "err");
    } finally {
      S.busy = false; renderApply();
    }
  }

  function bindUI() {
    $("btnApply").onclick = apply;
    $("fSearch").oninput = renderFolders;
    $("cSearch").oninput = renderCats;
    $("fSearch").onkeydown = function (e) {
      if (e.key === "Enter") { e.preventDefault(); var first = $("fList").querySelector(".frow"); if (first) first.click(); }
    };
    document.addEventListener("keydown", function (e) {
      var tag = (e.target.tagName || "").toLowerCase();
      if (e.key === "/" && tag !== "input") { e.preventDefault(); $("fSearch").focus(); }
      if (e.key === "Enter" && (e.ctrlKey || tag === "body" || e.target.id === "btnApply")) { e.preventDefault(); apply(); }
    });
    $("btnExport").onclick = function () {
      var txt = JSON.stringify(S.store);
      (navigator.clipboard ? navigator.clipboard.writeText(txt) : Promise.reject()).then(
        function () { status("Respaldo copiado al portapapeles.", "ok"); },
        function () { status("No pude copiar; tu navegador lo bloqueó.", "err"); });
    };
    $("btnReset").onclick = async function () {
      S.store = ENG.emptyStore(); await OL.saveStore(S.store).catch(function () {}); renderLearn(); status("Aprendizaje reiniciado.", "ok");
    };
  }

  function esc(s) { return String(s).replace(/[&<>"']/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]; }); }
})();
