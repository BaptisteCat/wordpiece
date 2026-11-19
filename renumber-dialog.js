/* global WPMsg */
// ============================================================
// LexRef — fenêtre « Renuméroter les pièces » (menu ☰ du volet).
// Le volet envoie les pièces (après le handshake « ready ») ; tout se règle ici en APERÇU, puis
// « Appliquer » renvoie le plan { tops: [{ id, num }], kids: [{ id, local }] } au volet, qui
// l'applique en une fois (WP.applyRenumber : le résultat est exactement cet aperçu).
//  - saisie une à une (champ « nouveau numéro ») ;
//  - TRANCHE : « de 5 à 9 → à partir de 20 » ;
//  - SÉLECTION par lot (clic, Ctrl+clic, Maj+clic, Ctrl+A) : numéroter à partir de…, décaler de ±k,
//    rétablir ;
//  - GLISSER-DÉPOSER (poignée ⋮⋮, une pièce ou toute la sélection) : les numéros restent attachés
//    aux PLACES, les pièces déplacées prennent le numéro de leur nouvelle place.
// Les sous-pièces suivent leur groupe ; leur indice (4.1 → 4.3) se modifie dans leur champ.
//  - RECHERCHE par nom (« SMS ») : la liste ne montre que les pièces correspondantes (un groupe reste
//    visible si l'une de ses sous-pièces correspond) ; « Tout sélectionner » / Ctrl+A ne prennent que
//    ce qui est affiché ; le glisser-déposer est suspendu tant qu'un filtre est actif.
//  - CADENAS cliquables (pièces et sous-pièces) : verrouiller / déverrouiller fait partie de l'aperçu
//    (plan.locks), appliqué avec le reste. La renumérotation elle-même ne verrouille jamais rien.
//  - SOUS-PIÈCES : poignée ⋮⋮ pour les réordonner DANS leur groupe (les indices restent aux places).
//  - « Appliquer » laisse la fenêtre ouverte : le volet renvoie la liste à jour.
//  - GROUPER la sélection en sous-pièces : 4, 5, 6, 7, 8 → 4.1, 4.2, 4.3, 4.4, 4.5 (le groupe prend
//    le numéro de la 1re pièce ; si la sélection contient un groupe existant, les autres pièces
//    s'ajoutent à la fin de ce groupe). « Resserrer la suite » : 9 → 5… Le plan envoyé porte alors
//    newGroups: [{ key, num }] et, pour chaque sous-pièce déplacée, kids[].parent (id ou clé).
// Les pièces verrouillées restent à leur numéro, sauf « Inclure les pièces verrouillées ».
// ============================================================
WPMsg.ready(() => {
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s == null ? "" : s).replace(/[&<>"']/g,
    (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  // Comparaison « naturelle » (1 < 1 bis < 2 < 10 ; 4.1 < 4.2 ; i.2 < i.10), comme le moteur.
  function naturalCompare(a, b) {
    const ax = String(a == null ? "" : a).match(/(\d+|\D+)/g) || [];
    const bx = String(b == null ? "" : b).match(/(\d+|\D+)/g) || [];
    for (let i = 0; i < Math.max(ax.length, bx.length); i++) {
      const an = ax[i], bn = bx[i];
      if (an === undefined) return -1;
      if (bn === undefined) return 1;
      if (/^\d+$/.test(an) && /^\d+$/.test(bn)) { const d = parseInt(an, 10) - parseInt(bn, 10); if (d) return d; }
      else if (an !== bn) return an < bn ? -1 : 1;
    }
    return 0;
  }
  const clean = (v) => String(v == null ? "" : v).replace(/\s+/g, " ").trim().replace(/(\d)([a-zA-Z])/g, "$1 $2");
  // Numéro de départ → suite : « 20 » → 20, 21… ; « i.1 » → i.1, i.2… ; « A » → A, B… ; « Annexe a » → a, b…
  function stepper(start) {
    const s = clean(start);
    let m = s.match(/^(.*?)(\d+)$/);
    if (m) { const pre = m[1], n = parseInt(m[2], 10); return (i) => pre + (n + i); }
    m = s.match(/^(.*?[^A-Za-z]|)([A-Za-z])$/);
    if (m) {
      const pre = m[1], up = m[2] === m[2].toUpperCase(), base = m[2].toLowerCase().charCodeAt(0) - 97;
      return (i) => {
        let x = base + i + 1, out = "";
        while (x > 0) { x--; out = String.fromCharCode(97 + (x % 26)) + out; x = Math.floor(x / 26); }
        return pre + (up ? out.toUpperCase() : out);
      };
    }
    return null;
  }
  // « i.12 » → { pre: "i.", n: 12 } ; null sans nombre final.
  const splitNum = (v) => { const m = clean(v).match(/^(.*?)(\d+)$/); return m ? { pre: m[1], n: parseInt(m[2], 10) } : null; };

  let rows = [];          // [{ id, old, num, name, locked, container, isNew?, kids: [{ id, oldLocal, local, name, locked, moved?, orig? }] }]
  let original = [];      // copie de l'état reçu (« Tout rétablir »)
  let newSeq = 0;         // clés des groupes créés dans la fenêtre (« new:1 »…)
  const cloneRows = (list) => JSON.parse(JSON.stringify(list));
  let selected = new Set();
  let anchor = null;      // dernière ligne cliquée (Maj+clic)
  let side = "own";
  // ---- Recherche ----
  // Texte normalisé : minuscules, sans accents (« Échange » ≈ « echange »).
  const fold = (v) => String(v == null ? "" : v).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/\s+/g, " ").trim();
  let query = "";
  const filtering = () => !!query;
  const nameHit = (name) => fold(name).includes(query);
  const matches = (r) => !query || nameHit(r.name) || r.kids.some((k) => nameHit(k.name));
  const visibleRows = () => rows.filter(matches);
  const incLocked = () => $("incLocked").checked;
  const movable = (r) => !r.locked || incLocked();

  function sortRows() {
    const idx = new Map(rows.map((r, i) => [r.id, i]));
    rows.sort((a, b) => naturalCompare(a.num, b.num) || idx.get(a.id) - idx.get(b.id));
  }
  function problems() {
    const errs = new Map(); // id → message
    const seen = new Map();
    for (const r of rows) {
      const v = clean(r.num);
      if (!v) errs.set(r.id, "numéro vide");
      else if (/^\d+\.\d+$/.test(v)) errs.set(r.id, `« ${v} » : une sous-pièce se range depuis le volet`);
      else if (seen.has(v)) { errs.set(r.id, `n°${v} en double`); errs.set(seen.get(v), `n°${v} en double`); }
      else seen.set(v, r.id);
      const ks = new Map();
      for (const k of r.kids) {
        const l = clean(k.local);
        if (!l) errs.set(k.id, "numéro vide");
        else if (ks.has(l)) { errs.set(k.id, `n°${v}.${l} en double`); errs.set(ks.get(l), `n°${v}.${l} en double`); }
        else ks.set(l, k.id);
      }
    }
    return errs;
  }
  const lockChanged = (x) => !!x.locked !== !!x.oldLocked;
  const changedCount = () => rows.reduce((n, r) => n + (r.isNew || clean(r.num) !== r.old ? 1 : 0) + (lockChanged(r) ? 1 : 0) +
    r.kids.reduce((m, k) => m + (clean(k.local) !== k.oldLocal ? 1 : 0) + (lockChanged(k) ? 1 : 0), 0), 0);

  const LOCK = '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect width="18" height="11" x="3" y="11" rx="2" ry="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg>';
  const UNLOCK = '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect width="18" height="11" x="3" y="11" rx="2" ry="2"/><path d="M7 11V7a5 5 0 0 1 9.9-1"/></svg>';
  const lockBtn = (x, attr) => `<button type="button" class="lock ${x.locked ? "on" : ""} ${lockChanged(x) ? "chg" : ""}" ${attr} ` +
    `title="${x.locked ? "Verrouillée : cliquer pour déverrouiller" : "Cliquer pour verrouiller"}${lockChanged(x) ? " (changement non appliqué)" : ""}" ` +
    `aria-label="${x.locked ? "Déverrouiller" : "Verrouiller"}">${x.locked ? LOCK : UNLOCK}</button>`;
  const GRIP = '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="9" cy="12" r="1"/><circle cx="9" cy="5" r="1"/><circle cx="9" cy="19" r="1"/><circle cx="15" cy="12" r="1"/><circle cx="15" cy="5" r="1"/><circle cx="15" cy="19" r="1"/></svg>';

  function render() {
    const errs = problems();
    const list = $("list");
    if (!rows.length) { list.innerHTML = `<div class="empty">Aucune pièce à renuméroter.</div>`; updateFoot(errs); return; }
    const vis = visibleRows();
    $("qCount").textContent = query ? `${vis.length} pièce${vis.length > 1 ? "s" : ""} sur ${rows.length}` : "";
    if (!vis.length) { list.innerHTML = `<div class="empty">Aucune pièce dont le nom contient « ${esc($("q").value.trim())} ».</div>`; updateFoot(errs); return; }
    let html = "";
    for (const r of vis) {
      const mov = movable(r), v = clean(r.num), dragOk = mov && !filtering();
      const cls = ["row", selected.has(r.id) ? "sel" : "", r.locked ? "locked" : "", r.isNew || v !== r.old ? "changed" : "", r.isNew ? "newgrp" : "", errs.has(r.id) ? "err" : ""].join(" ");
      const name = r.container ? (r.name || `Groupe · ${r.kids.length} sous-pièce${r.kids.length > 1 ? "s" : ""}`) : (r.name || "(sans nom)");
      html += `<div class="${cls}" data-id="${r.id}" title="${esc(errs.get(r.id) || "")}">` +
        `<input type="checkbox" class="pick" ${selected.has(r.id) ? "checked" : ""} ${mov ? "" : "disabled"} aria-label="Sélectionner" />` +
        `<span class="grip ${dragOk ? "" : "off"}" title="${filtering() ? "Effacez la recherche pour glisser" : "Glisser pour déplacer"}">${GRIP}</span>` +
        `<span class="oldnum" title="${r.isNew ? "Nouveau groupe" : "Numéro actuel"}">${r.isNew ? "nouveau" : esc(r.old)}</span><span class="arrow">→</span>` +
        `<input class="newnum" data-top="${r.id}" value="${esc(r.num)}" ${mov ? "" : "disabled"} aria-label="Nouveau numéro" />` +
        `<span class="name ${r.container ? "group" : ""}" title="${esc(name)}">${esc(name)}</span>` +
        (r.isNew ? `<span></span>` : lockBtn(r, `data-lock-top="${r.id}"`)) + `</div>`;
      for (const k of r.kids) {
        const kmov = !k.locked || incLocked(), kl = clean(k.local);
        const kdrag = kmov && !filtering();
        const kcls = ["row", "kid", query && nameHit(k.name) ? "hit" : "", k.locked ? "locked" : "", kl !== k.oldLocal ? "changed" : "", errs.has(k.id) ? "err" : ""].join(" ");
        html += `<div class="${kcls}" data-kid="${k.id}" title="${esc(errs.get(k.id) || "")}">` +
          `<span></span><span class="grip ${kdrag ? "" : "off"}" title="${filtering() ? "Effacez la recherche pour glisser" : "Glisser : dans le groupe, vers un autre groupe, ou hors du groupe"}">${GRIP}</span><span class="oldnum" title="Numéro actuel">${esc(k.moved ? k.orig.old : r.old + "." + k.oldLocal)}</span><span class="arrow">→</span>` +
          `<span class="kidnum"><span>${esc(v)}.</span><input class="newnum" data-kid="${k.id}" value="${esc(k.local)}" ${kmov ? "" : "disabled"} aria-label="Nouvel indice" /></span>` +
          `<span class="name" title="${esc(k.name)}">${esc(k.name || "(sans nom)")}</span>${lockBtn(k, `data-lock-kid="${k.id}"`)}</div>`;
      }
    }
    list.innerHTML = html;
    updateFoot(errs);
  }
  function updateFoot(errs) {
    const n = changedCount();
    const st = $("status");
    const first = [...(errs || problems()).values()][0];
    st.classList.toggle("err", !!first);
    st.classList.remove("ok");
    st.textContent = first ? "À corriger : " + first
      : n ? `${n} modification${n > 1 ? "s" : ""} — rien n'est appliqué avant « Appliquer ».`
        : "Aucun changement pour l'instant.";
    $("apply").disabled = !!first || !n;
    const k = selected.size;
    $("selTool").classList.toggle("off", !k);
    $("selCount").textContent = k ? `${k} sélectionnée${k > 1 ? "s" : ""}` : "Sélection";
  }

  // ---- Actions ----
  const selRows = () => rows.filter((r) => selected.has(r.id) && movable(r));
  function numberFrom(list, start) {
    const st = stepper(start);
    if (!st) return say("Numéro de départ invalide (ex. 20, i.1, A).");
    list.forEach((r, i) => { r.num = st(i); });
    sortRows(); render();
  }
  function say(msg) { const s = $("status"); s.textContent = msg; s.classList.add("err"); }
  function trancheRows() {
    const a = splitNum($("rFrom").value), b = splitNum($("rTo").value || $("rFrom").value);
    if (!a || !b) { say("Tranche : indiquez deux numéros, ex. de 5 à 9 (ou de i.1 à i.4)."); return null; }
    const lo = Math.min(a.n, b.n), hi = Math.max(a.n, b.n);
    const out = rows.filter((r) => { const x = splitNum(r.num); return movable(r) && x && x.pre === a.pre && x.n >= lo && x.n <= hi; });
    if (!out.length) say("Aucune pièce dans cette tranche.");
    return out.length ? out : null;
  }
  $("rApply").addEventListener("click", () => {
    const list = trancheRows();
    if (!list) return;
    if (!clean($("rStart").value)) return say("Indiquez le nouveau premier numéro (« à partir de »).");
    numberFrom(list, $("rStart").value);
  });
  $("rSelect").addEventListener("click", () => {
    const list = trancheRows();
    if (!list) return;
    selected = new Set(list.filter(matches).map((r) => r.id));
    render();
  });
  $("sNumber").addEventListener("click", () => {
    if (!clean($("sStart").value)) return say("Indiquez le premier numéro.");
    numberFrom(selRows(), $("sStart").value);
  });
  $("sShiftBtn").addEventListener("click", () => {
    const k = parseInt(String($("sShift").value).replace(/[−–]/g, "-").replace(/\s+/g, ""), 10);
    if (!k) return say("Décalage : un nombre entier, ex. +3 ou −2.");
    for (const r of selRows()) { const x = splitNum(r.num); if (x) r.num = x.pre + Math.max(1, x.n + k); }
    sortRows(); render();
  });
  $("sReset").addEventListener("click", () => {
    for (const r of selRows()) {
      ungroup(r);
      if (!rows.includes(r)) continue; // groupe créé ici : dissous
      r.num = r.old; for (const kd of r.kids) kd.local = kd.oldLocal;
    }
    sortRows(); render();
  });

  // ---- Grouper en sous-pièces ----
  // Défait un regroupement fait dans la fenêtre : les pièces déplacées redeviennent des pièces (à
  // leur numéro d'origine) ; un groupe créé ici disparaît.
  // Renvoie une sous-pièce dans son groupe d'origine (groupe recréé d'après l'état initial s'il a
  // disparu de la liste entre-temps).
  function sendHome(item) {
    let g = rows.find((x) => x.id === item.home.gid);
    if (!g) {
      const o = original.find((x) => x.id === item.home.gid);
      if (!o) return false;
      g = { ...cloneRows([o])[0], kids: [] };
      rows.push(g);
    }
    g.kids.push({ id: item.id, oldLocal: item.home.oldLocal, local: item.home.oldLocal, name: item.name, locked: item.locked, oldLocked: item.oldLocked });
    g.kids.sort((a, b) => naturalCompare(clean(a.local), clean(b.local)));
    return true;
  }
  function ungroup(r) {
    if (r.home) { // sous-pièce sortie de son groupe : elle y retourne, et les décalages s'annulent
      rows.splice(rows.indexOf(r), 1); selected.delete(r.id);
      revertShifts(r.undo);
      if (!sendHome(r)) rows.push(r);
      return;
    }
    const back = r.kids.filter((k) => k.moved);
    if (!back.length && !r.isNew) return;
    r.kids = r.kids.filter((k) => !k.moved);
    for (const k of back) { revertShifts(k.undo); if (!(k.home && sendHome(k))) rows.push({ ...k.orig, num: k.orig.old }); }
    // la suite resserrée reprend ses numéros (si on n'y a pas touché depuis)
    for (const sh of r.shifts || []) { const x = rows.find((y) => y.id === sh.id); if (x && clean(x.num) === sh.to) x.num = sh.from; }
    r.shifts = [];
    if (r.isNew) { rows.splice(rows.indexOf(r), 1); selected.delete(r.id); }
  }
  $("sGroup").addEventListener("click", () => {
    const sel = selRows(); // dans l'ordre affiché
    if (!sel.length) return say("Sélectionnez les pièces à grouper.");
    const heads = sel.filter((r) => r.container || r.kids.length);
    if (heads.length > 1) return say("Grouper : la sélection contient plusieurs groupes — n'en gardez qu'un (les autres pièces s'y ajouteront).");
    const target = heads[0] || null;
    const moving = sel.filter((r) => r !== target);
    if (!moving.length) return say("Grouper : sélectionnez au moins une autre pièce à ranger dans ce groupe.");
    const G = clean(target ? target.num : moving[0].num);
    if (!G || /^\d+\.\d+$/.test(G)) return say("Grouper : numéro de groupe invalide.");
    let group = target;
    if (!group) {
      group = { id: "new:" + (++newSeq), old: "", num: G, name: "", locked: false, container: true, isNew: true, kids: [] };
      rows.splice(rows.indexOf(moving[0]), 0, group);
    }
    // indices suivants : après le plus grand indice entier du groupe
    let next = group.kids.reduce((m, k) => { const n = parseInt(clean(k.local), 10); return /^\d+$/.test(clean(k.local)) ? Math.max(m, n) : m; }, 0);
    const freed = [];
    for (const r of moving) {
      const v = clean(r.num);
      if (/^\d+$/.test(v) && v !== G) freed.push(parseInt(v, 10));
      rows.splice(rows.indexOf(r), 1);
      selected.delete(r.id);
      const orig = { ...r, kids: [] };
      group.kids.push({ id: r.id, oldLocal: null, local: String(++next), name: r.name, locked: r.locked, moved: true, orig });
    }
    // Resserrer la suite : chaque pièce suivante (numéro entier) descend du nombre de numéros libérés avant elle.
    if ($("gCompact").checked && freed.length) compactAfter(group, freed);
    selected = new Set([group.id]);
    anchor = group.id;
    sortRows(); render();
  });
  $("sNone").addEventListener("click", () => { selected.clear(); render(); });
  $("resetAll").addEventListener("click", () => {
    rows = cloneRows(original);
    selected.clear();
    sortRows(); render();
  });
  $("incLocked").addEventListener("change", () => {
    if (!incLocked()) for (const r of rows) if (r.locked) { selected.delete(r.id); if (r.oldLocked) r.num = r.old; }
    sortRows(); render();
  });
  for (const id of ["rFrom", "rTo", "rStart"]) $(id).addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); $("rApply").click(); } });
  $("sStart").addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); $("sNumber").click(); } });
  $("sShift").addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); $("sShiftBtn").click(); } });

  // ---- Recherche : câblage ----
  function selectAllVisible() {
    selected = new Set(visibleRows().filter(movable).map((r) => r.id));
    render();
  }
  function setQuery(v) {
    $("q").value = v;
    query = fold(v);
    $("qClear").hidden = !v;
    // une pièce masquée ne reste pas sélectionnée : les actions ne portent que sur ce qu'on voit
    for (const id of [...selected]) { const r = rows.find((x) => x.id === id); if (!r || !matches(r)) selected.delete(id); }
    render();
  }
  $("q").addEventListener("input", () => setQuery($("q").value));
  $("q").addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); selectAllVisible(); } });
  $("qClear").addEventListener("click", () => { setQuery(""); $("q").focus(); });
  $("selAll").addEventListener("click", selectAllVisible);

  // ---- Cadenas ----
  $("list").addEventListener("click", (e) => {
    const b = e.target.closest(".lock[data-lock-top], .lock[data-lock-kid]");
    if (!b) return;
    e.stopPropagation();
    if (b.dataset.lockTop) {
      const r = rows.find((x) => x.id === b.dataset.lockTop);
      if (!r) return;
      r.locked = !r.locked;
      if (!movable(r)) selected.delete(r.id); // verrouillée : sort de la sélection
    } else {
      for (const r of rows) for (const k of r.kids) if (k.id === b.dataset.lockKid) k.locked = !k.locked;
    }
    render();
  }, true);

  // ---- Saisie une à une ----
  $("list").addEventListener("change", (e) => {
    const inp = e.target.closest(".newnum");
    if (inp && inp.dataset.top) {
      const r = rows.find((x) => x.id === inp.dataset.top);
      if (r) r.num = clean(inp.value);
      sortRows(); render();
      return;
    }
    if (inp && inp.dataset.kid) {
      for (const r of rows) for (const k of r.kids) if (k.id === inp.dataset.kid) k.local = clean(inp.value);
      render();
      return;
    }
    const cb = e.target.closest(".pick");
    if (cb) {
      const id = cb.closest(".row").dataset.id;
      if (cb.checked) selected.add(id); else selected.delete(id);
      anchor = id;
      render();
    }
  });
  $("list").addEventListener("keydown", (e) => {
    if (e.key === "Enter" && e.target.classList.contains("newnum")) { e.preventDefault(); e.target.blur(); }
  });

  // ---- Sélection au clic ----
  $("list").addEventListener("click", (e) => {
    if (e.target.closest("input, .grip")) return;
    const row = e.target.closest(".row[data-id]");
    if (!row) return;
    const r = rows.find((x) => x.id === row.dataset.id);
    if (!r || !movable(r)) return;
    if (e.shiftKey && anchor) {
      const vis = visibleRows(), ids = vis.map((x) => x.id);
      const a = ids.indexOf(anchor), b = ids.indexOf(r.id);
      if (a >= 0 && b >= 0) for (let i = Math.min(a, b); i <= Math.max(a, b); i++) if (movable(vis[i])) selected.add(vis[i].id);
    } else if (e.ctrlKey || e.metaKey) {
      if (selected.has(r.id)) selected.delete(r.id); else selected.add(r.id);
      anchor = r.id;
    } else {
      selected = new Set([r.id]);
      anchor = r.id;
    }
    render();
  });

  // ---- Glisser-déposer LIBRE (comme dans le volet) ----
  // L'étiquette suit le pointeur ; ce qu'on survole décide du résultat (affiché avant de lâcher) :
  //  - une sœur (même niveau / même groupe) → elle prend cette PLACE (les numéros restent aux places) ;
  //  - l'en-tête d'un groupe → la pièce y entre en fin de groupe (6.1…6.4 → 6.5) ;
  //  - une sous-pièce d'un autre groupe → la pièce s'insère à cet endroit du groupe ;
  //  - une pièce de 1er niveau, pour une sous-pièce → elle sort de son groupe.
  // Un groupe se déplace seulement parmi les pièces de 1er niveau (deux niveaux, pas plus).
  // « Resserrer la suite » s'applique aussi : une pièce qui entre dans un groupe libère son numéro.
  const kidMovable = (k) => !k.locked || incLocked();
  const isGroup = (r) => !!r && (r.container || r.kids.length > 0);
  const groupOfKid = (kid) => rows.find((r) => r.kids.includes(kid));
  const intLocal = (k) => (/^\d+$/.test(clean(k.local)) ? parseInt(clean(k.local), 10) : 0);
  let drag = null;

  function hint(text, e) {
    let h = document.getElementById("dropHint");
    if (!h) { h = document.createElement("div"); h.id = "dropHint"; h.className = "drop-hint"; document.body.appendChild(h); }
    h.hidden = !text;
    if (!text) return;
    h.textContent = text;
    h.style.left = Math.round(e.clientX + 14) + "px";
    h.style.top = Math.round(e.clientY + 12) + "px";
  }
  function clearDrop() {
    if (drag && drag.line) { drag.line.remove(); drag.line = null; }
    document.querySelectorAll(".row.drop-into").forEach((x) => x.classList.remove("drop-into"));
    hint("");
  }
  function lineAt(y) {
    const wrap = $("listWrap"), wb = wrap.getBoundingClientRect();
    if (!drag.line) { drag.line = document.createElement("div"); drag.line.className = "drop-line"; wrap.appendChild(drag.line); }
    drag.line.style.top = Math.round(y - wb.top + wrap.scrollTop) + "px";
  }
  function lightGroup(g) {
    document.querySelectorAll(".row").forEach((el) => {
      if (el.dataset.id === g.id || (el.dataset.kid && g.kids.some((k) => k.id === el.dataset.kid))) el.classList.add("drop-into");
    });
  }
  const rowEl = (id) => document.querySelector(`.row[data-id="${CSS.escape(id)}"]`);
  const kidEl = (id) => document.querySelector(`.row[data-kid="${CSS.escape(id)}"]`);
  function blockBottomEl(r) {
    let el = rowEl(r.id);
    for (const k of r.kids) { const ke = kidEl(k.id); if (ke) el = ke; }
    return el ? el.getBoundingClientRect().bottom : 0;
  }

  $("list").addEventListener("pointerdown", (e) => {
    const grip = e.target.closest(".grip");
    if (!grip || grip.classList.contains("off") || e.button !== 0) return;
    e.preventDefault();
    const krow = grip.closest(".row[data-kid]");
    let src;
    if (krow) {
      const g = rows.find((r) => r.kids.some((k) => k.id === krow.dataset.kid));
      if (!g) return;
      src = { kind: "kid", kid: g.kids.find((k) => k.id === krow.dataset.kid), group: g, el: krow };
    } else {
      const row = grip.closest(".row[data-id]");
      const id = row.dataset.id;
      if (!selected.has(id)) { selected = new Set([id]); anchor = id; render(); }
      const r = rows.find((x) => x.id === id);
      const moving = rows.filter((x) => selected.has(x.id) && movable(x));
      src = { kind: "top", row: r, moving, el: rowEl(id) };
    }
    drag = { pointerId: e.pointerId, src, x0: e.clientX, y0: e.clientY, active: false, zone: null, line: null };
    try { $("list").setPointerCapture(e.pointerId); } catch (err) { /* sans capture */ }
  });

  // Zone de dépôt sous le pointeur.
  function zoneAt(e) {
    const d = drag, src = d.src;
    const wrapB = $("list").getBoundingClientRect();
    const x = Math.min(Math.max(e.clientX, wrapB.left + 30), wrapB.right - 30);
    const hit = document.elementFromPoint(x, e.clientY);
    const el = hit && hit.closest(".row");
    if (!el || el === src.el) return null;
    const b = el.getBoundingClientRect(), rel = (e.clientY - b.top) / Math.max(1, b.height), lower = rel > 0.5;
    const meTopGroup = src.kind === "top" && (src.moving.some(isGroup));
    if (el.dataset.kid) {
      const g = rows.find((r) => r.kids.some((k) => k.id === el.dataset.kid));
      const k = g && g.kids.find((x) => x.id === el.dataset.kid);
      if (!g || !k) return null;
      if (meTopGroup) return src.moving.includes(g) ? null : topZone(g, e.clientY > (rowEl(g.id).getBoundingClientRect().top + blockBottomEl(g)) / 2);
      if (src.kind === "kid" && g === src.group) { // réordonner dans le groupe
        if (!kidMovable(k)) return null;
        const others = g.kids.filter((x) => x !== src.kid && kidMovable(x));
        let idx = others.indexOf(k) + (lower ? 1 : 0);
        return { kind: "kidReorder", group: g, index: idx, line: lower ? b.bottom + 1 : b.top - 2 };
      }
      const i = g.kids.indexOf(k) + (lower ? 1 : 0);
      const before = g.kids[i] || null;
      return { kind: "into", group: g, before, line: lower ? b.bottom + 1 : b.top - 2,
        text: before ? `Insérer dans le groupe ${clean(g.num)}` : `Ajouter au groupe → ${clean(g.num)}.${nextLocal(g)}` };
    }
    const r = rows.find((x) => x.id === el.dataset.id);
    if (!r || (src.kind === "top" && src.moving.includes(r))) return null;
    if (meTopGroup) return topZone(r, lower);
    if (isGroup(r)) {
      if (src.kind === "kid" && r === src.group) { // son propre en-tête : en tête de groupe
        const others = r.kids.filter((x) => x !== src.kid && kidMovable(x));
        return { kind: "kidReorder", group: r, index: 0, line: b.bottom + 1 };
      }
      if (rel < 0.28) return src.kind === "top" ? topZone(r, false) : outZone(r, false);
      return { kind: "into", group: r, before: null, text: `Ajouter au groupe → ${clean(r.num)}.${nextLocal(r)}` };
    }
    return src.kind === "top" ? topZone(r, lower) : outZone(r, lower);
  }
  function topZone(r, after) {
    if (!movable(r)) return null;
    const rest = rows.filter((x) => movable(x) && !drag.src.moving.includes(x));
    let idx = rest.indexOf(r);
    if (idx < 0) return null;
    if (after) idx++;
    const y = idx < rest.length ? rowEl(rest[idx].id).getBoundingClientRect().top - 2 : blockBottomEl(rest[rest.length - 1]) + 1;
    return { kind: "topReorder", index: idx, line: y };
  }
  function outZone(r, after) {
    const vis = visibleRows();
    let before = after ? vis[vis.indexOf(r) + 1] || null : r;
    const y = after ? blockBottomEl(r) + 1 : rowEl(r.id).getBoundingClientRect().top - 2;
    return { kind: "out", before, line: y, text: "Sortir du groupe" };
  }
  function nextLocal(g) { return g.kids.reduce((m, k) => Math.max(m, intLocal(k)), 0) + 1; }

  $("list").addEventListener("pointermove", (e) => {
    if (!drag || e.pointerId !== drag.pointerId) return;
    const dx = e.clientX - drag.x0, dy = e.clientY - drag.y0;
    if (!drag.active) {
      if (Math.hypot(dx, dy) < 4) return;
      drag.active = true;
      document.body.classList.add("is-dragging");
      if (drag.src.el) drag.src.el.classList.add("dragging");
    }
    if (drag.src.el) drag.src.el.style.transform = `translate(${dx}px, ${dy}px)`;
    clearDrop();
    const z = zoneAt(e);
    drag.zone = z;
    if (z) {
      if (z.kind === "into") { lightGroup(z.group); if (z.before) lineAt(z.line); }
      else lineAt(z.line);
      hint(z.text || "", e);
    }
    const wrap = $("listWrap"), wb = wrap.getBoundingClientRect();
    if (e.clientY < wb.top + 30) wrap.scrollTop -= 12;
    else if (e.clientY > wb.bottom - 30) wrap.scrollTop += 12;
  });
  const endDrag = (e, cancel) => {
    if (!drag || (e && e.pointerId !== drag.pointerId)) return;
    const d = drag;
    clearDrop();
    drag = null;
    document.body.classList.remove("is-dragging");
    if (d.src.el) { d.src.el.classList.remove("dragging"); d.src.el.style.transform = ""; }
    if (!d.active || cancel || !d.zone) { render(); return; }
    const z = d.zone;
    if (z.kind === "topReorder") {
      const mov = rows.filter((r) => movable(r));
      const slots = mov.map((r) => r.num).slice().sort(naturalCompare); // les PLACES (numéros) ne bougent pas
      const rest = mov.filter((r) => !d.src.moving.includes(r));
      const order = rest.slice(0, z.index).concat(d.src.moving, rest.slice(z.index));
      order.forEach((r, i) => { r.num = slots[i]; });
    } else if (z.kind === "kidReorder") {
      const g = z.group, mov = g.kids.filter(kidMovable);
      const slots = mov.map((k) => clean(k.local)).sort(naturalCompare);
      const rest = mov.filter((k) => k !== d.src.kid);
      rest.splice(Math.min(z.index, rest.length), 0, d.src.kid);
      rest.forEach((k, i) => { k.local = slots[i]; });
      g.kids.sort((a, b) => naturalCompare(clean(a.local), clean(b.local)));
    } else if (z.kind === "into") {
      const items = d.src.kind === "top" ? d.src.moving.filter((r) => !isGroup(r)) : [d.src.kid];
      if (items.length) intoGroup(items, d.src.kind === "kid" ? d.src.group : null, z.group, z.before);
    } else if (z.kind === "out") {
      outOfGroup(d.src.kid, d.src.group, z.before);
    }
    sortRows(); render();
  };
  $("list").addEventListener("pointerup", (e) => endDrag(e, false));
  $("list").addEventListener("pointercancel", (e) => endDrag(e, true));
  document.addEventListener("keydown", (e) => { if (e.key === "Escape" && drag) { e.stopImmediatePropagation(); endDrag(null, true); } }, true);

  // Range des pièces (lignes de 1er niveau, ou une sous-pièce venue d'un autre groupe) dans `g`.
  function intoGroup(items, fromGroup, g, before) {
    const freed = [];
    const newKids = items.map((it) => {
      if (fromGroup) { // sous-pièce d'un autre groupe
        fromGroup.kids.splice(fromGroup.kids.indexOf(it), 1);
        const undo = compactKids(fromGroup, it.local);
        dropEmptyGroup(fromGroup);
        if (it.moved) return { ...it, local: "", undo: (it.undo || []).concat(undo) };
        return { id: it.id, oldLocal: null, local: "", name: it.name, locked: it.locked, oldLocked: it.oldLocked, moved: true,
          orig: { id: it.id, old: clean(fromGroup.old) + "." + it.oldLocal, name: it.name }, home: { gid: fromGroup.id, oldLocal: it.oldLocal }, undo };
      }
      const v = clean(it.num);
      if (/^\d+$/.test(v)) freed.push(parseInt(v, 10));
      rows.splice(rows.indexOf(it), 1);
      selected.delete(it.id);
      if (it.home) { // sous-pièce sortie plus tôt, qu'on range à nouveau
        return { id: it.id, oldLocal: null, local: "", name: it.name, locked: it.locked, oldLocked: it.oldLocked, moved: true,
          orig: { id: it.id, old: it.old, name: it.name }, home: it.home, undo: it.undo || [] };
      }
      return { id: it.id, oldLocal: null, local: "", name: it.name, locked: it.locked, oldLocked: it.oldLocked, moved: true, orig: { ...it, kids: [] } };
    });
    if (!before) {
      let n = nextLocal(g) - 1;
      for (const k of newKids) { k.local = String(++n); g.kids.push(k); }
    } else {
      // insertion à une place : les indices restent aux PLACES (les suivants se décalent)
      const ints = g.kids.filter(kidMovable).map(intLocal);
      let n = Math.max(0, ...g.kids.map(intLocal));
      const slots = g.kids.filter(kidMovable).map((k) => clean(k.local));
      for (let i = 0; i < newKids.length; i++) slots.push(String(++n));
      g.kids.splice(g.kids.indexOf(before), 0, ...newKids);
      slots.sort(naturalCompare);
      g.kids.filter(kidMovable).forEach((k, i) => { k.local = slots[i]; });
    }
    // un ancien membre revenu dans son groupe d'origine retrouve son état
    for (const k of g.kids) if (k.home && k.home.gid === g.id && clean(k.local) === clean(k.home.oldLocal)) { k.moved = false; k.oldLocal = k.home.oldLocal; delete k.home; delete k.orig; }
    if ($("gCompact").checked && freed.length) compactAfter(g, freed);
  }
  // « Resserrer la suite » : chaque pièce entière plus loin descend du nombre de numéros libérés.
  // Toute pièce entière (non verrouillée) située après un numéro libéré descend d'autant — le groupe
  // compris, s'il venait après la pièce qui l'a rejoint. Les décalages sont notés pour « Rétablir ».
  function compactAfter(g, freed) {
    for (const r of rows) {
      if (!movable(r)) continue;
      const v = clean(r.num);
      if (!/^\d+$/.test(v)) continue;
      const n = parseInt(v, 10);
      const k = freed.filter((x) => x < n).length;
      if (k) { r.num = String(n - k); (g.shifts = g.shifts || []).push({ id: r.id, from: v, to: r.num }); }
    }
  }
  // Une sous-pièce quitte son groupe : les indices suivants du groupe se resserrent (7.3 → 7.2…).
  function compactKids(g, local) {
    const done = [];
    if (!$("gCompact").checked || !/^\d+$/.test(clean(local))) return done;
    const l0 = parseInt(clean(local), 10);
    for (const k of g.kids) {
      const n = intLocal(k);
      if (kidMovable(k) && n > l0) { done.push({ kid: k.id, from: k.local, to: String(n - 1) }); k.local = String(n - 1); }
    }
    return done;
  }
  // Annule les décalages notés (seulement là où l'on n'a pas retouché le numéro depuis).
  function revertShifts(list) {
    for (const sh of list || []) {
      if (sh.kid) {
        for (const r of rows) for (const k of r.kids) if (k.id === sh.kid && clean(k.local) === sh.to) k.local = sh.from;
      } else {
        const x = rows.find((y) => y.id === sh.id);
        if (x && clean(x.num) === sh.to) x.num = sh.from;
      }
    }
  }
  // Un groupe sans tête vidé de ses sous-pièces n'a plus de raison d'être.
  function dropEmptyGroup(g) {
    if (g.container && !g.kids.length && rows.includes(g)) { rows.splice(rows.indexOf(g), 1); selected.delete(g.id); }
  }
  // Sous-pièce → pièce de 1er niveau, avant la ligne `before` (ou en fin) : elle prend la place, les
  // pièces entières suivantes (non verrouillées) avancent d'un numéro.
  function outOfGroup(kid, g, before) {
    g.kids.splice(g.kids.indexOf(kid), 1);
    const undo = (kid.undo || []).concat(compactKids(g, kid.local));
    dropEmptyGroup(g);
    let num;
    const bn = before ? clean(before.num) : "";
    if (before && /^\d+$/.test(bn)) {
      const n0 = parseInt(bn, 10);
      for (const r of rows) {
        const v = clean(r.num);
        if (movable(r) && /^\d+$/.test(v) && parseInt(v, 10) >= n0) { r.num = String(parseInt(v, 10) + 1); undo.push({ id: r.id, from: v, to: r.num }); }
      }
      num = String(n0);
    } else {
      num = String(rows.reduce((m, r) => (/^\d+$/.test(clean(r.num)) ? Math.max(m, parseInt(clean(r.num), 10)) : m), 0) + 1);
    }
    const row = kid.moved && kid.orig && !kid.home
      ? { ...kid.orig, num, locked: kid.locked, kids: [] }
      : { id: kid.id, old: kid.moved ? kid.orig.old : clean(g.old) + "." + kid.oldLocal, num, name: kid.name, locked: kid.locked,
          oldLocked: kid.oldLocked, container: false, kids: [], home: kid.home || { gid: g.id, oldLocal: kid.oldLocal } };
    row.undo = undo;
    rows.push(row);
  }

  // ---- Clavier ----
  document.addEventListener("keydown", (e) => {
    const typing = e.target.closest && e.target.closest("input[type=text], .newnum");
    if (e.key === "Escape") {
      e.preventDefault();
      if (query || $("q").value) { setQuery(""); $("q").focus(); return; } // 1er Échap : efface la recherche
      cancel(); return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "a" && !typing) { e.preventDefault(); selectAllVisible(); }
  });

  // ---- Fin ----
  function cancel() { WPMsg.send(JSON.stringify({ cancelled: true })); }
  $("close").addEventListener("click", cancel);
  $("cancel").addEventListener("click", cancel);
  $("apply").addEventListener("click", () => {
    if ($("apply").disabled) return;
    const plan = {
      tops: rows.map((r) => ({ id: r.id, num: clean(r.num) })),
      kids: rows.flatMap((r) => r.kids.map((k) => (k.moved ? { id: k.id, local: clean(k.local), parent: r.id } : { id: k.id, local: clean(k.local) }))),
      newGroups: rows.filter((r) => r.isNew).map((r) => ({ key: r.id, num: clean(r.num) })),
      locks: rows.flatMap((r) => [r, ...r.kids]).filter((x) => !x.isNew && lockChanged(x)).map((x) => ({ id: x.id, locked: !!x.locked })),
    };
    $("apply").disabled = true;
    $("apply").textContent = "Application…";
    WPMsg.send(JSON.stringify({ apply: plan, side }));
  });

  WPMsg.onParent(
    (arg) => {
      let data = null;
      try { data = JSON.parse(arg.message); } catch (e) { data = null; }
      if (!data) return;
      if (data.error) { $("apply").textContent = "Appliquer"; say(data.error); return; } // refus du volet (rare : la liste a changé entre-temps)
      side = data.side === "adv" ? "adv" : "own";
      if (side === "adv") $("title").textContent = "Renuméroter les pièces adverses";
      rows = (data.rows || []).map((r) => ({
        id: r.id, old: r.num, num: r.num, name: r.name, locked: !!r.locked, oldLocked: !!r.locked, container: !!r.container,
        kids: (r.kids || []).map((k) => ({ id: k.id, oldLocal: k.local, local: k.local, name: k.name, locked: !!k.locked, oldLocked: !!k.locked })),
      }));
      original = cloneRows(rows);
      selected.clear();
      render();
      if (data.applied != null) {
        // Retour après « Appliquer » : la fenêtre reste ouverte, sur la liste à jour.
        const st = $("status"), n = data.applied;
        st.classList.remove("err"); st.classList.add("ok");
        st.textContent = n ? `✓ ${n} modification${n > 1 ? "s" : ""} appliquée${n > 1 ? "s" : ""} (↶ dans le volet pour annuler). Vous pouvez continuer.` : "Rien à appliquer.";
        $("cancel").textContent = "Fermer";
        $("apply").textContent = "Appliquer";
      } else {
        try { $("q").focus(); } catch (e) { /* rien */ }
      }
    },
    () => { WPMsg.send(JSON.stringify({ ready: true })); }
  );
  render();
});
