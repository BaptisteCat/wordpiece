/* global Office */
// Palette « citer une pièce par son nom ». Reçoit la liste des pièces du volet (parent) via
// messageChild après un handshake « ready », renvoie l'id choisi (ou { cancelled }) via messageParent.

Office.onReady().then(() => { // .then() : démarre toujours APRÈS l'évaluation complète du script
  const params = new URLSearchParams(location.search);
  let pieces = [];
  let filtered = [];
  let sel = 0;
  const q = document.getElementById("q");
  const list = document.getElementById("list");
  const scroller = document.getElementById("scroller");
  q.value = params.get("q") || "";

  const norm = (s) => String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
  const esc = (s) => String(s == null ? "" : s).replace(/[&<>"']/g,
    (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

  function render() {
    const raw = q.value.trim();
    if (!raw) {
      filtered = pieces.slice();
    } else if (/^\d+(?:\.\d+)*$/.test(raw)) {
      // Requête purement numérique (« 4 », « 13 », « 4.1 ») → correspondance EXACTE sur le numéro.
      // On n'inclut PAS les noms contenant ce nombre (ex. « 13 » ≠ « Contrat du 13 janvier »).
      filtered = pieces.filter((p) => String(p.num) === raw);
    } else {
      const f = norm(raw);
      filtered = pieces.filter((p) => norm(p.name).includes(f));
    }
    if (sel >= filtered.length) sel = filtered.length - 1;
    if (sel < 0) sel = 0;
    list.style.columnCount = "1";
    if (!filtered.length) {
      list.innerHTML = `<div class="empty">${pieces.length ? "Aucune pièce ne correspond." : "Aucune pièce dans le document."}</div>`;
      return;
    }
    list.innerHTML = filtered.map((p, i) =>
      `<li data-i="${i}" class="${i === sel ? "sel" : ""}"><span class="num ${p.locked ? "grey" : ""}">${esc(p.num)}</span><span class="nm">${esc(p.name || "(sans nom)")}</span></li>`
    ).join("");
    // Colonnes : DEUX au plus. Une seule (pleine largeur) si tout tient en hauteur ; sinon deux
    // colonnes équilibrées, et la zone défile verticalement si même deux ne suffisent pas.
    // Mesure réelle (les noms longs passent à la ligne, les lignes n'ont pas toutes la même hauteur).
    if (list.scrollHeight > scroller.clientHeight + 1) list.style.columnCount = "2";
    const selEl = list.querySelector("li.sel");
    if (selEl) selEl.scrollIntoView({ block: "nearest", inline: "nearest" });
  }

  function choose(i) {
    const p = filtered[i];
    if (p) Office.context.ui.messageParent(JSON.stringify({ id: p.id }));
  }
  function cancel() { Office.context.ui.messageParent(JSON.stringify({ cancelled: true })); }

  // Pièce de l'AUTRE colonne la plus proche en hauteur de la pièce sélectionnée (dir = +1 droite,
  // -1 gauche) ; -1 s'il n'y en a pas (une seule colonne, ou déjà au bord).
  function neighborInOtherColumn(dir) {
    const items = [...list.querySelectorAll("li[data-i]")];
    const cur = items[sel];
    if (!cur) return -1;
    const r = cur.getBoundingClientRect();
    let best = -1, bestDy = Infinity;
    items.forEach((li, i) => {
      const b = li.getBoundingClientRect();
      if (dir > 0 ? b.left > r.left + 5 : b.left < r.left - 5) {
        const dy = Math.abs(b.top - r.top);
        if (dy < bestDy) { bestDy = dy; best = i; }
      }
    });
    return best;
  }

  q.addEventListener("input", () => { sel = 0; render(); });
  q.addEventListener("keydown", (e) => {
    if (e.key === "ArrowDown") { e.preventDefault(); sel = Math.min(sel + 1, filtered.length - 1); render(); }
    else if (e.key === "ArrowUp") { e.preventDefault(); sel = Math.max(sel - 1, 0); render(); }
    else if (e.key === "ArrowRight" || e.key === "ArrowLeft") {
      // Passage à l'autre colonne ; s'il n'y en a pas, la flèche garde son rôle dans le champ texte.
      const n = neighborInOtherColumn(e.key === "ArrowRight" ? 1 : -1);
      if (n >= 0) { e.preventDefault(); sel = n; render(); }
    }
    else if (e.key === "Enter") { e.preventDefault(); choose(sel); }
    else if (e.key === "Escape") { e.preventDefault(); cancel(); }
  });
  list.addEventListener("click", (e) => {
    const li = e.target.closest("li[data-i]");
    if (li) choose(+li.dataset.i);
  });

  // On enregistre le récepteur AVANT de signaler « ready » (sinon on raterait la liste).
  Office.context.ui.addHandlerAsync(
    Office.EventType.DialogParentMessageReceived,
    (arg) => { try { pieces = JSON.parse(arg.message) || []; } catch (e) { pieces = []; } render(); },
    () => { Office.context.ui.messageParent(JSON.stringify({ ready: true })); }
  );
  q.focus();
});
