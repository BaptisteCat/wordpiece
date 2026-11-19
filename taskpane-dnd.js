/* global el, reorderPiece, movePieceToGroupUI, movePieceToTopUI, WP */
// ============================================================
// LexRef — volet : déplacer les pièces par GLISSER-DÉPOSER LIBRE.
// On saisit la poignée ⋮⋮ (à gauche de l'étiquette) : l'étiquette suit le pointeur dans les deux
// axes. Selon ce que l'on survole en maintenant le clic :
//  - une SŒUR (même niveau, même groupe)      → trait d'insertion : la pièce prend cette place ;
//  - un GROUPE (en-tête ou ses sous-pièces)   → le groupe s'illumine : la pièce y entre, en fin de
//    groupe sur l'en-tête (6.1…6.4 → 6.5), ou à la place visée parmi ses sous-pièces ;
//  - une pièce de 1er niveau, pour une sous-pièce → trait d'insertion : elle sort de son groupe.
// Une étiquette indique le résultat (« → 6.5 »). Échap annule. Les pièces verrouillées 🔒 ne bougent
// pas et ne comptent pas comme places. Un groupe (ou une pièce qui a des sous-pièces) ne se range
// pas dans un autre groupe : il se déplace parmi les pièces de 1er niveau.
// Au clavier : poignée sélectionnée, flèches ↑ ↓ (parmi les sœurs).
// Chargé AVANT taskpane.js (qui appelle bindDragReorder au démarrage).
// ============================================================
let drag = null; // glissement en cours

const pieceRows = () => [...el("pieceList").querySelectorAll(".piece[data-id]")];
// Lignes des sœurs déplaçables, dans l'ordre affiché (= ordre des numéros).
function movableSiblingRows(row) {
  const parent = row.dataset.parent || "";
  return pieceRows().filter((r) => (r.dataset.parent || "") === parent && r.dataset.locked !== "1");
}
// Bas d'un « bloc » : la ligne, plus les sous-pièces affichées juste dessous (groupe).
function blockBottom(row) {
  let last = row, n = row.nextElementSibling;
  while (n && n.dataset && n.dataset.parent === row.dataset.id) { last = n; n = n.nextElementSibling; }
  return last.getBoundingClientRect().bottom;
}
const isGroupRow = (r) => !!r && r.dataset.group === "1";
const isKidRow = (r) => !!r && !!r.dataset.parent;
// 1re pièce de 1er niveau APRÈS le bloc de `row` (pour « sortir » une sous-pièce sous ce bloc).
function nextTopRowAfter(row) {
  let n = row.nextElementSibling;
  while (n && n.dataset && (n.dataset.parent === row.dataset.id || !n.dataset.id)) n = n.nextElementSibling;
  return n && n.dataset && n.dataset.id && !n.dataset.parent ? n : null;
}
function nextKidRowAfter(row) {
  const n = row.nextElementSibling;
  return n && n.dataset && n.dataset.parent === row.dataset.parent ? n : null;
}

// Étiquette flottante qui annonce le résultat du dépôt.
function hintEl() {
  let h = document.getElementById("dropHint");
  if (!h) { h = document.createElement("div"); h.id = "dropHint"; h.className = "drop-hint hidden"; document.body.appendChild(h); }
  return h;
}
function clearTargets() {
  el("dropIndicator").classList.add("hidden");
  hintEl().classList.add("hidden");
  document.querySelectorAll(".piece.drop-into").forEach((r) => r.classList.remove("drop-into"));
}
function showLine(y) {
  const ind = el("dropIndicator");
  const list = el("pieceList").getBoundingClientRect();
  ind.style.top = Math.round(y) + "px";
  ind.style.left = Math.round(list.left) + "px";
  ind.style.width = Math.round(list.width) + "px";
  ind.classList.remove("hidden");
}
function showInto(gid) {
  pieceRows().forEach((r) => { if (r.dataset.id === gid || r.dataset.parent === gid) r.classList.add("drop-into"); });
}
function showHint(text, e) {
  const h = hintEl();
  if (!text) { h.classList.add("hidden"); return; }
  h.textContent = text;
  h.style.left = Math.round(e.clientX + 14) + "px";
  h.style.top = Math.round(e.clientY + 12) + "px";
  h.classList.remove("hidden");
}

function onGripDown(e) {
  const grip = e.target.closest(".grip");
  if (!grip || grip.disabled || e.button !== 0) return;
  e.preventDefault(); // pas de sélection de texte ; le focus reste utilisable au clavier
  grip.focus({ preventScroll: true });
  try { grip.setPointerCapture(e.pointerId); } catch (err) { /* capture indisponible : on suit quand même */ }
  const row = grip.closest(".piece");
  drag = {
    grip, row, id: grip.dataset.id, pointerId: e.pointerId,
    startX: e.clientX, startY: e.clientY, startScroll: window.scrollY, active: false,
    siblings: movableSiblingRows(row), zone: null,
  };
}

// Où lâcherait-on ? Renvoie { kind: "reorder", index } | { kind: "into", gid, beforeId } |
// { kind: "out", beforeId } | null, plus de quoi l'afficher (ligne ou groupe illuminé, texte).
function computeZone(e) {
  const d = drag, me = d.row;
  const meParent = me.dataset.parent || "";
  const meGroup = isGroupRow(me);
  const list = el("pieceList").getBoundingClientRect();
  // On repère la ligne par la HAUTEUR du pointeur (x ramené dans la liste) : l'étiquette peut être
  // tirée de côté sans perdre la cible. La ligne tirée ne capte pas le pointeur (CSS).
  const x = Math.min(Math.max(e.clientX, list.left + 8), list.right - 8);
  let R = null;
  const hit = document.elementFromPoint(x, e.clientY);
  if (hit) R = hit.closest(".piece[data-id]");
  if (!R) { // entre deux lignes : la plus proche
    let best = null, bd = Infinity;
    for (const r of pieceRows()) {
      if (r === me) continue;
      const b = r.getBoundingClientRect(), dd = Math.min(Math.abs(e.clientY - b.top), Math.abs(e.clientY - b.bottom));
      if (dd < bd) { bd = dd; best = r; }
    }
    R = best;
  }
  if (!R || R === me || R.dataset.parent === d.id) return null; // soi-même, ou ses propres sous-pièces
  const b = R.getBoundingClientRect();
  const rel = (e.clientY - b.top) / Math.max(1, b.height);
  const lower = rel > 0.5;
  const reorderAt = (row, after) => {
    const others = d.siblings.filter((r) => r !== me);
    let idx = others.indexOf(row);
    if (idx < 0) return null;
    if (after) idx++;
    const y = idx < others.length ? others[idx].getBoundingClientRect().top - 4 : blockBottom(others[others.length - 1]) + 3;
    return { kind: "reorder", index: idx, line: y };
  };
  const rGroup = isGroupRow(R), rKid = isKidRow(R);

  // Un groupe (ou une pièce qui a des sous-pièces) reste au 1er niveau : simple réordonnancement.
  if (meGroup) {
    const head = rKid ? pieceRows().find((r) => r.dataset.id === R.dataset.parent) : R;
    if (!head || head === me || head.dataset.locked === "1") return null;
    const hb = head.getBoundingClientRect();
    return reorderAt(head, e.clientY > (hb.top + blockBottom(head)) / 2);
  }
  // Survol d'un EN-TÊTE de groupe
  if (rGroup && !rKid) {
    if (R.dataset.id === meParent) { // son propre groupe : en tête de groupe
      const first = d.siblings.find((r) => r !== me);
      return first ? reorderAt(first, false) : null;
    }
    if (rel < 0.28) { // haut de l'en-tête : AVANT le groupe, au 1er niveau
      if (!meParent) return R.dataset.locked === "1" ? null : reorderAt(R, false);
      return { kind: "out", beforeId: R.dataset.id, line: b.top - 4, text: "Sortir du groupe" };
    }
    return { kind: "into", gid: R.dataset.id, beforeId: null, text: "Ajouter au groupe → " + WP.previewSubNumber(R.dataset.id) };
  }
  // Survol d'une SOUS-PIÈCE
  if (rKid) {
    if (R.dataset.parent === meParent) return R.dataset.locked === "1" ? null : reorderAt(R, lower);
    const before = lower ? nextKidRowAfter(R) : R;
    const head = pieceRows().find((r) => r.dataset.id === R.dataset.parent);
    const headNum = head ? (head.querySelector(".badge-input") || {}).value : "";
    return {
      kind: "into", gid: R.dataset.parent, beforeId: before ? before.dataset.id : null,
      line: lower ? b.bottom + 3 : b.top - 4,
      text: before ? `Insérer dans le groupe ${headNum || ""}`.trim() : "Ajouter au groupe → " + WP.previewSubNumber(R.dataset.parent),
    };
  }
  // Survol d'une pièce de 1er NIVEAU (sans sous-pièces)
  if (!meParent) return R.dataset.locked === "1" ? null : reorderAt(R, lower);
  const beforeRow = lower ? nextTopRowAfter(R) : R;
  return { kind: "out", beforeId: beforeRow ? beforeRow.dataset.id : null, line: lower ? blockBottom(R) + 3 : b.top - 4, text: "Sortir du groupe" };
}

function onGripMove(e) {
  if (!drag || e.pointerId !== drag.pointerId) return;
  const dx = e.clientX - drag.startX;
  const dy = e.clientY - drag.startY + (window.scrollY - drag.startScroll);
  if (!drag.active) {
    if (Math.hypot(dx, dy) < 4) return; // simple clic : pas de glissement
    drag.active = true;
    drag.row.classList.add("dragging");
    document.body.classList.add("is-dragging");
  }
  drag.row.style.transform = `translate(${dx}px, ${dy}px) rotate(${Math.max(-2, Math.min(2, dx / 60))}deg)`;
  clearTargets();
  const z = computeZone(e);
  drag.zone = z;
  if (z) {
    if (z.kind === "into") { showInto(z.gid); if (z.beforeId) showLine(z.line); } // insertion visée : trait
    else if (z.line != null) showLine(z.line);
    showHint(z.text || "", e);
  }
  // Défilement automatique près des bords (longues listes).
  if (e.clientY < 70) window.scrollBy(0, -14);
  else if (e.clientY > window.innerHeight - 50) window.scrollBy(0, 14);
}
function endDrag(e, cancel) {
  if (!drag || (e && e.pointerId !== drag.pointerId)) return;
  const d = drag;
  drag = null;
  try { d.grip.releasePointerCapture(d.pointerId); } catch (err) { /* déjà relâchée */ }
  clearTargets();
  document.body.classList.remove("is-dragging");
  d.row.classList.remove("dragging");
  d.row.style.transform = "";
  const z = d.zone;
  if (!d.active || cancel || !z) return;
  if (z.kind === "reorder") {
    const steps = z.index - d.siblings.indexOf(d.row);
    if (steps) reorderPiece(d.id, steps, { refocusGrip: true });
  } else if (z.kind === "into") movePieceToGroupUI(d.id, z.gid, z.beforeId);
  else if (z.kind === "out") movePieceToTopUI(d.id, z.beforeId);
}

function onGripKey(e) {
  const grip = e.target.closest(".grip");
  if (!grip || grip.disabled) return;
  if (e.key === "ArrowUp" || e.key === "ArrowDown") {
    e.preventDefault();
    reorderPiece(grip.dataset.id, e.key === "ArrowUp" ? -1 : 1, { refocusGrip: true });
  }
}

function bindDragReorder() {
  const list = el("pieceList");
  list.addEventListener("pointerdown", onGripDown);
  list.addEventListener("keydown", onGripKey);
  document.addEventListener("pointermove", onGripMove);
  document.addEventListener("pointerup", (e) => endDrag(e, false));
  document.addEventListener("pointercancel", (e) => endDrag(e, true));
  document.addEventListener("keydown", (e) => { if (e.key === "Escape" && drag) endDrag(null, true); });
}
