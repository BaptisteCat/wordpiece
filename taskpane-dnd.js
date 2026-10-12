/* global el, reorderPiece */
// ============================================================
// LexRef — volet : réordonner les pièces par GLISSER-DÉPOSER.
// On saisit la poignée ⋮⋮ d'une pièce (ou d'un groupe) et on la lâche à sa nouvelle place parmi ses
// SŒURS : même niveau, même groupe. Comme avec les anciennes flèches ▲▼, les pièces verrouillées 🔒
// ne bougent pas et ne comptent pas comme places. Au clavier : poignée sélectionnée, flèches ↑ ↓.
// Le déplacement lui-même passe par reorderPiece (taskpane.js) : renumérotation, annulation,
// animation et mise à jour du texte restent les mêmes.
// Chargé AVANT taskpane.js (qui appelle bindDragReorder au démarrage).
// ============================================================
let drag = null; // glissement en cours

// Lignes des sœurs déplaçables, dans l'ordre affiché (= ordre des numéros).
function movableSiblingRows(row) {
  const parent = row.dataset.parent || "";
  return [...el("pieceList").querySelectorAll(".piece[data-id]")]
    .filter((r) => (r.dataset.parent || "") === parent && r.dataset.locked !== "1");
}
// Bas d'un « bloc » : la ligne, plus les sous-pièces affichées juste dessous (groupe).
function blockBottom(row) {
  let last = row, n = row.nextElementSibling;
  while (n && n.dataset && n.dataset.parent === row.dataset.id) { last = n; n = n.nextElementSibling; }
  return last.getBoundingClientRect().bottom;
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
    startY: e.clientY, startScroll: window.scrollY, active: false,
    siblings: movableSiblingRows(row), target: null,
  };
}
function onGripMove(e) {
  if (!drag || e.pointerId !== drag.pointerId) return;
  const dy = e.clientY - drag.startY + (window.scrollY - drag.startScroll);
  if (!drag.active) {
    if (Math.abs(dy) < 4) return; // simple clic : pas de glissement
    drag.active = true;
    drag.row.classList.add("dragging");
    document.body.classList.add("is-dragging");
  }
  drag.row.style.transform = `translateY(${dy}px)`;
  const others = drag.siblings.filter((r) => r !== drag.row);
  let idx = 0;
  for (const r of others) { const b = r.getBoundingClientRect(); if (e.clientY > b.top + b.height / 2) idx++; }
  drag.target = idx;
  // Trait d'insertion entre deux lignes (ou après la dernière, sous ses éventuelles sous-pièces).
  const ind = el("dropIndicator");
  const list = el("pieceList").getBoundingClientRect();
  if (!others.length) { ind.classList.add("hidden"); return; }
  const y = idx < others.length ? others[idx].getBoundingClientRect().top - 4 : blockBottom(others[others.length - 1]) + 3;
  ind.style.top = Math.round(y) + "px";
  ind.style.left = Math.round(list.left) + "px";
  ind.style.width = Math.round(list.width) + "px";
  ind.classList.remove("hidden");
  // Défilement automatique près des bords (longues listes).
  if (e.clientY < 70) window.scrollBy(0, -14);
  else if (e.clientY > window.innerHeight - 50) window.scrollBy(0, 14);
}
function endDrag(e, cancel) {
  if (!drag || (e && e.pointerId !== drag.pointerId)) return;
  const d = drag;
  drag = null;
  try { d.grip.releasePointerCapture(d.pointerId); } catch (err) { /* déjà relâchée */ }
  el("dropIndicator").classList.add("hidden");
  document.body.classList.remove("is-dragging");
  d.row.classList.remove("dragging");
  d.row.style.transform = "";
  if (!d.active || cancel || d.target == null) return;
  const steps = d.target - d.siblings.indexOf(d.row);
  if (steps) reorderPiece(d.id, steps, { refocusGrip: true });
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
