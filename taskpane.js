/* global Office, Word, WP */
// ============================================================
// WordPiece — interface du volet (utilise le moteur window.WP)
// ============================================================

let ctxPieceId = null; // pièce sous le curseur (barre contextuelle)

// --- Renommage inline dans le document (bordereau + citations) ---
const NAME_TAG = "wp:name:";
let reconciling = false;
let reconcileTimer = null;
let runningSlash = false;
let lastSlashSig = null;   // texte déclencheur de la dernière commande /p traitée (anti-écho Ctrl+Z rapide)
let lastSlashTime = 0;
let selTimer = null;       // coalescence des DocumentSelectionChanged (rafales de Ctrl+Z)
let lastCreateSig = null;  // déclencheur de la dernière commande qui a CRÉÉ une pièce (anti-écho Ctrl+Z SANS limite de temps)
let lastCreateId = null;   // id de la pièce alors créée — sert à détecter un écho d'annulation
let pendingChain = []; // pièces /pN (ou /pA-B) en attente de nommage, dans l'ordre : [{id, num}]
let pendingChainAlign = null; // alignement d'origine de la ligne (pour reprendre le corps en fin de chaîne)
let pendingFormat = null; // /pf en cours : { baseAlign } — reprise du corps après Entrée
let confirmDeleteId = null; // pièce dont la suppression est en attente de confirmation (inline)
let confirmDeleteBordereau = false; // suppression du bordereau en attente de confirmation (inline)
const occIndex = new Map(); // navigation par occurrence : pièce -> index courant
let renameTimer = null; // propagation automatique du renommage (volet) après une courte pause
let disabled = false; // extension gelée POUR CE DOCUMENT (n'efface rien, coupe tous les automatismes)
let pieceFilter = ""; // filtre de recherche du volet (nom ou numéro)
let flashIds = new Set(); // pièces à faire clignoter au prochain rendu (feedback de réordonnancement)
let undoStack = []; // pile d'annulation : { model: instantané, label } capturés AVANT chaque action
let renameSnap = null; // instantané pris au FOCUS d'un champ de nom (pour annuler un renommage)
let hasErrors = false, hasWarnings = false; // présence d'alertes rouges / oranges (pour les pastilles)
let diagVisible = true; // le bloc d'alerte est-il visible à l'écran (sinon → pastilles dans le header)

// Pastilles d'alerte dans le header : affichées seulement si l'alerte correspondante existe ET que
// le bloc « À corriger / À vérifier » n'est plus visible (scroll).
function renderPips() {
  const pe = el("pipErr"), pw = el("pipWarn");
  if (pe) pe.classList.toggle("hidden", !(hasErrors && !diagVisible));
  if (pw) pw.classList.toggle("hidden", !(hasWarnings && !diagVisible));
}

// Icône poubelle (SVG plein, rendu identique partout — plus lisible que l'emoji 🗑).
const TRASH_SVG =
  '<svg class="ic-trash" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true">' +
  '<path d="M6 19c0 1.1.9 2 2 2h8c1.1 0 2-.9 2-2V7H6v12zM19 4h-3.5l-1-1h-5l-1 1H5v2h14V4z"/></svg>';

// Flèche retour (annuler) et coche (valider) pour la confirmation de suppression.
const BACK_SVG =
  '<svg class="ic-sq" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.3" ' +
  'stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
  '<path d="M9 14 4 9l5-5"/><path d="M4 9h10a6 6 0 0 1 0 12h-3"/></svg>';
const CHECK_SVG =
  '<svg class="ic-sq" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.8" ' +
  'stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M5 12.5 10 17.5 20 6.5"/></svg>';

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    try {
      hidePaneIfAutoStart(); // TOUT DE SUITE : referme le volet si chargement auto → l'écran bleu d'Office disparaît aussitôt
      show("app");
      hide("unsupported");
      WP.load();
      disabled = WP.isDisabled();
      bindUI();
      registerCommandActions(); // les handlers se neutralisent eux-mêmes si désactivé
      // Quand le volet redevient visible : on capture d'éventuelles éditions puis on rafraîchit.
      // (Non bloquant, ne se déclenche PAS sur un simple clic, et inerte si désactivé.)
      document.addEventListener("visibilitychange", () => {
        if (document.visibilityState === "visible") refreshFromDocument();
      });
      render();
      applyStartupBehavior(); // arme le chargement auto (icône + /p) pour les prochaines ouvertures
      if (disabled) return; // extension gelée : aucun automatisme, pas d'auto-démarrage
      registerSelectionHandler();
      // Filet : finalise un nommage en attente même si l'événement Word ne se déclenche pas.
      setInterval(pollPendingFinalize, 1200);
      // Rafraîchit le diagnostic (lecture seule) pour que les alertes suivent l'état réel du doc
      // même sans action de l'utilisateur (ex. citation supprimée à la main).
      setInterval(refreshDiag, 2500);
      // Au démarrage : capturer les éditions inline faites hors ligne, puis synchroniser.
      withBusy(async () => {
        await reconcileNames(false);
        await syncDoc();
      }).then(render).catch(reportError);
    } catch (e) {
      reportError(e); // affiche l'erreur de démarrage dans le bandeau (au lieu d'un volet muet)
    }
  } else {
    show("unsupported");
    hide("app");
  }
});


// Rafraîchissement non bloquant (pas de withBusy → ne bloque jamais les clics du volet).
// IMPORTANT : on ne fait PLUS de WP.reload() ici. Le modèle EN MÉMOIRE fait foi pendant la
// session (un seul runtime) ; recharger depuis les réglages du document pouvait EFFACER les
// pièces si l'enregistrement n'avait pas encore « pris » (cause du volet qui se vide).
async function refreshFromDocument() {
  if (disabled) return;
  try {
    await reconcileNames(false);
    await WP.sync();
    render();
  } catch (e) {
    /* silencieux */
  }
}

// force : réécrit tout (citations + bordereau) même à l'identique — nécessaire après un
// changement de réglages de mise en forme, que la synchro différentielle ne « voit » pas.
async function syncDoc(force) {
  return WP.sync(force ? { force: true } : undefined);
}

// Programme une réconciliation (anti-rebond) après un mouvement de curseur.
function scheduleReconcile() {
  clearTimeout(reconcileTimer);
  reconcileTimer = setTimeout(() => reconcileNames(true), 250);
}

// Cherche un nom édité par l'utilisateur dans le document et le propage à la pièce.
// skipActive : ignore le contrôle où se trouve le curseur (édition en cours) — sauf au démarrage.
async function reconcileNames(skipActive) {
  if (disabled || reconciling || runningSlash) return; // jamais pendant une commande /p… (évite les collisions)
  reconciling = true;
  try {
    let changedId = null;
    let changedName = null;
    let changedNumber = null;
    let forceResync = false; // pièce verrouillée dont le numéro a été modifié dans le texte → réécrire
    await Word.run(async (context) => {
      const sel = context.document.getSelection().parentContentControlOrNullObject;
      sel.load("id,isNullObject");
      const ccs = context.document.body.contentControls;
      ccs.load("items/tag,items/text,items/id");
      await context.sync();

      const skipId = skipActive && !sel.isNullObject ? sel.id : null;

      // Balaye aussi les contrôles imbriqués dans le bordereau (les noms).
      let candidates = ccs.items.slice();
      const bord = ccs.items.find((c) => c.tag === WP.BORDEREAU_TAG);
      if (bord) {
        const nested = bord.contentControls;
        nested.load("items/tag,items/text,items/id");
        await context.sync();
        candidates = candidates.concat(nested.items);
      }

      for (const cc of candidates) {
        if (!cc.tag) continue;
        if (skipId != null && cc.id === skipId) continue; // en cours d'édition
        if (cc.tag.startsWith(NAME_TAG)) {
          const pid = cc.tag.slice(NAME_TAG.length);
          const piece = WP.findPiece(pid);
          if (piece) {
            const t = (cc.text || "").trim();
            // On propage un nom VIDÉ seulement en édition interactive (pas au démarrage),
            // et JAMAIS pour une pièce scannée (elle ne doit pas se vider/disparaître seule).
            if ((t || (skipActive && !piece.scanned)) && t !== piece.name) { changedId = pid; changedName = t; break; }
          }
        } else if (cc.tag.startsWith(WP.TAG_PREFIX)) {
          const pid = cc.tag.slice(WP.TAG_PREFIX.length);
          const piece = WP.findPiece(pid);
          if (piece) {
            // 1) numéro modifié dans la citation ? (IGNORÉ si la pièce est verrouillée → la synchro
            //    réécrira le numéro verrouillé, annulant la modification manuelle dans le texte.)
            const nn = WP.extractNumberFromCitation(cc.text);
            const curNum = WP.stats.numbers.get(pid);
            if (nn != null && curNum != null && nn !== String(curNum)) {
              if (piece.locked) { forceResync = true; break; } // verrouillée : la synchro rétablira le numéro figé
              changedId = pid; changedNumber = nn; break;
            }
            // 2) sinon, nom modifié (vidé accepté en interactif → déclenche la purge) ?
            const nm = WP.extractNameFromCitation(cc.text, pid);
            if (nm != null && (nm !== "" || (skipActive && !piece.scanned)) && nm !== piece.name) { changedId = pid; changedName = nm; break; }
          }
        }
      }
    });

    if (changedId != null) {
      const piece = WP.findPiece(changedId);
      if (piece) {
        if (changedNumber != null) {
          await WP.setNumber(changedId, changedNumber);
          await WP.sync();
          render();
          toast(`Numéro → ${changedNumber}`);
        } else if (changedName != null) {
          piece.name = changedName;
          await WP.save();
          if (changedName.trim()) {
            await WP.sync();
            render();
            toast(`Renommée : « ${changedName} »`);
          }
          // Si le nom a été VIDÉ, la purge ci-dessous supprime la pièce (et ses citations).
        }
      }
    }
    // Pièce verrouillée renumérotée à la main dans le texte : on réécrit sans toucher au modèle.
    if (changedId == null && forceResync) {
      await WP.sync();
      render();
      toast("Pièce verrouillée — numéro rétabli");
    }
    // Purge des pièces au nom vide/espaces (hors nommage en cours) — édition interactive seulement.
    if (skipActive) await pruneNamelessPieces();
  } catch (e) {
    reportError(e);
  } finally {
    reconciling = false;
  }
}

// Supprime les pièces dont le nom est vide ou ne contient que des espaces
// (sauf celles en cours de nommage via /p). Retire aussi leurs éventuelles citations.
async function pruneNamelessPieces() {
  const pend = new Set(pendingChain.map((c) => c.id));
  const counts = (WP.stats && WP.stats.counts) || {};
  // SÉCURITÉ ANTI-PERTE : une pièce CITÉE n'est JAMAIS auto-supprimée (sinon vider son nom
  // détruirait ses citations). Elle reste affichée « Pièce n°X : » — visible et réparable.
  // Les pièces SCANNÉES (contenu réel pré-existant) ne sont jamais auto-supprimées non plus.
  const dead = WP.model.pieces.filter((p) => !(p.name && p.name.trim()) && !pend.has(p.id)
    && !p.scanned && !p.container && !(counts[p.id] > 0));
  if (!dead.length) return;
  for (const p of dead) await WP.deletePieceEverywhere(p.id);
  await WP.sync();
  render();
  toast(dead.length === 1 ? "Pièce sans nom supprimée" : `${dead.length} pièces sans nom supprimées`);
}

// ------------------------------------------------------------
// Actions (délèguent au moteur puis rafraîchissent l'UI)
// ------------------------------------------------------------
async function addPiece(name) {
  const snap = snapshotModel();
  const p = await WP.addPiece(name);
  if (!p) return;
  commitUndo(snap, "ajout d'une pièce");
  render();
  collapseAdd(); // on referme le champ après ajout (volet épuré)
}
function expandAdd() {
  el("addToggle").classList.add("hidden");
  el("addRow").classList.remove("hidden");
  el("newPieceName").value = "";
  el("newPieceName").focus();
}
function collapseAdd() {
  el("addRow").classList.add("hidden");
  el("addToggle").classList.remove("hidden");
  el("newPieceName").value = "";
}
async function renamePiece(id, name) {
  await WP.renamePiece(id, name);
  await withBusy(() => syncDoc());
  render();
}
// Propage le renommage au document PENDANT la frappe (après une courte pause),
// sans re-render le volet → le champ garde le focus, pas besoin de cliquer ailleurs.
async function livePropagateRename(id, value) {
  const piece = WP.findPiece(id);
  if (!piece || piece.name === value.trim()) return;
  if (!value.trim()) return; // on ne vide pas pendant la frappe ; la validation (sortie du champ) purge
  await WP.renamePiece(id, value);
  await WP.sync();
}
// Suppression confirmée EN LIGNE (plus de popup) : retire la pièce partout.
async function deletePiece(id) {
  confirmDeleteId = null;
  if (!WP.findPiece(id)) return;
  commitUndo(snapshotModel(), "suppression");
  await withBusy(async () => {
    await WP.deletePieceEverywhere(id);
    await syncDoc();
  });
  render();
}
// Fusionne des pièces DOUBLONS (même nom) après confirmation : redirige les citations vers keepId.
async function mergePiecesUI(keepId, otherIds) {
  const keep = WP.findPiece(keepId);
  if (!keep || !otherIds.length) return;
  const res = await openPrompt({ mode: "confirm", title: "Fusionner les pièces", label: `« ${keep.name} » — ${otherIds.length + 1} pièces réunies en une seule ?` });
  if (!res || !res.confirmed) return;
  commitUndo(snapshotModel(), "fusion");
  await withBusy(async () => {
    await WP.mergePieces(keepId, otherIds);
    await syncDoc();
  });
  render();
  toast("Pièces fusionnées");
}
// Navigation façon Ctrl+F : va à l'occurrence précédente/suivante de la pièce dans le texte.
async function gotoOccurrence(id, dir) {
  const count = WP.stats.counts[id] || 0;
  if (!count) { toast("Cette pièce n'est pas encore citée."); return; }
  const cur = occIndex.has(id) ? occIndex.get(id) : dir > 0 ? -1 : 0;
  const idx = (((cur + dir) % count) + count) % count;
  occIndex.set(id, idx);
  await WP.selectOccurrence(id, idx);
  toast(`Occurrence ${idx + 1} / ${count}`);
}
// Valide le champ numéro seulement s'il a changé (évite une synchro inutile).
function commitNumber(inp) {
  const id = inp.dataset.id;
  const cur = WP.stats.numbers.get(id);
  if (String(cur ?? "") !== inp.value.trim()) setPieceNumber(id, inp.value);
}
async function setPieceNumber(id, value) {
  commitUndo(snapshotModel(), "changement de numéro");
  await WP.setNumber(id, value);
  await runQuiet(() => syncDoc());
  render();
}
async function togglePieceLock(id) {
  const wasLocked = (WP.findPiece(id) || {}).locked;
  commitUndo(snapshotModel(), wasLocked ? "déverrouillage" : "verrouillage");
  await WP.toggleLock(id);
  await runQuiet(() => syncDoc());
  render();
}
// Capture la position verticale de chaque ligne (pour l'animation FLIP).
function captureRowTops() {
  const m = new Map();
  el("pieceList").querySelectorAll(".piece").forEach((row) => {
    if (row.dataset.id) m.set(row.dataset.id, row.getBoundingClientRect().top);
  });
  return m;
}
// Anime les lignes de leur ANCIENNE position vers la nouvelle (elles glissent = le mouvement se voit).
function flipRows(oldTops) {
  const rows = el("pieceList").querySelectorAll(".piece");
  const moving = [];
  rows.forEach((row) => {
    const id = row.dataset.id;
    if (!id || !oldTops.has(id)) return;
    const delta = oldTops.get(id) - row.getBoundingClientRect().top;
    if (!delta) return;
    row.style.transition = "none";
    row.style.transform = `translateY(${delta}px)`;
    moving.push(row);
  });
  if (!moving.length) return;
  requestAnimationFrame(() => requestAnimationFrame(() => {
    moving.forEach((row) => {
      row.style.transition = "transform 0.32s cubic-bezier(0.22, 1, 0.36, 1)";
      row.style.transform = "";
    });
  }));
}
async function reorderPiece(id, dir) {
  const snap = snapshotModel(); // point d'annulation (avant le déplacement)
  const before = new Map(WP.stats.numbers || new Map()); // numéros AVANT le déplacement
  const oldTops = captureRowTops(); // positions AVANT re-rendu
  const moved = await WP.movePieceStructured(id, dir);
  if (!moved) return; // bord de liste ou déplacement bloqué par un numéro verrouillé → aucun effet
  commitUndo(snap, "changement d'ordre");
  // Le moteur a déjà recalculé les numéros en mémoire → le volet se met à jour TOUT DE SUITE.
  const after = WP.stats.numbers || new Map();
  flashIds = new Set();
  for (const [pid, n] of after) if (String(before.get(pid)) !== String(n)) flashIds.add(pid);
  render();
  flipRows(oldTops); // fait GLISSER les lignes vers leur nouvelle place
  runQuiet(() => syncDoc()); // met à jour les citations du texte en arrière-plan (non bloquant)
}
// Renumérote les pièces selon leur ordre d'apparition dans le texte (action ponctuelle, confirmée).
async function sortByAppearance() {
  const choice = await showModal({
    title: "Numéroter par ordre d'apparition",
    msg: "Les pièces NON verrouillées seront renumérotées 1, 2, 3… selon l'ordre de leur première citation dans le texte.<br>Les pièces verrouillées 🔒 gardent leur numéro.",
    buttons: [
      { label: "Annuler", value: "cancel" },
      { label: "Renuméroter", value: "ok", primary: true },
    ],
  });
  if (choice !== "ok") return;
  commitUndo(snapshotModel(), "numérotation par ordre d'apparition");
  await WP.renumberByAppearance();
  await withBusy(() => syncDoc());
  render();
  toast("Pièces renumérotées par ordre d'apparition");
}
// Verrou général : verrouille tout si au moins une pièce (hors scan) est libre, sinon déverrouille tout.
async function toggleAllLocks() {
  const nonScanned = WP.model.pieces.filter((p) => !p.scanned && !p.container);
  const allLocked = nonScanned.length > 0 && nonScanned.every((p) => p.locked);
  await WP.setAllLocks(!allLocked);
  await withBusy(() => syncDoc());
  render();
}
function renderLockAll() {
  const btn = el("lockAllBtn");
  if (!btn) return;
  const nonScanned = WP.model.pieces.filter((p) => !p.scanned && !p.container);
  const show = nonScanned.length > 0;
  btn.classList.toggle("hidden", !show);
  if (!show) return;
  const allLocked = nonScanned.every((p) => p.locked);
  btn.textContent = allLocked ? "🔓 Tout déverrouiller" : "🔒 Tout verrouiller";
  btn.classList.toggle("on", allLocked);
}
async function insertCitation(id) {
  const piece = WP.findPiece(id);
  if (!piece) return;
  await withBusy(async () => {
    await WP.insertCitation(id);
    await syncDoc();
    toast(`« ${piece.name || "pièce"} » insérée`);
  });
  render();
}
async function gotoPiece(id) {
  await WP.gotoPiece(id).catch(reportError);
}
async function generateBordereau() {
  const existed = WP.stats && WP.stats.hasBordereau;
  await withBusy(async () => {
    await WP.generateBordereau();
    await syncDoc(true); // force : « Mettre à jour » doit régénérer même si rien n'a bougé
    toast(existed ? "Bordereau mis à jour" : "Bordereau généré");
  });
  render();
}

// Bouton bordereau variable : « Générer » (avant), « Mettre à jour » + croix de suppression
// (après), et confirmation en ligne (deux boutons) comme pour les pièces.
function renderBordereauAction() {
  const box = el("bordereauAction");
  if (!box) return;
  const has = !!(WP.stats && WP.stats.hasBordereau);
  if (confirmDeleteBordereau && has) {
    box.innerHTML =
      '<div class="piece-actions confirm">' +
      '<span class="confirm-q">Supprimer le bordereau ?</span>' +
      '<button class="mini sq red" data-bact="canceldel" title="Annuler">' + BACK_SVG + '</button>' +
      '<button class="mini sq green" data-bact="confirmdel" title="Valider la suppression">' + CHECK_SVG + '</button>' +
      '</div>';
    return;
  }
  if (!has) {
    box.innerHTML = '<button class="btn outline bordereau-btn" data-bact="generate">▤ Générer le bordereau</button>';
  } else {
    box.innerHTML =
      '<div class="bordereau-row">' +
      '<button class="btn outline bordereau-btn" data-bact="generate">▤ Mettre à jour</button>' +
      '<button class="mini trash" data-bact="delete" title="Supprimer le bordereau">' + TRASH_SVG + '</button>' +
      '</div>';
  }
}
async function onBordereauAction(e) {
  const b = e.target.closest("[data-bact]");
  if (!b) return;
  const act = b.dataset.bact;
  if (act === "generate") return generateBordereau();
  if (act === "delete") { confirmDeleteBordereau = true; renderBordereauAction(); return; }
  if (act === "canceldel") { confirmDeleteBordereau = false; renderBordereauAction(); return; }
  if (act === "confirmdel") {
    confirmDeleteBordereau = false;
    await withBusy(async () => { await WP.deleteBordereau(); await syncDoc(); });
    render();
    toast("Bordereau supprimé");
  }
}

// Détecte les pièces déjà présentes (« Pièce n° X : Nom ») et les rend gérées par l'extension.
async function scanExisting() {
  const snap = snapshotModel(); // état AVANT le scan (pour l'annulation)
  let r = { wrapped: 0, newPieces: 0 };
  await withBusy(async () => {
    const fmtInput = el("scanFormat");
    const fmt = fmtInput && !el("scanFmtRow").classList.contains("hidden") ? fmtInput.value : "";
    r = await WP.scanExistingPieces(fmt);
    if (r.wrapped) await syncDoc();
  });
  if (r.wrapped) {
    // Annulation du scan : on DÉBALLE les contrôles ajoutés (le texte reste) puis on restaure le modèle.
    const ccIds = r.wrappedCcIds || [];
    pushUndo("reprise des pièces (scan)", async () => {
      await WP.unwrapCitations(ccIds);
      WP.model.pieces = snap.pieces;
      await WP.save();
      await syncDoc();
    });
  }
  render();
  if (!r.wrapped) {
    toast("Aucune pièce « Pièce n° X : Nom » détectée");
  } else {
    const parts = [];
    if (r.newPieces) parts.push(`${r.newPieces} pièce${r.newPieces > 1 ? "s" : ""} ajoutée${r.newPieces > 1 ? "s" : ""}`);
    parts.push(`${r.wrapped} occurrence${r.wrapped > 1 ? "s" : ""} reliée${r.wrapped > 1 ? "s" : ""}`);
    if (r.conflicts) parts.push(`⚠ ${r.conflicts} conflit${r.conflicts > 1 ? "s" : ""} de numéro à arbitrer (voir « À vérifier »)`);
    toast(parts.join(" · ") + " (Ctrl+Z pour annuler)");
  }
}

// ------------------------------------------------------------
// Actions du menu clic droit (exécutées dans ce runtime partagé)
// ------------------------------------------------------------
function registerCommandActions() {
  if (!Office.actions || !Office.actions.associate) return;
  Office.actions.associate("wpRename", wpRenameCmd);
  Office.actions.associate("wpRenumber", wpRenumberCmd);
  Office.actions.associate("wpRemoveHere", wpRemoveHereCmd);
  Office.actions.associate("wpDeletePiece", wpDeletePieceCmd);
}

// Objectif : icône WordPiece présente + /p actif dès l'ouverture (le runtime se
// charge automatiquement = "load"), MAIS volet FERMÉ et SANS l'écran de chargement
// bleu d'Office. Comme Office affiche cet écran pendant qu'il ouvre le volet, on
// referme le volet LE PLUS TÔT POSSIBLE pour que l'écran disparaisse aussitôt.
//
// hidePaneIfAutoStart() : appelée en TOUTE PREMIÈRE ligne de onReady. On ne ferme le
// volet QUE si CE document a le chargement automatique activé (getStartupBehavior ===
// "load"), c.-à-d. que le runtime vient de se charger tout seul à l'ouverture. Signal
// PAR DOCUMENT (pas global) : sur un document neuf où l'utilisateur clique l'icône,
// getStartupBehavior n'est pas "load" → on NE ferme PAS, le volet s'ouvre normalement.
// getStartupBehavior est dans le même jeu d'API que hide() (déjà fonctionnel), donc
// disponible. On referme au plus tôt + quelques tentatives (le volet peut mettre un
// court instant à devenir « fermable ») → l'écran de chargement bleu disparaît vite.
async function hidePaneIfAutoStart() {
  try {
    if (!(Office.addin && Office.addin.hide && Office.addin.getStartupBehavior)) return;
    const prev = await Office.addin.getStartupBehavior();
    if (prev !== Office.StartupBehavior.load) return;  // doc neuf / ouverture manuelle → laisser ouvert
    const doHide = () => { try { Office.addin.hide().catch(() => {}); } catch (e) {} };
    doHide();
    setTimeout(doHide, 100);
    setTimeout(doHide, 300);
    setTimeout(doHide, 700);
  } catch (e) { /* API indisponible : sans effet */ }
}

// Arme le chargement automatique du runtime pour les prochaines ouvertures DE CE
// DOCUMENT (icône présente + /p actif dès l'ouverture). Appelée aussi après (dés)activation.
async function applyStartupBehavior() {
  try {
    if (Office.addin && Office.addin.setStartupBehavior) {
      await Office.addin.setStartupBehavior(Office.StartupBehavior.load);
    }
  } catch (e) { /* API indisponible : sans effet */ }
}

// Bascule activé/désactivé POUR CE DOCUMENT. Ne supprime rien : les pièces,
// citations et bordereau déjà en place restent intacts ; on ne fait que
// couper (ou rallumer) les automatismes de l'extension.
async function toggleDisabled() {
  const turnOff = !disabled;
  try {
    await WP.setDisabled(turnOff);
  } catch (e) { reportError(e); return; }
  disabled = turnOff;
  applyStartupBehavior();
  if (turnOff) {
    unregisterSelectionHandler();
    render();
    toast("WordPiece désactivé pour ce document");
  } else {
    registerSelectionHandler();
    await withBusy(async () => {
      await reconcileNames(false);
      await syncDoc();
    });
    render();
    toast("WordPiece réactivé pour ce document");
  }
}

function openPrompt(params, size) {
  return new Promise((resolve) => {
    const url = location.origin + "/dialog.html?" + new URLSearchParams(params).toString();
    const dim = { height: (size && size.height) || 40, width: (size && size.width) || 34, promptBeforeOpen: false };
    Office.context.ui.displayDialogAsync(url, dim, (res) => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) { resolve(null); return; }
      const dialog = res.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        let data = null; try { data = JSON.parse(arg.message); } catch (e) {}
        dialog.close();
        resolve(data);
      });
      dialog.addEventHandler(Office.EventType.DialogEventReceived, () => resolve(null));
    });
  });
}

// Palette de recherche « citer par nom » : ouvre une fenêtre centrée, lui envoie la liste des pièces
// (via messageChild après le handshake « ready »), et renvoie l'id choisi (ou null si annulé/fermé).
function openPiecePalette(query) {
  return new Promise((resolve) => {
    const numbers = WP.stats.numbers || new Map();
    const pieces = WP.model.pieces.filter((p) => !p.container).map((p) => ({
      id: p.id, name: p.name, num: String(numbers.get(p.id) ?? ""), locked: !!p.locked,
    })).sort((a, b) => WP.naturalCompare(a.num, b.num));
    const url = location.origin + "/search-piece.html?q=" + encodeURIComponent(query || "");
    // Largeur = celle du volet (les étiquettes) : displayDialogAsync ne prend qu'un % de l'écran,
    // donc on convertit la largeur en pixels du volet (window.innerWidth) en pourcentage d'écran.
    const scrW = (window.screen && window.screen.width) || 1920;
    const paneW = window.innerWidth || 340;
    const widthPct = Math.max(9, Math.min(22, Math.round((paneW / scrW) * 100 / 3)));
    Office.context.ui.displayDialogAsync(url, { height: 90, width: widthPct, promptBeforeOpen: false }, (res) => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) { resolve(null); return; }
      const dialog = res.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        let data = null; try { data = JSON.parse(arg.message); } catch (e) {}
        if (data && data.ready) { // le dialogue est prêt → on lui pousse la liste des pièces
          try { dialog.messageChild(JSON.stringify(pieces)); } catch (e) {}
          return;
        }
        dialog.close();
        resolve(data && data.id ? data.id : null);
      });
      dialog.addEventHandler(Office.EventType.DialogEventReceived, () => resolve(null));
    });
  });
}

// Ouvre les réglages dans une FENÊTRE CENTRALE (Office dialog, centrée sur tout Word). Le volet reste
// la source de vérité : il envoie l'état + le schéma, reçoit chaque changement { path, value }, applique.
let settingsDialog = null;
function openSettingsDialog() {
  if (settingsDialog) { try { settingsDialog.close(); } catch (e) {} settingsDialog = null; }
  const url = location.origin + "/settings-dialog.html";
  Office.context.ui.displayDialogAsync(url, { height: 78, width: 56, promptBeforeOpen: false }, (res) => {
    if (res.status !== Office.AsyncResultStatus.Succeeded) return;
    const dialog = res.value;
    settingsDialog = dialog;
    dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
      let data = null;
      try { data = JSON.parse(arg.message); } catch (e) { data = null; }
      if (!data) return;
      if (data.ready) { // la fenêtre est prête → on lui pousse les réglages + le schéma de liaison
        try { dialog.messageChild(JSON.stringify({ settings: WP.model.settings, bindings: SETTING_BINDINGS })); } catch (e) {}
        return;
      }
      if (data.close) { try { dialog.close(); } catch (e) {} settingsDialog = null; return; }
      if (data.path) { // un réglage a changé → on applique (enregistre + resynchronise le document)
        setPath(WP.model.settings, data.path, data.value);
        try { await applySettings(); } catch (e) { reportError(e); }
      }
    });
    dialog.addEventHandler(Office.EventType.DialogEventReceived, () => { settingsDialog = null; });
  });
}

async function wpRenameCmd(event) {
  if (disabled) { event.completed(); return; }
  try {
    const id = await WP.getPieceIdAtSelection();
    if (id) {
      const p = WP.findPiece(id);
      const res = await openPrompt({ mode: "rename", title: "Renommer", label: p ? p.name || "" : "" });
      if (res && res.value != null && !res.cancelled) { await WP.renamePiece(id, res.value); await syncDoc(); render(); }
    }
  } catch (e) { reportError(e); }
  event.completed();
}
async function wpRenumberCmd(event) {
  if (disabled) { event.completed(); return; }
  try {
    const id = await WP.getPieceIdAtSelection();
    if (id) {
      const cur = WP.stats.numbers.get(id);
      const res = await openPrompt({ mode: "renumber", title: "Numéroter", label: String(cur || "") });
      if (res && !res.cancelled) { await WP.setNumber(id, res.value); await syncDoc(); render(); }
    }
  } catch (e) { reportError(e); }
  event.completed();
}
async function wpRemoveHereCmd(event) {
  if (disabled) { event.completed(); return; }
  try {
    const removed = await WP.removeCitationAtSelection();
    if (removed) { await syncDoc(); render(); }
  } catch (e) { reportError(e); }
  event.completed();
}
async function wpDeletePieceCmd(event) {
  if (disabled) { event.completed(); return; }
  try {
    const id = await WP.getPieceIdAtSelection();
    if (id) {
      const p = WP.findPiece(id);
      const res = await openPrompt({ mode: "confirm", title: "Supprimer la pièce", label: p ? p.name || "cette pièce" : "cette pièce" });
      if (res && res.confirmed) { await WP.deletePieceEverywhere(id); await syncDoc(); render(); }
    }
  } catch (e) { reportError(e); }
  event.completed();
}

// ------------------------------------------------------------
// Barre contextuelle
// ------------------------------------------------------------
function registerSelectionHandler() {
  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChanged);
}
function unregisterSelectionHandler() {
  try {
    Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, { handler: onSelectionChanged });
  } catch (e) { /* sans effet */ }
}
// COALESCENCE : un Ctrl+Z maintenu (ou une frappe rapide) émet une RAFALE de
// DocumentSelectionChanged. Les traiter un par un multiplie les Word.run et déstabilise Word.
// On attend donc ~120 ms de calme avant de traiter une seule fois — bien plus sûr.
function onSelectionChanged() {
  if (disabled) return;
  clearTimeout(selTimer);
  selTimer = setTimeout(onSelectionSettled, 120);
}
async function onSelectionSettled() {
  if (disabled) return;
  // Commande /pN en deux temps — traitée en priorité (état le plus frais après la frappe).
  try {
    if (!runningSlash) await processSlash();
  } catch (e) {
    /* silencieux */
  }
  // Quand le curseur bouge, on tente de valider une éventuelle édition de nom terminée.
  scheduleReconcile();
  try {
    const id = await WP.getPieceIdAtSelection();
    if (id) showContextBar(id);
    else hideContextBar();
  } catch (e) {
    hideContextBar();
  }
}

async function processSlash() {
  runningSlash = true;
  try {
    // On regarde D'ABORD si une nouvelle commande « /p… » vient d'être tapée.
    const d = await WP.detectSlashAny();
    // ANTI-Ctrl+Z : une annulation restaure le texte « /pn » et re-déclencherait la MÊME commande
    // (→ pièce recréée pendant qu'on essaie d'annuler). Si la même signature revient dans un court
    // délai, c'est un écho d'annulation ou un doublon d'événement : on l'ignore.
    if (d && d._sig != null) {
      const now = Date.now();
      if (d._sig === lastSlashSig && now - lastSlashTime < 2500) return;
      lastSlashSig = d._sig; lastSlashTime = now;
    }
    // ANTI-Ctrl+Z ROBUSTE (sans limite de temps) pour les commandes qui CRÉENT une pièce.
    // Discriminateur fiable : une pièce créée par un /pn légitime est TOUJOURS citée (elle a son
    // invite/citation dans le texte). Si le même déclencheur revient (« /pn » restauré par une
    // annulation) ET que la pièce créée précédemment n'est plus citée (son insertion vient d'être
    // annulée) ET qu'elle n'est pas en cours de nommage → c'est un écho d'annulation : on NE recrée
    // rien et on SUPPRIME la pièce fantôme (au lieu de se battre avec le Ctrl+Z).
    const CREATE_KINDS = { new: 1, insert: 1, subnew: 1, subinsert: 1, single: 1 };
    if (d && d._sig != null && CREATE_KINDS[d.kind] && d._sig === lastCreateSig && lastCreateId) {
      await WP.refreshStats();
      const prev = WP.findPiece(lastCreateId);
      const inPending = pendingChain.some((c) => c.id === lastCreateId);
      if (prev && !inPending && !(WP.stats.counts[lastCreateId] > 0)) {
        await WP.deletePiece(lastCreateId);
        lastCreateSig = null; lastCreateId = null;
        pendingChain = []; pendingChainAlign = null; pendingFormat = null;
        await WP.refreshStats(); render();
        return;
      }
    }
    // Une commande /p modifie le document (nouveaux contrôles) → une annulation panneau antérieure
    // laisserait des citations orphelines : on vide la pile d'annulation par sécurité.
    if (d) { undoStack = []; renderUndo(); }
    // « /pf » : formate une ligne libre, sans rien capter.
    if (d && d.kind === "format") {
      if (pendingChain.length) await abandonPending();
      const res = await WP.startSlashFormat();
      pendingFormat = res ? { baseAlign: res.baseAlign } : null;
      return;
    }
    // Plage « /p4-8 » : cite si les pièces existent, sinon crée + nomme en chaîne.
    if (d && d.kind === "range") {
      if (pendingChain.length) await abandonPending();
      pendingFormat = null;
      const res = await WP.startSlashRange(d.start, d.end);
      if (res && res.error) { toast(`Plage invalide : ${res.error}`); pendingChain = []; pendingChainAlign = null; }
      else if (res && res.mode === "create") { pendingChain = res.chain; pendingChainAlign = res.baseAlign; }
      else { pendingChain = []; pendingChainAlign = null; }
      // Création + nommage en chaîne → rafraîchissement léger ; citation seule → sync complet.
      if (res && res.mode === "create") await WP.refreshStats(); else await WP.sync();
      render();
      return;
    }
    // « /p5+ » → « Pièce n°5 et s. » (citation simple, sans nommage).
    if (d && d.kind === "etseq") {
      if (pendingChain.length) await abandonPending();
      pendingChain = []; pendingChainAlign = null; pendingFormat = null;
      const res = await WP.startSlashEtSeq(d.start);
      if (res && res.error) toast(`Numéro introuvable : ${res.error}`);
      await WP.sync();
      render();
      return;
    }
    // « /p1,2,5 » → « Pièces n°1, 2 et 5 » (citation multiple ; pièces existantes uniquement).
    if (d && d.kind === "list") {
      if (pendingChain.length) await abandonPending();
      pendingChain = []; pendingChainAlign = null; pendingFormat = null;
      const res = await WP.startSlashList(d.list);
      if (res && res.error) toast(`Pièce n°${res.error} introuvable — aucune pièce créée.`);
      await WP.sync();
      render();
      return;
    }
    // « /p » (ou « /pbail ») → citer une pièce par son NOM. Si la recherche ne peut donner qu'UN
    // seul résultat, la pièce est citée immédiatement (aucune fenêtre) ; sinon on ouvre la palette.
    if (d && d.kind === "search") {
      if (pendingChain.length) await abandonPending();
      pendingChain = []; pendingChainAlign = null; pendingFormat = null;
      const token = "/p" + (d.query || "");
      const query = (d.query || "").trim();
      let id = null;
      if (query) {
        const norm = (s) => String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
        const f = norm(query);
        const hits = WP.model.pieces.filter((p) => !p.container && norm(p.name).includes(f));
        if (hits.length === 1) id = hits[0].id;          // un seul résultat → citation directe
      }
      if (id === null) id = await openPiecePalette(d.query); // 0 ou plusieurs → palette
      if (id) await WP.citeAtToken(token, id);
      else await WP.stripToken(token); // annulé / fermé → on retire le « /p… » tapé
      await WP.sync();
      render();
      return;
    }
    // « /pn », « /pnN », « /pn4. », « /pn4.2 » : création / insertion structurée + nommage.
    if (d && (d.kind === "new" || d.kind === "insert" || d.kind === "subnew" || d.kind === "subinsert")) {
      if (pendingChain.length) await abandonPending();
      pendingFormat = null;
      let res;
      let insertInfo = null;
      if (d.kind === "new") res = await WP.startSlashNew();
      else if (d.kind === "subnew") res = await WP.startSlashSubNew(d.parent);
      else if (d.kind === "subinsert") res = await WP.startSlashSubInsert(d.parent, d.sub);
      else insertInfo = WP.analyzeInsert(d.num); // /pnN : on décide selon l'analyse ci-dessous
      // « /pnN » sur un GROUPE (sous-pièces) ou une pièce VERROUILLÉE = conflit → pop-up À CHAQUE fois.
      if (insertInfo && insertInfo.status === "group") {
        const n = insertInfo.num, next = insertInfo.next;
        try { if (Office.addin && Office.addin.showAsTaskpane) await Office.addin.showAsTaskpane(); } catch (e) {}
        const choice = await showModal({
          title: `N° ${n} : groupe existant`,
          msg: `Le n°${n} est déjà utilisé par un groupe (${n}, ${n}.1, ${n}.2…).<br><b>Décaler</b> : le groupe existant devient ${n + 1}, ${n + 1}.1, ${n + 1}.2… et votre nouvelle pièce prend le n°${n}.<br><b>Ne pas décaler</b> : le groupe existant ne bouge pas et votre nouvelle pièce prend le n°${next}.`,
          buttons: [
            { label: "Ne pas décaler", value: "next" },
            { label: `Décaler ${n}→${n + 1}`, value: "shift", primary: true },
          ],
        });
        if (choice === "shift") res = insertInfo.locked ? await WP.startSlashInsertShift(d.num) : await WP.startSlashInsert(d.num);
        else if (choice === "next") res = await WP.startSlashNew("/pn" + d.num);
        else { pendingChain = []; pendingChainAlign = null; await WP.refreshStats(); render(); return; }
      } else if (insertInfo) {
        res = await WP.startSlashInsert(d.num); // « plain » (décalage silencieux) ou « range » (erreur)
      }
      if (res && res.error) {
        if (res.error === "locked") toast(`Le n°${res.num} est verrouillé — insertion impossible.`);
        else if (res.error === "range") toast(`Numéro ${res.num} hors suite — utilise /pn ou un numéro dans la série.`);
        else if (res.error === "noparent") toast(`Pièce n°${res.num} introuvable pour la sous-pièce.`);
        pendingChain = []; pendingChainAlign = null;
      } else if (res && res.naming) {
        pendingChain = [{ id: res.id, num: res.num }]; pendingChainAlign = res.baseAlign;
        lastCreateSig = d._sig; lastCreateId = res.id; // mémorise la commande créatrice (anti-écho Ctrl+Z)
      } else { pendingChain = []; pendingChainAlign = null; }
      // On NE reconstruit PAS le bordereau maintenant (lourd) : rafraîchissement LÉGER en lecture seule
      // pour que le volet montre les bons numéros. Le vrai sync (bordereau + citations) se fait à la
      // validation du nom → plus de course pendant la frappe.
      await WP.refreshStats();
      render();
      return;
    }
    if (d && d.kind === "single") {
      // Un nouveau /pN → on abandonne proprement un éventuel nommage resté en suspens.
      if (pendingChain.length) await abandonPending();
      pendingFormat = null;
      const res = await WP.startSlashPrompt(d.num);
      if (res && res.naming) {
        pendingChain = [{ id: res.id, num: res.num }]; pendingChainAlign = res.baseAlign;
        lastCreateSig = d._sig; lastCreateId = res.id; // mémorise la commande créatrice (anti-écho Ctrl+Z)
      } else { pendingChain = []; pendingChainAlign = null; }
      // Nommage d'une NOUVELLE pièce → rafraîchissement léger (pas de bordereau) ; citation d'une pièce
      // EXISTANTE → sync complet (la citation vient d'être insérée, il faut la mettre en forme).
      if (res && res.naming) await WP.refreshStats(); else await WP.sync();
      render();
      return;
    }
    // Sinon : reprise du corps après un /pf (curseur passé sur la ligne vide du dessous).
    if (pendingFormat) {
      const okFmt = await WP.finalizeFormat(pendingFormat.baseAlign);
      if (okFmt) pendingFormat = null;
      return;
    }
    // Sinon, on tente de finaliser le nommage en cours (curseur ayant quitté la ligne).
    if (pendingChain.length) {
      const cur = pendingChain[0];
      const next = pendingChain[1] || null;
      const nextName = next ? ((WP.findPiece(next.id) || {}).name || "") : "";
      const nm = await WP.finalizeSlashNamingChain(cur.num, cur.id, next ? next.num : null, nextName, pendingChainAlign);
      if (nm != null) {
        pendingChain.shift();
        if (pendingChain.length === 0) pendingChainAlign = null;
        await WP.sync();
        render();
        toast(`Pièce n°${cur.num}${nm ? " : " + nm : ""}`);
      }
    }
  } catch (e) {
    reportError(e);
  } finally {
    runningSlash = false;
  }
}

// Abandonne un nommage en suspens : supprime les pièces fantômes (vides, jamais citées) et réinitialise.
async function abandonPending() {
  const chain = pendingChain;
  pendingChain = [];
  pendingChainAlign = null;
  for (const item of chain) {
    const piece = WP.findPiece(item.id);
    if (piece && !piece.name && !(WP.stats.counts[item.id] > 0)) {
      await WP.deletePiece(item.id);
    }
  }
}

// FILET DE SÉCURITÉ : finalise un nommage /pN (ou une reprise /pf) en attente MÊME si
// l'événement Word « sélection modifiée » ne se déclenche pas après « Entrée » (peu fiable
// sur certains postes / au 1er chargement dans un doc existant). Appelé par un minuteur ;
// n'agit QUE si le curseur a quitté la ligne du prompt (même condition que la finalisation).
async function pollPendingFinalize() {
  if (disabled || runningSlash) return;
  if (!pendingChain.length && !pendingFormat) return;
  runningSlash = true;
  try {
    if (pendingFormat) {
      if (await WP.finalizeFormat(pendingFormat.baseAlign)) pendingFormat = null;
      return;
    }
    const cur = pendingChain[0];
    const next = pendingChain[1] || null;
    const nextName = next ? ((WP.findPiece(next.id) || {}).name || "") : "";
    // 1) tentative normale (gère l'enchaînement /pA-B ; curseur juste sous le prompt).
    let nm = await WP.finalizeSlashNamingChain(cur.num, cur.id, next ? next.num : null, nextName, pendingChainAlign);
    // 2) sinon, recherche robuste dans le document (curseur ailleurs) — pièce simple uniquement.
    if (nm == null && !next) nm = await WP.finalizePendingBySearch(cur.num, cur.id, pendingChainAlign);
    if (nm != null) {
      pendingChain.shift();
      if (pendingChain.length === 0) pendingChainAlign = null;
      await WP.sync();
      render();
      toast(`Pièce n°${cur.num}${nm ? " : " + nm : ""}`);
    } else if (!next) {
      // NETTOYEUR anti-fantôme : le nommage n'a pas pu se faire. Si l'invite « Pièce n°X : » n'est
      // PLUS dans le document (annulée par Ctrl+Z sans nouvelle commande) et que la pièce est
      // toujours sans nom et non citée, c'est une pièce fantôme → on la retire du volet.
      const piece = WP.findPiece(cur.id);
      const nameless = !(piece && piece.name && piece.name.trim());
      const uncited = !(WP.stats.counts[cur.id] > 0);
      if (piece && nameless && uncited && !(await WP.namingPromptPresent(cur.num))) {
        pendingChain = []; pendingChainAlign = null;
        lastCreateSig = null; lastCreateId = null;
        await WP.deletePiece(cur.id);
        await WP.refreshStats();
        render();
      }
    }
  } catch (e) {
    /* silencieux */
  } finally {
    runningSlash = false;
  }
}
// Rafraîchit le diagnostic depuis l'état RÉEL du document (lecture seule, ne salit pas le doc).
// N'intervient pas pendant une autre opération (nommage, réconciliation, ajout) ni si le volet est caché.
let refreshingStats = false;
async function refreshDiag() {
  if (disabled || refreshingStats || runningSlash || reconciling) return;
  if (pendingChain.length || pendingFormat) return;
  if (document.visibilityState !== "visible") return;
  refreshingStats = true;
  try {
    await WP.refreshStats();
    renderDiag(); // ne touche QUE le bloc #diag (ne perturbe pas les champs en cours d'édition)
  } catch (e) {
    /* silencieux */
  } finally {
    refreshingStats = false;
  }
}
function showContextBar(id) {
  const piece = WP.findPiece(id);
  ctxPieceId = id;
  const num = WP.stats.numbers ? WP.stats.numbers.get(id) : undefined;
  el("ctxLabel").textContent = piece
    ? `Curseur sur ${num ? "Pièce n°" + num : "une pièce"}${piece.name ? " : " + piece.name : ""}`
    : "Curseur sur une citation supprimée";
  el("ctxBar").classList.remove("hidden");
}
function hideContextBar() {
  ctxPieceId = null;
  el("ctxBar").classList.add("hidden");
}
async function removeCitationHere() {
  await withBusy(async () => {
    const removed = await WP.removeCitationAtSelection();
    if (removed) {
      hideContextBar();
      await syncDoc();
      toast("Citation retirée à cet endroit");
    } else {
      toast("Place le curseur sur une citation d'abord");
    }
  });
  render();
}

// ------------------------------------------------------------
// Réglages
// ------------------------------------------------------------
const SETTING_BINDINGS = [
  { id: "citationTemplate", path: ["citationTemplate"], type: "value" },
  { id: "citBold", path: ["citation", "bold"], type: "checked" },
  { id: "citItalic", path: ["citation", "italic"], type: "checked" },
  { id: "citUnderline", path: ["citation", "underline"], type: "checked" },
  { id: "citAlign", path: ["citation", "alignment"], type: "value" },
  { id: "citNewLine", path: ["citation", "newLine"], type: "checked" },
  { id: "rangeStyle", path: ["rangeStyle"], type: "value" },
  { id: "bordTitle", path: ["bordereau", "title"], type: "value" },
  { id: "bordTitleBold", path: ["bordereau", "titleBold"], type: "checked" },
  { id: "bordTitleUnderline", path: ["bordereau", "titleUnderline"], type: "checked" },
  { id: "bordTitleSize", path: ["bordereau", "titleSize"], type: "number" },
  { id: "bordTitleAlign", path: ["bordereau", "titleAlign"], type: "value" },
  { id: "bordListAlign", path: ["bordereau", "listAlign"], type: "value" },
  { id: "bordListSize", path: ["bordereau", "listSize"], type: "number" },
  { id: "bordLineSpacing", path: ["bordereau", "lineSpacing"], type: "value" },
  { id: "bordSpaceBefore", path: ["bordereau", "spaceBefore"], type: "checked" },
  { id: "bordSpaceAfter", path: ["bordereau", "spaceAfter"], type: "checked" },
  { id: "bordLayout", path: ["bordereau", "layout"], type: "value" },
  { id: "bordLabelTpl", path: ["bordereau", "labelTemplate"], type: "value" },
  { id: "bordSep", path: ["bordereau", "separator"], type: "raw" },
  { id: "bordLabelBold", path: ["bordereau", "labelBold"], type: "checked" },
  { id: "bordLabelItalic", path: ["bordereau", "labelItalic"], type: "checked" },
  { id: "bordLabelUnderline", path: ["bordereau", "labelUnderline"], type: "checked" },
  { id: "bordNameBold", path: ["bordereau", "nameBold"], type: "checked" },
  { id: "bordNameItalic", path: ["bordereau", "nameItalic"], type: "checked" },
  { id: "bordNameUnderline", path: ["bordereau", "nameUnderline"], type: "checked" },
];
function getPath(obj, path) { return path.reduce((o, k) => (o ? o[k] : undefined), obj); }
function setPath(obj, path, val) {
  let o = obj;
  for (let i = 0; i < path.length - 1; i++) { if (!o[path[i]]) o[path[i]] = {}; o = o[path[i]]; }
  o[path[path.length - 1]] = val;
}
async function applySettings() {
  await WP.save();
  WP.saveGlobal();
  await withBusy(() => syncDoc(true)); // force : un réglage de mise en forme peut changer sans changer les textes
  render();
}
function renderSettings() {
  for (const b of SETTING_BINDINGS) {
    const elm = el(b.id);
    if (!elm) continue;
    const v = getPath(WP.model.settings, b.path);
    if (b.type === "checked") elm.checked = !!v;
    else elm.value = v ?? "";
  }
  const isList = WP.model.settings.bordereau.layout === "list";
  document.querySelectorAll('[data-layout="list"]').forEach((n) => n.classList.toggle("hidden", !isList));
}

// ------------------------------------------------------------
// Rendu de la liste des pièces
// ------------------------------------------------------------
function applyDisabledUI() {
  const app = el("app");
  if (app) app.classList.toggle("disabled-mode", disabled);
  const banner = el("disabledBanner");
  if (banner) banner.classList.toggle("hidden", !disabled);
  const btn = el("disableBtn");
  if (btn) btn.textContent = disabled ? "Réactiver pour ce document" : "Désactiver pour ce document";
}

function render() {
  applyDisabledUI();
  renderBordereauAction();
  renderSettings();
  const list = el("pieceList");
  list.innerHTML = "";
  const realCount = WP.model.pieces.filter((p) => !p.container).length; // les conteneurs de groupe sont invisibles
  el("pieceCount").textContent = realCount ? `(${realCount})` : "";
  el("emptyState").classList.toggle("hidden", realCount > 0);
  el("listToolbar").classList.toggle("hidden", realCount === 0);
  renderLockAll();
  renderUndo();

  const numbers = WP.stats.numbers || new Map();
  // Tri par NUMÉRO (ordre naturel) — conteneurs de groupe EXCLUS (invisibles ; on ne montre que N.1, N.2…).
  const ordered = [...WP.model.pieces].filter((p) => !p.container).sort((a, b) => WP.naturalCompare(numbers.get(a.id), numbers.get(b.id)));

  // Champ de recherche : visible dès qu'il y a plusieurs pièces ; filtre par nom ou numéro.
  const showSearch = realCount >= 4;
  const search = el("pieceSearch");
  search.classList.toggle("hidden", !showSearch);
  if (!showSearch && pieceFilter) { pieceFilter = ""; search.value = ""; }
  const f = pieceFilter.toLowerCase();
  const shown = f
    ? ordered.filter((p) => (p.name || "").toLowerCase().includes(f) || String(numbers.get(p.id) ?? "").toLowerCase().includes(f))
    : ordered;

  if (f && shown.length === 0) {
    const note = document.createElement("div");
    note.className = "empty muted";
    note.textContent = "Aucune pièce ne correspond à « " + pieceFilter + " ».";
    list.appendChild(note);
  }

  for (const p of shown) {
    const num = numbers.get(p.id) ?? "?";
    const count = WP.stats.counts[p.id] || 0;
    const row = document.createElement("div");
    row.dataset.id = p.id; // pour l'animation de déplacement (FLIP)
    row.className = "piece" + (p.scanned ? " scanned" : "") + (flashIds.has(p.id) ? " flash" : ""); // sous-pièces N.x mises en page comme les autres (pas d'indentation)
    // Barre d'actions : Insérer · loupe (occurrence suivante) · ▲/▼ (réordonner+renuméroter) · verrou · poubelle.
    const noMove = p.locked ? "disabled" : ""; // pièce verrouillée : numéro figé → pas de réordonnancement
    const moveTitle = p.locked ? "Numéro verrouillé — déverrouille pour réordonner" : "";
    const reorderBtns =
      `<button class="mini nav" data-act="moveup" data-id="${p.id}" ${noMove} title="${moveTitle || "Monter la pièce (renumérote)"}" aria-label="Monter">▲</button>
             <button class="mini nav" data-act="movedown" data-id="${p.id}" ${noMove} title="${moveTitle || "Descendre la pièce (renumérote)"}" aria-label="Descendre">▼</button>`;
    const lockBtn =
      `<button class="mini lock push-right ${p.locked ? "on" : ""}" data-act="togglelock" data-id="${p.id}" title="${p.locked ? "Verrouillée (numéro figé, badge gris) — cliquer pour déverrouiller" : "Cliquer pour verrouiller : fige le numéro (badge gris)"}" aria-label="Verrou du numéro">${p.locked ? "🔒" : "🔓"}</button>`;
    const actions =
      confirmDeleteId === p.id
        ? `<div class="piece-actions confirm">
             <span class="confirm-q">Supprimer cette pièce ?</span>
             <button class="mini sq red" data-act="canceldelete" data-id="${p.id}" title="Annuler la suppression" aria-label="Annuler la suppression">${BACK_SVG}</button>
             <button class="mini sq green" data-act="confirmdelete" data-id="${p.id}" title="Valider la suppression" aria-label="Valider la suppression">${CHECK_SVG}</button>
           </div>`
        : `<div class="piece-actions">
             <button class="mini insert" data-act="insert" data-id="${p.id}">⤵ Insérer</button>
             <button class="mini nav" data-act="occ" data-id="${p.id}" ${count ? "" : "disabled"} title="Aller à l'occurrence suivante dans le texte" aria-label="Occurrence suivante">🔍</button>
             ${reorderBtns}
             ${lockBtn}
             <button class="mini trash" data-act="delete" data-id="${p.id}" title="Supprimer la pièce" aria-label="Supprimer la pièce">${TRASH_SVG}</button>
           </div>`;
    const badgeTitle = "Numéro de la pièce — le modifier le verrouille à ce numéro (badge gris) ; vider le champ = numéro automatique (badge bleu)";
    const badgeTitleLocked = "Numéro verrouillé — déverrouillez la pièce (🔓) pour pouvoir le changer";
    row.innerHTML = `
      <div class="piece-top">
        <input class="badge-input ${p.locked ? "locked" : ""}" data-act="setnum" data-id="${p.id}" value="${WP.escapeHtml(String(num))}" title="${p.locked ? badgeTitleLocked : badgeTitle}" ${p.locked ? "readonly" : ""} />
        <input class="piece-name" value="${WP.escapeHtml(p.name)}" data-act="rename" data-id="${p.id}" />
        ${p.scanned ? `<span class="scan-mark" title="Pièce déjà présente dans le document (détectée automatiquement). Protégée du verrou général ; déverrouille-la à la main pour la modifier.">📄</span>` : ""}
        <span class="count ${count ? "" : "zero"}">${count ? count + "×" : "0×"}</span>
      </div>
      ${actions}`;
    list.appendChild(row);
  }
  if (flashIds.size) flashIds = new Set(); // le flash n'est consommé qu'une fois
  renderDiag();
}

function renderDiag() {
  const box = el("diag");
  const numbers = WP.stats.numbers || new Map();
  const hasBord = !!WP.stats.hasBordereau;
  const pend = new Set(pendingChain.map((c) => c.id)); // pièces en cours de nommage : pas encore d'alerte
  const pieces = WP.model.pieces.filter((p) => !p.container);
  const uncited = pieces.filter((p) => !(WP.stats.counts[p.id] > 0) && !pend.has(p.id));

  const errors = [];   // ROUGE
  const warnings = []; // AMBRE

  // Suivi des modifications activé → simple information : WordPiece suspend le suivi le temps de
  // ses propres écritures (renumérotations, bordereau), qui ne polluent donc pas les révisions.
  if (WP.stats.tracking && WP.stats.tracking !== "Off") {
    warnings.push("Suivi des modifications activé — les mises à jour automatiques de WordPiece ne sont pas enregistrées comme révisions.");
  }

  // Citations orphelines (pointent vers une pièce supprimée) → erreur.
  if (WP.stats.orphanCCs > 0) {
    errors.push(`${WP.stats.orphanCCs} citation(s) pointent vers une pièce supprimée (marquées « ⚠ »).`);
  }
  // Pièce présente dans le bordereau mais jamais citée ailleurs dans le document → erreur (rouge).
  if (hasBord && uncited.length > 0) {
    const names = uncited.map((p) => `n°${numbers.get(p.id) ?? "?"} ${p.name || "sans nom"}`).join(", ");
    errors.push(`Pièce(s) présente(s) au bordereau mais jamais citée(s) dans le texte : ${names}.`);
  } else if (uncited.length > 0 && pieces.length > 0) {
    // Sans bordereau : simple rappel (ambre) que des pièces ne sont pas encore citées.
    warnings.push(`${uncited.length} pièce(s) pas encore citée(s) : ${uncited.map((p) => p.name || "sans nom").join(", ")}.`);
  }

  // Doublons de numéro (deux pièces portant le même numéro) → à vérifier.
  const byNum = new Map();
  for (const p of pieces) {
    const n = String(numbers.get(p.id) ?? "").trim();
    if (!n) continue;
    if (!byNum.has(n)) byNum.set(n, []);
    byNum.get(n).push(p);
  }
  const dups = [...byNum.entries()].filter(([, arr]) => arr.length > 1).map(([n]) => n);
  if (dups.length) warnings.push(`Numéro(s) en double : ${dups.map((n) => "n°" + n).join(", ")}.`);
  // Trous dans la numérotation (entiers seulement).
  const ints = [...byNum.keys()].filter((n) => /^\d+$/.test(n)).map(Number).sort((a, b) => a - b);
  if (ints.length >= 2) {
    const missing = [];
    for (let i = ints[0]; i <= ints[ints.length - 1]; i++) { if (!ints.includes(i)) missing.push(i); }
    if (missing.length) warnings.push(`Numéro(s) manquant(s) dans la suite : ${missing.map((n) => "n°" + n).join(", ")}.`);
  }
  // Doublons de NOM → proposition de fusion (dans le bloc ambre).
  const norm = (s) => String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().trim().replace(/\s+/g, " ");
  const byName = new Map();
  for (const p of pieces) {
    const k = norm(p.name);
    if (!k) continue;
    if (!byName.has(k)) byName.set(k, []);
    byName.get(k).push(p);
  }
  const mergeHtml = [...byName.values()].filter((arr) => arr.length > 1).map((arr) => {
    const sorted = [...arr].sort((a, b) => WP.naturalCompare(numbers.get(a.id), numbers.get(b.id)));
    const keep = sorted[0], others = sorted.slice(1);
    const nums = sorted.map((p) => "n°" + (numbers.get(p.id) ?? "?")).join(", ");
    return `<div class="dup-row"><span>Pièces identiques : « ${WP.escapeHtml(keep.name)} » (${WP.escapeHtml(nums)})</span>` +
      `<button class="mini" data-act="merge" data-keep="${keep.id}" data-others="${others.map((p) => p.id).join(",")}">Fusionner</button></div>`;
  }).join("");

  hasErrors = errors.length > 0;
  hasWarnings = warnings.length > 0 || !!mergeHtml;
  if (!hasErrors && !hasWarnings) {
    box.className = "diag hidden";
    box.innerHTML = "";
    renderPips();
    return;
  }
  const li = (a) => "<ul>" + a.map((x) => `<li>${WP.escapeHtml(x)}</li>`).join("") + "</ul>";
  let html = "";
  if (errors.length) html += `<div class="diag-block err"><strong>⚠ À corriger</strong>${li(errors)}</div>`;
  if (warnings.length || mergeHtml) html += `<div class="diag-block warn"><strong>À vérifier</strong>${warnings.length ? li(warnings) : ""}${mergeHtml}</div>`;
  box.className = "diag";
  box.innerHTML = html;
  renderPips();
}

// ------------------------------------------------------------
// Liaison des événements
// ------------------------------------------------------------
function bindUI() {
  el("addToggle").addEventListener("click", expandAdd);
  el("addCancel").addEventListener("click", collapseAdd);
  el("addPieceBtn").addEventListener("click", () => addPiece(el("newPieceName").value));
  el("newPieceName").addEventListener("keydown", (e) => {
    if (e.key === "Enter") addPiece(el("newPieceName").value);
    else if (e.key === "Escape") collapseAdd();
  });

  el("bordereauAction").addEventListener("click", onBordereauAction);
  el("scanBtn").addEventListener("click", scanExisting);
  el("scanFmtToggle").addEventListener("click", () => {
    const row = el("scanFmtRow");
    row.classList.toggle("hidden");
    if (!row.classList.contains("hidden")) el("scanFormat").focus();
  });
  el("ctxRemoveBtn").addEventListener("click", removeCitationHere);
  el("disableBtn").addEventListener("click", toggleDisabled);

  // Croix maison (haut gauche) : referme le volet via l'API Office (shared runtime).
  el("paneClose").addEventListener("click", () => {
    try { if (Office.addin && Office.addin.hide) Office.addin.hide().catch(() => {}); } catch (e) {}
  });

  // Aide (notice) : le « ? » ouvre/ferme le panneau ; la croix le ferme.
  const helpPanel = el("helpPanel");
  el("helpToggle").addEventListener("click", () => helpPanel.classList.toggle("hidden"));
  el("helpClose").addEventListener("click", () => helpPanel.classList.add("hidden"));

  // Réglages : l'engrenage ouvre la FENÊTRE CENTRALE (Office dialog), pas un panneau du volet.
  el("settingsToggle").addEventListener("click", openSettingsDialog);
  document.addEventListener("keydown", (e) => { if (e.key === "Escape") helpPanel.classList.add("hidden"); });
  // Info « i » à côté de « Désactiver » : ouvre/ferme l'explication.
  el("disableInfo").addEventListener("click", () => el("disableInfoPanel").classList.toggle("hidden"));
  el("pieceSearch").addEventListener("input", (e) => { pieceFilter = e.target.value; render(); });
  el("lockAllBtn").addEventListener("click", toggleAllLocks);

  for (const b of SETTING_BINDINGS) {
    const elm = el(b.id);
    if (!elm) continue;
    elm.addEventListener("change", async () => {
      let v;
      if (b.type === "checked") v = elm.checked;
      else if (b.type === "number") v = parseInt(elm.value, 10) || 14;
      else if (b.type === "raw") v = elm.value;
      else v = elm.value;
      setPath(WP.model.settings, b.path, v);
      renderSettings();
      await applySettings();
    });
  }
  el("appearanceBtn").addEventListener("click", sortByAppearance);
  el("undoBtn").addEventListener("click", doUndo);
  el("modal").addEventListener("click", (e) => {
    const b = e.target.closest("[data-mi]");
    if (b) { closeModal((modalButtons[+b.dataset.mi] || {}).value ?? null); return; }
    if (e.target === el("modal")) closeModal(null); // clic sur le fond flouté → ferme
  });
  // Pastilles d'alerte : apparaissent quand le bloc de diagnostic passe SOUS le header (hors vue).
  if (window.IntersectionObserver) {
    const io = new IntersectionObserver((entries) => {
      diagVisible = entries[0].isIntersecting;
      renderPips();
    }, { threshold: 0, rootMargin: "-44px 0px 0px 0px" }); // -44px = hauteur du header sticky
    io.observe(el("diag"));
  }
  const scrollToDiag = () => el("diag").scrollIntoView({ behavior: "smooth", block: "start" });
  el("pipErr").addEventListener("click", scrollToDiag);
  el("pipWarn").addEventListener("click", scrollToDiag);

  el("diag").addEventListener("click", (e) => {
    const btn = e.target.closest('[data-act="merge"]');
    if (!btn) return;
    mergePiecesUI(btn.dataset.keep, (btn.dataset.others || "").split(",").filter(Boolean));
  });
  el("pieceList").addEventListener("click", (e) => {
    const btn = e.target.closest("[data-act]");
    if (!btn || btn.tagName === "INPUT") return;
    const { act, id } = btn.dataset;
    if (act === "insert") insertCitation(id);
    else if (act === "occ") gotoOccurrence(id, +1);
    else if (act === "delete") { confirmDeleteId = id; render(); }
    else if (act === "confirmdelete") deletePiece(id);
    else if (act === "canceldelete") { confirmDeleteId = null; render(); }
    else if (act === "togglelock") togglePieceLock(id);
    else if (act === "moveup") reorderPiece(id, -1);
    else if (act === "movedown") reorderPiece(id, +1);
  });
  // Renommage : propagation automatique 750 ms après la dernière frappe (sans clic).
  el("pieceList").addEventListener("input", (e) => {
    const inp = e.target.closest('input[data-act="rename"]');
    if (!inp) return;
    clearTimeout(renameTimer);
    const id = inp.dataset.id, value = inp.value;
    renameTimer = setTimeout(() => livePropagateRename(id, value), 750);
  });
  // Instantané d'annulation pris quand on ENTRE dans un champ de nom (avant l'édition).
  el("pieceList").addEventListener("focusin", (e) => {
    if (e.target.closest('input[data-act="rename"]')) renameSnap = snapshotModel();
  });
  el("pieceList").addEventListener("change", async (e) => {
    const inp = e.target.closest('input[data-act="rename"]');
    if (!inp) return;
    clearTimeout(renameTimer);
    const snap = renameSnap; renameSnap = null;
    const snapPiece = snap && snap.pieces.find((p) => p.id === inp.dataset.id);
    const oldName = snapPiece ? (snapPiece.name || "") : null;
    if (inp.value.trim()) {
      if (snap && inp.value.trim() !== oldName) commitUndo(snap, "renommage");
      await renamePiece(inp.dataset.id, inp.value);
    } else {
      // Nom vidé et validé → suppression de la pièce (et de ses citations).
      if (snap && oldName) commitUndo(snap, "suppression");
      const piece = WP.findPiece(inp.dataset.id);
      if (piece) { piece.name = ""; await WP.save(); await pruneNamelessPieces(); }
    }
  });
  // Numéro : on valide à la sortie du champ (focusout) et sur Entrée.
  el("pieceList").addEventListener("focusout", (e) => {
    const inp = e.target.closest('input[data-act="setnum"]');
    if (inp) commitNumber(inp);
  });
  el("pieceList").addEventListener("keydown", (e) => {
    if (e.key !== "Enter") return;
    const inp = e.target.closest('input[data-act="setnum"], input[data-act="rename"]');
    if (inp) { e.preventDefault(); inp.blur(); }
  });
}

// ------------------------------------------------------------
// Helpers UI
// ------------------------------------------------------------
function el(id) { return document.getElementById(id); }
function show(id) { el(id).classList.remove("hidden"); }
function hide(id) { el(id).classList.add("hidden"); }

let toastTimer = null;
function toast(msg) {
  let t = el("toast");
  if (!t) { t = document.createElement("div"); t.id = "toast"; t.className = "toast"; document.body.appendChild(t); }
  t.textContent = msg;
  t.classList.add("show");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.remove("show"), 2200);
}
async function withBusy(fn) {
  const app = el("app");
  app.classList.add("busy");
  try { return await fn(); }
  catch (e) { reportError(e); }
  finally { app.classList.remove("busy"); }
}
// Variante SANS grisage du volet : pour les micro-actions rapides (réordonner, verrou, badge…)
// où le voile « occupé » clignotait inutilement.
async function runQuiet(fn) {
  try { return await fn(); }
  catch (e) { reportError(e); }
}

// ---- Annulation (pile d'instantanés du modèle) ----
function snapshotModel() {
  try { return JSON.parse(JSON.stringify(WP.model)); } catch (e) { return null; }
}
// Empile un point d'annulation avec une FONCTION d'annulation quelconque.
function pushUndo(label, undoFn) {
  undoStack.push({ label, undo: undoFn });
  if (undoStack.length > 30) undoStack.shift();
  renderUndo();
}
// Annulation par restauration d'un INSTANTANÉ du modèle (pris AVANT l'action) — pièces uniquement.
function commitUndo(snap, label) {
  if (!snap) return;
  pushUndo(label, async () => {
    WP.model.pieces = snap.pieces; // restaure UNIQUEMENT les pièces (pas les réglages)
    await WP.save();
    await syncDoc();               // réécrit citations + bordereau depuis le modèle restauré
  });
}
function renderUndo() {
  const btn = el("undoBtn");
  if (!btn) return;
  const has = undoStack.length > 0;
  btn.disabled = !has;
  btn.title = has ? "Annuler : " + undoStack[undoStack.length - 1].label : "Rien à annuler";
}
// ---- Pop-up intégrée au volet (modal, fond flouté) ----
let modalButtons = [];
let modalResolve = null;
function showModal({ title, msg, buttons }) {
  return new Promise((resolve) => {
    modalResolve = resolve;
    modalButtons = buttons || [];
    el("modalCard").innerHTML =
      (title ? `<div class="modal-title">${title}</div>` : "") +
      `<div class="modal-msg">${msg}</div>` +
      `<div class="modal-btns">` +
      modalButtons.map((b, i) => `<button class="btn ${b.primary ? "primary" : "outline"}" data-mi="${i}">${b.label}</button>`).join("") +
      `</div>`;
    el("modal").classList.remove("hidden");
  });
}
function closeModal(value) {
  el("modal").classList.add("hidden");
  const r = modalResolve; modalResolve = null; modalButtons = [];
  if (r) r(value);
}

async function doUndo() {
  if (!undoStack.length) return;
  const entry = undoStack.pop();
  await withBusy(() => entry.undo());
  render();
  renderUndo();
  const note = /suppression/i.test(entry.label)
    ? " (Ctrl+Z pour restaurer aussi le texte effacé)"
    : "";
  toast("Annulé : " + entry.label + note);
}
function reportError(e) {
  console.error(e);
  const info = e && (e.debugInfo || e.message || e);
  const msg = (info && info.message ? info.message : String(info));
  toast("Erreur : " + msg);
  // Bandeau PERSISTANT (ne disparaît pas) pour qu'on puisse lire/copier l'erreur.
  const b = el("errBanner");
  if (b) {
    const details = e && e.debugInfo ? JSON.stringify(e.debugInfo) : "";
    b.textContent = "⚠ Erreur : " + msg + (details ? " — " + details : "");
    b.classList.remove("hidden");
  }
}
