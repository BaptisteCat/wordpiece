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
let pendingFormat = null; // /pf en cours : { baseAlign, restore, anchor, endAnchor } — reprise après Entrée
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

// Icônes Lucide (lucide.dev), comme dans Juritel : viewBox 24, trait 2, extrémités arrondies, currentColor.
const ICON_PATHS = {
  lock: '<rect width="18" height="11" x="3" y="11" rx="2" ry="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/>',
  unlock: '<rect width="18" height="11" x="3" y="11" rx="2" ry="2"/><path d="M7 11V7a5 5 0 0 1 9.9-1"/>',
  search: '<circle cx="11" cy="11" r="8"/><path d="m21 21-4.3-4.3"/>',
  // Curseur de texte (I, tiré de Lucide « text-cursor-input ») + « + » = insérer une pièce au curseur
  insert: '<path d="M4 4h1a3 3 0 0 1 3 3 3 3 0 0 1 3-3h1"/><path d="M12 20h-1a3 3 0 0 1-3-3 3 3 0 0 1-3 3H4"/>' +
    '<path d="M8 7v10"/><path d="M18 9v6"/><path d="M15 12h6"/>',
  up: '<path d="m18 15-6-6-6 6"/>',
  down: '<path d="m6 9 6 6 6-6"/>',
  file: '<path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/><path d="M14 2v4a2 2 0 0 0 2 2h4"/><path d="m9 15 2 2 4-4"/>',
  list: '<path d="M10 12h11"/><path d="M10 18h11"/><path d="M10 6h11"/><path d="M4 10h2"/><path d="M4 6h1v4"/><path d="M6 18H4c0-1 2-2 2-3s-1-1.5-2-1"/>',
};
function icon(name, size = 14) {
  return `<svg class="ic" width="${size}" height="${size}" viewBox="0 0 24 24" fill="none" stroke="currentColor" ` +
    `stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICON_PATHS[name]}</svg>`;
}

// Icône poubelle (Lucide « trash-2 »).
const TRASH_SVG =
  '<svg class="ic-trash" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" ' +
  'stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
  '<path d="M3 6h18"/><path d="M19 6v14c0 1-1 2-2 2H7c-1 0-2-1-2-2V6"/><path d="M8 6V4c0-1 1-2 2-2h4c1 0 2 1 2 2v2"/>' +
  '<line x1="10" x2="10" y1="11" y2="17"/><line x1="14" x2="14" y1="11" y2="17"/></svg>';

// Flèche retour (annuler) et coche (valider) pour la confirmation de suppression.
const BACK_SVG =
  '<svg class="ic-sq" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.3" ' +
  'stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
  '<path d="M9 14 4 9l5-5"/><path d="M4 9h10a6 6 0 0 1 0 12h-3"/></svg>';
const CHECK_SVG =
  '<svg class="ic-sq" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.8" ' +
  'stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M5 12.5 10 17.5 20 6.5"/></svg>';

// .then() et non onReady(callback) : si Office est DÉJÀ prêt quand ce script s'exécute (script lent à
// arriver, ex. servi depuis GitHub Pages), le callback serait appelé tout de suite, AVANT la fin de
// l'évaluation du fichier → les const déclarées plus bas (SETTING_BINDINGS…) seraient encore
// inaccessibles (« cannot access … before initialization »). Une promesse attend toujours la fin du script.
Office.onReady().then((info) => {
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
    let changedExtra = null;   // complément local (« , page 8 ») ajouté/retiré sur UNE citation
    let changedExtraCc = null; // id du contrôle concerné
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
          const { id: pid, extra } = WP.splitPieceTag(cc.tag);
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
            // 2) sinon, texte modifié : précision LOCALE ou véritable renommage ?
            const nm = WP.extractNameFromCitation(cc.text, pid);
            if (nm != null && nm !== (piece.name || "") + extra) {
              // a) « Nom, page 8 » → précision propre à CETTE citation : on l'enregistre dans le
              //    tag du contrôle. Le nom de la pièce (bordereau, autres citations) ne bouge pas.
              const r = WP.extraAfter(piece.name, nm);
              if (r != null) { changedId = pid; changedExtraCc = cc.id; changedExtra = r; break; }
              // b) précision effacée par l'utilisateur → on retire le complément (sinon la synchro
              //    la remettrait aussitôt).
              if (extra && nm === (piece.name || "")) { changedId = pid; changedExtraCc = cc.id; changedExtra = ""; break; }
              // c) sinon : véritable renommage (on ne réabsorbe pas le complément dans le nom).
              let newName = nm;
              if (extra && newName.endsWith(extra)) newName = newName.slice(0, newName.length - extra.length);
              if (newName !== piece.name && (newName !== "" || (skipActive && !piece.scanned))) {
                changedId = pid; changedName = newName; break;
              }
            }
          }
        }
      }
    });

    if (changedId != null) {
      const piece = WP.findPiece(changedId);
      if (piece) {
        if (changedExtra != null) {
          await WP.setCitationExtra(changedExtraCc, changedExtra);
          await WP.sync();
          render();
          toast(changedExtra ? `Précision enregistrée pour cette citation : « ${changedExtra.trim()} »` : "Précision retirée de cette citation");
        } else if (changedNumber != null) {
          const r = await WP.setNumber(changedId, changedNumber);
          await WP.sync(); // en cas de refus, la synchro rétablit le numéro dans le texte
          render();
          toast(numberError(r) || `Numéro → ${r && r.num != null ? r.num : changedNumber}`);
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
  const snap = snapshotModel();
  const res = await WP.setNumber(id, value);
  const err = numberError(res);
  if (err) { toast(err); render(); return; } // render remet l'ancien numéro dans le badge
  commitUndo(snap, "changement de numéro");
  await runQuiet(() => syncDoc());
  render();
  const typed = String(value || "").trim();
  if (typed && res && res.num != null && String(res.num) !== typed) toast(`Numéro → ${res.num}`);
}
// Message pour un changement de numéro refusé par le moteur (null si accepté).
function numberError(res) {
  if (!res || !res.error) return null;
  if (res.error === "locked") return "Pièce verrouillée — déverrouille-la (🔓) pour changer son numéro.";
  if (res.error === "hasSubs") return "Cette pièce a des sous-pièces : elle ne peut pas devenir elle-même une sous-pièce.";
  if (res.error === "groupTaken") return `Le n°${res.num} est déjà un groupe de sous-pièces.`;
  return null;
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
  btn.innerHTML = allLocked ? icon("unlock", 13) + " Tout déverrouiller" : icon("lock", 13) + " Tout verrouiller";
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
    box.innerHTML = '<button class="btn outline bordereau-btn" data-bact="generate">' + icon("list", 14) + ' Générer le bordereau</button>';
  } else {
    box.innerHTML =
      '<div class="bordereau-row">' +
      '<button class="btn outline bordereau-btn" data-bact="generate">' + icon("list", 14) + ' Mettre à jour</button>' +
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
    if (r.extras) parts.push(`${r.extras} précision${r.extras > 1 ? "s" : ""} de citation conservée${r.extras > 1 ? "s" : ""}`);
    if (r.harmonized) parts.push(`${r.harmonized} variante${r.harmonized > 1 ? "s" : ""} de nom rattachée${r.harmonized > 1 ? "s" : ""} à la même pièce`);
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
    // Relatif à la page du volet (et non à la racine du domaine) : fonctionne en local
    // (https://localhost:3002/) comme sur GitHub Pages (https://…github.io/wordpiece/).
    const url = new URL("dialog.html?" + new URLSearchParams(params).toString(), location.href).href;
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
    const url = new URL("search-piece.html?q=" + encodeURIComponent(query || ""), location.href).href;
    // Largeur ≈ 600 px : deux colonnes où les noms courts tiennent en entier (les longs passent à
    // la ligne). displayDialogAsync ne prend qu'un % de l'écran → conversion, bornée pour les
    // petits écrans (45 %) comme pour les très grands (26 %).
    const scrW = (window.screen && window.screen.width) || 1920;
    const widthPct = Math.max(26, Math.min(45, Math.round((600 / scrW) * 100)));
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
  const url = new URL("settings-dialog.html", location.href).href;
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
      if (res && !res.cancelled) {
        const r = await WP.setNumber(id, res.value);
        await syncDoc();
        render();
        const err = numberError(r);
        if (err) toast(err);
      }
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
        pendingChain = []; pendingChainAlign = null; discardPendingFormat();
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
      discardPendingFormat(); // un /pf précédent resté en suspens
      pendingFormat = (await WP.startSlashFormat()) || null;
      return;
    }
    // Plage « /p4-8 » : cite si les pièces existent, sinon crée + nomme en chaîne.
    if (d && d.kind === "range") {
      if (pendingChain.length) await abandonPending();
      discardPendingFormat();
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
      pendingChain = []; pendingChainAlign = null; discardPendingFormat();
      const res = await WP.startSlashEtSeq(d.start);
      if (res && res.error) toast(`Numéro introuvable : ${res.error}`);
      await WP.sync();
      render();
      return;
    }
    // « /p1,2,5 » → « Pièces n°1, 2 et 5 » (citation multiple ; pièces existantes uniquement).
    if (d && d.kind === "list") {
      if (pendingChain.length) await abandonPending();
      pendingChain = []; pendingChainAlign = null; discardPendingFormat();
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
      pendingChain = []; pendingChainAlign = null; discardPendingFormat();
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
      discardPendingFormat();
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
      discardPendingFormat();
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
      const okFmt = await WP.finalizeFormat(pendingFormat);
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

// Abandonne une reprise /pf en attente et retire ses repères invisibles (signets cachés) du document.
function discardPendingFormat() {
  const pf = pendingFormat;
  pendingFormat = null;
  if (pf && pf.anchor) WP.dropFormatAnchor(pf);
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
      if (await WP.finalizeFormat(pendingFormat)) pendingFormat = null;
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
  // Tri par NUMÉRO (ordre naturel). Un GROUPE sans tête (conteneur) apparaît comme une ligne d'en-tête
  // juste avant ses sous-pièces N.1, N.2… : on peut ainsi renuméroter, déplacer ou verrouiller le groupe.
  const ordered = [...WP.model.pieces].sort((a, b) => WP.naturalCompare(numbers.get(a.id), numbers.get(b.id)));

  // Champ de recherche : visible dès qu'il y a plusieurs pièces ; filtre par nom ou numéro.
  const showSearch = realCount >= 4;
  const search = el("pieceSearch");
  search.classList.toggle("hidden", !showSearch);
  if (!showSearch && pieceFilter) { pieceFilter = ""; search.value = ""; }
  const f = pieceFilter.toLowerCase();
  const matches = (p) => (p.name || "").toLowerCase().includes(f) || String(numbers.get(p.id) ?? "").toLowerCase().includes(f);
  const shown = f
    ? ordered.filter((p) => (p.container ? WP.model.pieces.some((k) => k.parentId === p.id && matches(k)) : matches(p)))
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
    if (p.container) { // en-tête de groupe : numéro + déplacement + verrou (les actions « pièce » n'ont pas de sens)
      const kids = WP.model.pieces.filter((k) => k.parentId === p.id).length;
      row.className = "piece group-head" + (flashIds.has(p.id) ? " flash" : "");
      const dis = p.locked ? "disabled" : "";
      row.innerHTML = `
        <div class="piece-top">
          <input class="badge-input ${p.locked ? "locked" : ""}" data-act="setnum" data-id="${p.id}" value="${WP.escapeHtml(String(num))}" ${p.locked ? "readonly" : ""}
            title="${p.locked ? "Groupe verrouillé — déverrouillez-le (🔓) pour changer son numéro" : "Numéro du groupe — le modifier renumérote toutes ses sous-pièces ; vider le champ = numéro automatique"}" />
          <span class="group-label">Groupe · ${kids} sous-pièce${kids > 1 ? "s" : ""}</span>
          <button class="mini nav" data-act="moveup" data-id="${p.id}" ${dis} title="Monter le groupe (renumérote)" aria-label="Monter le groupe">${icon("up", 15)}</button>
          <button class="mini nav" data-act="movedown" data-id="${p.id}" ${dis} title="Descendre le groupe (renumérote)" aria-label="Descendre le groupe">${icon("down", 15)}</button>
          <button class="mini lock ${p.locked ? "on" : ""}" data-act="togglelock" data-id="${p.id}" title="${p.locked ? "Groupe verrouillé — cliquer pour déverrouiller" : "Cliquer pour verrouiller le numéro du groupe"}" aria-label="Verrou du groupe">${icon(p.locked ? "lock" : "unlock", 14)}</button>
        </div>`;
      list.appendChild(row);
      continue;
    }
    row.className = "piece" + (p.scanned ? " scanned" : "") + (flashIds.has(p.id) ? " flash" : ""); // sous-pièces N.x mises en page comme les autres (pas d'indentation)
    // Barre d'actions : Insérer · loupe (occurrence suivante) · ▲/▼ (réordonner+renuméroter) · verrou · poubelle.
    const noMove = p.locked ? "disabled" : ""; // pièce verrouillée : numéro figé → pas de réordonnancement
    const moveTitle = p.locked ? "Numéro verrouillé — déverrouille pour réordonner" : "";
    const reorderBtns =
      `<button class="mini nav" data-act="moveup" data-id="${p.id}" ${noMove} title="${moveTitle || "Monter la pièce (renumérote)"}" aria-label="Monter">${icon("up", 15)}</button>
             <button class="mini nav" data-act="movedown" data-id="${p.id}" ${noMove} title="${moveTitle || "Descendre la pièce (renumérote)"}" aria-label="Descendre">${icon("down", 15)}</button>`;
    const lockBtn =
      `<button class="mini lock ${p.locked ? "on" : ""}" data-act="togglelock" data-id="${p.id}" title="${p.locked ? "Verrouillée (numéro figé, badge hachuré) — cliquer pour déverrouiller" : "Cliquer pour verrouiller : fige le numéro (badge hachuré)"}" aria-label="Verrou du numéro">${icon(p.locked ? "lock" : "unlock", 14)}</button>`;
    const actions =
      confirmDeleteId === p.id
        ? `<div class="piece-actions confirm">
             <span class="confirm-q">Supprimer cette pièce ?</span>
             <button class="mini sq red" data-act="canceldelete" data-id="${p.id}" title="Annuler la suppression" aria-label="Annuler la suppression">${BACK_SVG}</button>
             <button class="mini sq green" data-act="confirmdelete" data-id="${p.id}" title="Valider la suppression" aria-label="Valider la suppression">${CHECK_SVG}</button>
           </div>`
        : `<div class="piece-actions">
             ${lockBtn}
             <button class="mini insert" data-act="insert" data-id="${p.id}" title="Insérer la citation à l'endroit du curseur" aria-label="Insérer la citation à l'endroit du curseur">${icon("insert", 17)}</button>
             <button class="mini nav" data-act="occ" data-id="${p.id}" ${count ? "" : "disabled"} title="Aller à l'occurrence suivante dans le texte" aria-label="Occurrence suivante">${icon("search", 14)}</button>
             ${reorderBtns}
             <button class="mini trash push-right" data-act="delete" data-id="${p.id}" title="Supprimer la pièce" aria-label="Supprimer la pièce">${TRASH_SVG}</button>
           </div>`;
    const badgeTitle = "Numéro de la pièce — le modifier le verrouille à ce numéro (badge hachuré) ; vider le champ = numéro automatique (badge plein)";
    const badgeTitleLocked = "Numéro verrouillé — déverrouillez la pièce (🔓) pour pouvoir le changer";
    row.innerHTML = `
      <div class="piece-top">
        <input class="badge-input ${p.locked ? "locked" : ""}" data-act="setnum" data-id="${p.id}" value="${WP.escapeHtml(String(num))}" title="${p.locked ? badgeTitleLocked : badgeTitle}" ${p.locked ? "readonly" : ""} />
        <input class="piece-name" value="${WP.escapeHtml(p.name)}" data-act="rename" data-id="${p.id}" />
        ${p.scanned ? `<span class="scan-mark" title="Pièce déjà présente dans le document (détectée automatiquement). Protégée du verrou général ; déverrouille-la à la main pour la modifier.">${icon("file", 14)}</span>` : ""}
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

  // Chaque point de diagnostic porte une CLÉ stable : elle permet de le masquer durablement
  // (« Ignorer ») sans masquer les autres, et de le faire réapparaître si la situation change
  // (la clé inclut ce qui caractérise le point : numéro concerné, liste de pièces, etc.).
  const errors = [];   // ROUGE   — { key, text }
  const warnings = []; // AMBRE   — { key, text }
  const merges = [];   // AMBRE   — { key, html } (proposition de fusion)

  // Suivi des modifications activé → simple information : WordPiece suspend le suivi le temps de
  // ses propres écritures (renumérotations, bordereau), qui ne polluent donc pas les révisions.
  if (WP.stats.tracking && WP.stats.tracking !== "Off") {
    warnings.push({ key: "tracking", text: "Suivi des modifications activé — les mises à jour automatiques de WordPiece ne sont pas enregistrées comme révisions." });
  }

  // Citations orphelines (pointent vers une pièce supprimée) → erreur.
  if (WP.stats.orphanCCs > 0) {
    errors.push({ key: "orphans:" + WP.stats.orphanCCs, text: `${WP.stats.orphanCCs} citation(s) pointent vers une pièce supprimée (marquées « ⚠ »).` });
  }
  // Pièce présente dans le bordereau mais jamais citée ailleurs dans le document → erreur (rouge).
  if (hasBord && uncited.length > 0) {
    const names = uncited.map((p) => `n°${numbers.get(p.id) ?? "?"} ${p.name || "sans nom"}`).join(", ");
    errors.push({ key: "uncited:" + uncited.map((p) => p.id).sort().join(","), text: `Pièce(s) présente(s) au bordereau mais jamais citée(s) dans le texte : ${names}.` });
  } else if (uncited.length > 0 && pieces.length > 0) {
    // Sans bordereau : simple rappel (ambre) que des pièces ne sont pas encore citées.
    warnings.push({ key: "uncited:" + uncited.map((p) => p.id).sort().join(","), text: `${uncited.length} pièce(s) pas encore citée(s) : ${uncited.map((p) => p.name || "sans nom").join(", ")}.` });
  }

  // Doublons de numéro : UN point par numéro concerné (on peut donc en ignorer un seul). Les groupes
  // comptent aussi (un groupe 4 et une pièce 4 distincte = doublon, même si le groupe n'a pas de nom).
  const byNum = new Map();
  for (const p of WP.model.pieces) {
    const n = String(numbers.get(p.id) ?? "").trim();
    if (!n) continue;
    if (!byNum.has(n)) byNum.set(n, []);
    byNum.get(n).push(p);
  }
  const label = (p) => (p.container ? "groupe de sous-pièces" : p.name || "sans nom");
  for (const [n, arr] of byNum) {
    if (arr.length < 2) continue;
    const text = `Numéro en double : n°${n} (${arr.map(label).join(" / ")}).`;
    if (arr.some((p) => p.container)) { warnings.push({ key: "dupnum:" + n, text }); continue; }
    // Même numéro = souvent la même pièce saisie deux fois : fusion en un clic, en gardant la plus citée.
    const cnt = (p) => WP.stats.counts[p.id] || 0;
    const keep = [...arr].sort((a, b) => cnt(b) - cnt(a) || (b.name ? 1 : 0) - (a.name ? 1 : 0))[0];
    const others = arr.filter((p) => p !== keep);
    merges.push({
      key: "dupnum:" + n,
      html: `<span>${WP.escapeHtml(text)}</span>` +
        `<button class="mini" data-act="merge" data-keep="${keep.id}" data-others="${others.map((p) => p.id).join(",")}" title="Garder « ${WP.escapeHtml(keep.name || "sans nom")} » : les citations des autres la rejoignent">Fusionner</button>`,
    });
  }
  // Trous dans la numérotation : UN point par numéro manquant.
  const ints = [...byNum.keys()].filter((n) => /^\d+$/.test(n)).map(Number).sort((a, b) => a - b);
  if (ints.length >= 2) {
    for (let i = ints[0]; i <= ints[ints.length - 1]; i++) {
      if (!ints.includes(i)) warnings.push({ key: "gap:" + i, text: `Numéro manquant dans la suite : n°${i}.` });
    }
  }
  // Doublons de NOM → proposition de fusion.
  const norm = (s) => String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().trim().replace(/\s+/g, " ");
  const byName = new Map();
  for (const p of pieces) {
    const k = norm(p.name);
    if (!k) continue;
    if (!byName.has(k)) byName.set(k, []);
    byName.get(k).push(p);
  }
  for (const arr of byName.values()) {
    if (arr.length < 2) continue;
    const sorted = [...arr].sort((a, b) => WP.naturalCompare(numbers.get(a.id), numbers.get(b.id)));
    const keep = sorted[0], others = sorted.slice(1);
    const nums = sorted.map((p) => "n°" + (numbers.get(p.id) ?? "?")).join(", ");
    merges.push({
      key: "dupname:" + norm(keep.name),
      html: `<span>Pièces identiques : « ${WP.escapeHtml(keep.name)} » (${WP.escapeHtml(nums)})</span>` +
        `<button class="mini" data-act="merge" data-keep="${keep.id}" data-others="${others.map((p) => p.id).join(",")}">Fusionner</button>`,
    });
  }

  // Filtrage : on retire les points que l'utilisateur a explicitement ignorés pour ce document.
  const keep = (it) => !WP.isWarningIgnored(it.key);
  const shownErrors = errors.filter(keep);
  const shownWarnings = warnings.filter(keep);
  const shownMerges = merges.filter(keep);
  const nIgnored = WP.ignoredCount();

  hasErrors = shownErrors.length > 0;
  hasWarnings = shownWarnings.length > 0 || shownMerges.length > 0;
  if (!hasErrors && !hasWarnings && !nIgnored) {
    box.className = "diag hidden";
    box.innerHTML = "";
    renderPips();
    return;
  }

  // Chaque ligne porte son bouton « Ignorer » (masquage durable, réversible).
  const ignoreBtn = (key) =>
    `<button class="mini diag-ignore" data-act="ignorewarn" data-key="${WP.escapeHtml(key)}" title="Ne plus signaler ce point pour ce document">Ignorer</button>`;
  const li = (items) => "<ul>" + items.map((it) =>
    `<li><span class="diag-txt">${WP.escapeHtml(it.text)}</span>${ignoreBtn(it.key)}</li>`).join("") + "</ul>";
  const mergeRows = (items) => items.map((it) =>
    `<div class="dup-row">${it.html}${ignoreBtn(it.key)}</div>`).join("");

  let html = "";
  if (shownErrors.length) html += `<div class="diag-block err"><strong>⚠ À corriger</strong>${li(shownErrors)}</div>`;
  if (shownWarnings.length || shownMerges.length) {
    html += `<div class="diag-block warn"><strong>À vérifier</strong>${shownWarnings.length ? li(shownWarnings) : ""}${mergeRows(shownMerges)}</div>`;
  }
  if (nIgnored) {
    html += `<div class="diag-ignored">${nIgnored} avertissement${nIgnored > 1 ? "s" : ""} ignoré${nIgnored > 1 ? "s" : ""} ` +
      `<button class="link-btn" data-act="restorewarn">Tout réafficher</button></div>`;
  }
  box.className = "diag";
  box.innerHTML = html;
  renderPips();
}

// Masque durablement un point de diagnostic (stocké dans le document, réversible).
async function ignoreWarning(key) {
  await WP.ignoreWarning(key);
  renderDiag();
  toast("Avertissement ignoré pour ce document");
}
async function restoreWarnings() {
  await WP.restoreWarnings();
  renderDiag();
  toast("Avertissements réaffichés");
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
    const btn = e.target.closest("[data-act]");
    if (!btn) return;
    const act = btn.dataset.act;
    if (act === "merge") mergePiecesUI(btn.dataset.keep, (btn.dataset.others || "").split(",").filter(Boolean));
    else if (act === "ignorewarn") ignoreWarning(btn.dataset.key);
    else if (act === "restorewarn") restoreWarnings();
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
