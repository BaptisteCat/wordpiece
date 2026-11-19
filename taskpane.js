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
let selectedPieceId = null; // pièce sélectionnée d'un clic dans le volet (touche Suppr → suppression)
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
// Mode du volet (V9) : "own" = nos pièces (défaut) ; "adv" = gestion des pièces adverses.
let paneSide = "own";
let focusWarned = false; // avertissement « API de focus absente » affiché une seule fois par session
const isAdvMode = () => paneSide === "adv";
// Libellé d'une pièce selon son côté (« Pièce » / « Pièce adverse »).
const sideLabel = (id) => (WP.sideOf(id) === "adv" ? "Pièce adverse" : isLetterDoc() ? "Pièce jointe" : "Pièce");
// Type du document : acte (défaut) ou courrier. Un courrier parle de « pièces jointes » et son
// bloc de fin est un simple récapitulatif — même mécanique que le bordereau, autre vocabulaire.
const isLetterDoc = () => WP.isLetter();
const bordWord = () => (isLetterDoc() ? "récapitulatif" : "bordereau");
const BordWord = () => (isLetterDoc() ? "Récapitulatif" : "Bordereau");

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
  file: '<path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/><path d="M14 2v4a2 2 0 0 0 2 2h4"/><path d="m9 15 2 2 4-4"/>',
  check: '<path d="M20 6 9 17l-5-5"/>',
  list: '<path d="M10 12h11"/><path d="M10 18h11"/><path d="M10 6h11"/><path d="M4 10h2"/><path d="M4 6h1v4"/><path d="M6 18H4c0-1 2-2 2-3s-1-1.5-2-1"/>',
  // Poignée de glisser-déposer (« grip-vertical »)
  grip: '<circle cx="9" cy="12" r="1"/><circle cx="9" cy="5" r="1"/><circle cx="9" cy="19" r="1"/><circle cx="15" cy="12" r="1"/><circle cx="15" cy="5" r="1"/><circle cx="15" cy="19" r="1"/>',
  // Pièce communiquée (« mail-check ») / pas encore (« mail »)
  mailCheck: '<path d="M22 13V6a2 2 0 0 0-2-2H4a2 2 0 0 0-2 2v12c0 1.1.9 2 2 2h8"/><path d="m22 7-8.97 5.7a1.94 1.94 0 0 1-2.06 0L2 7"/><path d="m16 19 2 2 4-4"/>',
  mail: '<rect width="20" height="16" x="2" y="4" rx="2"/><path d="m22 7-8.97 5.7a1.94 1.94 0 0 1-2.06 0L2 7"/>',
  plus: '<path d="M5 12h14"/><path d="M12 5v14"/>',
  x: '<path d="M18 6 6 18"/><path d="m6 6 12 12"/>',
};
function icon(name, size = 14, strokeWidth = 2) {
  return `<svg class="ic" width="${size}" height="${size}" viewBox="0 0 24 24" fill="none" stroke="currentColor" ` +
    `stroke-width="${strokeWidth}" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICON_PATHS[name]}</svg>`;
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
      applyUiPrefs(); // thème, densité, onglets : avant le premier rendu
      disabled = WP.isDisabled();
      bindUI();
      registerCommandActions(); // les handlers se neutralisent eux-mêmes si désactivé
      // Quand le volet redevient visible : on capture d'éventuelles éditions puis on rafraîchit.
      // (Non bloquant, ne se déclenche PAS sur un simple clic, et inerte si désactivé.)
      document.addEventListener("visibilitychange", () => {
        if (document.visibilityState === "visible") refreshFromDocument();
      });
      render();
      applyStartupBehavior(); // retire tout chargement automatique hérité (document ouvrable partout sans erreur)
      if (disabled) return; // extension gelée : aucun automatisme, pas d'auto-démarrage
      registerSelectionHandler();
      // Filet : finalise un nommage en attente même si l'événement Word ne se déclenche pas.
      setInterval(pollPendingFinalize, 1200);
      // Rafraîchit le diagnostic (lecture seule) pour que les alertes suivent l'état réel du doc
      // même sans action de l'utilisateur (ex. citation supprimée à la main).
      setInterval(refreshDiag, 2500);
      // Au démarrage : capturer les éditions inline faites hors ligne, puis synchroniser.
      withBusy(async () => {
        await upgradeDocumentOnce();
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
    let ambiguous = null;    // « Nom + suite » sur une pièce citée ailleurs : précision ou renommage ? (on demande)
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
        } else if (WP.isSingleCiteTag(cc.tag)) {
          const { id: pid, extra, side } = WP.splitPieceTag(cc.tag);
          const piece = WP.findPiece(pid);
          if (piece) {
            // 1) numéro modifié dans la citation ? (IGNORÉ si la pièce est verrouillée → la synchro
            //    réécrira le numéro verrouillé, annulant la modification manuelle dans le texte.)
            //    Pièce adverse : son numéro se corrige librement (c'est celui de l'adversaire).
            const nn = WP.extractNumberFromCitation(cc.text, side);
            const curNum = WP.numberOf(pid);
            if (nn != null && curNum != null && nn !== String(curNum)) {
              if (piece.locked) { forceResync = true; break; } // verrouillée : la synchro rétablira le numéro figé
              changedId = pid; changedNumber = nn; break;
            }
            // 2) sinon, texte modifié : précision LOCALE ou véritable renommage ?
            const nm = WP.extractNameFromCitation(cc.text, pid, side);
            if (nm != null && nm !== (piece.name || "") + extra) {
              // a) « Nom, page 8 » → précision propre à CETTE citation : on l'enregistre dans le
              //    tag du contrôle. Le nom de la pièce (bordereau, autres citations) ne bouge pas.
              const r = WP.splitPrecision(piece.name, nm); // « , page 12 », « page 12 », « § 4 »…
              if (r != null) { changedId = pid; changedExtraCc = cc.id; changedExtra = r; break; }
              // b) précision effacée par l'utilisateur → on retire le complément (sinon la synchro
              //    la remettrait aussitôt).
              if (extra && nm === (piece.name || "")) { changedId = pid; changedExtraCc = cc.id; changedExtra = ""; break; }
              // c) « Nom + suite » sans mot de précision reconnu, alors que la pièce est citée ailleurs :
              //    on DEMANDE (précision pour cette citation, ou nouveau nom partout).
              const base = piece.name || "";
              if (skipActive && base && nm.length > base.length && nm.slice(0, base.length).toLowerCase() === base.toLowerCase()
                && (WP.stats.counts[pid] || 0) > 1) {
                ambiguous = { pid, ccId: cc.id, base, nm, rest: nm.slice(base.length), side };
                break;
              }
              // d) sinon : véritable renommage (on ne réabsorbe pas le complément dans le nom).
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

    if (ambiguous) {
      const a = ambiguous, esc = WP.escapeHtml;
      const num = WP.numberOf(a.pid), n = WP.stats.counts[a.pid] || 0;
      const choice = await showModal({
        title: `${a.side === "adv" ? "Pièce adverse" : "Pièce"} n°${esc(String(num ?? "?"))}`,
        msg: `Vous avez écrit « ${esc(a.base)}<b>${esc(a.rest)}</b> ».<br>Est-ce une <b>précision</b> pour cette citation seulement, ` +
          `ou le <b>nouveau nom</b> de la pièce (citée ${n} fois, et au bordereau) ?`,
        buttons: [
          { label: "Annuler", value: "cancel" },
          { label: "Renommer partout", value: "rename" },
          { label: "Précision ici", value: "prec", primary: true },
        ],
      });
      if (choice === "prec") { changedId = a.pid; changedExtraCc = a.ccId; changedExtra = a.rest; }
      else if (choice === "rename") { changedId = a.pid; changedName = a.nm; }
      else { await WP.sync(); render(); toast("Citation rétablie"); }
    }
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
  const dead = WP.model.pieces.concat(WP.model.adverse || []).filter((p) => !(p.name && p.name.trim()) && !pend.has(p.id)
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
  const p = await WP.addPiece(name, paneSide);
  if (!p) return;
  commitUndo(snap, isAdvMode() ? "ajout d'une pièce adverse" : isLetterDoc() ? "ajout d'une pièce jointe" : "ajout d'une pièce");
  if (isAdvMode()) await runQuiet(() => WP.refreshStats()); // numéros adverses à jour pour l'affichage
  // Courrier : une pièce jointe n'a pas à être citée — elle rejoint tout de suite le récapitulatif.
  else if (isLetterDoc()) await withBusy(() => syncDoc());
  render();
  collapseAdd(); // on referme le champ après ajout (volet épuré)
}
// « + » d'un groupe : ligne de saisie sous sa dernière sous-pièce ; Entrée crée la sous-pièce suivante.
let addSubFor = null; // id du groupe (conteneur ou pièce de tête) en cours d'ajout
function openAddSub(id) {
  addSubFor = id;
  render();
  const inp = document.getElementById("addSubName");
  if (inp) { inp.focus(); if (inp.scrollIntoView) inp.scrollIntoView({ block: "nearest" }); }
}
function closeAddSub() { addSubFor = null; render(); }
async function submitAddSub() {
  const inp = document.getElementById("addSubName");
  const parentId = addSubFor;
  if (!inp || !parentId) return;
  const name = inp.value.trim();
  if (!name) { inp.focus(); return; }
  const snap = snapshotModel();
  const p = await WP.addSubPiece(parentId, name);
  if (!p) { closeAddSub(); return; }
  commitUndo(snap, "ajout d'une sous-pièce");
  addSubFor = null;
  flashIds.add(p.id);
  if (isLetterDoc()) await withBusy(() => syncDoc());
  else await runQuiet(() => WP.refreshStats());
  render();
}
function addSubBtn(p) {
  return `<button class="mini addsub" data-act="addsub" data-id="${p.id}" title="Ajouter une sous-pièce à la fin de ce groupe (n°${WP.escapeHtml(WP.previewSubNumber(p.id))})" aria-label="Ajouter une sous-pièce">${icon("plus", 15)}</button>`;
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

// ---- Créer des pièces à partir du TEXTE SÉLECTIONNÉ dans le document ----
// Une ligne sélectionnée = une pièce (son intitulé). Une liste NUMÉROTÉE (« 1. Contrat », « Pièce n°3 :
// Bail »…) ouvre l'aperçu d'import, qui reprend les numéros et signale les pièces déjà présentes.
// Le texte du document n'est pas modifié.
function cleanSelName(s) {
  return String(s || "").replace(/\s+/g, " ").trim()
    .replace(/^[«"“'‘(\[\s]+|[»"”'’)\]\s]+$/g, "") // guillemets / parenthèses autour
    .replace(/[\s,;:.]+$/, "")
    .slice(0, 200).trim();
}
async function createFromSelection() {
  let text = "";
  try { text = await WP.readSelectionText(); } catch (e) { reportError(e); return; }
  const lines = text.split("\n").map((l) => l.replace(/\s+/g, " ").trim()).filter(Boolean);
  if (!lines.length) { toast("Sélectionnez d'abord, dans le document, le texte qui doit devenir une pièce."); return; }
  const numbered = WP.parseAdverseList(lines.join("\n"));
  // Liste numérotée : la plupart des lignes portent un numéro (une ligne seule : seulement « Pièce n°… »,
  // pour qu'un intitulé comme « 3 mois de loyers impayés » reste un simple nom).
  const isList = lines.length > 1 ? numbered.length >= Math.ceil(lines.length / 2) : numbered.length === 1 && /^\s*pi[eè]ce/i.test(lines[0]);
  if (isList) { openImportWithText(lines.join("\n"), paneSide); return; }
  if (lines.length > 50) { toast("Sélection trop longue : 50 lignes au plus (une ligne = une pièce)."); return; }
  const existing = new Set(WP.sidePieces(paneSide).map((p) => (p.name || "").trim().toLowerCase()));
  const names = [];
  let dup = 0;
  for (const l of lines) {
    const n = cleanSelName(l);
    if (!n) continue;
    if (existing.has(n.toLowerCase())) { dup++; continue; }
    existing.add(n.toLowerCase());
    names.push(n);
  }
  if (!names.length) { toast(dup ? "Cette pièce existe déjà dans la liste." : "Aucun intitulé utilisable dans la sélection."); return; }
  const snap = snapshotModel();
  const ids = [];
  await withBusy(async () => {
    for (const n of names) { const p = await WP.addPiece(n, paneSide); if (p) ids.push(p.id); }
    if (isAdvMode()) await WP.refreshStats();
    else if (isLetterDoc()) await syncDoc();
    else await WP.refreshStats();
  });
  if (!ids.length) return;
  commitUndo(snap, ids.length > 1 ? "création de pièces depuis la sélection" : "création d'une pièce depuis la sélection");
  flashIds = new Set(ids);
  selectedPieceId = ids[ids.length - 1];
  render();
  const s = ids.length > 1 ? "s" : "";
  toast(`${ids.length} pièce${s} créée${s} depuis la sélection` + (dup ? ` · ${dup} déjà présente${dup > 1 ? "s" : ""}` : "") + " — icône d'insertion pour la citer");
}

// ---- Interrupteur « Manuel » : numérotation manuelle de nos pièces (réglage du document) ----
function renderManualSwitch() {
  const sw = el("manualSwitch");
  if (!sw) return;
  const on = WP.isManual();
  sw.setAttribute("aria-checked", String(on));
  sw.classList.toggle("on", on);
  sw.title = on
    ? "Numérotation MANUELLE : chaque pièce garde son numéro (supprimer, déplacer ou déverrouiller ne renumérote rien). Cliquer pour revenir à l'automatique."
    : "Numérotation automatique (1, 2, 3…). Cliquer pour passer en numérotation MANUELLE : chaque pièce garde le numéro que vous lui donnez.";
}
async function toggleManualNumbering() {
  const on = !WP.isManual();
  if (!on) {
    const res = await openPrompt({ mode: "confirm", title: "Numérotation automatique",
      label: "Les pièces non verrouillées seront renumérotées automatiquement (1, 2, 3… ; les séries comme « i.1 » restent des séries). Verrouillez d'abord celles dont le numéro doit rester. Continuer ?" });
    if (!res || !res.confirmed) return;
  }
  const snap = snapshotModel();
  const wasManual = !on;
  await WP.setManualNumbering(on);
  pushUndo(on ? "numérotation manuelle" : "numérotation automatique", async () => {
    WP.model.pieces = snap.pieces;
    WP.model.adverse = snap.adverse || [];
    WP.model.settings.manualNumbering = wasManual;
    await WP.save();
    await syncDoc();
  });
  await runQuiet(() => syncDoc());
  render();
  toast(on ? "Numérotation manuelle : les numéros ne bougeront plus d'eux-mêmes" : "Numérotation automatique rétablie");
}

// ---- Suppression au clavier : clic sur une pièce, puis Suppr (confirmation : Suppr ou Entrée) ----
function setSelectedPiece(id) {
  selectedPieceId = id || null;
  el("pieceList").querySelectorAll(".piece[data-id]").forEach((r) => r.classList.toggle("selected", r.dataset.id === selectedPieceId));
}
// Le champ texte a-t-il le curseur EN FIN de texte, sans sélection ? (Suppr n'y ferait rien.)
function caretAtEnd(inp) {
  try { return inp.selectionStart === inp.selectionEnd && inp.selectionEnd === inp.value.length; } catch (e) { return false; }
}
function onDeleteKey(e) {
  if (e.key === "Escape" && confirmDeleteId) { // annule la suppression en attente
    const id = confirmDeleteId;
    confirmDeleteId = null;
    render();
    const row = el("pieceList").querySelector(`.piece[data-id="${id}"]`);
    if (row) { row.tabIndex = -1; row.focus({ preventScroll: true }); }
    return;
  }
  if (e.key !== "Delete" && e.key !== "Del") return;
  if (e.ctrlKey || e.altKey || e.metaKey || e.shiftKey) return;
  const a = document.activeElement;
  const inList = a && a.closest && a.closest("#pieceList");
  const typing = a && (a.tagName === "INPUT" || a.tagName === "TEXTAREA" || a.isContentEditable);
  if (typing && !(inList && caretAtEnd(a))) return; // on laisse Suppr effacer du texte
  if (modalResolve || warnModalOpen) return; // une pop-up est ouverte
  const row = inList ? a.closest(".piece[data-id]") : null;
  const id = (row && row.dataset.id) || selectedPieceId;
  const p = id && WP.findPiece(id);
  if (!p || p.container) return; // un groupe se supprime avec ses sous-pièces : pas au clavier
  e.preventDefault();
  if (confirmDeleteId === id) { deletePiece(id); return; } // 2e Suppr = confirmation
  selectedPieceId = id;
  confirmDeleteId = id;
  render();
  const ok = el("pieceList").querySelector(`.piece[data-id="${id}"] [data-act="confirmdelete"]`);
  if (ok) ok.focus(); // Entrée valide, Échap annule
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
  if (selectedPieceId === id) selectedPieceId = null;
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
  const cur = WP.numberOf(id);
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
  if (res.error === "advEmpty") return "Une pièce adverse garde toujours un numéro (celui de l'adversaire).";
  if (res.error === "manualEmpty") return "Numérotation manuelle : chaque pièce garde un numéro — tapez-en un (ou repassez en automatique).";
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
// steps : crans parmi les sœurs déplaçables (±1 au clavier, davantage au glisser-déposer).
// opts.refocusGrip : rend le focus à la poignée de la pièce après le re-rendu (clavier).
async function reorderPiece(id, steps, opts) {
  const snap = snapshotModel(); // point d'annulation (avant le déplacement)
  const before = new Map(WP.stats.numbers || new Map()); // numéros AVANT le déplacement
  const oldTops = captureRowTops(); // positions AVANT re-rendu
  const moved = await WP.movePieceBy(id, steps);
  if (!moved) return; // bord de liste ou déplacement bloqué par un numéro verrouillé → aucun effet
  commitUndo(snap, "changement d'ordre");
  // Le moteur a déjà recalculé les numéros en mémoire → le volet se met à jour TOUT DE SUITE.
  const after = WP.stats.numbers || new Map();
  flashIds = new Set();
  for (const [pid, n] of after) if (String(before.get(pid)) !== String(n)) flashIds.add(pid);
  render();
  flipRows(oldTops); // fait GLISSER les lignes vers leur nouvelle place
  if (opts && opts.refocusGrip) {
    const g = el("pieceList").querySelector(`.grip[data-id="${id}"]`);
    if (g) g.focus({ preventScroll: true });
  }
  runQuiet(() => syncDoc()); // met à jour les citations du texte en arrière-plan (non bloquant)
}

// Glisser-déposer libre : la pièce entre dans un groupe (en fin, ou avant une sous-pièce) / en sort.
async function moveAcrossGroups(label, fn, id) {
  const snap = snapshotModel();
  const before = new Map(WP.stats.numbers || new Map());
  const oldTops = captureRowTops();
  const num = await fn();
  if (num == null) { toast("Déplacement impossible ici"); return; }
  commitUndo(snap, label);
  const after = WP.stats.numbers || new Map();
  flashIds = new Set([id]);
  for (const [pid, n] of after) if (String(before.get(pid)) !== String(n)) flashIds.add(pid);
  render();
  flipRows(oldTops);
  toast(`Pièce n°${num} (flèche ↶ pour annuler)`);
  runQuiet(() => syncDoc());
}
function movePieceToGroupUI(id, gid, beforeId) {
  return moveAcrossGroups("ajout au groupe", () => WP.movePieceToGroup(id, gid, beforeId), id).catch(reportError);
}
function movePieceToTopUI(id, beforeId) {
  return moveAcrossGroups("sortie du groupe", () => WP.movePieceToTop(id, beforeId), id).catch(reportError);
}

// Pièce communiquée à l'adversaire, ou nouvelle (bouton enveloppe de la ligne).
async function toggleCommunicated(id) {
  const p = WP.findPiece(id);
  if (!p) return;
  commitUndo(snapshotModel(), p.communicated ? "pièce marquée nouvelle" : "pièce marquée communiquée");
  await WP.setCommunicated([id], !p.communicated);
  render();
  runQuiet(() => syncDoc()); // la mention « pièce nouvelle » du bordereau suit
}
// Menu ☰ : toutes les pièces pas encore communiquées le deviennent (date du jour).
async function markAllCommunicated() {
  const todo = WP.model.pieces.filter((p) => !p.container && !p.communicated).length;
  if (!todo) { toast("Toutes les pièces sont déjà marquées communiquées"); return; }
  const snap = snapshotModel();
  const n = await WP.setCommunicated([], true);
  commitUndo(snap, "pièces marquées communiquées");
  render();
  runQuiet(() => syncDoc());
  toast(`${n} pièce${n > 1 ? "s" : ""} marquée${n > 1 ? "s" : ""} communiquée${n > 1 ? "s" : ""} (flèche ↶ pour annuler)`);
}
// Menu ☰ : copie du document sans contrôles ni données LexRef, dans une nouvelle fenêtre.
async function openCleanCopy() {
  let ok = false;
  await withBusy(async () => { await WP.makeCleanCopy(); ok = true; });
  if (ok) toast("Copie propre ouverte dans une nouvelle fenêtre — enregistrez-la sous le nom de votre choix");
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
// ---- Bouton « Verrous » : sélection de GROUPES (Scannées, Créées, Toutes), puis « Valider » ----
// Au clic sur « Verrous » : le reste du volet se floute, les groupes apparaissent au-dessus du
// bouton, qui devient « Valider » (coche). Un clic sur un groupe INVERSE son état VOULU (verrouillé
// ↔ déverrouillé ; un groupe en partie verrouillé part de « déverrouillé ») ; « Toutes » inverse les
// deux à la fois. Rien n'est écrit tant qu'on n'a pas validé ; clic sur le flou ou Échap = annuler.
const LOCK_GROUPS = [
  { scope: "manual", label: "Pièces créées", test: (p) => !p.scanned },
  { scope: "scanned", label: "Pièces existantes", test: (p) => !!p.scanned },
  { scope: "all", label: "Toutes les pièces", test: () => true },
];
function lockGroupPieces(scope) {
  const g = LOCK_GROUPS.find((x) => x.scope === scope);
  return WP.model.pieces.filter((p) => !p.container && g.test(p));
}
// "locked" (toutes verrouillées), "unlocked" (aucune), "mixed" (certaines), "empty" (aucune pièce)
function lockGroupState(scope) {
  const ps = lockGroupPieces(scope);
  if (!ps.length) return "empty";
  const n = ps.filter((p) => p.locked).length;
  return n === ps.length ? "locked" : n === 0 ? "unlocked" : "mixed";
}
// Sélection en cours (null = menu fermé) : état VOULU de chaque groupe réel.
// initial = état affiché à l'ouverture (true = verrouillé) ; want = choix de l'utilisateur.
let lockSel = null;
function lockMenuOpen() { return !!lockSel; }
function groupWantLocked(scope) {
  if (scope !== "all") return lockSel[scope] ? lockSel[scope].want : null;
  const real = ["scanned", "manual"].filter((s) => lockSel[s]);
  return real.length ? real.every((s) => lockSel[s].want) : null;
}
// Les groupes s'empilent AU-DESSUS du bouton, même largeur et même style que lui : cadenas +
// libellé, GRIS quand le groupe sera verrouillé, VERT quand il sera déverrouillé.
function renderLockMenu() {
  const menu = el("lockMenu");
  if (!menu || !lockSel) return;
  const btn = el("lockAllBtn");
  // Au moins la largeur du bouton ; plus large si un libellé l'exige (toutes les pilules alignées).
  if (btn) { menu.style.width = ""; menu.style.minWidth = btn.offsetWidth + "px"; }
  menu.innerHTML = LOCK_GROUPS.map((g) => {
    const want = groupWantLocked(g.scope);
    const n = lockGroupPieces(g.scope).length;
    const empty = want === null;
    const tip = empty ? `${g.label} : aucune pièce`
      : `${g.label} (${n}) : ${want ? "verrouillées" : "déverrouillées"} après validation — cliquer pour inverser`;
    return `<button class="drop-item ${want ? "is-locked" : "is-unlocked"}" role="menuitemcheckbox"` +
      ` aria-checked="${!!want}" data-scope="${g.scope}" title="${tip}" ${empty ? "disabled" : ""}>` +
      `${icon(want ? "lock" : "unlock", 15, 2.7)} ${g.label}</button>`;
  }).join("");
}
// Clic sur « Verrous » (menu fermé) : ouvre la sélection ; clic sur « Valider » : applique.
function toggleAllLocks() {
  if (lockMenuOpen()) { validateLockSelection().catch(reportError); return; }
  if (!WP.model.pieces.some((p) => !p.container)) return;
  lockSel = {};
  for (const s of ["scanned", "manual"]) {
    const st = lockGroupState(s);
    if (st !== "empty") lockSel[s] = { initial: st === "locked", want: st === "locked" };
  }
  renderLockMenu();
  el("lockMenu").classList.remove("hidden");
  el("lockVeil").classList.remove("hidden");
  const btn = el("lockAllBtn");
  btn.setAttribute("aria-expanded", "true");
  btn.classList.add("is-validate");
  renderLockAll();
  const first = el("lockMenu").querySelector(".drop-item:not([disabled])");
  if (first) first.focus();
}
// Ferme sans rien appliquer (clic sur le flou, Échap, défilement).
function closeLockMenu() {
  if (!lockMenuOpen()) return;
  lockSel = null;
  el("lockMenu").classList.add("hidden");
  el("lockVeil").classList.add("hidden");
  const btn = el("lockAllBtn");
  btn.setAttribute("aria-expanded", "false");
  btn.classList.remove("is-validate");
  renderLockAll();
}
// Clic sur un groupe : inverse son état VOULU (« Toutes » inverse les deux groupes).
function toggleLockGroup(scope) {
  if (!lockMenuOpen()) return;
  const want = groupWantLocked(scope);
  if (want === null) return;
  const targets = scope === "all" ? ["scanned", "manual"] : [scope];
  for (const s of targets) if (lockSel[s]) lockSel[s].want = !want;
  renderLockMenu();
  const again = el("lockMenu").querySelector(`[data-scope="${scope}"]`);
  if (again) again.focus();
}
// « Valider » : applique les groupes dont l'état voulu diffère de l'état de départ.
async function validateLockSelection() {
  const sel = lockSel;
  closeLockMenu();
  const changes = Object.entries(sel || {}).filter(([, v]) => v.want !== v.initial);
  if (!changes.length) return;
  let locked = 0, unlocked = 0;
  for (const [scope, v] of changes) {
    const n = lockGroupPieces(scope).length;
    await WP.setAllLocks(v.want, scope);
    if (v.want) locked += n; else unlocked += n;
  }
  await withBusy(() => syncDoc());
  render();
  const part = (n, w) => `${n} pièce${n > 1 ? "s" : ""} ${w}${n > 1 ? "s" : ""}`;
  toast([locked ? part(locked, "verrouillée") : "", unlocked ? part(unlocked, "déverrouillée") : ""].filter(Boolean).join(" · "));
}
// Toujours visible (dès qu'il y a des pièces : la barre d'outils est masquée sans pièce).
function renderLockAll() {
  const btn = el("lockAllBtn");
  if (!btn) return;
  btn.classList.remove("hidden");
  btn.innerHTML = lockMenuOpen() ? icon("check", 13) + " Valider" : icon("lock", 13) + " Verrous";
}
async function insertCitation(id) {
  const piece = WP.findPiece(id);
  if (!piece) return;
  await withBusy(async () => {
    await WP.insertCitation(id);
    await syncDoc();
    toast(`« ${piece.name || sideLabel(id).toLowerCase()} » insérée`);
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
    toast(BordWord() + (existed ? " mis à jour" : " généré"));
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
      '<span class="confirm-q">Supprimer le ' + bordWord() + ' ?</span>' +
      '<button class="mini sq red" data-bact="canceldel" title="Annuler">' + BACK_SVG + '</button>' +
      '<button class="mini sq green" data-bact="confirmdel" title="Valider la suppression">' + CHECK_SVG + '</button>' +
      '</div>';
    return;
  }
  if (!has) {
    box.innerHTML = '<button class="btn outline bordereau-btn" data-bact="generate">' + icon("list", 14) + ' Générer le ' + bordWord() + '</button>';
  } else {
    box.innerHTML =
      '<div class="bordereau-row">' +
      '<button class="btn outline bordereau-btn" data-bact="generate">' + icon("list", 14) + ' Mettre à jour</button>' +
      '<button class="mini trash" data-bact="delete" title="Supprimer le ' + bordWord() + '">' + TRASH_SVG + '</button>' +
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
    toast(BordWord() + " supprimé");
  }
}

// Détecte les pièces déjà présentes (« Pièce n° X : Nom ») et les rend gérées par l'extension.
function scanFormatValue() {
  const fmtInput = el("scanFormat");
  return fmtInput && !el("scanFmtRow").classList.contains("hidden") ? fmtInput.value : "";
}

async function scanExisting() {
  const fmt = scanFormatValue();

  // ÉTAPE 1 — repérage seul, AVANT toute modification. Si l'acte se termine par un bloc de
  // pièces, Word l'amène à l'écran et le surligne : l'utilisateur répond en le voyant.
  let bordInfo = null, bordChoice = null;
  await withBusy(async () => { bordInfo = await WP.detectBordereau(fmt, { reveal: true }); });
  if (bordInfo && bordInfo.found) {
    bordChoice = await askAboutBordereau(bordInfo);
  }

  // ÉTAPE 2 — le scan, avec la réponse. « cite » : le bloc n'était pas un bordereau, ses
  // lignes sont reliées comme n'importe quelle citation (comportement d'avant).
  const snap = snapshotModel(); // état AVANT le scan (pour l'annulation)
  let r = { wrapped: 0, newPieces: 0 };
  await withBusy(async () => {
    r = await WP.scanExistingPieces(fmt, { bordereau: bordChoice === "cite" ? "cite" : "skip" });
    if (r.wrapped || r.newPieces) await syncDoc();
  });
  if (r.wrapped || r.newPieces) {
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
  // ÉTAPE 3 — reprise du bordereau, si c'est ce qui a été demandé.
  if (bordChoice === "adopt") await adoptBordereau();
  if (!r.wrapped && !r.newPieces) {
    toast("Aucune pièce « Pièce n° X : Nom » détectée");
  } else {
    const parts = [];
    if (r.newPieces) parts.push(`${r.newPieces} pièce${r.newPieces > 1 ? "s" : ""} ajoutée${r.newPieces > 1 ? "s" : ""}`);
    if (r.wrapped) parts.push(`${r.wrapped} occurrence${r.wrapped > 1 ? "s" : ""} reliée${r.wrapped > 1 ? "s" : ""}`);
    if (r.bordListed) parts.push(`${r.bordListed} ligne${r.bordListed > 1 ? "s" : ""} de bordereau (non comptée${r.bordListed > 1 ? "s" : ""} comme citation${r.bordListed > 1 ? "s" : ""})`);
    if (r.extras) parts.push(`${r.extras} précision${r.extras > 1 ? "s" : ""} de citation conservée${r.extras > 1 ? "s" : ""}`);
    if (r.harmonized) parts.push(`${r.harmonized} variante${r.harmonized > 1 ? "s" : ""} de nom rattachée${r.harmonized > 1 ? "s" : ""} à la même pièce`);
    if (r.conflicts) parts.push(`⚠ ${r.conflicts} conflit${r.conflicts > 1 ? "s" : ""} de numéro à arbitrer (voir « À vérifier »)`);
    // Ctrl+Z (Word) n'annule pas un scan : seule la flèche d'annulation du volet le défait proprement.
    toast(parts.join(" · ") + " (flèche ↶ du volet pour annuler)");
  }
}

// Bordereau tapé à la main repéré en fin d'acte par le scan. On ne le reprend JAMAIS d'office :
// c'est un bloc que l'utilisateur a mis en forme, et la reprise le réécrit selon les réglages.
// Refusé, il reste intact — simplement, ses lignes ne sont plus prises pour des citations.
// Le bloc est SOUS LES YEUX de l'utilisateur (sélectionné par detectBordereau) : la question
// se pose donc dans la pop-up INTÉGRÉE AU VOLET, qui ne couvre que le volet. Une fenêtre Office
// se serait posée par-dessus le document — c'est-à-dire par-dessus ce qu'on lui montre.
async function askAboutBordereau(info) {
  const n = info.rows;
  const cite = (t) => `<i>${WP.escapeHtml(t.length > 52 ? t.slice(0, 52) + "…" : t)}</i>`;
  const titre = info.hasTitle ? `, sous le titre ${cite(info.titleText)}` : "";
  return showModal({
    title: "Est-ce votre bordereau ?",
    msg: `Votre acte se termine par <b>${n} ligne${n > 1 ? "s" : ""} de pièces</b> qui se suivent${titre}.<br>` +
      `De ${cite(info.firstText)} à ${cite(info.lastText)}.<br>` +
      `Le bloc est surligné dans le document.<br><br>` +
      `<b>Reprendre</b> : il devient un bordereau dynamique — il suivra vos numéros, vos renommages ` +
      `et vos réglages de présentation. Sa mise en forme actuelle sera remplacée par celle des réglages.<br>` +
      `<b>Laisser tel quel</b> : LexRef n'y touche pas, mais « Générer le bordereau » en ajouterait un second.<br>` +
      `<b>Ce sont des citations</b> : ce n'est pas un bordereau — ces lignes sont reliées comme les autres citations.`,
    buttons: [
      { label: "Ce sont des citations", value: "cite" },
      { label: "Laisser tel quel", value: "keep" },
      { label: "Reprendre", value: "adopt", primary: true },
    ],
  });
}

// Reprise proprement dite. Appelée seulement si l'utilisateur a répondu « Reprendre ».
async function adoptBordereau() {
  let out = null;
  await withBusy(async () => {
    out = await WP.adoptExistingBordereau(scanFormatValue());
    if (out && out.ok) await syncDoc(true); // force : le bloc repris doit être réécrit
  });
  if (!out || !out.ok) { toast("Le bordereau n'a pas pu être repris — il reste inchangé"); return; }

  // Annulation : on remet le bloc d'ORIGINE (son OOXML, mise en forme comprise) et on retire
  // le contrôle. Ctrl+Z de Word ne sait pas défaire cela.
  const ooxml = out.ooxml;
  pushUndo("reprise du bordereau existant", async () => {
    await WP.restoreExistingBordereau(ooxml);
    await WP.refreshStats();
  });
  render();
  toast(`Bordereau repris (${out.rows} pièces) — il se met à jour tout seul (flèche ↶ pour annuler)`);
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

// MISE À NIVEAU UNIQUE d'un document ouvert pour la première fois avec cette version :
// retire les contrôles « wp:… » d'anciennes versions (texte conservé), enregistre le modèle des
// pièces dans son format actuel et note le format dans le document (plus jamais refait ensuite).
const DOC_FORMAT = 8;
async function upgradeDocumentOnce() {
  if (disabled || WP.docFormat() >= DOC_FORMAT) return;
  await WP.unwrapLegacyControls();
  await WP.markDocFormat(DOC_FORMAT);
}

// RÈGLE DEPUIS LA V8 : WordPiece N'INSCRIT JAMAIS de chargement automatique dans un document.
// Un document qui demande le chargement d'un complément absent affiche « Erreur relative au
// complément » sur tout poste qui ne l'a pas (destinataire, confrère, autre version…). Sans cette
// inscription, le document s'ouvre partout sans message ; chez l'utilisateur, WordPiece démarre
// au clic sur « Pièces ». Si une ancienne version avait armé le chargement automatique dans CE
// document, on le retire (écriture unique ; rien n'est écrit si le document est déjà propre).
async function applyStartupBehavior() {
  try {
    if (Office.addin && Office.addin.setStartupBehavior && Office.addin.getStartupBehavior) {
      const cur = await Office.addin.getStartupBehavior();
      if (cur === Office.StartupBehavior.load) {
        await Office.addin.setStartupBehavior(Office.StartupBehavior.none);
      }
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
    toast("LexRef désactivé pour ce document");
  } else {
    registerSelectionHandler();
    await withBusy(async () => {
      await reconcileNames(false);
      await syncDoc();
    });
    render();
    toast("LexRef réactivé pour ce document");
  }
}

// OUVERTURE DES FENÊTRES (invites, palette /p, réglages, aide) — deux chemins :
//   - Windows / Mac : vraie FENÊTRE Office (displayDialogAsync), centrée sur Word.
//   - Word pour le web : les fenêtres Office y sont des popups du navigateur, refusées dès que celui-ci
//     bloque le stockage tiers (« paramètres de sécurité… »). La page s'affiche alors DANS LE VOLET,
//     en plein volet (#paneModal), et dialogue avec lui par postMessage (voir embed.js).
function isWordOnline() {
  try { return Office.context.platform === Office.PlatformType.OfficeOnline; } catch (e) { return false; }
}

// onMessage(texteJSON, api) ; api = { messageChild(texte), close() }. onClosed() n'est appelé QUE si la
// fenêtre n'a pas pu s'ouvrir ou si l'utilisateur l'a fermée lui-même — jamais par api.close(), sans quoi
// l'appelant recevrait « annulé » avant son résultat.
let paneModalApi = null;
function openWindow(url, dim, onMessage, onClosed) {
  if (isWordOnline()) {
    if (paneModalApi) paneModalApi.close(); // une seule à la fois
    const wrap = el("paneModal");
    const frame = el("paneModalFrame");
    let closed = false;
    const onMsg = (e) => {
      if (e.source !== frame.contentWindow || e.origin !== location.origin) return;
      onMessage(typeof e.data === "string" ? e.data : JSON.stringify(e.data), api);
    };
    const api = {
      messageChild(text) { try { frame.contentWindow.postMessage(text, location.origin); } catch (e) {} },
      close() {
        if (closed) return;
        closed = true;
        window.removeEventListener("message", onMsg);
        frame.src = "about:blank";
        wrap.classList.add("hidden");
        if (paneModalApi === api) paneModalApi = null;
      },
    };
    paneModalApi = api;
    window.addEventListener("message", onMsg);
    frame.src = url + (url.indexOf("?") >= 0 ? "&" : "?") + "embed=1";
    wrap.classList.remove("hidden");
    return;
  }
  Office.context.ui.displayDialogAsync(url, dim, (res) => {
    if (res.status !== Office.AsyncResultStatus.Succeeded) { if (onClosed) onClosed(true); return; }
    const dialog = res.value;
    const api = {
      messageChild(text) { try { dialog.messageChild(text); } catch (e) {} },
      close() { try { dialog.close(); } catch (e) {} },
    };
    dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => onMessage(arg.message, api));
    dialog.addEventHandler(Office.EventType.DialogEventReceived, () => { if (onClosed) onClosed(true); });
  });
}

function openPrompt(params, size) {
  return new Promise((resolve) => {
    // Relatif à la page du volet (et non à la racine du domaine) : fonctionne en local
    // (https://localhost:3002/) comme sur GitHub Pages (https://…github.io/wordpiece/).
    const url = new URL("dialog.html?" + new URLSearchParams(params).toString(), location.href).href;
    const dim = { height: (size && size.height) || 40, width: (size && size.width) || 34, promptBeforeOpen: false };
    openWindow(url, dim,
      (msg, api) => {
        let data = null; try { data = JSON.parse(msg); } catch (e) {}
        api.close();
        resolve(data);
      },
      () => resolve(null));
  });
}

// Palette de recherche « citer par nom » : envoie la liste des pièces (après le handshake « ready »),
// renvoie l'id choisi (ou null si annulé/fermé).
function openPiecePalette(query, side) {
  return new Promise((resolve) => {
    const adv = side === "adv";
    const pieces = WP.sidePieces(adv ? "adv" : "own").filter((p) => !p.container).map((p) => ({
      id: p.id, name: p.name, num: String(WP.numberOf(p.id) ?? ""), locked: !adv && !!p.locked, // adverse : pas de notion de verrou
    })).sort((a, b) => WP.naturalCompare(a.num, b.num));
    const url = new URL("search-piece.html?q=" + encodeURIComponent(query || "") + (adv ? "&side=adv" : ""), location.href).href;
    // Largeur ≈ 600 px : deux colonnes où les noms courts tiennent en entier (les longs passent à
    // la ligne). displayDialogAsync ne prend qu'un % de l'écran → conversion, bornée pour les
    // petits écrans (45 %) comme pour les très grands (26 %).
    const scrW = (window.screen && window.screen.width) || 1920;
    const widthPct = Math.max(26, Math.min(45, Math.round((600 / scrW) * 100)));
    openWindow(url, { height: 90, width: widthPct, promptBeforeOpen: false },
      (msg, api) => {
        let data = null; try { data = JSON.parse(msg); } catch (e) {}
        if (data && data.ready) { api.messageChild(JSON.stringify(pieces)); return; }
        api.close();
        resolve(data && data.id ? data.id : null);
      },
      () => resolve(null));
  });
}

// Réglages. Le volet reste la source de vérité : il envoie l'état + le schéma, reçoit chaque
// changement { path, value }, applique.
function openSettingsDialog() {
  const url = new URL("settings-dialog.html", location.href).href;
  openWindow(url, { height: 78, width: 56, promptBeforeOpen: false }, async (msg, api) => {
    let data = null;
    try { data = JSON.parse(msg); } catch (e) { data = null; }
    if (!data) return;
    if (data.ready) { // la fenêtre est prête → on lui pousse les réglages + le schéma de liaison
      api.messageChild(JSON.stringify({ settings: WP.model.settings, bindings: SETTING_BINDINGS }));
      return;
    }
    if (data.close) { api.close(); return; }
    if (data.path) { // un réglage a changé → on applique (enregistre + resynchronise le document)
      try { await changeSetting(data.path, data.value, isUiOnly(data.path)); } catch (e) { reportError(e); }
    }
  });
}

// Renuméroter les pièces (liste affichée : nos pièces ou pièces adverses). La fenêtre travaille en
// APERÇU et renvoie un plan complet ; le volet l'applique en une fois (annulable par ↶).
function openRenumberDialog() {
  const side = paneSide;
  const url = new URL("renumber-dialog.html", location.href).href;
  openWindow(url, { height: 85, width: 58, promptBeforeOpen: false }, async (msg, api) => {
    let data = null;
    try { data = JSON.parse(msg); } catch (e) { data = null; }
    if (!data) return;
    if (data.ready) { api.messageChild(JSON.stringify({ side, rows: WP.renumberRows(side) })); return; }
    if (data.cancelled) { api.close(); return; }
    if (data.apply) {
      const snap = snapshotModel();
      let res = null;
      try {
        await withBusy(async () => {
          res = await WP.applyRenumber(side, data.apply);
          if (res && !res.error && res.changed) await syncDoc();
        });
      } catch (e) { reportError(e); }
      if (!res) return;
      if (res.error) {
        const why = res.error === "dup" ? `le n°${res.num} est utilisé deux fois`
          : res.error === "subFormat" ? `« ${res.num} » désigne une sous-pièce (utilisez « Grouper en sous-pièces »)`
          : res.error === "hasSubs" ? `la pièce n°${res.num} a déjà des sous-pièces et ne peut pas devenir une sous-pièce`
          : res.error === "badGroup" ? "le groupe visé n'existe plus (rouvrez la fenêtre)" : "un numéro est vide";
        api.messageChild(JSON.stringify({ error: "Non appliqué : " + why + "." }));
        return;
      }
      // La fenêtre RESTE OUVERTE : elle reçoit la liste à jour pour enchaîner d'autres opérations.
      if (res.changed) {
        commitUndo(snap, "renumérotation");
        if (side !== paneSide) setPaneSide(side);
        render();
        toast(`${res.changed} modification${res.changed > 1 ? "s" : ""} appliquée${res.changed > 1 ? "s" : ""} (flèche ↶ pour annuler)`);
      }
      try { api.messageChild(JSON.stringify({ side, rows: WP.renumberRows(side), applied: res.changed || 0 })); } catch (e) { reportError(e); }
    }
  });
}

// Mode d'emploi (contenu statique ; la page demande seulement sa fermeture).
function openHelpDialog() {
  const url = new URL("help-dialog.html", location.href).href;
  openWindow(url, { height: 80, width: 60, promptBeforeOpen: false }, (msg, api) => {
    let data = null;
    try { data = JSON.parse(msg); } catch (e) { data = null; }
    if (data && data.close) api.close();
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
      const cur = WP.numberOf(id);
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
    // Information curseur décochée dans les réglages : on n'interroge même pas Word à
    // chaque déplacement du curseur (une lecture de moins par frappe).
    if (!contextBarEnabled()) { hideContextBar(); return; }
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
    // Commandes décochées dans les réglages : on ne LIT pas le document pour en chercher
    // (une lecture de Word en moins à chaque déplacement du curseur) et « /p… » reste du
    // texte ordinaire — mais on continue plus bas, pour laisser se terminer un nommage
    // déjà engagé, sinon une pièce resterait en attente de nom.
    const d = slashCommandsEnabled() ? await WP.detectSlashAny() : null;
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
    const side = (d && d.side) || "own";
    const label = side === "adv" ? "Pièce adverse" : "Pièce";
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
      const res = await WP.startSlashRange(d.start, d.end, side);
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
      const res = await WP.startSlashEtSeq(d.start, side);
      if (res && res.error) toast(`Numéro introuvable : ${res.error}`);
      await WP.sync();
      render();
      return;
    }
    // « /p1,2,5 » → « Pièces n°1, 2 et 5 » (citation multiple ; pièces existantes uniquement).
    if (d && d.kind === "list") {
      if (pendingChain.length) await abandonPending();
      pendingChain = []; pendingChainAlign = null; discardPendingFormat();
      const res = await WP.startSlashList(d.list, side);
      if (res && res.error) toast(`${label} n°${res.error} introuvable — aucune pièce créée.`);
      await WP.sync();
      render();
      return;
    }
    // « /p » (ou « /pbail ») → citer une pièce par son NOM. Si la recherche ne peut donner qu'UN
    // seul résultat, la pièce est citée immédiatement (aucune fenêtre) ; sinon on ouvre la palette.
    if (d && d.kind === "search") {
      if (pendingChain.length) await abandonPending();
      pendingChain = []; pendingChainAlign = null; discardPendingFormat();
      const token = (side === "adv" ? "/pa" : "/p") + (d.query || "");
      const query = (d.query || "").trim();
      let id = null;
      if (query) {
        const norm = (s) => String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
        const f = norm(query);
        const hits = WP.sidePieces(side).filter((p) => !p.container && norm(p.name).includes(f));
        if (hits.length === 1) id = hits[0].id;          // un seul résultat → citation directe
      }
      if (id === null) id = await openPiecePalette(d.query, side); // 0 ou plusieurs → palette
      if (id) await WP.citeAtToken(token, id);
      else await WP.stripToken(token); // annulé / fermé → on retire le « /p… » tapé
      await WP.sync();
      render();
      // Rend le clavier au document (la fermeture de la fenêtre l'a donné au volet — voir focusSelection).
      const focusOk = await WP.focusSelection();
      if (!focusOk && !focusWarned) {
        focusWarned = true;
        toast("Cette version de Word ne permet pas de rendre le clavier au texte (API WordApiDesktop 1.4 absente) : cliquez dans le document.");
      }
      return;
    }
    // « /pn », « /pnN », « /pn4. », « /pn4.2 » : création / insertion structurée + nommage.
    if (d && (d.kind === "new" || d.kind === "insert" || d.kind === "subnew" || d.kind === "subinsert")) {
      if (pendingChain.length) await abandonPending();
      discardPendingFormat();
      let res;
      let insertInfo = null;
      if (d.kind === "new") res = await WP.startSlashNew(null, side);
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
      const res = await WP.startSlashPrompt(d.num, side, d.token);
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
        toast(`${sideLabel(cur.id)} n°${cur.num}${nm ? " : " + nm : ""}`);
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
      toast(`${sideLabel(cur.id)} n°${cur.num}${nm ? " : " + nm : ""}`);
    } else if (!next) {
      // NETTOYEUR anti-fantôme : le nommage n'a pas pu se faire. Si l'invite « Pièce n°X : » n'est
      // PLUS dans le document (annulée par Ctrl+Z sans nouvelle commande) et que la pièce est
      // toujours sans nom et non citée, c'est une pièce fantôme → on la retire du volet.
      const piece = WP.findPiece(cur.id);
      const nameless = !(piece && piece.name && piece.name.trim());
      const uncited = !(WP.stats.counts[cur.id] > 0);
      if (piece && nameless && uncited && !(await WP.namingPromptPresent(cur.num, WP.sideOf(cur.id)))) {
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
// Préférences d'affichage du poste. Vraies par défaut, y compris sur un document
// enregistré avant leur arrivée (mergeSettings pose les défauts).
function uiPrefs() {
  return (WP.model && WP.model.settings && WP.model.settings.ui) || {};
}
function contextBarEnabled() { return uiPrefs().contextBar !== false; }
function slashCommandsEnabled() { return uiPrefs().slashCommands !== false; }

// Applique les préférences au DOM. theme.js garde sa propre copie du thème et de la
// densité — il doit agir avant tout script du volet pour éviter le flash de couleur —
// mais c'est le modèle qui fait foi : on la lui pousse ici.
function applyUiPrefs() {
  const ui = uiPrefs();
  if (window.JT) {
    JT.setTheme(ui.theme || "auto");
    JT.setDensity(ui.density || "standard");
  }
  // Onglet adverse masqué alors qu'on s'y trouvait : on revient sur nos pièces.
  if (ui.adverseTab === false && paneSide === "adv" && !WP.sidePieces("adv").length) setPaneSide("own");
}
function showContextBar(id) {
  if (!contextBarEnabled()) return hideContextBar();
  const piece = WP.findPiece(id);
  ctxPieceId = id;
  const num = piece ? WP.numberOf(id) : undefined;
  el("ctxLabel").textContent = piece
    ? `Curseur sur ${num ? sideLabel(id) + " n°" + num : "une pièce"}${piece.name ? " : " + piece.name : ""}`
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
  { id: "numStart", path: ["startNumber"], type: "number", fallback: 1 },
  // Propres au document (menu ☰ du volet) : « Numéroter à partir de » ci-dessus, pièces jointes numérotées ou non.
  // Le type acte / courrier passe par l'interrupteur du menu (onAppMenuClick).
  { id: "letNumbered", path: ["letterNumbered"], type: "checked" },
  { id: "bordNoTitle", path: ["bordereauNoTitle"], type: "checked" },
  // Présentation des pièces jointes d'un courrier (fenêtre Réglages).
  { id: "letTemplate", path: ["lettre", "template"], type: "value" },
  { id: "letTemplatePlain", path: ["lettre", "templatePlain"], type: "value" },
  { id: "letBold", path: ["lettre", "bold"], type: "checked" },
  { id: "letItalic", path: ["lettre", "italic"], type: "checked" },
  { id: "letUnderline", path: ["lettre", "underline"], type: "checked" },
  { id: "letAlign", path: ["lettre", "alignment"], type: "value" },
  { id: "letNewLine", path: ["lettre", "newLine"], type: "checked" },
  { id: "recPlacement", path: ["lettre", "recap", "placement"], type: "value" },
  { id: "recTitle", path: ["lettre", "recap", "title"], type: "value" },
  { id: "recTitleBold", path: ["lettre", "recap", "titleBold"], type: "checked" },
  { id: "recTitleItalic", path: ["lettre", "recap", "titleItalic"], type: "checked" },
  { id: "recTitleUnderline", path: ["lettre", "recap", "titleUnderline"], type: "checked" },
  { id: "recTitleSize", path: ["lettre", "recap", "titleSize"], type: "number" },
  { id: "recTitleAlign", path: ["lettre", "recap", "titleAlign"], type: "value" },
  { id: "recLabelTpl", path: ["lettre", "recap", "labelTemplate"], type: "value" },
  { id: "recBullet", path: ["lettre", "recap", "bullet"], type: "value" },
  { id: "recLabelBold", path: ["lettre", "recap", "labelBold"], type: "checked" },
  { id: "recLabelItalic", path: ["lettre", "recap", "labelItalic"], type: "checked" },
  { id: "recLabelUnderline", path: ["lettre", "recap", "labelUnderline"], type: "checked" },
  { id: "recNameBold", path: ["lettre", "recap", "nameBold"], type: "checked" },
  { id: "recNameItalic", path: ["lettre", "recap", "nameItalic"], type: "checked" },
  { id: "recNameUnderline", path: ["lettre", "recap", "nameUnderline"], type: "checked" },
  { id: "recListSize", path: ["lettre", "recap", "listSize"], type: "number" },
  { id: "recListAlign", path: ["lettre", "recap", "listAlign"], type: "value" },
  { id: "recLineSpacing", path: ["lettre", "recap", "lineSpacing"], type: "value" },
  { id: "recSpaceBefore", path: ["lettre", "recap", "spaceBefore"], type: "checked" },
  { id: "recSpaceAfter", path: ["lettre", "recap", "spaceAfter"], type: "checked" },
  { id: "advSame", path: ["adverse", "sameAsOwn"], type: "checked" },
  { id: "advTemplate", path: ["adverse", "template"], type: "value" },
  { id: "advBold", path: ["adverse", "bold"], type: "checked" },
  { id: "advItalic", path: ["adverse", "italic"], type: "checked" },
  { id: "advUnderline", path: ["adverse", "underline"], type: "checked" },
  { id: "advAlign", path: ["adverse", "alignment"], type: "value" },
  { id: "advNewLine", path: ["adverse", "newLine"], type: "checked" },
  { id: "bordPlacement", path: ["bordereau", "placement"], type: "value" },
  { id: "bordTitle", path: ["bordereau", "title"], type: "value" },
  { id: "bordTitleBold", path: ["bordereau", "titleBold"], type: "checked" },
  { id: "bordTitleItalic", path: ["bordereau", "titleItalic"], type: "checked" },
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
  { id: "bordMarkNew", path: ["bordereau", "markNew"], type: "checked" },
  { id: "bordNewMark", path: ["bordereau", "newMark"], type: "value" },
  { id: "bordGroupNames", path: ["bordereau", "showGroupNames"], type: "checked" },
  { id: "recGroupNames", path: ["bordereau", "showGroupNames"], type: "checked" }, // même réglage, vue Courriers
  // Mise en forme des noms de groupes (intertitres du bordereau / récapitulatif) — mêmes réglages dans les deux vues.
  { id: "bordGroupBold", path: ["bordereau", "groupBold"], type: "checked" },
  { id: "bordGroupItalic", path: ["bordereau", "groupItalic"], type: "checked" },
  { id: "bordGroupUnderline", path: ["bordereau", "groupUnderline"], type: "checked" },
  { id: "bordGroupSize", path: ["bordereau", "groupSize"], type: "value" },
  { id: "bordGroupAlign", path: ["bordereau", "groupAlign"], type: "value" },
  { id: "bordGroupIndent", path: ["bordereau", "groupIndent"], type: "value" },
  { id: "bordGroupKidsIndent", path: ["bordereau", "groupKidsIndent"], type: "value" },
  { id: "bordGroupSpaceBefore", path: ["bordereau", "groupSpaceBefore"], type: "checked" },
  { id: "bordGroupSpaceAfter", path: ["bordereau", "groupSpaceAfter"], type: "checked" },
  { id: "recGroupBold", path: ["bordereau", "groupBold"], type: "checked" },
  { id: "recGroupItalic", path: ["bordereau", "groupItalic"], type: "checked" },
  { id: "recGroupUnderline", path: ["bordereau", "groupUnderline"], type: "checked" },
  { id: "recGroupSize", path: ["bordereau", "groupSize"], type: "value" },
  { id: "recGroupAlign", path: ["bordereau", "groupAlign"], type: "value" },
  { id: "recGroupIndent", path: ["bordereau", "groupIndent"], type: "value" },
  { id: "recGroupKidsIndent", path: ["bordereau", "groupKidsIndent"], type: "value" },
  { id: "recGroupSpaceBefore", path: ["bordereau", "groupSpaceBefore"], type: "checked" },
  { id: "recGroupSpaceAfter", path: ["bordereau", "groupSpaceAfter"], type: "checked" },
  // uiOnly : réglage d'affichage du volet — ne touche pas le document, donc pas de
  // resynchronisation forcée quand il change (voir applySettings).
  { id: "uiContextBar", path: ["ui", "contextBar"], type: "checked", uiOnly: true },
  { id: "uiSlashCommands", path: ["ui", "slashCommands"], type: "checked", uiOnly: true },
  { id: "uiAdverseTab", path: ["ui", "adverseTab"], type: "checked", uiOnly: true },
  { id: "uiTheme", path: ["ui", "theme"], type: "value", uiOnly: true },
  { id: "uiDensity", path: ["ui", "density"], type: "value", uiOnly: true },
  { id: "uiPieceNames", path: ["ui", "pieceNames"], type: "value", uiOnly: true },
];
function getPath(obj, path) { return path.reduce((o, k) => (o ? o[k] : undefined), obj); }
function setPath(obj, path, val) {
  let o = obj;
  for (let i = 0; i < path.length - 1; i++) { if (!o[path[i]]) o[path[i]] = {}; o = o[path[i]]; }
  o[path[path.length - 1]] = val;
}
// Un réglage d'affichage ne touche pas le document : on l'enregistre et on rafraîchit
// le volet, sans la resynchronisation forcée (qui réécrit citations et bordereau).
function isUiOnly(path) {
  const key = path.join(".");
  return SETTING_BINDINGS.some((b) => b.uiOnly && b.path.join(".") === key);
}
// Emplacement du bordereau : le changer DÉPLACE le bordereau déjà présent (les autres réglages
// se contentent de le réécrire sur place). À appeler avant applySettings.
// Changer le TYPE de document change aussi l'emplacement voulu (bordereau d'acte / récapitulatif de
// courrier ont chacun le leur).
async function moveBordereauIfNeeded(path) {
  const key = path.join(".");
  const own = isLetterDoc() ? "lettre.recap.placement" : "bordereau.placement";
  if (disabled || (key !== own && key !== "docType") || !(WP.stats && WP.stats.hasBordereau)) return;
  const moved = await withBusy(() => WP.placeBordereau());
  if (moved === false) toast(BordWord() + " laissé à sa place : déplacez-le à la main");
}
// Point d'entrée unique d'un changement de réglage depuis le volet (champs liés, menu ☰).
async function changeSetting(path, value, uiOnly) {
  setPath(WP.model.settings, path, value);
  renderSettings();
  await moveBordereauIfNeeded(path);
  await applySettings(uiOnly);
}

// ---- Menu ☰ de l'en-tête ----
function appMenuOpen() { return !el("appMenu").classList.contains("hidden"); }
function openAppMenu() {
  renderSettings(); // valeurs à jour (elles ont pu changer dans la fenêtre Réglages)
  el("appMenu").classList.remove("hidden");
  el("menuToggle").setAttribute("aria-expanded", "true");
}
function closeAppMenu() {
  const menu = el("appMenu");
  if (!menu || menu.classList.contains("hidden")) return;
  menu.classList.add("hidden");
  el("menuToggle").setAttribute("aria-expanded", "false");
}
function toggleAppMenu() { appMenuOpen() ? closeAppMenu() : openAppMenu(); }
async function onAppMenuClick(e) {
  const item = e.target.closest("[data-menu]");
  if (item) {
    closeAppMenu();
    const act = item.dataset.menu;
    if (act === "settings") openSettingsDialog();
    else if (act === "help") openHelpDialog();
    else if (act === "warnings") openWarningsModal();
    else if (act === "import") openImport(paneSide);
    else if (act === "renumber") openRenumberDialog();
    else if (act === "markcomm") await markAllCommunicated();
    else if (act === "cleancopy") await openCleanCopy();
    else if (act === "disable") await toggleDisabled();
    return;
  }
  // Acte / Courrier : le menu reste ouvert, pour voir la bascule se faire.
  const seg = e.target.closest("[data-doctype]");
  if (!seg || seg.dataset.doctype === (WP.model.settings.docType || "acte")) return;
  confirmDeleteBordereau = false;
  try {
    await changeSetting(["docType"], seg.dataset.doctype, false);
    toast(seg.dataset.doctype === "lettre" ? "Document : courrier — pièces jointes et récapitulatif" : "Document : acte — pièces citées et bordereau");
  } catch (err) { reportError(err); }
}

async function applySettings(uiOnly) {
  await WP.save();
  WP.saveGlobal();
  if (uiOnly) {
    // Décoché alors que le bandeau est affiché : il disparaît tout de suite. Recoché :
    // il reviendra au prochain déplacement du curseur.
    if (!contextBarEnabled()) hideContextBar();
    applyUiPrefs(); // thème, densité, onglet adverse
    render();
    return;
  }
  // Désactivé pour ce document : on n'y écrit rien ; la synchro se fera à la réactivation.
  if (!disabled) await withBusy(() => syncDoc(true)); // force : un réglage de mise en forme peut changer sans changer les textes
  render();
}
function renderSettings() {
  for (const b of SETTING_BINDINGS) {
    const elm = el(b.id);
    if (!elm || elm === document.activeElement) continue; // ne pas écraser une saisie en cours
    const v = getPath(WP.model.settings, b.path);
    if (b.type === "checked") elm.checked = !!v;
    else elm.value = v ?? "";
  }
  const docType = WP.model.settings.docType === "lettre" ? "lettre" : "acte";
  document.querySelectorAll(".seg-btn[data-doctype]").forEach((b) => b.setAttribute("aria-pressed", String(b.dataset.doctype === docType)));
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
  const label = el("disableLabel");
  if (label) label.textContent = disabled ? "Réactiver pour ce document" : "Désactiver pour ce document";
}

// Onglets « Nos pièces » / « Pièces adverses » : le volet ne montre qu'une liste à la fois.
function renderSideTabs() {
  const count = (side) => WP.sidePieces(side).filter((p) => !p.container).length;
  // Onglet adverse masqué à la demande — mais JAMAIS quand le document contient déjà des
  // pièces adverses : un réglage d'affichage ne doit pas rendre des données inaccessibles.
  // Il réapparaît donc de lui-même dès qu'une pièce adverse existe.
  const nav = el("sideTabs");
  if (nav) nav.classList.toggle("hidden", uiPrefs().adverseTab === false && count("adv") === 0);
  for (const [id, side] of [["tabOwn", "own"], ["tabAdv", "adv"]]) {
    const t = el(id);
    if (!t) continue;
    const on = paneSide === side;
    t.classList.toggle("is-active", on);
    t.setAttribute("aria-selected", String(on));
    if (side === "own") {
      t.firstChild.nodeValue = isLetterDoc() ? "Pièces jointes " : "Nos pièces ";
      t.title = isLetterDoc() ? "Pièces jointes : numérotation, récapitulatif, verrous" : "Nos pièces : numérotation, bordereau, verrous";
    }
    const n = count(side);
    t.querySelector(".side-count").textContent = n ? String(n) : "";
  }
}
function setPaneSide(side) {
  if (paneSide === side) return;
  closeLockMenu();
  paneSide = side;
  pieceFilter = "";
  el("pieceSearch").value = "";
  confirmDeleteId = null;
  collapseAdd();
  render();
  window.scrollTo({ top: 0 });
}

function render() {
  applyDisabledUI();
  renderBordereauAction();
  renderSettings();
  const adv = isAdvMode();
  const letter = isLetterDoc();
  const wrapNames = namesWrap();
  el("app").classList.toggle("adv-mode", adv);
  el("app").classList.toggle("names-wrap", wrapNames);
  el("bordNoTitleLabel").textContent = letter ? "Récapitulatif sans titre" : "Bordereau sans titre";
  el("app").classList.toggle("letter-mode", letter); // masque ce qui ne vaut que pour un acte (.acte-only)
  renderSideTabs();
  const all = WP.sidePieces(paneSide);
  const list = el("pieceList");
  list.innerHTML = "";
  const realCount = all.filter((p) => !p.container).length; // les conteneurs de groupe sont invisibles
  el("pieceCount").textContent = realCount ? `(${realCount})` : "";
  el("listTitle").textContent = adv ? "Pièces adverses" : letter ? "Pièces jointes" : "Pièces";
  el("addToggleLabel").textContent = adv ? "Nouvelle pièce adverse" : letter ? "Nouvelle pièce jointe" : "Nouvelle pièce";
  el("addToggle").title = el("addToggleLabel").textContent;
  el("addToggle").setAttribute("aria-label", el("addToggleLabel").textContent);
  el("newPieceName").placeholder = adv ? "Intitulé de la pièce adverse (ex. Assignation…)"
    : letter ? "Intitulé de la pièce jointe (ex. Contrat de bail…)" : "Intitulé de la pièce (ex. Contrat de bail…)";
  el("emptyState").textContent = adv
    ? "Aucune pièce adverse pour l'instant. Tapez /pa1 dans le texte pour citer (et créer) la pièce adverse n°1, ou ajoutez-en une ci-dessus."
    : letter
      ? "Aucune pièce jointe pour l'instant. Ajoutez-en une ci-dessus : elle figurera au récapitulatif de fin de courrier, sans avoir à être citée dans le texte."
      : "Aucune pièce pour l'instant. Ajoutez-en une ci-dessus, puis cliquez sur l'icône d'insertion (curseur de texte) pour la citer à l'endroit du curseur.";
  el("emptyState").classList.toggle("hidden", realCount > 0);
  el("listToolbar").classList.toggle("hidden", realCount === 0 || adv);
  renderLockAll();
  renderManualSwitch();
  renderUndo();
  const manual = !adv && WP.isManual();

  const numbers = (adv ? WP.stats.advNumbers : WP.stats.numbers) || new Map();
  // Tri par NUMÉRO (ordre naturel). Un GROUPE sans tête (conteneur) apparaît comme une ligne d'en-tête
  // juste avant ses sous-pièces N.1, N.2… : on peut ainsi renuméroter, déplacer ou verrouiller le groupe.
  const ordered = [...all].sort((a, b) => WP.naturalCompare(numbers.get(a.id), numbers.get(b.id)));

  // Champ de recherche : visible dès qu'il y a plusieurs pièces ; filtre par nom ou numéro.
  const showSearch = realCount >= 4;
  const search = el("pieceSearch");
  el("pieceSearchWrap").classList.toggle("hidden", !showSearch);
  if (!showSearch && pieceFilter) { pieceFilter = ""; search.value = ""; }
  el("pieceSearchClear").classList.toggle("hidden", !search.value);
  const f = pieceFilter.toLowerCase();
  const matches = (p) => (p.name || "").toLowerCase().includes(f) || String(numbers.get(p.id) ?? "").toLowerCase().includes(f);
  const shown = f
    ? ordered.filter((p) => (p.container ? matches(p) || all.some((k) => k.parentId === p.id && matches(k)) : matches(p)))
    : ordered;

  if (f && shown.length === 0) {
    const note = document.createElement("div");
    note.className = "empty muted";
    note.textContent = (adv ? "Aucune pièce adverse" : "Aucune pièce") + " ne correspond à « " + pieceFilter + " ».";
    list.appendChild(note);
  }

  for (const p of shown) {
    const num = numbers.get(p.id) ?? "?";
    const count = WP.stats.counts[p.id] || 0;
    const row = document.createElement("div");
    row.dataset.id = p.id; // pour l'animation de déplacement (FLIP)
    row.dataset.parent = p.parentId || "";  // glisser-déposer : sœurs = même niveau, même groupe
    row.dataset.locked = p.locked ? "1" : "";
    row.dataset.group = !p.parentId && (p.container || all.some((k) => k.parentId === p.id)) ? "1" : ""; // cible « dans le groupe »
    if (p.container && adv) { // groupe adverse (n°4.1, 4.2…) : simple en-tête, numéro modifiable
      const kids = all.filter((k) => k.parentId === p.id).length;
      row.className = "piece group-head";
      row.innerHTML = `
        <div class="piece-top">
          <input class="badge-input adv" data-act="setnum" data-id="${p.id}" value="${WP.escapeHtml(String(num))}" title="Numéro du groupe adverse — le modifier renumérote ses sous-pièces" />
          ${nameField(p, wrapNames, `Groupe · ${kids} sous-pièce${kids > 1 ? "s" : ""} — ajouter un nom`, "Nom du groupe (facultatif)")}
          ${addSubBtn(p)}
        </div>`;
      list.appendChild(row);
      continue;
    }
    if (p.container) { // en-tête de groupe : numéro + déplacement + verrou (les actions « pièce » n'ont pas de sens)
      const kids = all.filter((k) => k.parentId === p.id).length;
      row.className = "piece group-head has-grip" + (flashIds.has(p.id) ? " flash" : "");
      row.innerHTML = `
        <div class="piece-top">
          ${gripBtn(p, "le groupe")}
          <input class="badge-input ${p.locked ? "locked" : ""}" data-act="setnum" data-id="${p.id}" value="${WP.escapeHtml(String(num))}" ${p.locked ? "readonly" : ""}
            title="${p.locked ? "Groupe verrouillé — déverrouillez-le (🔓) pour changer son numéro" : "Numéro du groupe — le modifier renumérote toutes ses sous-pièces ; vider le champ = numéro automatique"}" />
          ${nameField(p, wrapNames, `Groupe · ${kids} sous-pièce${kids > 1 ? "s" : ""} — ajouter un nom`, letter
            ? "Nom du groupe (facultatif) — repère dans le volet ; figure seul, en intertitre, au récapitulatif si « Afficher les noms des groupes » est coché (Réglages)"
            : "Nom du groupe (facultatif) — repère dans le volet ; figure seul, en intertitre, au bordereau si « Afficher les noms des groupes » est coché (Réglages)")}
          ${addSubBtn(p)}
          <button class="mini lock ${p.locked ? "on" : ""}" data-act="togglelock" data-id="${p.id}" title="${p.locked ? "Groupe verrouillé — cliquer pour déverrouiller" : "Cliquer pour verrouiller le numéro du groupe"}" aria-label="Verrou du groupe">${icon(p.locked ? "lock" : "unlock", 14)}</button>
        </div>`;
      list.appendChild(row);
      continue;
    }
    row.className = "piece" + (adv ? "" : " has-grip") + (p.scanned ? " scanned" : "") + (flashIds.has(p.id) ? " flash" : "") + (selectedPieceId === p.id ? " selected" : ""); // sous-pièces N.x mises en page comme les autres (pas d'indentation)
    // Barre d'actions : Insérer · loupe (occurrence suivante) · ▲/▼ (réordonner+renuméroter) · verrou · poubelle.
    if (adv) row.classList.add("adv");
    const reorderBtns = gripBtn(p, "la pièce");
    // Enveloppe : pièce déjà communiquée (pleine, verte) ou nouvelle — actes seulement.
    const commBtn = letter ? "" : (p.communicated
      ? `<button class="mini comm on" data-act="togglecomm" data-id="${p.id}" title="Communiquée${p.communicatedOn ? " le " + frDate(p.communicatedOn) : ""} — cliquer pour la marquer comme nouvelle" aria-label="Pièce communiquée">${icon("mailCheck", 15)}</button>`
      : `<button class="mini comm" data-act="togglecomm" data-id="${p.id}" title="Pas encore communiquée — cliquer pour la marquer comme communiquée" aria-label="Pièce pas encore communiquée">${icon("mail", 15)}</button>`);
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
             ${adv ? "" : commBtn}
             <span class="count ${count ? "" : "zero"}" title="Nombre de citations dans le texte">${count}×</span>
             <button class="mini trash" data-act="delete" data-id="${p.id}" title="Supprimer la pièce" aria-label="Supprimer la pièce">${TRASH_SVG}</button>
           </div>`;
    const badgeTitle = manual
      ? "Numéro de la pièce (numérotation manuelle) — tapez le numéro voulu : 12, 4.1, 2 bis, i.3, A-2… Il ne changera plus de lui-même."
      : "Numéro de la pièce — le modifier le verrouille à ce numéro (badge hachuré) ; vider le champ = numéro automatique (badge plein) ; « i. » = série automatique i.1, i.2…";
    const badgeTitleLocked = "Numéro verrouillé — déverrouillez la pièce (🔓) pour pouvoir le changer";
    // Pièce adverse : numéro de l'adversaire, toujours modifiable (pas de verrou ni de renumérotation).
    const badge = adv
      ? `<input class="badge-input adv ${p.locked ? "locked" : ""}" data-act="setnum" data-id="${p.id}" value="${WP.escapeHtml(String(num))}" ${p.locked ? "readonly" : ""} title="${p.locked ? "Numéro verrouillé — déverrouillez la pièce (🔓) pour le corriger" : "Numéro de la pièce adverse (celui du bordereau adverse) — modifiable"}" />`
      : `<input class="badge-input ${p.locked ? "locked" : ""}" data-act="setnum" data-id="${p.id}" value="${WP.escapeHtml(String(num))}" title="${p.locked ? badgeTitleLocked : badgeTitle}" ${p.locked ? "readonly" : ""} />`;
    row.innerHTML = `
      <div class="piece-top">
        ${adv ? "" : reorderBtns}
        ${badge}
        ${nameField(p, wrapNames)}
        ${!p.parentId && all.some((k) => k.parentId === p.id) ? addSubBtn(p) : ""}
        ${p.scanned ? `<span class="scan-mark" title="Pièce déjà présente dans le document (détectée automatiquement). Protégée du verrou général ; déverrouille-la à la main pour la modifier.">${icon("file", 14)}</span>` : ""}
      </div>
      ${actions}`;
    list.appendChild(row);
  }
  if (addSubFor) {
    // Ligne de saisie de la sous-pièce suivante, sous la dernière sous-pièce affichée du groupe.
    const head = WP.findPiece(addSubFor);
    const rowsEls = [...list.querySelectorAll(".piece[data-id]")];
    const anchorRow = rowsEls.filter((r) => r.dataset.parent === addSubFor).pop() || rowsEls.find((r) => r.dataset.id === addSubFor);
    if (!head || !anchorRow) addSubFor = null;
    else {
      const next = WP.previewSubNumber(addSubFor);
      const add = document.createElement("div");
      add.className = "piece add-sub";
      add.innerHTML = `
        <div class="piece-top">
          <span class="badge next" title="Numéro de la nouvelle sous-pièce">${WP.escapeHtml(next)}</span>
          <input id="addSubName" class="piece-name" placeholder="Intitulé de la pièce ${WP.escapeHtml(next)}" aria-label="Intitulé de la nouvelle sous-pièce" autocomplete="off" />
          <button class="mini addsub-ok" data-act="addsubok" title="Créer la pièce ${WP.escapeHtml(next)} (Entrée)" aria-label="Créer">${icon("check", 15)}</button>
          <button class="mini addsub-cancel" data-act="addsubcancel" title="Annuler (Échap)" aria-label="Annuler">${icon("x", 15)}</button>
        </div>`;
      anchorRow.after(add);
    }
  }
  if (flashIds.size) flashIds = new Set(); // le flash n'est consommé qu'une fois
  if (wrapNames) autosizeNames();
  renderDiag();
}


// Six points (2 colonnes × 3 rangées) répartis sur la hauteur de l'étiquette ; la poignée occupe toute sa hauteur, à gauche.
const GRIP_TALL = '<span class="grip-dots" aria-hidden="true"><i></i><i></i><i></i><i></i><i></i><i></i></span>';
// (points en CSS, dimensions entières : rendu net et parfaitement aligné, quel que soit le zoom)
// Poignée ⋮⋮ de glisser-déposer (taskpane-dnd.js). Inactive si la pièce est verrouillée, ou pendant
// une recherche (une partie des pièces est masquée : la place visée serait ambiguë).
function gripBtn(p, what) {
  const why = p.locked ? "Numéro verrouillé — déverrouillez pour réordonner"
    : pieceFilter ? "Effacez la recherche pour réordonner" : "";
  return `<button class="mini grip" data-act="grip" data-id="${p.id}" ${why ? "disabled" : ""} ` +
    `title="${why || "Glisser pour déplacer " + what + " — librement, y compris dans un groupe (flèches ↑ ↓ au clavier)"}" aria-label="Déplacer ${what}">${GRIP_TALL}</button>`;
}
function frDate(iso) {
  const m = String(iso || "").match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? `${m[3]}/${m[2]}/${m[1]}` : "";
}

let diagCollapsed = false; // panneau « À vérifier » replié (le temps de la session)
function renderDiag() {
  const box = el("diag");
  const esc = WP.escapeHtml;
  const numbers = WP.stats.numbers || new Map();
  const counts = WP.stats.counts || {};
  const hasBord = !!WP.stats.hasBordereau;
  const pend = new Set(pendingChain.map((c) => c.id)); // pièces en cours de nommage : pas encore d'alerte
  const pieces = WP.model.pieces.filter((p) => !p.container);
  const uncited = pieces.filter((p) => !(counts[p.id] > 0) && !pend.has(p.id));

  // Chaque point est une CARTE : { key, kind: "err" | "warn" | "info", title, detail (HTML), text (texte
  // brut, pour la fenêtre des avertissements), actions: [{ act, label, attrs, primary, title }] }.
  // La CLÉ est stable : « Ignorer » masque ce point durablement sans masquer les autres, et il
  // réapparaît si la situation change (la clé inclut ce qui le caractérise).
  const items = [];
  const numOf = (p) => String(numbers.get(p.id) ?? "?");
  // Pastille cliquable d'une pièce : numéro, nom, nombre de citations (clic = la montrer dans la liste).
  const chip = (p, opts) => {
    const n = counts[p.id] || 0;
    const name = p.container ? (p.name || "groupe de sous-pièces") : (p.name || "sans nom");
    return `<button type="button" class="dchip" data-act="showpiece" data-id="${p.id}" title="Montrer cette pièce dans la liste">` +
      `<b>n°${esc(numOf(p))}</b><span class="dchip-name">${esc(name)}</span>` +
      (opts && opts.count === false ? "" : `<span class="dchip-n">${n}×</span>`) + `</button>`;
  };
  const chips = (arr, opts) => `<div class="dchips">${arr.map((p) => chip(p, opts)).join("")}</div>`;

  // Suivi des modifications → simple information.
  if (WP.stats.tracking && WP.stats.tracking !== "Off") {
    items.push({ key: "tracking", kind: "info", title: "Suivi des modifications activé",
      detail: "LexRef le suspend le temps de ses propres mises à jour : renumérotations et bordereau ne deviennent pas des révisions.",
      text: "Suivi des modifications activé — les mises à jour automatiques de LexRef ne sont pas enregistrées comme révisions." });
  }
  // Citations orphelines (pointent vers une pièce supprimée).
  if (WP.stats.orphanCCs > 0) {
    const k = WP.stats.orphanCCs, s = k > 1 ? "s" : "";
    items.push({ key: "orphans:" + k, kind: "err", title: `${k} citation${s} sans pièce`,
      detail: `Elle${s} renvoie${k > 1 ? "nt" : ""} à une pièce supprimée et ${k > 1 ? "sont marquées" : "est marquée"} « ⚠ » dans le texte : corrigez-la${s} ou supprimez-la${s}.`,
      text: `${k} citation(s) pointent vers une pièce supprimée (marquées « ⚠ »).` });
  }
  // Pièces jamais citées (courrier : sans objet — une pièce jointe n'a pas à être citée).
  if (!isLetterDoc() && uncited.length) {
    const k = uncited.length, s = k > 1 ? "s" : "";
    const key = "uncited:" + uncited.map((p) => p.id).sort().join(",");
    const names = uncited.map((p) => `n°${numOf(p)} ${p.name || "sans nom"}`).join(", ");
    items.push(hasBord
      ? { key, kind: "err", title: `${k} pièce${s} au bordereau, jamais citée${s}`,
        detail: `Citez-l${k > 1 ? "es" : "a"} dans le texte (icône d'insertion) ou retirez-l${k > 1 ? "es" : "a"} de la liste.` + chips(uncited, { count: false }),
        text: `Pièce(s) présente(s) au bordereau mais jamais citée(s) dans le texte : ${names}.` }
      : { key, kind: "warn", title: `${k} pièce${s} pas encore citée${s}`,
        detail: chips(uncited, { count: false }), text: `${k} pièce(s) pas encore citée(s) : ${names}.` });
  }

  // Même NUMÉRO pour plusieurs pièces : UN point par numéro. Trois issues possibles :
  //  - « Même pièce + précision » : l'une est la pièce, l'autre la même avec une précision
  //    (« Contrat, page 12 ») → ses citations rejoignent la pièce en gardant la précision ;
  //  - « Fusionner » : même pièce, libellés harmonisés ;
  //  - « Pièces différentes » : l'une reçoit un autre numéro.
  // (Courrier sans numérotation : les numéros ne paraissent nulle part → rien à signaler.)
  const byNum = new Map();
  for (const p of (WP.letterPlain() ? [] : WP.model.pieces)) {
    const n = String(numbers.get(p.id) ?? "").trim();
    if (!n) continue;
    if (!byNum.has(n)) byNum.set(n, []);
    byNum.get(n).push(p);
  }
  for (const [n, arr] of byNum) {
    if (arr.length < 2) continue;
    const key = "dupnum:" + n;
    if (arr.some((p) => p.container)) {
      items.push({ key, kind: "warn", title: `N°${n} : un groupe et une pièce`, detail: "Donnez un autre numéro à l'un des deux." + chips(arr),
        text: `Numéro en double : n°${n}.` });
      continue;
    }
    const pair = arr.length === 2 ? dupPair(arr[0], arr[1]) : null;
    const cnt = (p) => counts[p.id] || 0;
    const keep = pair ? pair.keep : [...arr].sort((a, b) => cnt(b) - cnt(a) || (b.name ? 1 : 0) - (a.name ? 1 : 0))[0];
    const others = arr.filter((p) => p !== keep);
    const actions = [];
    if (pair) {
      actions.push({ act: "sameprec", label: "Même pièce + précision", primary: !!pair.precision,
        attrs: { keep: keep.id, other: pair.other.id },
        title: `Les citations de « ${pair.other.name || "sans nom"} » deviennent des citations de « ${keep.name || "sans nom"} », avec leur précision (page, §…)` });
    }
    actions.push({ act: "merge", label: "Fusionner", primary: !pair || !pair.precision, attrs: { keep: keep.id, others: others.map((p) => p.id).join(",") },
      title: `Une seule pièce « ${keep.name || "sans nom"} » : les citations des autres la rejoignent` });
    if (pair) actions.push({ act: "separate", label: "Pièces différentes", attrs: { id: pair.other.id },
      title: `« ${pair.other.name || "sans nom"} » reçoit un autre numéro` });
    items.push({ key, kind: "warn",
      title: `${arr.length} pièces portent le n°${n}`,
      detail: (pair && pair.precision
        ? `Même pièce citée avec une précision (« ${esc(pair.precision.trim())} ») ?`
        : "Même pièce sous deux libellés, ou deux pièces différentes ?") + chips(arr),
      text: `Numéro en double : n°${n} (${arr.map((p) => p.name || "sans nom").join(" / ")}).`, actions });
  }
  // Numéros manquants : un seul point pour toute la suite.
  const ints = [...byNum.keys()].filter((n) => /^\d+$/.test(n)).map(Number).sort((a, b) => a - b);
  const gaps = [];
  if (ints.length >= 2) for (let i = ints[0]; i <= ints[ints.length - 1]; i++) if (!ints.includes(i)) gaps.push(i);
  if (gaps.length) {
    const g = gaps.length > 1;
    items.push({ key: "gaps:" + gaps.join(","), kind: "warn",
      title: g ? `Numéros manquants : ${gaps.slice(0, 8).join(", ")}${gaps.length > 8 ? "…" : ""}` : `Numéro manquant : n°${gaps[0]}`,
      detail: WP.isManual && WP.isManual() ? "Numérotation manuelle : renumérotez les pièces suivantes si ce trou n'est pas voulu." : "Une pièce verrouillée ou scannée garde un numéro plus loin dans la suite.",
      text: `Numéro(s) manquant(s) dans la suite : ${gaps.join(", ")}.` });
  }
  // Même NOM sous plusieurs numéros → fusion proposée.
  const norm = (x) => String(x || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().trim().replace(/\s+/g, " ");
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
    items.push({ key: "dupname:" + norm(keep.name), kind: "warn",
      title: `« ${keep.name} » existe ${arr.length} fois`,
      detail: "Même pièce enregistrée sous plusieurs numéros ?" + chips(sorted),
      text: `Pièces identiques : « ${keep.name} » (${sorted.map((p) => "n°" + numOf(p)).join(", ")})`,
      actions: [{ act: "merge", label: `Garder le n°${numOf(keep)}`, primary: true, attrs: { keep: keep.id, others: others.map((p) => p.id).join(",") },
        title: "Une seule pièce : les citations des autres la rejoignent" }] });
  }

  // Filtrage : on retire les points que l'utilisateur a explicitement ignorés pour ce document.
  const shown = items.filter((it) => !WP.isWarningIgnored(it.key));
  lastDiagItems = items.map((it) => ({ key: it.key, kind: it.kind === "err" ? "err" : "warn", text: it.text }));
  const nIgnored = items.length - shown.length;
  el("menuWarnings").classList.toggle("hidden", !nIgnored);
  el("menuWarningsLabel").textContent = `Avertissement${nIgnored > 1 ? "s" : ""} ignoré${nIgnored > 1 ? "s" : ""} (${nIgnored})`;

  const nErr = shown.filter((it) => it.kind === "err").length;
  const nWarn = shown.filter((it) => it.kind === "warn").length;
  hasErrors = nErr > 0;
  hasWarnings = nWarn > 0;
  if (!shown.length && !nIgnored) {
    box.className = "diag hidden";
    box.innerHTML = "";
    renderPips();
    if (warnModalOpen) renderWarningsModal();
    return;
  }

  const attrs = (o) => Object.entries(o || {}).map(([k, v]) => ` data-${k}="${esc(String(v))}"`).join("");
  const card = (it) =>
    `<div class="dcard ${it.kind}">` +
      `<div class="dcard-head"><span class="dcard-ico" aria-hidden="true">${it.kind === "err" ? "!" : it.kind === "info" ? "i" : "?"}</span>` +
      `<span class="dcard-title">${esc(it.title)}</span></div>` +
      (it.detail ? `<div class="dcard-detail">${it.detail}</div>` : "") +
      `<div class="dcard-actions">` +
        (it.actions || []).map((a) => `<button type="button" class="dbtn${a.primary ? " primary" : ""}" data-act="${a.act}"${attrs(a.attrs)}${a.title ? ` title="${esc(a.title)}"` : ""}>${esc(a.label)}</button>`).join("") +
        `<button type="button" class="dbtn ghost" data-act="ignorewarn" data-key="${esc(it.key)}" title="Ne plus signaler ce point pour ce document (réversible : ☰ ▸ Avertissements ignorés)">Ignorer</button>` +
      `</div></div>`;
  // Ordre : à corriger, puis à vérifier, puis informations.
  const rank = { err: 0, warn: 1, info: 2 };
  const ordered = [...shown].sort((a, b) => rank[a.kind] - rank[b.kind]);
  const summary = [];
  if (nErr) summary.push(`<span class="dsum err">${nErr} à corriger</span>`);
  if (nWarn) summary.push(`<span class="dsum warn">${nWarn} à vérifier</span>`);
  const nInfo = shown.length - nErr - nWarn;
  if (nInfo) summary.push(`<span class="dsum info">${nInfo} info</span>`);
  let html = "";
  if (shown.length) {
    html += `<button type="button" class="diag-head" data-act="togglediag" aria-expanded="${!diagCollapsed}">` +
      `<span class="diag-head-title">Contrôles</span>${summary.join("")}` +
      `<svg class="diag-chev" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="m6 9 6 6 6-6"/></svg></button>`;
    if (!diagCollapsed) html += `<div class="diag-cards">${ordered.map(card).join("")}</div>`;
  }
  if (nIgnored) {
    const s = nIgnored > 1 ? "s" : "";
    html += `<div class="diag-ignored">${nIgnored} point${s} ignoré${s} · ` +
      `<button class="link-btn" data-act="openwarnings">Gérer…</button></div>`;
  }
  box.className = "diag" + (diagCollapsed ? " collapsed" : "");
  box.innerHTML = html;
  renderPips();
  if (warnModalOpen) renderWarningsModal(); // la pop-up suit l'état réel si elle est ouverte
}
// Deux pièces au même numéro : laquelle est « la » pièce, et l'autre en est-elle une variante avec
// précision (« Contrat de bail » / « Contrat de bail, page 12 ») ?
function dupPair(a, b) {
  const pa = WP.splitPrecision(a.name, b.name);
  if (pa != null) return { keep: a, other: b, precision: pa };
  const pb = WP.splitPrecision(b.name, a.name);
  if (pb != null) return { keep: b, other: a, precision: pb };
  const cnt = (p) => (WP.stats.counts || {})[p.id] || 0;
  const keep = cnt(b) > cnt(a) || (!a.name && b.name) ? b : a;
  return { keep, other: keep === a ? b : a, precision: null };
}
// « Même pièce + précision » : demande la précision (pré-remplie), puis rattache les citations de
// `otherId` à `keepId` en leur gardant cette précision. Le nom de la pièce et le bordereau ne changent pas.
async function samePieceWithPrecision(keepId, otherId) {
  const keep = WP.findPiece(keepId), other = WP.findPiece(otherId);
  if (!keep || !other) return;
  const esc = WP.escapeHtml;
  let guess = WP.splitPrecision(keep.name, other.name);
  if (guess == null) {
    const t = WP.precisionTail(other.name);
    guess = t ? t.extra : "";
  }
  const num = WP.numberOf(keepId);
  const res = await showModal({
    title: "Même pièce, avec une précision",
    msg: `Les citations de « ${esc(other.name || "sans nom")} » deviennent des citations de la pièce n°${esc(String(num ?? "?"))} ` +
      `« <b>${esc(keep.name || "sans nom")}</b> », suivies de la précision ci-dessous. Elle ne vaut que pour ces citations : ` +
      `le nom de la pièce, le bordereau et les autres citations ne changent pas.`,
    input: { label: "Précision", value: guess.replace(/^\s*[,;]\s*/, ""), placeholder: "ex. page 12, § 4, annexe 2" },
    buttons: [{ label: "Annuler", value: "cancel" }, { label: "Valider", value: "ok", primary: true }],
  });
  if (!res || res.value !== "ok") return;
  let extra = String(res.text || "").trim();
  if (extra && !/^[,;:(\[\-–—]/.test(extra)) extra = ", " + extra; // « page 12 » → « , page 12 »
  else if (extra && /^[(\[\-–—]/.test(extra)) extra = " " + extra;
  commitUndo(snapshotModel(), "même pièce + précision");
  await withBusy(async () => {
    await WP.mergePieces(keepId, [otherId], { extras: { [otherId]: extra } });
    await syncDoc();
  });
  render();
  toast(extra ? `Rattachée à la pièce n°${num} avec la précision « ${extra.replace(/^[\s,;]+/, "")} »` : `Rattachée à la pièce n°${num}`);
}
// « Pièces différentes » : la pièce reçoit un numéro à elle.
async function separateDuplicate(id) {
  const p = WP.findPiece(id);
  if (!p) return;
  commitUndo(snapshotModel(), "nouveau numéro");
  let num = null;
  await withBusy(async () => {
    num = await WP.separatePiece(id);
    await syncDoc();
  });
  selectedPieceId = id;
  flashIds = new Set([id]);
  render();
  toast(`« ${p.name || "sans nom"} » → n°${num ?? "?"} (badge modifiable pour choisir un autre numéro)`);
}
// Clic sur une pastille de pièce dans les contrôles : la montrer (et la sélectionner) dans la liste.
function showPieceInList(id) {
  const p = WP.findPiece(id);
  if (!p) return;
  const side = WP.sideOf(id);
  if (side !== paneSide) setPaneSide(side);
  if (pieceFilter) { pieceFilter = ""; el("pieceSearch").value = ""; render(); }
  setSelectedPiece(id);
  const row = el("pieceList").querySelector(`.piece[data-id="${id}"]`);
  if (row) {
    row.scrollIntoView({ block: "center", behavior: "smooth" });
    row.classList.remove("flash"); void row.offsetWidth; row.classList.add("flash");
  }
}

// ---- Pop-up « Tout réafficher » : gérer les avertissements ignorés ----
// Section 1 : les points ignorés (bouton « Réafficher ») ; section 2 : les points affichés (bouton
// « Ignorer »). Chaque clic s'applique tout de suite ; l'utilisateur ferme la pop-up quand il a fini.
let lastDiagItems = [];   // tous les points de diagnostic d'actualité (calculés par renderDiag)
let warnModalOpen = false;
function renderWarningsModal() {
  const ignored = lastDiagItems.filter((it) => WP.isWarningIgnored(it.key));
  const shown = lastDiagItems.filter((it) => !WP.isWarningIgnored(it.key));
  const row = (it, isIgnored) =>
    `<li class="wm-row ${it.kind}${isIgnored ? " is-ignored" : ""}">` +
    `<span class="wm-dot" aria-hidden="true"></span><span class="wm-txt">${WP.escapeHtml(it.text || "")}</span>` +
    `<button class="mini diag-ignore" data-wm="${isIgnored ? "unignore" : "ignore"}" data-key="${WP.escapeHtml(it.key)}">` +
    `${isIgnored ? "Réafficher" : "Ignorer"}</button></li>`;
  const section = (title, items, isIgnored, empty) =>
    `<div class="wm-sec"><div class="wm-head">${title} <span class="wm-count">${items.length}</span></div>` +
    (items.length ? `<ul class="wm-list">${items.map((it) => row(it, isIgnored)).join("")}</ul>` : `<div class="wm-empty">${empty}</div>`) +
    `</div>`;
  el("modalCard").innerHTML =
    `<div class="modal-title">Avertissements</div>` +
    `<div class="modal-msg">Réaffichez un point ignoré, ou ignorez un point affiché. Les choix valent pour ce document.</div>` +
    section("Ignorés", ignored, true, "Aucun avertissement ignoré.") +
    section("Affichés", shown, false, "Aucun avertissement affiché.") +
    `<div class="modal-btns">` +
    (ignored.length > 1 ? `<button class="btn outline" data-wm="unignoreall">Tout réafficher</button>` : "") +
    `<button class="btn primary" data-wm="close">Fermer</button></div>`;
}
function openWarningsModal() {
  warnModalOpen = true;
  renderWarningsModal();
  el("modalCard").classList.add("wide");
  el("modal").classList.remove("hidden");
}
function closeWarningsModal() {
  warnModalOpen = false;
  el("modalCard").classList.remove("wide");
  el("modal").classList.add("hidden");
}
async function onWarningsModalClick(e) {
  const b = e.target.closest("[data-wm]");
  if (!b) return false;
  const a = b.dataset.wm;
  if (a === "close") closeWarningsModal();
  else if (a === "ignore") { await WP.ignoreWarning(b.dataset.key); renderDiag(); }
  else if (a === "unignore") { await WP.unignoreWarning(b.dataset.key); renderDiag(); }
  else if (a === "unignoreall") { await WP.restoreWarnings(); renderDiag(); }
  return true;
}

// Masque durablement un point de diagnostic (stocké dans le document, réversible).
async function ignoreWarning(key) {
  await WP.ignoreWarning(key);
  renderDiag();
  toast("Avertissement ignoré pour ce document");
}

// ------------------------------------------------------------
// Liaison des événements
// ------------------------------------------------------------
function bindUI() {
  el("addToggle").addEventListener("click", expandAdd);
  el("fromSelBtn").addEventListener("click", () => createFromSelection().catch(reportError));
  el("manualSwitch").addEventListener("click", () => toggleManualNumbering().catch(reportError));
  document.addEventListener("keydown", onDeleteKey);
  // Clic sur une pièce = sélection (Suppr la supprime) ; clic ailleurs = plus de sélection.
  document.addEventListener("pointerdown", (e) => {
    const row = e.target.closest && e.target.closest("#pieceList .piece[data-id]:not(.group-head)");
    if (row) setSelectedPiece(row.dataset.id);
    else if (!(e.target.closest && e.target.closest("#pieceList"))) setSelectedPiece(null);
  });
  el("tabOwn").addEventListener("click", () => setPaneSide("own"));
  bindImportUI();     // taskpane-import.js
  bindDragReorder();  // taskpane-dnd.js
  el("tabAdv").addEventListener("click", () => setPaneSide("adv"));
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
  el("reenableBtn").addEventListener("click", toggleDisabled);

  // Croix maison (haut gauche) : referme le volet via l'API Office (shared runtime).
  el("paneClose").addEventListener("click", () => {
    try { if (Office.addin && Office.addin.hide) Office.addin.hide().catch(() => {}); } catch (e) {}
  });

  // Menu ☰ : réglages et aide (FENÊTRES CENTRALES), plus les réglages propres à CE document.
  el("menuToggle").addEventListener("click", (e) => { e.stopPropagation(); toggleAppMenu(); });
  el("appMenu").addEventListener("click", onAppMenuClick);
  document.addEventListener("click", (e) => { if (!e.target.closest(".menu-wrap")) closeAppMenu(); });
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && appMenuOpen()) { closeAppMenu(); el("menuToggle").focus(); }
  });

  el("pieceSearch").addEventListener("input", (e) => { pieceFilter = e.target.value; render(); });
  // Effacer la recherche : croix (le focus reste dans le champ) ou Échap (quitte le champ).
  const clearPieceSearch = (leave) => {
    const s = el("pieceSearch");
    s.value = "";
    pieceFilter = "";
    render();
    if (leave) s.blur(); else s.focus();
  };
  el("pieceSearchClear").addEventListener("click", () => clearPieceSearch(false));
  el("pieceSearch").addEventListener("keydown", (e) => {
    if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); clearPieceSearch(true); }
  });
  el("lockAllBtn").addEventListener("click", (e) => { e.stopPropagation(); toggleAllLocks(); });
  el("lockMenu").addEventListener("click", (e) => {
    e.stopPropagation();
    const it = e.target.closest(".drop-item");
    if (it && !it.disabled) toggleLockGroup(it.dataset.scope);
  });
  // Sélection des verrous : se referme SANS appliquer au clic ailleurs (le flou), avec Échap, ou au
  // défilement du volet.
  document.addEventListener("click", closeLockMenu);
  document.addEventListener("keydown", (e) => { if (e.key === "Escape") closeLockMenu(); });
  window.addEventListener("scroll", closeLockMenu, { passive: true });

  for (const b of SETTING_BINDINGS) {
    const elm = el(b.id);
    if (!elm) continue;
    elm.addEventListener("change", async () => {
      let v;
      if (b.type === "checked") v = elm.checked;
      else if (b.type === "number") {
        const n = parseInt(elm.value, 10);
        v = b.fallback != null ? (n >= b.fallback ? n : b.fallback) : (n || 14); // numéro de départ : jamais < 1
        elm.value = v; // valeur retenue affichée (renderSettings saute le champ qui a le focus)
      }
      else if (b.type === "raw") v = elm.value;
      else v = elm.value;
      await changeSetting(b.path, v, b.uiOnly);
    });
  }
  el("appearanceBtn").addEventListener("click", sortByAppearance);
  el("undoBtn").addEventListener("click", doUndo);
  el("modal").addEventListener("click", async (e) => {
    if (warnModalOpen) {
      if (e.target === el("modal")) { closeWarningsModal(); return; } // clic sur le fond → ferme
      try { await onWarningsModalClick(e); } catch (err) { reportError(err); }
      return;
    }
    const b = e.target.closest("[data-mi]");
    if (b) { closeModal((modalButtons[+b.dataset.mi] || {}).value ?? null); return; }
    if (e.target === el("modal")) closeModal(null); // clic sur le fond flouté → ferme
  });
  document.addEventListener("keydown", (e) => { if (e.key === "Escape" && warnModalOpen) closeWarningsModal(); });
  // Pop-up de choix : Échap annule, Entrée dans le champ valide (bouton principal).
  document.addEventListener("keydown", (e) => {
    if (!modalResolve) return;
    if (e.key === "Escape") { e.preventDefault(); closeModal(null); }
    else if (e.key === "Enter" && e.target.classList && e.target.classList.contains("modal-input")) {
      e.preventDefault();
      const i = modalButtons.findIndex((b) => b.primary);
      closeModal(i >= 0 ? modalButtons[i].value : null);
    }
  });
  // Pastilles d'alerte : apparaissent quand le bloc de diagnostic passe SOUS le header (hors vue).
  if (window.IntersectionObserver) {
    const io = new IntersectionObserver((entries) => {
      diagVisible = entries[0].isIntersecting;
      renderPips();
    }, { threshold: 0, rootMargin: "-44px 0px 0px 0px" }); // -44px = hauteur du header sticky
    io.observe(el("diag"));
  }
  // Les alertes concernent nos pièces : depuis l'onglet adverse, on revient d'abord sur « Nos pièces ».
  const scrollToDiag = () => {
    if (isAdvMode()) setPaneSide("own");
    el("diag").scrollIntoView({ behavior: "smooth", block: "start" });
  };
  el("pipErr").addEventListener("click", scrollToDiag);
  el("pipWarn").addEventListener("click", scrollToDiag);

  el("diag").addEventListener("click", (e) => {
    const btn = e.target.closest("[data-act]");
    if (!btn) return;
    const act = btn.dataset.act;
    if (act === "merge") mergePiecesUI(btn.dataset.keep, (btn.dataset.others || "").split(",").filter(Boolean));
    else if (act === "sameprec") samePieceWithPrecision(btn.dataset.keep, btn.dataset.other).catch(reportError);
    else if (act === "separate") separateDuplicate(btn.dataset.id).catch(reportError);
    else if (act === "showpiece") showPieceInList(btn.dataset.id);
    else if (act === "togglediag") { diagCollapsed = !diagCollapsed; renderDiag(); }
    else if (act === "ignorewarn") ignoreWarning(btn.dataset.key);
    else if (act === "openwarnings") openWarningsModal();
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
    else if (act === "togglecomm") toggleCommunicated(id);
    else if (act === "addsub") { if (addSubFor === id) closeAddSub(); else openAddSub(id); }
    else if (act === "addsubok") submitAddSub().catch(reportError);
    else if (act === "addsubcancel") closeAddSub();
  });
  el("pieceList").addEventListener("keydown", (e) => {
    if (e.target.id !== "addSubName") return;
    if (e.key === "Enter") { e.preventDefault(); e.stopPropagation(); submitAddSub().catch(reportError); }
    else if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); closeAddSub(); }
  }, true);
  // Renommage : propagation automatique 750 ms après la dernière frappe (sans clic).
  el("pieceList").addEventListener("input", (e) => {
    const inp = e.target.closest('[data-act="rename"]');
    if (!inp) return;
    if (inp.tagName === "TEXTAREA") {
      // Un nom tient sur UNE ligne logique : un retour à la ligne collé devient une espace.
      if (/[\r\n]/.test(inp.value)) {
        const pos = inp.selectionStart;
        inp.value = inp.value.replace(/\r\n|[\r\n]/g, " ");
        inp.setSelectionRange(pos, pos);
      }
      autosizeName(inp);
    }
    clearTimeout(renameTimer);
    const id = inp.dataset.id, value = inp.value;
    renameTimer = setTimeout(() => livePropagateRename(id, value), 750);
  });
  // Instantané d'annulation pris quand on ENTRE dans un champ de nom (avant l'édition).
  el("pieceList").addEventListener("focusin", (e) => {
    if (e.target.closest('[data-act="rename"]')) renameSnap = snapshotModel();
  });
  el("pieceList").addEventListener("change", async (e) => {
    const inp = e.target.closest('[data-act="rename"]');
    if (!inp) return;
    clearTimeout(renameTimer);
    const snap = renameSnap; renameSnap = null;
    const snapPiece = snap && snap.pieces.concat(snap.adverse || []).find((p) => p.id === inp.dataset.id);
    const oldName = snapPiece ? (snapPiece.name || "") : null;
    const target = WP.findPiece(inp.dataset.id);
    if (target && target.container) {
      if (snap && inp.value.trim() !== (oldName || "")) commitUndo(snap, "nom du groupe");
      await renamePiece(inp.dataset.id, inp.value); // vide = groupe sans nom (pas de ligne au bordereau)
      return;
    }
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
    const inp = e.target.closest('input[data-act="setnum"], [data-act="rename"]');
    if (inp) { e.preventDefault(); inp.blur(); } // Entrée valide, y compris dans un nom affiché en entier
  });
  // Noms en entier : la largeur du volet change → la hauteur des noms aussi.
  let resizeTimer = null;
  window.addEventListener("resize", () => {
    clearTimeout(resizeTimer);
    resizeTimer = setTimeout(() => { if (namesWrap()) autosizeNames(); }, 120);
  });
}

// Champ du NOM d'une pièce ou d'un groupe : une ligne, ou zone de texte qui s'allonge (noms en entier).
function nameField(p, wrap, placeholder, title) {
  const extra = (placeholder ? ` placeholder="${WP.escapeHtml(placeholder)}"` : "") + (title ? ` title="${WP.escapeHtml(title)}"` : "");
  return wrap
    ? `<textarea class="piece-name" rows="1" spellcheck="false" data-act="rename" data-id="${p.id}"${extra}>${WP.escapeHtml(p.name || "")}</textarea>`
    : `<input class="piece-name" value="${WP.escapeHtml(p.name || "")}" data-act="rename" data-id="${p.id}"${extra} />`;
}

// Noms des pièces EN ENTIER (Réglages ▸ Volet LexRef ▸ Noms des pièces) : chaque nom est une zone de
// texte qui prend la hauteur de son contenu. Sinon, champ d'une ligne (les noms longs sont coupés).
function namesWrap() { return uiPrefs().pieceNames === "wrap"; }
function autosizeName(t) {
  t.style.height = "auto";
  t.style.height = t.scrollHeight + "px";
}
function autosizeNames() {
  el("pieceList").querySelectorAll("textarea.piece-name").forEach(autosizeName);
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
    WP.model.adverse = snap.adverse || [];
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
// input : { value, placeholder, label } → un champ de saisie ; la promesse rend alors { value, text }.
let modalHasInput = false;
function showModal({ title, msg, buttons, input }) {
  return new Promise((resolve) => {
    modalResolve = resolve;
    modalButtons = buttons || [];
    modalHasInput = !!input;
    el("modalCard").innerHTML =
      (title ? `<div class="modal-title">${title}</div>` : "") +
      `<div class="modal-msg">${msg}</div>` +
      (input ? `<label class="modal-field">${input.label ? `<span>${input.label}</span>` : ""}` +
        `<input class="modal-input" type="text" value="${WP.escapeHtml(input.value || "")}" placeholder="${WP.escapeHtml(input.placeholder || "")}" /></label>` : "") +
      `<div class="modal-btns">` +
      modalButtons.map((b, i) => `<button class="btn ${b.primary ? "primary" : "outline"}" data-mi="${i}">${b.label}</button>`).join("") +
      `</div>`;
    el("modal").classList.remove("hidden");
    const inp = el("modalCard").querySelector(".modal-input");
    if (inp) { inp.focus(); inp.select(); }
    else { const pb = el("modalCard").querySelector(".btn.primary"); if (pb) pb.focus(); }
  });
}
function closeModal(value) {
  const inp = el("modalCard").querySelector(".modal-input");
  const text = inp ? inp.value : "";
  el("modal").classList.add("hidden");
  if (document.activeElement && el("modal").contains(document.activeElement)) document.activeElement.blur();
  const r = modalResolve; modalResolve = null; modalButtons = [];
  const withInput = modalHasInput; modalHasInput = false;
  if (r) r(withInput ? (value == null ? null : { value, text }) : value);
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
