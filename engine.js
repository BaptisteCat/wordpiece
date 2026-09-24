/* global Office, Word */
// ============================================================
// WordPiece — moteur partagé (volet + menu contextuel)
// Expose window.WP. Aucune dépendance à l'UI.
// ============================================================
(function () {
  const TAG_PREFIX = "wp:piece:";
  const RANGE_PREFIX = "wp:range:"; // citation de plage : wp:range:<startId>~<endId>
  const ETSEQ_PREFIX = "wp:etseq:"; // citation « et suivantes » : wp:etseq:<startId>
  const LIST_PREFIX = "wp:list:";   // citation multiple : wp:list:<id1>~<id2>~<id3> (« Pièces n°1, 2 et 3 »)
  // PIÈCES ADVERSES (V9) : même mécanique, balises distinctes, liste séparée (model.adverse).
  const ADV_PREFIX = "wp:adv:";             // wp:adv:<id>[#complément]
  const ADV_RANGE_PREFIX = "wp:advrange:";  // wp:advrange:<startId>~<endId>
  const ADV_ETSEQ_PREFIX = "wp:advetseq:";  // wp:advetseq:<startId>
  const ADV_LIST_PREFIX = "wp:advlist:";    // wp:advlist:<id1>~<id2>…
  const BORDEREAU_TAG = "wp:bordereau";
  const MODEL_KEY = "wordpiece.model.v1";
  const LS_SETTINGS_KEY = "wordpiece.settings.v2";
  const DISABLED_KEY = "wordpiece.disabled.v1";

  let model = null;
  let stats = { counts: {}, orphanCCs: 0, numbers: new Map(), hasBordereau: false };

  // ---------------- Réglages / modèle ----------------
  function defaultSettings() {
    return {
      citationTemplate: "Pièce n°{num} : {nom}",
      citation: { bold: true, italic: true, underline: false, alignment: "right", newLine: true },
      // Pièces adverses : par défaut (sameAsOwn) elles reprennent le format et la mise en forme de nos
      // pièces, avec « adverse » ajouté au modèle. Décoché, ces réglages-ci s'appliquent.
      adverse: {
        sameAsOwn: true,
        template: "Pièce adverse n°{num} : {nom}",
        bold: true, italic: true, underline: false, alignment: "right", newLine: true,
      },
      rangeStyle: "stacked", // citation d'une plage /pA-B : "stacked" (une par ligne, défaut) ou "inline" (Pièces n°A à B)
      bordereau: {
        title: "Liste des pièces",
        titleBold: true, titleUnderline: false, titleSize: 12, titleAlign: "left",
        layout: "list",
        labelTemplate: "Pièce n°{num}",
        separator: " : ",
        labelBold: true, labelItalic: false, labelUnderline: true,
        nameBold: false, nameItalic: false, nameUnderline: false,
        listAlign: "left", listSize: 12, lineSpacing: "single", spaceBefore: false, spaceAfter: false,
      },
      // Préférences d'AFFICHAGE, propres au POSTE (voir docSettings/load) : elles ne
      // touchent jamais le document, donc les modifier ne déclenche aucune
      // resynchronisation (voir uiOnly côté volet).
      ui: {
        contextBar: true,   // bandeau « Curseur sur Pièce n°… » quand le curseur est sur une citation
        slashCommands: true, // commandes « /p… » tapées directement dans le texte
        adverseTab: true,   // onglet « Pièces adverses » dans le volet
        theme: "auto",      // auto (comme Word) | light | dark
        density: "standard", // compact | standard | comfortable
      },
    };
  }
  function defaultModel() {
    return { version: 3, settings: defaultSettings(), pieces: [], adverse: [], ignored: [] };
  }

  function mergeSettings(base, over) {
    const out = Array.isArray(base) ? base.slice() : { ...base };
    if (!over || typeof over !== "object") return out;
    for (const k of Object.keys(base)) {
      if (base[k] && typeof base[k] === "object" && !Array.isArray(base[k])) {
        out[k] = mergeSettings(base[k], over[k]);
      } else if (over[k] !== undefined) {
        out[k] = over[k];
      }
    }
    return out;
  }
  function normalizeLegacy(s) {
    if (!s || typeof s !== "object") return {};
    const out = { ...s };
    if (s.bordereauTitle && !s.bordereau) {
      out.bordereau = { title: s.bordereauTitle, layout: s.bordereauLayout || "list" };
    }
    return out;
  }

  function load() {
    const raw = Office.context.document.settings.get(MODEL_KEY);
    model = raw && typeof raw === "object" ? raw : defaultModel();
    if (!Array.isArray(model.pieces)) model.pieces = [];
    if (!Array.isArray(model.adverse)) model.adverse = []; // pièces adverses (V9)
    if (!Array.isArray(model.ignored)) model.ignored = []; // avertissements masqués par l'utilisateur
    // Réglages à DEUX couches : ceux du document s'ils existent, sinon les prefs globales (localStorage).
    const source = raw && raw.settings ? raw.settings : loadGlobal();
    model.settings = mergeSettings(defaultSettings(), normalizeLegacy(source));
    // Affichage : TOUJOURS les préférences du poste, même sur un document qui en porte
    // (actes enregistrés avant que ces réglages ne deviennent locaux).
    model.settings.ui = mergeSettings(defaultSettings().ui, (loadGlobal() || {}).ui);
    migratePieces(); // reprend les vieux docs au modèle structuré (pièces figées, non destructif)
    normalizeStructure(); // répare la hiérarchie pièces / sous-pièces (numéros affichés inchangés)
    onSide("adv", normalizeStructure);
  }

  // Migration : passe du modèle « chaîne libre » (champ number) au modèle structuré
  // (parentId/locked/fixedNumber). NON DESTRUCTIF : chaque pièce déjà numérotée est reprise
  // VERROUILLÉE à son numéro courant → rien ne bouge à l'ouverture d'un acte existant.
  // Ne PAS save() ici (n'altère pas le doc à l'ouverture) ; le champ legacy `number` est
  // retiré du modèle en mémoire et disparaîtra du doc au prochain enregistrement.
  function migratePieces() {
    const needs = model.pieces.some((p) => p.parentId === undefined && p.locked === undefined);
    if (needs) {
      // Ordre du tableau := ordre d'affichage actuel (tri naturel) pour un rendu identique après migration.
      model.pieces.sort((a, b) => naturalCompare(a.number, b.number));
      for (const p of model.pieces) {
        if (p.parentId === undefined) p.parentId = null;
        if (p.locked === undefined) {
          const num = p.number != null && p.number !== "" ? String(p.number) : "";
          p.locked = num !== "";
          p.fixedNumber = num !== "" ? num : null;
        }
      }
    }
    for (const p of model.pieces) delete p.number; // fixedNumber + calcul font désormais foi
  }

  // Recharge en rafraîchissant le cache (utile après une action d'un autre runtime).
  function reload() {
    return new Promise((resolve) => {
      try {
        Office.context.document.settings.refreshAsync(() => { load(); resolve(); });
      } catch (e) {
        load();
        resolve();
      }
    });
  }

  // Le document ne reçoit PAS les préférences d'affichage : un acte transmis à un confrère
  // lui imposerait sinon notre thème, notre densité et nos choix de volet. Elles vivent
  // uniquement dans le localStorage du poste (saveGlobal / loadGlobal).
  function docModel() {
    const s = { ...model.settings };
    delete s.ui;
    return { ...model, settings: s };
  }
  function save() {
    Office.context.document.settings.set(MODEL_KEY, docModel());
    return new Promise((resolve, reject) => {
      Office.context.document.settings.saveAsync((res) => {
        res.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(res.error);
      });
    });
  }

  // Désactivation PAR DOCUMENT (drapeau persistant, n'altère ni le modèle ni le texte).
  function isDisabled() {
    try { return Office.context.document.settings.get(DISABLED_KEY) === true; } catch (e) { return false; }
  }
  function setDisabled(v) {
    Office.context.document.settings.set(DISABLED_KEY, !!v);
    return new Promise((resolve, reject) => {
      Office.context.document.settings.saveAsync((res) => {
        res.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(res.error);
      });
    });
  }
  // Prefs GLOBALES (par poste) : localStorage uniquement. Les réglages d'un document donné
  // voyagent, eux, dans le document (document.settings) et priment au chargement.
  function saveGlobal() {
    try { localStorage.setItem(LS_SETTINGS_KEY, JSON.stringify(model.settings)); } catch (e) {}
  }
  function loadGlobal() {
    try { const r = localStorage.getItem(LS_SETTINGS_KEY); return r ? JSON.parse(r) : null; } catch (e) { return null; }
  }

  // ---------------- Compléments locaux de citation ----------------
  // Une citation peut porter une PRÉCISION propre à cet endroit du texte (« , page 8 », « (p. 12) »,
  // « — art. 5 »). Elle est stockée DANS LE TAG du contrôle : wp:piece:<id>#<complément>, donc elle
  // survit à la renumérotation et au renommage, et elle ne concerne QUE cette occurrence.
  // Le nom de la pièce (bordereau, autres citations) reste inchangé.
  const CITE_EXTRA_SEP = "#"; // les ids (« p_… ») ne contiennent jamais « # » : séparateur sûr
  // Un complément doit commencer par une ponctuation de renvoi ou un mot de localisation :
  // on ne confond donc pas « Contrat, page 8 » (complément) avec « Contrat de vente » (renommage).
  const CITE_EXTRA_RE = /^\s*(?:[,;:(\[\-–—]|p\.|pp\.|pages?\b|art\.|articles?\b|§|n°)/i;
  function looksLikeExtra(s) { return !!s && CITE_EXTRA_RE.test(s); }
  // Si `full` = `base` + un complément, renvoie ce complément (texte brut) ; sinon null.
  function extraAfter(base, full) {
    const b = String(base == null ? "" : base).trim();
    const f = String(full == null ? "" : full).trim();
    if (!b || f.length <= b.length) return null;
    if (f.slice(0, b.length).toLowerCase() !== b.toLowerCase()) return null;
    const rest = f.slice(b.length);
    return looksLikeExtra(rest) ? rest : null;
  }
  // Citation simple d'une pièce (nôtre « wp:piece: » ou adverse « wp:adv: ») ?
  function isSingleCiteTag(tag) {
    return !!tag && (tag.indexOf(TAG_PREFIX) === 0 || tag.indexOf(ADV_PREFIX) === 0);
  }
  // Découpe un tag « wp:piece:<id>[#<complément>] » ou « wp:adv:<id>[#<complément>] ».
  function splitPieceTag(tag) {
    const t = String(tag || "");
    const adv = t.indexOf(ADV_PREFIX) === 0;
    const raw = t.slice(adv ? ADV_PREFIX.length : TAG_PREFIX.length);
    const i = raw.indexOf(CITE_EXTRA_SEP);
    const side = adv ? "adv" : "own";
    return i < 0 ? { id: raw, extra: "", side } : { id: raw.slice(0, i), extra: raw.slice(i + 1), side };
  }
  function pieceIdFromTag(tag) { return splitPieceTag(tag).id; }
  function pieceTagFor(id, extra, side) {
    return sideCfg(side || sideOf(id)).cite + id + (extra ? CITE_EXTRA_SEP + extra : "");
  }
  // Le tag désigne-t-il une citation de CETTE pièce (avec ou sans complément) ?
  function isTagOfPiece(tag, id) {
    return isSingleCiteTag(tag) && splitPieceTag(tag).id === id;
  }

  // ---------------- Côtés : nos pièces / pièces adverses ----------------
  // Les deux listes partagent TOUTE la mécanique de numérotation (qui travaille sur model.pieces).
  // onSide("adv", fn) exécute fn SYNCHRONE avec la liste adverse à la place de model.pieces, puis
  // remet tout en place — sûr car JavaScript ne s'interrompt pas pendant un appel synchrone.
  // Ne JAMAIS y passer une fonction async (la bascule prendrait fin au premier await).
  const SIDES = {
    own: { key: "own", cite: TAG_PREFIX, range: RANGE_PREFIX, etseq: ETSEQ_PREFIX, list: LIST_PREFIX, one: "Pièce", many: "Pièces", token: "/p" },
    adv: { key: "adv", cite: ADV_PREFIX, range: ADV_RANGE_PREFIX, etseq: ADV_ETSEQ_PREFIX, list: ADV_LIST_PREFIX, one: "Pièce adverse", many: "Pièces adverses", token: "/pa" },
  };
  function sideCfg(side) { return SIDES[side === "adv" ? "adv" : "own"]; }
  function onSide(side, fn) {
    if (side !== "adv") return fn();
    const own = model.pieces;
    model.pieces = model.adverse;
    try { return fn(); }
    finally { model.adverse = model.pieces; model.pieces = own; }
  }
  function sideOf(id) {
    return model && Array.isArray(model.adverse) && model.adverse.some((p) => p.id === id) ? "adv" : "own";
  }
  function sidePieces(side) { return side === "adv" ? model.adverse : model.pieces; }

  // ---------------- Utilitaires ----------------
  function uid() { return "p_" + Date.now().toString(36) + Math.random().toString(36).slice(2, 7); }
  // Cherche dans les DEUX listes (les ids sont uniques) : renommer, citer, supprimer… marchent pour les deux.
  function findPiece(id) {
    return model.pieces.find((p) => p.id === id) || (model.adverse || []).find((p) => p.id === id);
  }
  // FABRIQUE unique : toute pièce du modèle naît ici (forme canonique + surcharges).
  function makePiece(over) {
    return { id: uid(), name: "", note: "", parentId: null, locked: false, fixedNumber: null, ...over };
  }

  // ---------------- Suivi des modifications ----------------
  // Toute ÉCRITURE dans le document passe par runUntracked : si le suivi des modifications est
  // activé, on le suspend le temps de nos écritures puis on le RESTAURE. Les renumérotations et
  // mises à jour mécaniques de WordPiece ne polluent donc jamais les révisions de l'utilisateur.
  // (API WordApi 1.4 ; si indisponible, on écrit normalement.)
  let trackingMode = null; // dernier mode observé ("Off" / "TrackAll" / "TrackMineOnly" / null = API inconnue)
  // opts.lazy : NE suspend PAS le suivi d'emblée. Le corps appelle context.wpUntracked() juste avant
  // de mettre en file sa 1re écriture (même lot, exécuté dans l'ordre). Une synchro qui n'a rien à
  // écrire ne touche donc jamais au mode de suivi → le document n'est pas marqué « modifié ».
  function runUntracked(body, opts) {
    const lazy = !!(opts && opts.lazy);
    return Word.run(async (context) => {
      let restore = null;
      context.wpUntracked = () => {};
      try {
        context.document.load("changeTrackingMode");
        await context.sync();
        trackingMode = context.document.changeTrackingMode || "Off";
        if (trackingMode !== "Off") {
          const mode = trackingMode;
          const suspend = () => {
            if (restore) return;
            restore = mode;
            context.document.changeTrackingMode = "Off";
          };
          if (lazy) context.wpUntracked = suspend;
          else { suspend(); await context.sync(); }
        }
      } catch (e) { trackingMode = null; }
      try {
        return await body(context);
      } finally {
        if (restore) {
          try { context.document.changeTrackingMode = restore; await context.sync(); } catch (e) {}
        }
      }
    });
  }
  // Ensemble d'API Word disponible ? (les écritures d'une API absente feraient échouer tout le lot)
  function apiOK(version) {
    try { return Office.context.requirements.isSetSupported("WordApi", version); } catch (e) { return false; }
  }
  function fillTemplate(tpl, num, name) {
    return (tpl || "")
      .replace(/\{num\}/gi, String(num))
      .replace(/\{nom\}/gi, name || "")
      .replace(/\{name\}/gi, name || "")
      .replace(/\s+/g, " ")
      .trim();
  }
  // Modèle de citation d'un côté. Adverse : dérivé du modèle réglé (« Pièce n°{num} : {nom} » →
  // « Pièce adverse n°{num} : {nom} ») pour garder la même présentation ; repli si le modèle ne
  // commence pas par « Pièce ».
  function citeTemplate(side) {
    const tpl = model.settings.citationTemplate || "Pièce n°{num}";
    if (side !== "adv") return tpl;
    const a = model.settings.adverse || {};
    if (a.sameAsOwn === false) return a.template || "Pièce adverse n°{num} : {nom}";
    const m = tpl.match(/^(\s*)(pi[eè]ce)(s?)(?=\s|$)/i);
    if (m) return m[1] + m[2] + m[3] + " adverse" + m[3] + tpl.slice(m[0].length);
    return "Pièce adverse n°{num} : {nom}";
  }
  // Mise en forme (gras / italique / souligné / alignement / ligne dédiée) du côté demandé.
  function citeSettings(side) {
    const a = model.settings.adverse || {};
    return side === "adv" && a.sameAsOwn === false ? a : model.settings.citation;
  }
  function formatCitation(num, name, side) {
    return fillTemplate(citeTemplate(side), num, name);
  }
  // Citations multiples (plages / « et suivantes ») — phrasé juridique français fixe.
  function formatRange(startNum, endNum, side) {
    const c = sideCfg(side);
    if (String(startNum) === String(endNum)) return c.one + " n°" + startNum;
    return c.many + " n°" + startNum + " à " + endNum;
  }
  function formatEtSeq(num, side) {
    return sideCfg(side).many + " n°" + num + " et suivantes";
  }
  // Citation multiple « Pièces n°1, 2 et 3 » (numéros triés ; « Pièces n°1 et 2 » à deux).
  function formatList(nums, side) {
    const c = sideCfg(side);
    const s = [...nums].sort(naturalCompare);
    if (s.length === 1) return c.one + " n°" + s[0];
    return c.many + " n°" + s.slice(0, -1).join(", ") + " et " + s[s.length - 1];
  }
  // Retrouve l'id de la pièce portant ce numéro (chaîne), via les numéros calculés (toujours frais :
  // les stats peuvent retarder d'une opération). Les GROUPES sans tête (conteneurs) sont ignorés,
  // sauf withContainers : on ne cite jamais un conteneur, mais on peut vouloir le retrouver.
  function pieceIdByNumber(numStr, withContainers) {
    const target = String(numStr).trim();
    for (const [id, n] of computeStructuredNumbers()) {
      if (String(n) !== target) continue;
      const p = findPiece(id);
      if (p && (withContainers || !p.container)) return id;
    }
    return null;
  }
  // Numéro AFFICHÉ d'une pièce (stats si disponibles, sinon calcul à la volée).
  function numberOf(id) {
    const side = sideOf(id);
    const map = side === "adv" ? stats.advNumbers : stats.numbers;
    const n = map && map.get(id);
    return n != null ? n : onSide(side, computeStructuredNumbers).get(id);
  }
  // Prochain numéro d'une pièce ADVERSE créée sans numéro (/pan, « Nouvelle pièce adverse ») : le
  // numéro entier qui suit le plus grand (les numéros adverses sont ceux de l'adversaire, jamais recalculés).
  function nextAdverseNumber() {
    const nums = onSide("adv", computeStructuredNumbers);
    let max = 0;
    for (const p of model.adverse) {
      if (p.parentId) continue;
      const m = String(nums.get(p.id) ?? "").match(/^(\d+)/);
      if (m) max = Math.max(max, parseInt(m[1], 10));
    }
    return String(max + 1);
  }
  function mapAlignment(a) {
    return { left: "Left", center: "Centered", right: "Right", justify: "Justified" }[a] || "Left";
  }
  // Interligne en points selon le mode choisi (null = ne pas toucher / interligne simple).
  function lineSpacingPts(mode, size) {
    if (mode === "1.5") return (size || 12) * 1.5;
    if (mode === "double") return (size || 12) * 2;
    return null;
  }
  function escapeHtml(s) {
    return String(s || "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  }
  // Normalise un nom pour comparaison (accents/casse/espaces ignorés) — détection de conflits/doublons.
  function normScan(s) {
    return String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().trim().replace(/\s+/g, " ");
  }
  // Clé de comparaison d'un NOM : sans accents, casse, guillemets ni ponctuation.
  function nameKey(s) {
    return normScan(s).replace(/[«»“”"'’‘`]/g, " ").replace(/[.,;:!?()[\]/\\\-–—_]/g, " ").replace(/\s+/g, " ").trim();
  }
  function editDistance(a, b) {
    let prev = Array.from({ length: b.length + 1 }, (_, j) => j);
    for (let i = 1; i <= a.length; i++) {
      const cur = [i];
      for (let j = 1; j <= b.length; j++) {
        cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
      }
      prev = cur;
    }
    return prev[b.length];
  }
  const NAME_STOP = new Set(["de", "du", "des", "la", "le", "les", "et", "en", "au", "aux", "un", "une", "sur", "par", "pour", "avec"]);
  // Deux libellés portant le MÊME numéro désignent-ils vraisemblablement la même pièce ?
  // Oui si identiques à la ponctuation/casse/accents près, si l'un prolonge l'autre (« Contrat de bail »
  // / « Contrat de bail commercial »), si l'écart est une faute de frappe, ou s'ils partagent l'essentiel
  // de leurs mots. Des nombres différents (dates, montants) signent en revanche deux pièces distinctes.
  function sameDocName(a, b) {
    const x = nameKey(a), y = nameKey(b);
    if (!x || !y || x === y) return true;
    if (x.startsWith(y + " ") || y.startsWith(x + " ")) return true;
    const digits = (s) => (s.match(/\d+/g) || []).sort().join(" ");
    if (digits(x) !== digits(y)) return false;
    const long = Math.max(x.length, y.length);
    if (long >= 8 && editDistance(x, y) / long <= 0.15) return true;
    const toks = (s) => new Set(s.split(" ").filter((w) => w.length > 1 && !NAME_STOP.has(w)));
    const A = toks(x), B = toks(y);
    if (!A.size || !B.size) return false;
    let inter = 0;
    for (const w of A) if (B.has(w)) inter++;
    return (2 * inter) / (A.size + B.size) >= 0.6;
  }
  // Si `full` = `base` + une suite (« , page 8 », « commercial », « . »), renvoie cette suite telle
  // quelle ; sinon null. Sert à conserver le texte exact d'une variante citée.
  function nameRemainder(base, full) {
    const b = String(base || "").trim(), f = String(full || "").trim();
    if (!b || f.length <= b.length) return null;
    if (f.slice(0, b.length).toLowerCase() !== b.toLowerCase()) return null;
    const rest = f.slice(b.length);
    return /^[\s,;:.(\[\-–—]/.test(rest) ? rest : null;
  }

  // id -> numéro (entier en mode auto, chaîne libre en mode manuel).
  // Reflow d'un groupe de FRÈRES (top-level ou sous-pièces d'un même parent), dans l'ordre du tableau.
  // Les pièces verrouillées gardent fixedNumber ; les coulantes consomment 1,2,3,… en EXCLUANT les
  // entiers verrouillés. `prefix` = "" (top-level) ou "P." (sous-pièces de la pièce affichant P).
  function reflowGroup(pieces, prefix) {
    const map = new Map();
    const lockedInts = new Set();
    for (const p of pieces) {
      if (p.fixedNumber != null) {
        map.set(p.id, prefix + p.fixedNumber);
        if (/^\d+$/.test(String(p.fixedNumber))) lockedInts.add(parseInt(p.fixedNumber, 10));
      }
    }
    let n = 0;
    const nextFree = () => { do { n++; } while (lockedInts.has(n)); return n; };
    for (const p of pieces) {
      if (!(p.fixedNumber != null)) map.set(p.id, prefix + nextFree());
    }
    return map;
  }

  // Numérotation structurée : top-level par ordre du tableau + reflow ; sous-pièces = parent.k.
  function computeStructuredNumbers() {
    const tops = model.pieces.filter((p) => !p.parentId);
    const map = reflowGroup(tops, "");
    for (const parent of tops) {
      const kids = model.pieces.filter((k) => k.parentId === parent.id);
      if (!kids.length) continue;
      const sub = reflowGroup(kids, map.get(parent.id) + ".");
      for (const [id, v] of sub) map.set(id, v);
    }
    return map;
  }

  // ---- Résolveurs structurés (utilisés par les commandes /pn…) ----
  function topPieces() { return model.pieces.filter((p) => !p.parentId); }
  function childrenOf(parentId) { return model.pieces.filter((k) => k.parentId === parentId); }
  function lockedIntSet(pieces) {
    const s = new Set();
    for (const p of pieces) if (p.fixedNumber != null && /^\d+$/.test(String(p.fixedNumber))) s.add(parseInt(p.fixedNumber, 10));
    return s;
  }
  // k-ième entier libre (hors verrouillés) — sert au « prochain » numéro d'un groupe de frères.
  function nthFreeInt(lockedInts, k) {
    let n = 0, got = 0;
    while (got < k) { n++; if (!lockedInts.has(n)) got++; }
    return n;
  }
  // Numéro qu'obtiendrait une NOUVELLE pièce coulante ajoutée en fin de top-level.
  function nextTopNumber() {
    const tops = topPieces();
    const flow = tops.filter((p) => !(p.fixedNumber != null)).length;
    return nthFreeInt(lockedIntSet(tops), flow + 1);
  }
  // Indice de sous-pièce qu'obtiendrait un NOUVEL enfant coulant ajouté en fin de fratrie.
  function nextSubIndex(parentId) {
    const kids = childrenOf(parentId);
    const flow = kids.filter((p) => !(p.fixedNumber != null)).length;
    return nthFreeInt(lockedIntSet(kids), flow + 1);
  }
  // Insère une pièce dans model.pieces juste AVANT beforeId (ou en fin si null).
  function insertPieceBefore(piece, beforeId) {
    if (beforeId == null) { model.pieces.push(piece); return; }
    const idx = model.pieces.findIndex((p) => p.id === beforeId);
    if (idx < 0) model.pieces.push(piece); else model.pieces.splice(idx, 0, piece);
  }
  // Id de la 1re pièce APRÈS le bloc [parent + ses enfants] (pour ajouter un enfant en fin de fratrie).
  function afterChildrenBlockId(parentId) {
    const idx = model.pieces.findIndex((p) => p.id === parentId);
    if (idx < 0) return null;
    let j = idx + 1;
    while (j < model.pieces.length && model.pieces[j].parentId === parentId) j++;
    return j < model.pieces.length ? model.pieces[j].id : null;
  }

  // ---- Invariants de STRUCTURE (2 niveaux : pièces / sous-pièces) ----
  // Répare et range le modèle après toute opération qui touche à la hiérarchie :
  //  1. sous-pièce dont le parent a disparu → le groupe est recréé (conteneur, même id) : ses
  //     sous-pièces gardent leur numéro N.k au lieu de n'en avoir plus aucun (« ? ») ;
  //  2. jamais de 3e niveau : une sous-sous-pièce remonte dans le groupe de son grand-parent ;
  //  3. « fausse » sous-pièce = pièce de 1er niveau épinglée « N.k » (scan, ancienne saisie) →
  //     vraie sous-pièce k du groupe N (créé au besoin). Même numéro affiché, mais la structure
  //     est cohérente : /pnN., les verrous et les décalages traitent enfin le groupe entier ;
  //  4. groupe (conteneur) vide → supprimé (il occupait un numéro invisible) ;
  //  5. contiguïté : chaque parent est immédiatement suivi de ses sous-pièces.
  function normalizeStructure() {
    const ids = new Map(model.pieces.map((p) => [p.id, p]));
    // 1. parents disparus
    for (const p of model.pieces.slice()) {
      if (p.parentId === p.id) p.parentId = null;
      if (p.parentId && !ids.has(p.parentId)) {
        const c = makePiece({ id: p.parentId, container: true });
        model.pieces.splice(model.pieces.indexOf(p), 0, c);
        ids.set(c.id, c);
      }
    }
    // 2. deux niveaux maximum (un groupe est toujours de 1er niveau)
    for (const p of model.pieces) if (p.container) p.parentId = null;
    for (const p of model.pieces) {
      let par = p.parentId ? ids.get(p.parentId) : null;
      for (let guard = 0; par && par.parentId && guard < 10; guard++) {
        p.parentId = par.parentId;
        par = ids.get(p.parentId);
      }
    }
    // 3. fausses sous-pièces « N.k » → vraies sous-pièces du groupe N
    for (const p of model.pieces.slice()) {
      if (p.parentId || p.container || p.fixedNumber == null) continue;
      const m = String(p.fixedNumber).trim().match(/^(\d+)\.(\d+)$/);
      if (!m) continue;
      if (model.pieces.some((k) => k.parentId === p.id)) continue; // a déjà ses propres sous-pièces
      const nums = computeStructuredNumbers();
      const tops = model.pieces.filter((t) => t !== p && !t.parentId && String(nums.get(t.id)) === m[1]);
      let g = tops.find((t) => t.container) || tops[0];
      if (!g) {
        // Aucun n°N : le groupe est créé ÉPINGLÉ à N (personne n'affichait N → rien ne bouge).
        g = makePiece({ container: true, fixedNumber: m[1], locked: !!p.locked, scanned: !!p.scanned });
        model.pieces.splice(model.pieces.indexOf(p), 0, g);
      }
      p.parentId = g.id;
      p.fixedNumber = m[2];
    }
    // 4. groupes vides
    const hasKids = new Set(model.pieces.filter((p) => p.parentId).map((p) => p.parentId));
    model.pieces = model.pieces.filter((p) => !p.container || hasKids.has(p.id));
    // 5. contiguïté parent → sous-pièces (ordre relatif conservé ; aucune pièce n'est jamais perdue)
    const out = [];
    const seen = new Set();
    for (const t of model.pieces) {
      if (t.parentId) continue;
      out.push(t); seen.add(t);
      for (const k of model.pieces) if (k.parentId === t.id) { out.push(k); seen.add(k); }
    }
    for (const p of model.pieces) if (!seen.has(p)) out.push(p);
    model.pieces = out;
  }

  // Retire des pièces du modèle SANS casser la hiérarchie : une pièce qui a encore des sous-pièces
  // devient un simple groupe (conteneur) — ses sous-pièces gardent leur numéro N.k ; un groupe
  // devenu vide disparaît.
  function removeFromModel(idList) {
    const rm = new Set(idList);
    for (const p of model.pieces) {
      if (!rm.has(p.id) || p.container) continue;
      if (model.pieces.some((k) => k.parentId === p.id && !rm.has(k.id))) {
        rm.delete(p.id);
        p.container = true;
        p.name = "";
        delete p.scanned;
        delete p.recovered;
      }
    }
    model.pieces = model.pieces.filter((p) => !rm.has(p.id));
    normalizeStructure();
  }

  // La pièce `p` devient la TÊTE du groupe `container` : elle prend sa place dans la liste (donc
  // son numéro), les sous-pièces lui sont rattachées et le conteneur disparaît.
  function takeOverContainer(container, p) {
    const pi = model.pieces.indexOf(p);
    if (pi >= 0) model.pieces.splice(pi, 1);
    model.pieces.splice(model.pieces.indexOf(container), 1, p);
    for (const k of model.pieces) if (k.parentId === container.id && k !== p) k.parentId = p.id;
    p.parentId = null;
  }

  // Range une pièce de 1er niveau (déjà retirée de la liste) avant la 1re pièce de 1er niveau dont le
  // numéro affiché dépasse `num` : l'ordre reste cohérent si elle est plus tard déverrouillée.
  function insertTopByNumber(piece, num, nums) {
    const next = model.pieces.find((t) => !t.parentId && t !== piece && nums.has(t.id) && naturalCompare(nums.get(t.id), num) > 0);
    insertPieceBefore(piece, next ? next.id : null);
  }

  // Crée une pièce au numéro EXPLICITE demandé (« 7 », « 2 bis », « 4.1 »), verrouillée à ce numéro :
  //  - « N.k » → sous-pièce k du groupe N (groupe créé au besoin ; une pièce N existante sert de tête) ;
  //  - N est un groupe sans tête → la nouvelle pièce en devient la tête (le groupe garde son numéro) ;
  //  - sinon → nouvelle pièce de 1er niveau épinglée à N.
  function createPieceAt(numStr) {
    const num = String(numStr).trim();
    const sub = num.match(/^(\d+)\.(\d+)$/);
    if (sub) {
      let g = findPiece(pieceIdByNumber(sub[1], true));
      if (g && g.parentId) g = null;
      if (!g) {
        const flowing = String(nextTopNumber()) === sub[1];
        g = makePiece({ container: true, fixedNumber: flowing ? null : sub[1] });
        model.pieces.push(g);
      }
      const child = makePiece({ parentId: g.id, locked: true, fixedNumber: sub[2] });
      insertPieceBefore(child, afterChildrenBlockId(g.id));
      return child;
    }
    const piece = makePiece({ locked: true, fixedNumber: num });
    const c = findPiece(pieceIdByNumber(num, true));
    if (c && c.container) takeOverContainer(c, piece);
    else model.pieces.push(piece);
    return piece;
  }

  // Place une pièce EXISTANTE au numéro saisi par l'utilisateur, en changeant de niveau si besoin :
  //  - « N.k » : devient (ou reste) la sous-pièce k du groupe N — groupe créé si N n'existe pas ;
  //    « 4 » → « 4.k » transforme la pièce 4 en groupe dont elle est une sous-pièce ;
  //  - numéro de 1er niveau : une sous-pièce est PROMUE pièce ; si ce numéro est un groupe sans tête,
  //    la pièce en devient la tête (ex. la 4.1 restée seule retapée « 4 » redevient la pièce 4).
  // Le numéro saisi est garanti : la pièce est épinglée et verrouillée ; si le groupe cible a glissé
  // (pièces coulantes resserrées par le déplacement), il est épinglé à son numéro.
  function placeAt(p, v) {
    const nums = computeStructuredNumbers(); // numéros tels que l'utilisateur les voit
    const hasKids = model.pieces.some((k) => k.parentId === p.id);
    const tops = topPieces();
    const detach = () => { const i = model.pieces.indexOf(p); if (i >= 0) model.pieces.splice(i, 1); };
    const sub = v.match(/^(\d+)\.(\d+)$/);
    if (sub) {
      if (p.container || hasKids) return { error: "hasSubs" };
      const G = sub[1];
      const cands = tops.filter((t) => String(nums.get(t.id)) === G);
      let g = cands.find((t) => t.container) || cands[0] || null;
      if (g === p) {
        // « 4 » → « 4.k » : un groupe prend la place (et le numéro) de la pièce, qui y entre.
        g = makePiece({ container: true, fixedNumber: p.fixedNumber });
        insertPieceBefore(g, p.id);
        p.parentId = g.id;
      } else if (!(g && p.parentId === g.id)) {
        detach();
        if (!g) {
          g = makePiece({ container: true, fixedNumber: G });
          insertTopByNumber(g, G, nums);
        }
        p.parentId = g.id;
        insertPieceBefore(p, afterChildrenBlockId(g.id));
      }
      p.fixedNumber = sub[2];
      p.locked = true;
      normalizeStructure();
      if (findPiece(g.id) && String(computeStructuredNumbers().get(g.id)) !== G) g.fixedNumber = G;
      return { ok: true };
    }
    const cont = tops.find((t) => t !== p && t.container && String(nums.get(t.id)) === v);
    if (cont) {
      if (p.container || hasKids) return { error: "groupTaken", num: v };
      takeOverContainer(cont, p);
    } else if (p.parentId) {
      detach();
      p.parentId = null;
      insertTopByNumber(p, v, nums);
    }
    p.fixedNumber = v;
    p.locked = true;
    normalizeStructure();
    return { ok: true };
  }
  // Comparaison « naturelle » de numéros libres : 1 < 1 bis < 2 < 10 ; gère 1.2.4, 3.2, 2 bis…
  function naturalCompare(a, b) {
    const ax = String(a == null ? "" : a).match(/(\d+|\D+)/g) || [];
    const bx = String(b == null ? "" : b).match(/(\d+|\D+)/g) || [];
    const n = Math.max(ax.length, bx.length);
    for (let i = 0; i < n; i++) {
      const an = ax[i], bn = bx[i];
      if (an === undefined) return -1;
      if (bn === undefined) return 1;
      const aIsNum = /^\d+$/.test(an), bIsNum = /^\d+$/.test(bn);
      if (aIsNum && bIsNum) {
        const d = parseInt(an, 10) - parseInt(bn, 10);
        if (d !== 0) return d;
      } else if (an !== bn) {
        return an < bn ? -1 : 1;
      }
    }
    return 0;
  }

  // Tri par NUMÉRO (ordre numérique naturel), en manuel comme en auto.
  function orderedPieces(numbers) {
    return [...model.pieces].sort((a, b) => naturalCompare(numbers.get(a.id), numbers.get(b.id)));
  }

  // ---------------- Bordereau (HTML) ----------------
  function wrapFmt(text, fmt) {
    let t = text;
    if (fmt.underline) t = `<u>${t}</u>`;
    if (fmt.italic) t = `<i>${t}</i>`;
    if (fmt.bold) t = `<b>${t}</b>`;
    return t;
  }
  function applyFont(range, fmt) {
    range.font.bold = !!fmt.bold;
    range.font.italic = !!fmt.italic;
    range.font.underline = fmt.underline ? "Single" : "None";
  }

  // Le bordereau ne doit dépendre QUE des réglages de la section Bordereau : un paragraphe inséré
  // hérite sinon du dernier paragraphe du document (puces, retraits, alignement, gras…).
  // → style Normal, aucun retrait ; la sortie de liste se fait après coup (detachBordereauLists).
  function resetBordPara(para) {
    if (apiOK("1.3")) para.styleBuiltIn = "Normal";
    para.leftIndent = 0;
    para.firstLineIndent = 0;
    para.rightIndent = 0;
  }
  // Police d'un morceau du bordereau : aucun attribut hérité (barré, exposant, surlignage…) ;
  // police et couleur = celles du style Normal du document quand on sait les lire.
  function resetBordFont(range, normalFont) {
    const f = range.font;
    f.strikeThrough = false;
    f.doubleStrikeThrough = false;
    f.superscript = false;
    f.subscript = false;
    f.highlightColor = null;
    if (normalFont && normalFont.name) f.name = normalFont.name;
    if (normalFont && normalFont.color) f.color = normalFont.color;
  }
  async function loadNormalFont(context) {
    if (!apiOK("1.5")) return null;
    try {
      const st = context.document.getStyles().getByNameOrNullObject("Normal");
      st.load("isNullObject");
      st.font.load("name,color");
      await context.sync();
      return st.isNullObject ? null : { name: st.font.name, color: st.font.color };
    } catch (e) { return null; }
  }
  async function detachBordereauLists(context, bord) {
    if (!apiOK("1.3")) return;
    try {
      const ps = bord.paragraphs;
      ps.load("items/isListItem");
      await context.sync();
      let n = 0;
      for (const p of ps.items) if (p.isListItem) { p.detachFromList(); n++; }
      if (n) await context.sync();
    } catch (e) { /* au mieux : ne jamais casser la synchro */ }
  }

  // Bordereau « liste » avec chaque NOM dans son propre content control (wp:name:<id>),
  // ce qui le rend éditable et détectable pour le renommage inline.
  function fillBordereauListEditable(bord, numbers, normalFont) {
    const s = model.settings.bordereau;
    const ordered = orderedPieces(numbers);
    bord.clear();
    const title = bord.insertParagraph(s.title || "Bordereau de pièces", "Start");
    resetBordPara(title);
    resetBordFont(title, normalFont);
    title.font.italic = false;
    title.font.bold = !!s.titleBold;
    title.font.size = s.titleSize || 12;
    title.font.underline = s.titleUnderline ? "Single" : "None";
    title.alignment = mapAlignment(s.titleAlign || "left");

    const labelFmt = { bold: s.labelBold, italic: s.labelItalic, underline: s.labelUnderline };
    const nameFmt = { bold: s.nameBold, italic: s.nameItalic, underline: s.nameUnderline };
    const listSize = s.listSize || 12; // taille des pièces, INDÉPENDANTE du titre
    const listAlign = mapAlignment(s.listAlign || "left");
    const spBefore = s.spaceBefore ? 6 : 0;
    const spAfter = s.spaceAfter ? 6 : 0;
    const lineSp = lineSpacingPts(s.lineSpacing, listSize);

    for (const p of ordered) {
      if (p.container) continue; // conteneur de groupe : pas de ligne propre (seules ses sous-pièces)
      const num = numbers.get(p.id);
      const para = bord.insertParagraph("", "End");
      resetBordPara(para);
      para.alignment = listAlign;
      if (p.parentId) { try { para.leftIndent = 18; } catch (e) {} } // sous-pièce : décalée sous son parent
      para.spaceBefore = spBefore;
      para.spaceAfter = spAfter;
      if (lineSp != null) para.lineSpacing = lineSp;
      const sep = p.name ? (s.separator ?? " : ") : "";
      const trail = sep.match(/\s*$/)[0]; // espace(s) de fin du séparateur
      const sepMain = sep.slice(0, sep.length - trail.length);
      // Étiquette + séparateur SANS l'espace final → mise en forme étiquette (souligné s'arrête au « : »).
      const labelRange = para.insertText(fillTemplate(s.labelTemplate, num, p.name) + sepMain, "End");
      resetBordFont(labelRange, normalFont);
      applyFont(labelRange, labelFmt);
      labelRange.font.size = listSize;
      // L'espace qui sépare du nom n'est ni souligné ni gras/italique.
      if (trail) {
        const tRange = para.insertText(trail, "End");
        resetBordFont(tRange, normalFont);
        tRange.font.underline = "None";
        tRange.font.bold = false;
        tRange.font.italic = false;
        tRange.font.size = listSize;
      }
      if (p.name) {
        const nameRange = para.insertText(p.name, "End");
        resetBordFont(nameRange, normalFont);
        applyFont(nameRange, nameFmt);
        nameRange.font.size = listSize;
        const nameCC = nameRange.insertContentControl();
        nameCC.tag = "wp:name:" + p.id;
        nameCC.title = "Nom pièce " + num;
        nameCC.appearance = "Hidden"; // le nom reste éditable mais sans encadré
      }
    }
  }

  function escapeRegex(s) {
    return String(s || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }

  // Extrait le NUMÉRO tapé dans une citation (ex. « Pièce n°7 » → 7), selon le modèle.
  function extractNumberFromCitation(text, side) {
    const tpl = citeTemplate(side);
    if (!/\{num\}/i.test(tpl)) return null;
    let pattern = "";
    for (const tk of tpl.split(/(\{num\}|\{nom\}|\{name\})/i)) {
      if (/^\{num\}$/i.test(tk)) pattern += "(.+?)"; // numéro libre (2 bis, 1.2.4, …)
      else if (/^\{nom\}$/i.test(tk) || /^\{name\}$/i.test(tk)) pattern += "(?:.*?)";
      else pattern += escapeRegex(tk).replace(/\s+/g, "\\s+");
    }
    const m = (text || "").trim().match(new RegExp("^\\s*" + pattern + "\\s*$", "i"));
    return m && m[1] != null ? m[1].trim() : null;
  }

  // Extrait le nom depuis le texte d'une citation, selon le modèle de citation.
  // Renvoie null si le modèle n'affiche pas {nom} (rien à éditer dans le texte).
  function extractNameFromCitation(text, pieceId, side) {
    const tpl = model.settings.citationTemplate ? citeTemplate(side || sideOf(pieceId)) : "";
    if (!/\{nom\}/i.test(tpl)) return null;
    // On construit un motif à partir du modèle, SANS dépendre du numéro courant
    // (au rechargement, stats.numbers n'est pas encore peuplé → l'ancienne
    //  approche par préfixe échouait et doublait la citation).
    let pattern = "";
    for (const tk of tpl.split(/(\{num\}|\{nom\}|\{name\})/i)) {
      if (/^\{num\}$/i.test(tk)) pattern += "(?:.+?)"; // n'importe quel numéro libre
      else if (/^\{nom\}$/i.test(tk) || /^\{name\}$/i.test(tk)) pattern += "(.*)"; // le nom (glouton)
      else pattern += escapeRegex(tk).replace(/\s+/g, "\\s+");
    }
    const m = (text || "").trim().match(new RegExp("^\\s*" + pattern + "\\s*$", "i"));
    return m && m[1] != null ? m[1].trim() : null;
  }

  function buildBordereauHtml(numbers) {
    const s = model.settings.bordereau;
    const ordered = orderedPieces(numbers);
    const titleHtml = wrapFmt(escapeHtml(s.title || "Bordereau de pièces"), { bold: s.titleBold, underline: s.titleUnderline });
    const cssAlign = { left: "left", center: "center", right: "right", justify: "justify" }[s.titleAlign || "left"] || "left";
    let html = `<p style="font-size:${s.titleSize || 12}pt; text-align:${cssAlign}; margin:0 0 8pt 0;">${titleHtml}</p>`;
    const labelFmt = { bold: s.labelBold, italic: s.labelItalic, underline: s.labelUnderline };
    const nameFmt = { bold: s.nameBold, italic: s.nameItalic, underline: s.nameUnderline };
    if (s.layout === "table") {
      html += `<table border="1" cellspacing="0" cellpadding="4" style="border-collapse:collapse;">`;
      html += `<tr><td>${wrapFmt("N°", { bold: true })}</td><td>${wrapFmt("Intitulé", { bold: true })}</td></tr>`;
      for (const p of ordered) {
        if (p.container) continue;
        const num = wrapFmt(escapeHtml(String(numbers.get(p.id) ?? "")), labelFmt);
        const name = wrapFmt(escapeHtml(p.name || ""), nameFmt);
        const pad = p.parentId ? ' style="padding-left:18pt"' : "";
        html += `<tr><td${pad}>${num}</td><td>${name}</td></tr>`;
      }
      html += `</table>`;
    } else {
      for (const p of ordered) {
        if (p.container) continue;
        const label = wrapFmt(escapeHtml(fillTemplate(s.labelTemplate, numbers.get(p.id), p.name)), labelFmt);
        const name = p.name ? wrapFmt(escapeHtml(p.name), nameFmt) : "";
        const sep = name ? escapeHtml(s.separator ?? " : ") : "";
        const indent = p.parentId ? "margin-left:18pt; " : "";
        html += `<p style="${indent}margin-top:0; margin-bottom:3pt;">${label}${sep}${name}</p>`;
      }
    }
    return html;
  }

  // ---------------- Synchronisation document ----------------
  // Sérialise les synchronisations (jamais deux en parallèle) et réessaie UNE fois si une
  // référence de contrôle a été invalidée par une opération concurrente (GeneralException).
  // opts.force : réécrit tout même à l'identique (après un changement de réglages de mise en forme).
  let syncChain = Promise.resolve();
  function sync(opts) {
    const attempt = async () => {
      try { return await doSync(opts); }
      catch (e) { return await doSync(opts); } // 2e essai : rechargement frais (exclut le contrôle disparu)
    };
    syncChain = syncChain.then(attempt, attempt);
    return syncChain;
  }

  const CITE_PREFIXES = [TAG_PREFIX, RANGE_PREFIX, ETSEQ_PREFIX, LIST_PREFIX, ADV_PREFIX, ADV_RANGE_PREFIX, ADV_ETSEQ_PREFIX, ADV_LIST_PREFIX];
  const isCiteTag = (t) => !!t && CITE_PREFIXES.some((p) => t.startsWith(p));
  // Signature du bordereau tel qu'il DEVRAIT être rendu : tant qu'elle ne change pas, on ne le
  // reconstruit pas (synchro différentielle). Mémorisée AUSSI dans le document (écrite seulement
  // après une reconstruction, qui modifie déjà le document) : à la réouverture d'un acte à jour,
  // le bordereau n'est pas reconstruit et Word ne demande pas d'enregistrer.
  const BORD_SIG_KEY = "wordpiece.bordSig";
  let lastBordSig = null;
  function storedBordSig() {
    try { return Office.context.document.settings.get(BORD_SIG_KEY) || null; } catch (e) { return null; }
  }
  function storeBordSig(sig) {
    try {
      if (Office.context.document.settings.get(BORD_SIG_KEY) === sig) return;
      Office.context.document.settings.set(BORD_SIG_KEY, sig);
      Office.context.document.settings.saveAsync(() => {});
    } catch (e) { /* sans effet : reconstruction à la prochaine session */ }
  }
  async function doSync(opts) {
    const force = !!(opts && opts.force);
    await runUntracked(async (context) => {
      const ccs = context.document.body.contentControls;
      // text : auto-réparation + SYNCHRO DIFFÉRENTIELLE (on ne réécrit que ce qui change) ;
      // color : évite de re-marquer un orphelin déjà signalé.
      ccs.load("items/tag,items/text,items/color");
      await context.sync();

      // Détecte les citations IMBRIQUÉES (un contrôle dans un autre, dû à un double-emballage).
      // On les SAUTE : traiter le contrôle externe (insertText "Replace") supprime l'interne →
      // référence périmée → GeneralException qui ferait planter toute la synchro. En sautant
      // l'interne, l'externe l'absorbe proprement (auto-réparation).
      const citeAll = ccs.items.filter((cc) => isCiteTag(cc.tag));
      const parents = citeAll.map((cc) => cc.parentContentControlOrNullObject);
      parents.forEach((p) => p.load("tag,isNullObject"));
      await context.sync();
      const nested = new Set();
      citeAll.forEach((cc, i) => { if (!parents[i].isNullObject && isCiteTag(parents[i].tag)) nested.add(cc); });

      const pieceCCs = ccs.items.filter((cc) => cc.tag && cc.tag.startsWith(TAG_PREFIX) && !nested.has(cc));
      const appearanceOrderIds = [];
      for (const cc of pieceCCs) {
        const id = pieceIdFromTag(cc.tag);
        if (!appearanceOrderIds.includes(id)) appearanceOrderIds.push(id);
      }

      // ---- AUTO-RÉPARATION (anti-perte de travail) ----------------------------------------
      // Si des citations pointent vers des pièces ABSENTES du modèle (typiquement : réglages du
      // document perdus après un crash/récupération de Word, alors que le corps du texte, lui, est
      // restauré), on RECONSTRUIT ces pièces depuis le document lui-même — au lieu de détruire le
      // texte. Sources : le texte des citations (« Pièce n°X : Nom ») puis les noms du bordereau.
      try {
        let healed = 0;
        for (const cc of pieceCCs) {
          const id = pieceIdFromTag(cc.tag);
          if (findPiece(id)) continue;
          const num = extractNumberFromCitation(cc.text);
          const name = extractNameFromCitation(cc.text, id);
          if ((num == null || num === "") && (name == null || name === "")) continue; // texte inexploitable
          const fixed = num != null && String(num).trim() !== "" ? String(num).trim() : null;
          model.pieces.push(makePiece({ id, name: (name || "").trim(), locked: fixed != null, fixedNumber: fixed, recovered: true }));
          healed++;
        }
        // Pièces présentes UNIQUEMENT au bordereau (jamais citées dans le corps). On NE fait ce
        // 2e passage (coûteux : un aller-retour Word supplémentaire) QUE si une perte est probable
        // (on a déjà réparé une citation, ou le modèle est totalement vide) — jamais en régime normal.
        const bordForHeal = (healed > 0 || model.pieces.length === 0)
          ? ccs.items.find((c) => c.tag === BORDEREAU_TAG) : null;
        if (bordForHeal) {
          const nn = bordForHeal.contentControls;
          nn.load("items/tag,items/text");
          await context.sync();
          for (const nc of nn.items) {
            if (!nc.tag || !nc.tag.startsWith("wp:name:")) continue;
            const id = nc.tag.slice("wp:name:".length);
            if (findPiece(id)) continue;
            const name = (nc.text || "").trim();
            if (!name) continue;
            model.pieces.push(makePiece({ id, name, recovered: true }));
            healed++;
          }
        }
        // Même filet pour les pièces ADVERSES citées mais absentes du modèle.
        for (const cc of ccs.items) {
          if (!cc.tag || cc.tag.indexOf(ADV_PREFIX) !== 0 || nested.has(cc)) continue;
          const id = pieceIdFromTag(cc.tag);
          if (findPiece(id)) continue;
          const num = extractNumberFromCitation(cc.text, "adv");
          const name = extractNameFromCitation(cc.text, id, "adv");
          if ((num == null || num === "") && (name == null || name === "")) continue;
          const fixed = num != null && String(num).trim() !== "" ? String(num).trim() : null;
          model.adverse.push(makePiece({ id, name: (name || "").trim(), locked: fixed != null, fixedNumber: fixed, recovered: true }));
          healed++;
        }
        if (healed) save();
      } catch (e) { /* la réparation est au mieux : ne jamais casser la synchro */ }

      const numbers = computeStructuredNumbers();
      const advNumbers = onSide("adv", computeStructuredNumbers);

      // Bordereau DIFFÉRENTIEL : reconstruit seulement si son contenu attendu (pièces, numéros,
      // noms, réglages de présentation) a changé depuis la dernière reconstruction — ou en force.
      // Décidé AVANT toute écriture : la lecture du style Normal (un aller-retour) ne se mêle
      // ainsi jamais aux écritures en attente.
      const bord = ccs.items.find((cc) => cc.tag === BORDEREAU_TAG);
      let bordSig = null, rebuildBord = false, normalFont = null;
      if (bord) {
        bordSig = JSON.stringify({
          s: model.settings.bordereau,
          rows: orderedPieces(numbers).filter((p) => !p.container)
            .map((p) => [String(numbers.get(p.id) ?? ""), p.name || "", p.parentId ? 1 : 0]),
        });
        if (lastBordSig === null) lastBordSig = storedBordSig(); // 1re synchro de la session
        rebuildBord = force || bordSig !== lastBordSig;
        if (rebuildBord) normalFont = await loadNormalFont(context);
      }

      const counts = {};
      let orphanCCs = 0;
      // SYNCHRO DIFFÉRENTIELLE : on ne réécrit une citation QUE si son texte doit changer (ou en
      // mode force, après un changement de réglages de mise en forme). Un document déjà à jour
      // n'est donc PAS modifié : pas de révisions parasites, pas de « document modifié », et la
      // pile Ctrl+Z de l'utilisateur reste intacte.
      const writeCite = (cc, text, title, side) => {
        if (!force && (cc.text || "") === text) return;
        const cite = citeSettings(side);
        context.wpUntracked();
        cc.insertText(text, "Replace");
        cc.title = title;
        cc.appearance = "Hidden"; // texte naturel, sans encadré
        cc.font.bold = !!cite.bold;
        cc.font.italic = !!cite.italic;
        cc.font.underline = cite.underline ? "Single" : "None";
        if (cite.alignment && cite.alignment !== "none") {
          cc.paragraphs.getFirst().alignment = mapAlignment(cite.alignment);
        }
      };
      const markOrphan = (cc) => {
        orphanCCs++;
        // NON DESTRUCTIF : on préserve le texte existant et on signale seulement (une seule fois :
        // un orphelin déjà marqué en rouge n'est pas re-marqué → le doc n'est pas re-sali).
        if ((cc.color || "").toLowerCase() === "#dc2626") return;
        context.wpUntracked();
        cc.title = "Pièce supprimée";
        cc.appearance = "BoundingBox";
        cc.color = "#dc2626";
      };
      const countCite = (id) => { counts[id] = (counts[id] || 0) + 1; };

      // Citations des DEUX côtés (nos pièces, pièces adverses) : même traitement, balises distinctes.
      const syncSide = (side, nums) => {
        const c = sideCfg(side);
        const has = (id) => sidePieces(side).some((p) => p.id === id);
        for (const cc of ccs.items.filter((x) => x.tag && x.tag.startsWith(c.cite) && !nested.has(x))) {
          const { id, extra } = splitPieceTag(cc.tag);
          const piece = has(id) ? findPiece(id) : null;
          if (piece) {
            countCite(id);
            const num = nums.get(id);
            // `extra` = précision propre à CETTE occurrence (« , page 8 ») : conservée telle quelle.
            writeCite(cc, formatCitation(num, piece.name, side) + extra, c.one + " " + num, side);
          } else markOrphan(cc);
        }
        // Citations multiples : plages « Pièces n°X à Y », « et s. », listes.
        for (const cc of ccs.items.filter((x) => x.tag && x.tag.startsWith(c.range) && !nested.has(x))) {
          const [aId, bId] = cc.tag.slice(c.range.length).split("~");
          if (has(aId) && has(bId)) {
            countCite(aId); countCite(bId);
            let an = nums.get(aId), bn = nums.get(bId);
            if (naturalCompare(an, bn) > 0) { const t = an; an = bn; bn = t; }
            writeCite(cc, formatRange(String(an), String(bn), side), c.many + " n°" + an + " à " + bn, side);
          } else markOrphan(cc);
        }
        for (const cc of ccs.items.filter((x) => x.tag && x.tag.startsWith(c.etseq) && !nested.has(x))) {
          const aId = cc.tag.slice(c.etseq.length);
          if (has(aId)) {
            countCite(aId);
            const an = nums.get(aId);
            writeCite(cc, formatEtSeq(String(an), side), c.many + " n°" + an + " et suivantes", side);
          } else markOrphan(cc);
        }
        for (const cc of ccs.items.filter((x) => x.tag && x.tag.startsWith(c.list) && !nested.has(x))) {
          const ids = cc.tag.slice(c.list.length).split("~");
          if (ids.every(has)) {
            ids.forEach(countCite);
            writeCite(cc, formatList(ids.map((id) => String(nums.get(id))), side), "Citation multiple", side);
          } else markOrphan(cc);
        }
      };
      syncSide("own", numbers);
      syncSide("adv", advNumbers);

      if (rebuildBord) {
        context.wpUntracked();
        if (model.settings.bordereau.layout === "table") bord.insertHtml(buildBordereauHtml(numbers), "Replace");
        else fillBordereauListEditable(bord, numbers, normalFont);
      }

      await context.sync();
      if (rebuildBord) await detachBordereauLists(context, bord); // aucune puce héritée
      lastBordSig = bordSig; // mémorisé APRÈS un sync réussi (sinon le 2e essai sauterait le rendu)
      if (rebuildBord && bordSig) storeBordSig(bordSig);

      // Ordre chronologique : rang selon la 1re apparition dans le document ;
      // les pièces non citées viennent après (ordre de création).
      const orderRank = new Map();
      let rank = 0;
      for (const id of appearanceOrderIds) {
        if (findPiece(id) && !orderRank.has(id)) orderRank.set(id, rank++);
      }
      for (const p of model.pieces) {
        if (!orderRank.has(p.id)) orderRank.set(p.id, rank++);
      }

      stats = { counts, orphanCCs, numbers, advNumbers, hasBordereau: !!bord, orderRank, tracking: trackingMode };
    }, { lazy: true });
    return stats;
  }

  // Rafraîchissement LECTURE SEULE des stats (counts / orphelins / présence du bordereau) : ne
  // réécrit RIEN dans le document (ne le salit pas). Sert à garder le diagnostic à jour sans sync.
  async function refreshStats() {
    const counts = {};
    let orphanCCs = 0;
    let hasBordereau = false;
    const appearanceOrderIds = [];
    const bump = (id) => { if (findPiece(id)) counts[id] = (counts[id] || 0) + 1; else orphanCCs++; };
    await Word.run(async (context) => {
      // Détection (lecture seule) du suivi des modifications, pour le bandeau du volet.
      try {
        context.document.load("changeTrackingMode");
        await context.sync();
        trackingMode = context.document.changeTrackingMode || "Off";
      } catch (e) { trackingMode = null; }
      const ccs = context.document.body.contentControls;
      ccs.load("items/tag");
      await context.sync();
      for (const cc of ccs.items) {
        const tag = cc.tag || "";
        if (tag === BORDEREAU_TAG) { hasBordereau = true; continue; }
        if (tag.indexOf(TAG_PREFIX) === 0) {
          const id = pieceIdFromTag(tag);
          if (findPiece(id) && !appearanceOrderIds.includes(id)) appearanceOrderIds.push(id);
          bump(id);
        } else if (tag.indexOf(ADV_PREFIX) === 0) {
          bump(pieceIdFromTag(tag));
        } else if (tag.indexOf(ETSEQ_PREFIX) === 0 || tag.indexOf(ADV_ETSEQ_PREFIX) === 0) {
          bump(tag.slice(tag.indexOf(ETSEQ_PREFIX) === 0 ? ETSEQ_PREFIX.length : ADV_ETSEQ_PREFIX.length));
        } else if (tag.indexOf(RANGE_PREFIX) === 0 || tag.indexOf(ADV_RANGE_PREFIX) === 0) {
          const [a, b] = tag.slice(tag.indexOf(RANGE_PREFIX) === 0 ? RANGE_PREFIX.length : ADV_RANGE_PREFIX.length).split("~");
          bump(a); if (b && b !== a) bump(b);
        } else if (tag.indexOf(LIST_PREFIX) === 0 || tag.indexOf(ADV_LIST_PREFIX) === 0) {
          tag.slice(tag.indexOf(LIST_PREFIX) === 0 ? LIST_PREFIX.length : ADV_LIST_PREFIX.length).split("~").forEach(bump);
        }
      }
    });
    const numbers = computeStructuredNumbers();
    const orderRank = new Map();
    let rank = 0;
    for (const id of appearanceOrderIds) if (findPiece(id) && !orderRank.has(id)) orderRank.set(id, rank++);
    for (const p of model.pieces) if (!orderRank.has(p.id)) orderRank.set(p.id, rank++);
    const advNumbers = onSide("adv", computeStructuredNumbers);
    stats = { counts, orphanCCs, numbers, advNumbers, hasBordereau, orderRank, tracking: trackingMode };
    return stats;
  }

  // ---------------- Actions document ----------------
  async function insertCitation(id) {
    const piece = findPiece(id);
    if (!piece) return;
    const side = sideOf(id);
    const c = sideCfg(side);
    const onNewLine = citeSettings(side).newLine !== false;
    await runUntracked(async (context) => {
      if (onNewLine) {
        // La citation va dans SON PROPRE paragraphe (l'alignement ne touche donc pas la phrase),
        // et le curseur repart sur une ligne normale en dessous.
        const para = context.document.getSelection().paragraphs.getLast();
        para.load("alignment");
        await context.sync();
        const baseAlign = para.alignment;

        const citePara = para.insertParagraph(formatCitation("?", piece.name, side), "After");
        const cc = citePara.getRange("Content").insertContentControl();
        cc.tag = c.cite + id;
        cc.title = c.one;
        cc.appearance = "Hidden";
        await context.sync();
        // Une ligne vide avant et après, ni plus ni moins (curseur sur un paragraphe vide, entre deux
        // paragraphes, en fin de document…).
        await ensureBlankBefore(context, citePara, baseAlign || "Left");
        const afterPara = await ensureBlankAfter(context, citePara, baseAlign || "Left");
        afterPara.getRange("Start").select();
        await context.sync();
      } else {
        const cc = context.document.getSelection().insertContentControl();
        cc.tag = c.cite + id;
        cc.title = c.one;
        cc.appearance = "Hidden";
        cc.insertText(formatCitation("?", piece.name, side), "Replace");
        await context.sync();
      }
    });
  }

  // Rend le FOCUS CLAVIER au document après la fermeture d'une fenêtre Office (palette /p).
  // Cause : la fenêtre appartient au volet ; à sa fermeture, Windows rend le clavier au volet. Re-sélectionner
  // le texte replace le point d'insertion mais NE rend PAS le clavier (OfficeDev/office-js#316). Depuis
  // WordApiDesktop 1.4, Word.Window.activate() / setFocus() (équivalents VBA) activent la fenêtre du
  // document et lui rendent le clavier. Renvoie false si cette API n'existe pas dans ce Word.
  function focusApiSupported() {
    try { return Office.context.requirements.isSetSupported("WordApiDesktop", "1.4"); } catch (e) { return false; }
  }
  async function focusSelection() {
    const supported = focusApiSupported();
    const once = async () => {
      try {
        await Word.run(async (context) => {
          context.document.getSelection().select();
          await context.sync();
          if (!supported) return;
          const w = context.document.activeWindow;
          try { w.activate(); await context.sync(); } catch (e) { /* ignoré */ }
          try { w.setFocus(); await context.sync(); } catch (e) { /* ignoré */ }
        });
      } catch (e) { /* sans effet */ }
    };
    await once();
    // Windows peut rendre le focus au volet un court instant APRÈS la fermeture : on réessaie.
    setTimeout(once, 200);
    setTimeout(once, 600);
    return supported;
  }

  async function gotoPiece(id) {
    await Word.run(async (context) => {
      const ccs = context.document.body.contentControls;
      ccs.load("items/tag");
      await context.sync();
      const cc = ccs.items.find((c) => isTagOfPiece(c.tag, id));
      if (cc) { cc.select(); await context.sync(); }
    });
  }

  // Sélectionne la N-ième occurrence (citation) d'une pièce dans le document.
  // Renvoie le nombre total d'occurrences (pour la navigation cyclique).
  async function selectOccurrence(pieceId, index) {
    let total = 0;
    await Word.run(async (context) => {
      const ccs = context.document.body.contentControls;
      ccs.load("items/tag");
      await context.sync();
      const matching = ccs.items.filter((c) => isTagOfPiece(c.tag, pieceId));
      total = matching.length;
      if (!total) return;
      const i = ((index % total) + total) % total; // cyclique
      matching[i].select();
      await context.sync();
    });
    return total;
  }

  async function generateBordereau() {
    await runUntracked(async (context) => {
      const ccs = context.document.body.contentControls;
      ccs.load("items/tag");
      await context.sync();
      let bord = ccs.items.find((cc) => cc.tag === BORDEREAU_TAG);
      if (!bord) {
        const body = context.document.body;
        body.insertBreak(Word.BreakType.page, "End");
        const para = body.insertParagraph("", "End");
        resetBordPara(para); // ne pas hériter du dernier paragraphe (puces, alignement…)
        const cc = para.insertContentControl();
        cc.tag = BORDEREAU_TAG;
        cc.title = "Bordereau de pièces";
        cc.appearance = "Hidden";
      }
      await context.sync();
    });
  }

  // Supprime le bordereau du document (le contrôle de contenu ET son texte).
  async function deleteBordereau() {
    await runUntracked(async (context) => {
      const ccs = context.document.body.contentControls;
      ccs.load("items/tag");
      await context.sync();
      const bord = ccs.items.find((cc) => cc.tag === BORDEREAU_TAG);
      if (!bord) return;
      bord.delete(false); // supprime le contrôle et tout son contenu
      await context.sync();
    });
  }

  // Reconnaissance de l'EXISTANT : détecte les paragraphes entièrement d'une forme « Pièce … » et
  // enveloppe chaque occurrence dans un contrôle géré (wp:piece) pour la rendre dynamique.
  // TOLÉRANT (autonome) : « Pièce »/« Pieces »/« Piece » (accent/pluriel), « n° »/« n »/« nº »/« no »
  //   OPTIONNEL ; numéro entier, sous-numéro « 1.1 »/« 2.3.1 », ou « 1 bis/ter/quater » ; séparateur =
  //   toute suite de « ) ] . : _ · • - – — » (donc « : », « ) », « ). », « - », « — »…). Le numéro doit
  //   suivre « Pièce » et un vrai séparateur doit précéder le nom → très peu de faux positifs.
  const SCAN_RE = /^\s*pi[eè]ces?\s+(?:n\s*[°ºo]?\s*)?(\d+(?:\.\d+)*(?:\s*(?:bis|ter|quater))?)\s*[)\].:_·•\-–—]+\s*(.+?)\s*$/i;
  // Numéro détecté → forme canonique : « 05 » → « 5 », « 1bis » → « 1 bis », espaces normalisés.
  // Sans cela, « Pièce n°05 » et « Pièce n°5 » donnaient deux pièces distinctes.
  function canonScanNum(raw) {
    return String(raw).replace(/\s+/g, " ").trim()
      .replace(/(^|\.)0+(\d)/g, "$1$2")
      .replace(/(\d)\s*(bis|ter|quater)$/i, (m, d, w) => d + " " + w.toLowerCase());
  }
  // Construit une regex de scan à partir d'un GABARIT personnalisé « … {num} … {nom} … » (repli manuel
  // pour les formats exotiques). {num} avant {nom}. Renvoie null si le gabarit est inutilisable.
  function buildScanRegex(tpl) {
    if (!/\{num\}/i.test(tpl) || !/\{nom\}|\{name\}/i.test(tpl)) return null;
    let pattern = "";
    for (const tk of String(tpl).split(/(\{num\}|\{nom\}|\{name\})/i)) {
      if (/^\{num\}$/i.test(tk)) pattern += "(\\d+(?:\\.\\d+)*(?:\\s*(?:bis|ter|quater))?)";
      else if (/^\{nom\}$/i.test(tk) || /^\{name\}$/i.test(tk)) pattern += "(.+?)";
      else pattern += escapeRegex(tk).replace(/\s+/g, "\\s+");
    }
    try { return new RegExp("^\\s*" + pattern + "\\s*$", "i"); } catch (e) { return null; }
  }
  async function scanExistingPieces(customTemplate) {
    const re = (customTemplate && customTemplate.trim()) ? (buildScanRegex(customTemplate) || SCAN_RE) : SCAN_RE;
    const result = { wrapped: 0, newPieces: 0, conflicts: 0, extras: 0, harmonized: 0 };
    const wrappedCCs = []; // contrôles créés par ce scan (pour l'annulation)
    const newIds = [];     // pièces créées par ce scan (pour l'annulation)
    const byNum = new Map(); // numéro -> [ids] (pièces existantes + créées pendant le scan)
    const curNums = computeStructuredNumbers();
    const addByNum = (n, id) => { if (!byNum.has(n)) byNum.set(n, []); byNum.get(n).push(id); };
    for (const p of model.pieces) {
      if (p.container) continue;
      const n = curNums.get(p.id);
      if (n != null && n !== "") addByNum(String(n), p.id);
    }

    await runUntracked(async (context) => {
      const paras = context.document.body.paragraphs;
      paras.load("items/text");
      await context.sync();
      const items = paras.items;
      // Détection FIABLE du « déjà géré » : on regarde les contrôles CONTENUS dans le paragraphe
      // (le bord de paragraphe via getRange("Start") était ambigu et provoquait des imbrications).
      const innerCCs = items.map((p) => p.contentControls);
      innerCCs.forEach((cl) => cl.load("items/tag"));
      await context.sync();

      // ---- PASSE 1 : relever toutes les citations détectées (sans rien modifier) ----
      const hits = [];
      for (let i = 0; i < items.length; i++) {
        // Déjà géré si le paragraphe contient un contrôle wp: (citation ou nom de bordereau).
        if (innerCCs[i].items.some((c) => c.tag && c.tag.indexOf("wp:") === 0)) continue;
        const m = (items[i].text || "").match(re);
        if (!m) continue;
        hits.push({
          i,
          rawNum: canonScanNum(m[1]),
          name: m[2].replace(/\s+/g, " ").trim(),
        });
      }

      // Nom CANONIQUE d'une occurrence = le plus COURT des libellés SEMBLABLES portant le même numéro
      // (le tronc commun : « Contrat de bail » plutôt que « Contrat de bail, page 8 »).
      const canonicalFor = (h) => {
        let best = h.name;
        for (const o of hits) {
          if (o.rawNum === h.rawNum && o.name && o.name.length < best.length && sameDocName(o.name, h.name)) best = o.name;
        }
        return best;
      };

      // ---- PASSE 2 : rattacher chaque occurrence à UNE pièce, puis emballer la citation ----
      // Même numéro dans le document ⇒ en principe la même pièce : on rattache dès que les libellés se
      // ressemblent (variante, précision, faute de frappe, guillemets, ponctuation…). Un vrai conflit
      // (libellés sans rapport) crée UNE pièce distincte PAR LIBELLÉ — jamais une par occurrence.
      for (const h of hits) {
        const cands = (byNum.get(h.rawNum) || []).map(findPiece).filter(Boolean);
        let pc = cands.find((c) => c.name && c.name.trim() && sameDocName(c.name, h.name))
          || cands.find((c) => !(c.name && c.name.trim())) || null;
        if (!pc) {
          // Pièce détectée avec un numéro explicite → reprise ÉPINGLÉE + PROTÉGÉE à ce numéro.
          // `scanned` : exclue du verrou général (elle reste figée à son numéro d'origine).
          if (cands.length) result.conflicts++;
          pc = makePiece({ name: canonicalFor(h), locked: true, fixedNumber: h.rawNum, scanned: true });
          model.pieces.push(pc);
          addByNum(h.rawNum, pc.id);
          newIds.push(pc.id);
          result.newPieces++;
        } else if (!(pc.name && pc.name.trim())) {
          pc.name = canonicalFor(h); // même pièce : complète un nom vide
        }
        // Libellé de CETTE occurrence ≠ nom de la pièce : « nom + suite » → précision locale (texte
        // conservé à l'identique) ; autre variante → harmonisée au nom de la pièce à la synchro.
        let extra = "";
        if (h.name && nameKey(h.name) !== nameKey(pc.name)) {
          const r = nameRemainder(pc.name, h.name);
          if (r != null) { extra = r; result.extras++; }
          else result.harmonized++;
        } else if (h.name && h.name !== pc.name) {
          const r = nameRemainder(pc.name, h.name); // simple ponctuation finale (« Contrat de bail. »)
          if (r != null) extra = r;
        }
        const id = pc.id;
        const cc = items[h.i].getRange("Content").insertContentControl();
        cc.tag = pieceTagFor(id, extra);
        cc.title = "Pièce";
        cc.appearance = "Hidden";
        wrappedCCs.push(cc);
        result.wrapped++;
      }
      await context.sync();
      // Récupère les ids des contrôles créés (pour pouvoir annuler le scan : les déballer).
      wrappedCCs.forEach((c) => c.load("id"));
      await context.sync();
      result.wrappedCcIds = wrappedCCs.map((c) => c.id);
      result.newPieceIds = newIds.slice();
    });

    if (result.wrapped) {
      normalizeStructure(); // « 3.1 », « 3.2 » détectées → vraies sous-pièces du groupe 3
      await save();
    }
    return result;
  }

  // Déballe (supprime le content control en gardant le texte) des citations par id — sert à ANNULER un scan.
  async function unwrapCitations(ids) {
    if (!ids || !ids.length) return;
    await runUntracked(async (context) => {
      const ccs = ids.map((id) => context.document.contentControls.getByIdOrNullObject(id));
      ccs.forEach((cc) => cc.load("isNullObject"));
      await context.sync();
      // delete(TRUE) = retire le contrôle en GARDANT son texte (delete(false) effacerait le texte : c'était
      // le bug qui faisait disparaître du document toutes les pièces reprises par le scan).
      ccs.forEach((cc) => { if (!cc.isNullObject) cc.delete(true); });
      await context.sync();
    });
  }

  async function getPieceIdAtSelection() {
    let id = null;
    await Word.run(async (context) => {
      const cc = context.document.getSelection().parentContentControlOrNullObject;
      cc.load("tag,isNullObject");
      await context.sync();
      if (!cc.isNullObject && isSingleCiteTag(cc.tag)) id = pieceIdFromTag(cc.tag);
    });
    return id;
  }

  async function removeCitationAtSelection() {
    let removed = false;
    await runUntracked(async (context) => {
      const cc = context.document.getSelection().parentContentControlOrNullObject;
      cc.load("tag,isNullObject");
      await context.sync();
      const isCite = !cc.isNullObject && cc.tag && isCiteTag(cc.tag);
      if (isCite) {
        const para = cc.getRange("Whole").paragraphs.getFirst();
        cc.load("text");
        para.load("text");
        await context.sync();
        const ccText = (cc.text || "").trim();
        // Citation seule sur sa ligne → la ligne part aussi (sans laisser deux lignes vides) ;
        // sinon on retire seulement la citation (contrôle ET texte).
        if (ccText && (para.text || "").trim() === ccText) await deleteCitationLine(context, para);
        else { cc.delete(false); await context.sync(); }
        removed = true;
      }
    });
    return removed;
  }

  // Contexte du curseur : pièce + id du content control sous le curseur.
  async function getSelectionContext() {
    const out = { pieceId: null, ccId: null };
    await Word.run(async (context) => {
      const cc = context.document.getSelection().parentContentControlOrNullObject;
      cc.load("tag,id,isNullObject");
      await context.sync();
      if (!cc.isNullObject && isSingleCiteTag(cc.tag)) {
        out.pieceId = pieceIdFromTag(cc.tag);
        out.ccId = cc.id;
      }
    });
    return out;
  }

  // Liste des occurrences d'une pièce : id du CC, index, aperçu du paragraphe.
  async function listOccurrences(pieceId) {
    const list = [];
    await Word.run(async (context) => {
      const ccs = context.document.body.contentControls;
      ccs.load("items/tag,items/id");
      await context.sync();
      const matching = ccs.items.filter((cc) => isTagOfPiece(cc.tag, pieceId));
      const paras = matching.map((cc) => cc.getRange("Whole").paragraphs.getFirst());
      paras.forEach((p) => p.load("text"));
      await context.sync();
      matching.forEach((cc, i) => {
        let preview = (paras[i].text || "").replace(/\s+/g, " ").trim();
        if (preview.length > 70) preview = preview.slice(0, 67) + "…";
        list.push({ ccId: cc.id, index: i + 1, preview });
      });
    });
    return list;
  }

  async function selectCcById(id) {
    await Word.run(async (context) => {
      const cc = context.document.contentControls.getByIdOrNullObject(id);
      cc.load("isNullObject");
      await context.sync();
      if (!cc.isNullObject) { cc.select(); await context.sync(); }
    });
  }

  async function removeCcById(id) {
    await runUntracked(async (context) => {
      const cc = context.document.contentControls.getByIdOrNullObject(id);
      cc.load("isNullObject");
      await context.sync();
      if (!cc.isNullObject) { cc.delete(false); await context.sync(); }
    });
  }

  // ---------------- Import d'une liste de pièces adverses (copier-coller) ----------------
  // L'utilisateur colle la liste des pièces qui clôt les conclusions / l'assignation adverses. Formats
  // reconnus, ligne par ligne : « 1. Assignation », « 1) … », « 1 - … », « 1 : … », « N°1 … »,
  // « Pièce n°1 : … », « PIECE 1 – … », puces « - », « • », sous-numéros « 4.1 », « 2 bis »,
  // tableau collé depuis Word (« 1⇥Contrat de bail⇥12/03/2019 »). Les lignes sans numéro avant la
  // 1re pièce (titre « Bordereau de pièces »…) sont ignorées ; une ligne sans numéro qui commence par
  // une minuscule ou une parenthèse prolonge le nom précédent (intitulé coupé sur deux lignes).
  const ADV_LINE_RE = /^\s*(?:[-–—•·*▪►]\s*)?(?:pi[eè]ces?\s*)?(?:n\s*[°ºo]\.?\s*)?(\d{1,4}(?:\.\d{1,3})*(?:\s*(?:bis|ter|quater))?)(?![\d/])\s*(?:[)\].:_·•\-–—\t|]+\s*|\s+)(.+?)\s*$/i;
  function parseAdverseList(text) {
    const out = [];
    for (const raw of String(text || "").split(/\r\n|\r|\n|\u2028|\u2029/)) {
      const line = raw.replace(/\u00a0/g, " ");
      if (!line.trim()) continue;
      const m = line.match(ADV_LINE_RE);
      if (m) {
        const name = m[2].replace(/\t+/g, " – ").replace(/\s+/g, " ").replace(/\s*[;,]+$/, "").replace(/(?<!\.\.)(?<!\b[A-Za-z])\.$/, "").trim();
        if (!name) continue;
        out.push({ num: canonScanNum(m[1]), name });
      } else if (out.length && /^\s*[a-zà-ÿ(«"“]/.test(line)) {
        const last = out[out.length - 1];
        last.name = (last.name + " " + line.trim()).replace(/\s*[;,]+$/, "");
      }
    }
    return out;
  }
  // Plan d'import (LECTURE SEULE) : pour chaque ligne détectée, ce qui se passera.
  //  "new"    → nouvelle pièce adverse (cochée) ;
  //  "same"   → déjà présente sous ce numéro avec un nom équivalent (rien à faire) ;
  //  "rename" → numéro déjà pris par une pièce au nom différent (décochée : on remplacera son nom) ;
  //  "dup"    → numéro répété dans la liste collée (seule la 1re occurrence compte).
  function planAdverseImport(parsed) {
    const seen = new Set();
    return parsed.map((it) => {
      if (seen.has(it.num)) return { ...it, status: "dup", checked: false };
      seen.add(it.num);
      const id = onSide("adv", () => pieceIdByNumber(it.num));
      const ex = id ? findPiece(id) : null;
      if (!ex) return { ...it, status: "new", checked: true };
      if (!ex.name || sameDocName(ex.name, it.name)) {
        return { ...it, status: ex.name ? "same" : "rename", oldName: ex.name || "", checked: !ex.name };
      }
      return { ...it, status: "rename", oldName: ex.name, checked: false };
    });
  }
  // Applique les lignes COCHÉES du plan. Pièces créées VERROUILLÉES à leur numéro (celui du bordereau
  // adverse) — déverrouillables ensuite dans le volet. Renvoie { created, renamed, ids }.
  async function importAdversePieces(items) {
    const res = { created: 0, renamed: 0, ids: [] };
    onSide("adv", () => {
      for (const it of items || []) {
        if (!it.checked || !it.num || !it.name) continue;
        const id = pieceIdByNumber(it.num);
        const ex = id ? findPiece(id) : null;
        if (ex) {
          if (ex.name !== it.name) { ex.name = it.name; res.renamed++; res.ids.push(ex.id); }
        } else {
          const p = createPieceAt(it.num);
          p.name = it.name;
          p.locked = true;
          p.imported = true;
          res.created++;
          res.ids.push(p.id);
        }
      }
    });
    if (res.created || res.renamed) await save();
    return res;
  }

  // ---------------- Mutations du modèle (persistées) ----------------
  // side "adv" : pièce adverse au numéro suivant le plus grand (modifiable ensuite dans le volet).
  async function addPiece(name, side) {
    const clean = (name || "").trim();
    if (!clean) return null;
    let piece;
    if (side === "adv") {
      piece = onSide("adv", () => createPieceAt(nextAdverseNumber()));
      piece.name = clean;
    } else {
      piece = makePiece({ name: clean });
      model.pieces.push(piece);
    }
    await save();
    return piece;
  }
  async function renamePiece(id, name) {
    const p = findPiece(id);
    if (!p) return;
    p.name = (name || "").trim();
    await save();
  }
  async function deletePiece(id) {
    onSide(sideOf(id), () => removeFromModel([id]));
    await save();
  }

  // Supprime la pièce PARTOUT : retire ses citations du corps du texte, puis du modèle
  // (le bordereau et le volet se mettent à jour à la synchro suivante).
  // Un GROUPE (conteneur) emporte ses sous-pièces ; une pièce qui a des sous-pièces laisse un groupe
  // (les sous-pièces gardent leur numéro) ; un groupe vidé disparaît (removeFromModel).
  async function deletePieceEverywhere(id) {
    const target = findPiece(id);
    const side = sideOf(id);
    const removeIds = [id];
    if (target && target.container) for (const kid of onSide(side, () => childrenOf(id))) removeIds.push(kid.id);
    await runUntracked(async (context) => {
      const ccs = context.document.body.contentControls;
      ccs.load("items/tag,items/id");
      await context.sync();
      const ids = ccs.items.filter((cc) => removeIds.some((rid) => isTagOfPiece(cc.tag, rid))).map((cc) => cc.id);
      // Une citation à la fois (voisinage relu à chaque fois) : des citations empilées supprimées
      // ensemble ne laissent pas deux lignes vides accolées.
      for (const ccId of ids) {
        const cc = context.document.contentControls.getByIdOrNullObject(ccId);
        cc.load("isNullObject,text");
        await context.sync();
        if (cc.isNullObject) continue;
        const para = cc.getRange("Whole").paragraphs.getFirst();
        para.load("text");
        await context.sync();
        const ccText = (cc.text || "").trim();
        // Citation seule sur sa ligne → on supprime la ligne ; sinon juste la citation.
        if (ccText.length > 0 && (para.text || "").trim() === ccText) await deleteCitationLine(context, para);
        else { cc.delete(false); await context.sync(); }
      }
    });
    onSide(side, () => removeFromModel(removeIds));
    await save();
  }

  // Détection unifiée : plage « /p3à7 », « et s. » « /p5+ », ou citation simple « /p1 ».
  // « à » et « + » ne font pas partie des numéros libres → aucune ambiguïté.
  async function detectSlashAny() {
    let out = null;
    await Word.run(async (context) => {
      const sel = context.document.getSelection();
      const before = sel.paragraphs.getFirst().getRange("Start").expandTo(sel);
      before.load("text");
      await context.sync();
      const t = before.text || "";
      let m;
      // « /pf » : applique seulement le format citation à une ligne libre (aucune saisie captée).
      if (/(?:^|\s)\/pf\s$/.test(t)) out = { kind: "format" };
      // ---- PIÈCES ADVERSES « /pa… » (testées AVANT « /p<nom> », qui les capterait sinon) ----
      // « /pan » : pièce adverse au numéro suivant le plus grand ; « /pan3 » = « /pa3 ».
      else if (/(?:^|\s)\/pan\s$/.test(t)) out = { kind: "new", side: "adv" };
      else if ((m = t.match(/(?:^|\s)\/pan(\d[\w.]*)\s$/))) out = { kind: "single", num: m[1], side: "adv", token: "/pan" + m[1] };
      else if ((m = t.match(/(?:^|\s)\/pa(\d+)-(\d+)\s$/))) out = { kind: "range", start: m[1], end: m[2], side: "adv" };
      else if ((m = t.match(/(?:^|\s)\/pa(\d[\w.]*)\+\s$/))) out = { kind: "etseq", start: m[1], side: "adv" };
      else if ((m = t.match(/(?:^|\s)\/pa(\d[\w.]*(?:\s*,\s*\d[\w.]*)+)\s$/))) out = { kind: "list", list: m[1], side: "adv" };
      else if ((m = t.match(/(?:^|\s)\/pa(\d[\w.]*)\s$/))) out = { kind: "single", num: m[1], side: "adv" };
      else if (/(?:^|\s)\/pa\s$/.test(t)) out = { kind: "search", query: "", side: "adv" };
      else if ((m = t.match(/(?:^|\s)\/pa([A-Za-zÀ-ÿ][^\n]*?)\s$/))) out = { kind: "search", query: m[1], side: "adv" };
      // « /pn » : nouvelle pièce au PROCHAIN numéro top-level.  (À tester AVANT « /p… ».)
      else if (/(?:^|\s)\/pn\s$/.test(t)) out = { kind: "new" };
      // « /pn4.2 » : insère une SOUS-pièce au rang 2 sous la pièce 4 (décale les frères suivants).
      else if ((m = t.match(/(?:^|\s)\/pn(\d+)\.(\d+)\s$/))) out = { kind: "subinsert", parent: m[1], sub: m[2] };
      // « /pn4. » : ajoute la PROCHAINE sous-pièce sous la pièce 4 (4.1, 4.2…).
      else if ((m = t.match(/(?:^|\s)\/pn(\d+)\.\s$/))) out = { kind: "subnew", parent: m[1] };
      // « /pn2 » : INSÈRE une nouvelle pièce au n°2 et DÉCALE les suivantes.
      else if ((m = t.match(/(?:^|\s)\/pn(\d+)\s$/))) out = { kind: "insert", num: m[1] };
      // Plage « /p4-8 » : bornes ENTIÈRES (permet d'énumérer 4,5,6,7,8).
      else if ((m = t.match(/(?:^|\s)\/p(\d+)-(\d+)\s$/))) out = { kind: "range", start: m[1], end: m[2] };
      // « et suivantes » « /p5+ ».
      else if ((m = t.match(/(?:^|\s)\/p(\d[\w.]*)\+\s$/))) out = { kind: "etseq", start: m[1] };
      // Liste « /p1,2,5 » : plusieurs pièces séparées par des virgules (au moins deux).
      else if ((m = t.match(/(?:^|\s)\/p(\d[\w.]*(?:\s*,\s*\d[\w.]*)+)\s$/))) out = { kind: "list", list: m[1] };
      // Citation simple « /p1 », « /p1bis », « /p1.1 » (cite l'existante, sinon crée AU n° indiqué).
      else if ((m = t.match(/(?:^|\s)\/p(\d[\w.]*)\s$/))) out = { kind: "single", num: m[1] };
      // Recherche par NOM : « /p » seul → palette vide ; « /pbail » → palette pré-filtrée « bail ».
      else if (/(?:^|\s)\/p\s$/.test(t)) out = { kind: "search", query: "" };
      else if ((m = t.match(/(?:^|\s)\/p([A-Za-zÀ-ÿ][^\n]*?)\s$/))) out = { kind: "search", query: m[1] };
      // Signature = texte déclencheur exact : sert au volet à ignorer un écho d'annulation (Ctrl+Z),
      // qui restaure le même « /pn » et re-déclencherait sinon la même commande.
      if (out) { out._sig = t; if (!out.side) out.side = "own"; }
    });
    return out;
  }

  // ---------------- Espacement autour des citations ----------------
  // RÈGLE : un bloc de citation(s) est TOUJOURS encadré d'exactement UNE ligne vide :
  //   paragraphe · ligne vide · citation(s) · ligne vide · paragraphe suivant.
  // Quel que soit le chemin (commande tapée en fin de paragraphe ou sur une ligne déjà sautée, insertion
  // depuis le volet entre deux paragraphes, suppression d'une citation…), on COMPTE les lignes vides de
  // part et d'autre : aucune → on en crée une ; plusieurs → on retire les lignes en trop.
  // « Ligne vide » = paragraphe sans texte (espaces tolérés), hors tableau, sans image, sans contrôle de
  // contenu, sans saut de page ni de section : on ne supprime JAMAIS autre chose qu'une vraie ligne vide.
  const BLANK_SCAN_MAX = 12; // au-delà, on laisse en l'état (mise en page volontaire)
  async function isBlankPara(context, para) {
    para.load("isNullObject,text" + (apiOK("1.3") ? ",tableNestingLevel" : ""));
    await context.sync();
    if (para.isNullObject) return null; // bord du document
    if (!/^[ \t ]*$/.test(para.text || "")) return false;
    if (para.tableNestingLevel > 0) return false;
    const pics = para.inlinePictures;
    const ccs = para.contentControls;
    pics.load("items"); ccs.load("items");
    try { await context.sync(); } catch (e) { return false; } // doute → on ne touche à rien
    if (pics.items.length || ccs.items.length) return false;
    try {
      const brk = para.search("^m");   // saut de page manuel
      const sec = para.search("^b");   // saut de section
      brk.load("items"); sec.load("items");
      await context.sync();
      if (brk.items.length || sec.items.length) return false;
    } catch (e) { /* recherche des caractères spéciaux indisponible : le texte vide fait foi */ }
    return true;
  }
  // Lignes vides consécutives à partir de `para` (dir -1 : au-dessus, +1 : en dessous), de la plus
  // proche à la plus lointaine. edge = on a atteint le début / la fin du document.
  async function blanksAround(context, para, dir) {
    const blanks = [];
    let cur = para;
    for (let i = 0; i < BLANK_SCAN_MAX; i++) {
      const n = dir < 0 ? cur.getPreviousOrNullObject() : cur.getNextOrNullObject();
      const b = await isBlankPara(context, n);
      if (b === null) return { blanks, edge: true };
      if (!b) return { blanks, edge: false };
      blanks.push(n);
      cur = n;
    }
    return { blanks, edge: false, capped: true };
  }
  function resetBlankFormat(para, baseAlign) {
    para.font.bold = false; para.font.italic = false; para.font.underline = "None";
    if (baseAlign) para.alignment = baseAlign;
  }

  // Exactement UNE ligne vide juste AVANT le paragraphe `p` (aucune en tout début de document).
  // `context` doit être synchronisé (p en place). Fait ses propres synchros.
  async function ensureBlankBefore(context, p, baseAlign) {
    const { blanks, edge, capped } = await blanksAround(context, p, -1);
    if (edge) return; // début de document : pas de ligne vide à imposer (ni à retirer)
    if (!blanks.length) {
      resetBlankFormat(p.insertParagraph("", "Before"), baseAlign);
      await context.sync();
      return;
    }
    if (!capped && blanks.length > 1) {
      for (const b of blanks.slice(1)) b.delete(); // on garde la plus proche de la citation
      await context.sync();
    }
  }

  // Exactement UNE ligne vide juste APRÈS le paragraphe `p` ; la renvoie (le curseur y est replacé).
  // En fin de document, on garde la DERNIÈRE ligne (le document se termine toujours par un paragraphe).
  async function ensureBlankAfter(context, p, baseAlign) {
    const { blanks, edge, capped } = await blanksAround(context, p, +1);
    let after;
    if (!blanks.length) {
      after = p.insertParagraph("", "After");
    } else if (capped || blanks.length === 1) {
      after = blanks[0];
    } else if (edge) {
      after = blanks[blanks.length - 1];
      for (const b of blanks.slice(0, -1)) b.delete();
    } else {
      after = blanks[0];
      for (const b of blanks.slice(1)) b.delete();
    }
    resetBlankFormat(after, baseAlign);
    await context.sync();
    return after;
  }

  // Supprime la LIGNE d'une citation (paragraphe qui ne contient qu'elle) sans laisser deux lignes vides
  // accolées : si la ligne au-dessus ET celle en dessous sont vides, celle du dessous part aussi.
  async function deleteCitationLine(context, para) {
    const prev = para.getPreviousOrNullObject();
    const prevBlank = await isBlankPara(context, prev);
    const next = para.getNextOrNullObject();
    const nextBlank = await isBlankPara(context, next);
    para.delete();
    if (prevBlank && nextBlank) {
      // On retire celle du dessous… sauf si c'est la toute dernière ligne du document (indélébile) :
      // on retire alors celle du dessus.
      const after = next.getNextOrNullObject();
      after.load("isNullObject");
      await context.sync();
      if (after.isNullObject) prev.delete(); else next.delete();
    } else if (prevBlank === null && nextBlank) {
      next.delete(); // citation en tête de document : pas de ligne vide orpheline au début
    }
    await context.sync();
  }

  // Étape 1 : remplace « /pN » par « Pièce n°N :  » DÉJÀ FORMATÉE, curseur prêt pour le nom.
  // Si la pièce existe déjà, insère simplement une citation normale (pas de nommage).
  // side "adv" : même chose pour une pièce adverse (« /pa1 » → « Pièce adverse n°1 : »).
  // token : texte tapé à remplacer quand il diffère de « /p<N> » / « /pa<N> » (ex. « /pan3 »).
  async function startSlashPrompt(rawNum, side, token) {
    const c = sideCfg(side);
    const tok = token || c.token + rawNum;
    // Numéro affiché : « 1bis » → « 1 bis » (espace entre chiffre et lettres). « 1.1 » inchangé.
    const numberStr = String(rawNum).replace(/(\d)([a-zA-Z])/g, "$1 $2");
    // Résolution par NUMÉRO CALCULÉ (structuré) : cite la pièce affichant ce numéro si elle existe.
    let piece = onSide(side, () => { const id = pieceIdByNumber(numberStr); return id ? findPiece(id) : null; });
    // On (re)propose le nom si la pièce n'existe pas OU existe mais n'a pas encore de nom.
    const naming = !piece || !piece.name;
    if (!piece) {
      // Création AU numéro demandé → pièce VERROUILLÉE (n'entraîne aucun décalage). « 4.1 » crée une
      // vraie sous-pièce du groupe 4 ; « 4 » sur un groupe sans tête en devient la tête.
      piece = onSide(side, () => createPieceAt(numberStr));
      await save();
    }
    if (naming) {
      const baseAlign = await insertNamingPromptFor(tok, numberStr, side);
      return { id: piece.id, num: numberStr, naming: true, baseAlign };
    }
    // Pièce existante déjà nommée → citation directe, au texte FINAL (la synchro n'aura rien à réécrire).
    await insertCitationsAtToken(tok, [
      { tag: c.cite + piece.id, title: c.one, text: formatCitation(numberStr, piece.name, side) },
    ], side);
    return { id: piece.id, num: numberStr, naming: false, baseAlign: "Left" };
  }

  // Écrit le prompt de nommage « Pièce n°<disp> :  » à la place du token, formaté, curseur prêt.
  // Factorisé pour /pn, /pnN, /pn4., /pn4.2 (le token complet est passé tel quel).
  async function insertNamingPromptFor(rawToken, dispNum, side) {
    const label = sideCfg(side).one + " n°";
    const applyFmt = citeFmtApplier(side);
    let baseAlign = "Left";
    await runUntracked(async (context) => {
      const sel = context.document.getSelection();
      const para = sel.paragraphs.getFirst();
      para.load("text,alignment,style,leftIndent,firstLineIndent,rightIndent,lineSpacing,spaceBefore,spaceAfter");
      para.font.load("name,size,bold,italic,underline,color,highlightColor");
      const found = para.search(rawToken + " ", { matchCase: true });
      found.load("items");
      await context.sync();
      baseAlign = para.alignment || "Left";
      namingRestore = captureParaFormat(para); // format du texte d'AVANT la commande, rendu après Entrée
      if (!found.items.length) return;
      const range = found.items[found.items.length - 1];
      const tokenAlone = (para.text || "").trim() === rawToken;
      let promptPara, promptRange;
      if (tokenAlone) {
        promptRange = range.insertText(label + dispNum + " : ", "Replace");
        promptPara = para;
      } else {
        range.insertText("", "Replace");
        promptPara = para.insertParagraph(label + dispNum + " : ", "After");
        promptRange = promptPara.getRange("Content");
      }
      applyFmt(promptPara);
      await context.sync();
      await ensureBlankBefore(context, promptPara, baseAlign);
      promptRange.select("End");
      await context.sync();
    });
    return baseAlign;
  }

  // « /pn » : nouvelle pièce COULANTE au prochain numéro top-level, puis nommage.
  // `token` : jeton à remplacer dans le texte (défaut « /pn » ; « /pn3 » quand on refuse le décalage).
  async function startSlashNew(token, side) {
    if (side === "adv") {
      // Pièce adverse : numéro suivant le plus grand, épinglé (jamais de décalage chez l'adversaire).
      const num = nextAdverseNumber();
      const adv = onSide("adv", () => createPieceAt(num));
      await save();
      const align = await insertNamingPromptFor(token || "/pan", num, "adv");
      return { id: adv.id, num, naming: true, baseAlign: align };
    }
    const piece = makePiece({});
    model.pieces.push(piece);
    await save();
    const disp = computeStructuredNumbers().get(piece.id);
    const baseAlign = await insertNamingPromptFor(token || "/pn", disp);
    return { id: piece.id, num: disp, naming: true, baseAlign };
  }

  // Analyse (LECTURE SEULE) de « /pnN » : dit s'il faut proposer un choix à l'utilisateur.
  //  - "range"  : N hors suite (pas de pièce N, et N ≠ prochain numéro) → erreur.
  //  - "plain"  : N libre en fin, ou occupé par une pièce SIMPLE coulante → décalage silencieux.
  //  - "group"  : N est un GROUPE (a des sous-pièces) OU est verrouillé → on demande (décaler / prochain n°).
  //               `locked` indique s'il faut le décalage par ré-épinglage (startSlashInsertShift).
  function analyzeInsert(rawN) {
    const N = parseInt(rawN, 10);
    const nums = computeStructuredNumbers();
    const tops = topPieces();
    const target = tops.find((p) => nums.get(p.id) === String(N));
    const hasSub = model.pieces.some((p) => (nums.get(p.id) || "").indexOf(String(N) + ".") === 0);
    const locked = lockedIntSet(tops).has(N) || !!(target && target.fixedNumber != null);
    if (!target && !hasSub) {
      if (String(N) !== String(nextTopNumber())) return { status: "range", num: N };
      return { status: "plain", num: N };
    }
    if (locked || hasSub) return { status: "group", num: N, next: nextTopNumber(), locked: locked };
    return { status: "plain", num: N };
  }

  // « /pnN » : INSÈRE une pièce coulante au n°N et décale les suivantes (verrous respectés).
  async function startSlashInsert(rawN) {
    const N = parseInt(rawN, 10);
    const tops = topPieces();
    const nums = computeStructuredNumbers();
    // N occupé par une pièce/groupe VERROUILLÉ → on ne décale pas d'office : on propose le décalage.
    if (lockedIntSet(tops).has(N)) return { error: "lockedGroup", num: N, next: nextTopNumber() };
    let insertBeforeId;
    const target = tops.find((p) => nums.get(p.id) === String(N));
    if (target) {
      if (target.fixedNumber != null) return { error: "lockedGroup", num: N, next: nextTopNumber() };
      insertBeforeId = target.id;
    } else {
      if (String(N) !== String(nextTopNumber())) return { error: "range", num: N }; // hors suite
      insertBeforeId = null; // ajout en fin
    }
    const piece = makePiece({});
    insertPieceBefore(piece, insertBeforeId);
    await save();
    const disp = computeStructuredNumbers().get(piece.id);
    const baseAlign = await insertNamingPromptFor("/pn" + rawN, disp);
    return { id: piece.id, num: disp, naming: true, baseAlign };
  }

  // « /pnN » lorsque N est occupé par un groupe VERROUILLÉ (après confirmation) : DÉCALE d'un cran le
  // numéro de tête de toutes les pièces top-level ÉPINGLÉES ≥ N (« 3 »→« 4 », « 3.1 »→« 4.1 », cascade
  // sur 4→5…), puis insère une nouvelle pièce COULANTE au n°N (qui vient de se libérer) + nommage.
  async function startSlashInsertShift(rawN) {
    const N = parseInt(rawN, 10);
    const nums = computeStructuredNumbers();
    const target = topPieces().find((p) => nums.get(p.id) === String(N)); // pièce/conteneur affichant N
    // Incrémente le numéro de tête entier des pièces TOP-LEVEL épinglées ≥ N (les enfants suivent leur parent).
    for (const p of model.pieces) {
      if (p.parentId || p.fixedNumber == null) continue;
      const m = String(p.fixedNumber).match(/^(\d+)(.*)$/);
      if (!m) continue;
      const lead = parseInt(m[1], 10);
      if (lead >= N) p.fixedNumber = (lead + 1) + m[2];
    }
    const piece = makePiece({});
    insertPieceBefore(piece, target ? target.id : null); // avant l'ancienne « N » (désormais « N+1 »)
    await save();
    const disp = computeStructuredNumbers().get(piece.id);
    const baseAlign = await insertNamingPromptFor("/pn" + rawN, disp);
    return { id: piece.id, num: disp, naming: true, baseAlign };
  }

  // Renvoie le CONTENEUR (groupe) du numéro N, en CONVERTISSANT si besoin une pièce N pleine en
  // sous-pièce N.1 — après quoi il ne subsiste plus de pièce N nue. Crée le groupe si N est libre.
  // Renvoie null si N est déjà une SOUS-pièce (cible invalide).
  function ensureGroup(rawN) {
    const N = String(rawN);
    const id = pieceIdByNumber(N, true);
    const p = id ? findPiece(id) : null;
    if (p && p.parentId) return null;      // N est une sous-pièce → cible invalide
    if (p && p.container) return p;         // groupe déjà existant → on ajoute simplement dedans
    // Pièce N ayant DÉJÀ des sous-pièces (4, 4.1…) → c'est déjà un groupe : on ajoute dedans. (La
    // convertir en 4.1 ferait de ses sous-pièces des sous-sous-pièces, sans numéro.)
    if (p && childrenOf(p.id).length) return p;
    if (p) {
      // CAS B : pièce N pleine → un conteneur reprend le numéro N ; p devient sa 1re sous-pièce (N.1).
      const container = makePiece({ container: true, locked: !!p.locked, fixedNumber: p.fixedNumber != null ? p.fixedNumber : null });
      insertPieceBefore(container, p.id);
      p.parentId = container.id;
      p.locked = false; p.fixedNumber = null; // sous-pièce coulante → N.1
      return container;
    }
    // CAS A : aucune pièce N → nouveau groupe (coulant si N = prochain numéro, sinon épinglé à N).
    const flowing = String(nextTopNumber()) === N;
    const container = makePiece({ container: true, fixedNumber: flowing ? null : N });
    model.pieces.push(container);
    return container;
  }

  // « /pn4. » : N devient un GROUPE. Si la pièce 4 existait, elle devient 4.1 et la nouvelle est 4.2 ;
  // sinon on commence à 4.1. Il ne subsiste jamais de pièce 4 nue.
  async function startSlashSubNew(rawParent) {
    const container = ensureGroup(rawParent);
    if (!container) return { error: "noparent", num: rawParent };
    const child = makePiece({ parentId: container.id });
    insertPieceBefore(child, afterChildrenBlockId(container.id));
    await save();
    const disp = computeStructuredNumbers().get(child.id);
    const baseAlign = await insertNamingPromptFor("/pn" + rawParent + ".", disp);
    return { id: child.id, num: disp, naming: true, baseAlign };
  }

  // « /pn4.2 » : INSÈRE une sous-pièce au rang 2 du groupe 4 (décale les frères suivants).
  async function startSlashSubInsert(rawParent, rawSub) {
    // ensureGroup peut transformer la pièce N en groupe : en cas de refus, on RESTAURE le modèle
    // (sinon la conversion — ou un groupe vide invisible — resterait en mémoire et serait enregistrée).
    const snap = JSON.stringify(model.pieces);
    const fail = (err) => { model.pieces = JSON.parse(snap); return err; };
    const container = ensureGroup(rawParent);
    if (!container) return fail({ error: "noparent", num: rawParent + "." + rawSub });
    const S = parseInt(rawSub, 10);
    const kids = childrenOf(container.id);
    const nums = computeStructuredNumbers();
    const parentDisp = nums.get(container.id);
    if (lockedIntSet(kids).has(S)) return fail({ error: "locked", num: parentDisp + "." + S });
    let beforeId;
    const target = kids.find((k) => nums.get(k.id) === parentDisp + "." + S);
    if (target) {
      if (target.fixedNumber != null) return fail({ error: "locked", num: parentDisp + "." + S });
      beforeId = target.id;
    } else {
      if (S !== nextSubIndex(container.id)) return fail({ error: "range", num: parentDisp + "." + S }); // hors suite
      beforeId = afterChildrenBlockId(container.id);
    }
    const child = makePiece({ parentId: container.id });
    insertPieceBefore(child, beforeId);
    await save();
    const disp = computeStructuredNumbers().get(child.id);
    const baseAlign = await insertNamingPromptFor("/pn" + rawParent + "." + rawSub, disp);
    return { id: child.id, num: disp, naming: true, baseAlign };
  }

  // Formatage caractère d'une citation (gras/italique/souligné + alignement paragraphe).
  // Format du paragraphe où une commande de création (/pn, /pnN, /pn4., plage…) a été tapée ; rendu à la
  // ligne qui suit la citation une fois le nom saisi (voir finalizeSlashNamingChain).
  let namingRestore = null;
  function citeFmtApplier(side) {
    const cite = citeSettings(side);
    return (p) => {
      p.font.bold = !!cite.bold;
      p.font.italic = !!cite.italic;
      p.font.underline = cite.underline ? "Single" : "None";
      if (cite.alignment && cite.alignment !== "none") p.alignment = mapAlignment(cite.alignment);
    };
  }

  // HELPER UNIQUE de citation par token : remplace le token « /p… » tapé par une ou plusieurs
  // citations gérées (une par entrée de `cites` : { tag, title, text }). Une seule → sur la ligne
  // du token (ou sa propre ligne si le token termine une phrase) ; plusieurs → paragraphes empilés.
  // Garantit UNE ligne vide avant/après le bloc, applique le format « citation », replace le curseur.
  async function insertCitationsAtToken(rawToken, cites, side) {
    const token = rawToken + " ";
    const applyFmt = citeFmtApplier(side);
    await runUntracked(async (context) => {
      const sel = context.document.getSelection();
      const para = sel.paragraphs.getFirst();
      para.load("text,alignment");
      const found = para.search(token, { matchCase: false });
      found.load("items");
      await context.sync();
      const baseAlign = para.alignment || "Left";
      if (!found.items.length) return;
      const range = found.items[found.items.length - 1];
      const tokenAlone = (para.text || "").trim().toLowerCase() === rawToken.trim().toLowerCase();
      range.insertText("", "Replace");
      let firstPara = null, lastPara = null;
      for (let i = 0; i < cites.length; i++) {
        const cpara = i === 0 ? (tokenAlone ? para : para.insertParagraph("", "After")) : lastPara.insertParagraph("", "After");
        const cc = cpara.getRange("Start").insertContentControl();
        cc.tag = cites[i].tag;
        cc.title = cites[i].title || "Citation";
        cc.appearance = "Hidden";
        cc.insertText(cites[i].text, "Replace");
        applyFmt(cpara);
        if (!firstPara) firstPara = cpara;
        lastPara = cpara;
      }
      await context.sync(); // citations en place → on peut inspecter les paragraphes voisins
      await ensureBlankBefore(context, firstPara, baseAlign);
      const afterPara = await ensureBlankAfter(context, lastPara, baseAlign);
      afterPara.getRange("Start").select();
      await context.sync();
    });
  }

  // « /p5+ » → « Pièce n°5 et s. » (borne unique suivie par id → suit la renumérotation).
  async function startSlashEtSeq(rawStart, side) {
    const c = sideCfg(side);
    const dispStart = String(rawStart).replace(/(\d)([a-zA-Z])/g, "$1 $2");
    const startId = onSide(side, () => pieceIdByNumber(dispStart));
    if (!startId) return { error: dispStart };
    const num = String(numberOf(startId));
    await insertCitationsAtToken(c.token + rawStart + "+", [
      { tag: c.etseq + startId, title: c.many + " n°" + num + " et suivantes", text: formatEtSeq(num, side) },
    ], side);
    return { ok: true };
  }

  // « /p1,2,5 » : citation MULTIPLE de plusieurs pièces EXISTANTES → « Pièces n°1, 2 et 5 » (sans les
  // noms). Si une pièce n'existe pas → erreur, AUCUNE création. Suit la renumérotation (id-tracké).
  async function startSlashList(rawList, side) {
    const c = sideCfg(side);
    const toks = String(rawList).split(",").map((s) => s.trim().replace(/(\d)([a-zA-Z])/g, "$1 $2")).filter(Boolean);
    const ids = [];
    for (const t of toks) {
      const id = onSide(side, () => pieceIdByNumber(t));
      if (!id) return { error: t }; // pièce inexistante → on n'insère rien et on ne crée rien
      if (!ids.includes(id)) ids.push(id);
    }
    if (ids.length < 2) return { error: rawList };
    await insertCitationsAtToken(c.token + rawList, [
      { tag: c.list + ids.join("~"), title: "Citation multiple", text: formatList(ids.map((id) => String(numberOf(id))), side) },
    ], side);
    return { ok: true };
  }

  // Insère une citation d'une pièce EXISTANTE à la place du token tapé (« /p », « /pbail »…) — sert à
  // la citation par NOM depuis la palette.
  async function citeAtToken(rawToken, pieceId) {
    const piece = findPiece(pieceId);
    if (!piece) return { error: "notfound" };
    const side = sideOf(pieceId);
    const c = sideCfg(side);
    const num = numberOf(pieceId);
    await insertCitationsAtToken(rawToken, [
      { tag: c.cite + pieceId, title: c.one, text: formatCitation(num != null ? num : "?", piece.name, side) },
    ], side);
    return { ok: true };
  }
  // Vérifie si l'invite de nommage « Pièce n°<num> » d'une pièce en attente est ENCORE présente
  // dans le document. Sert au volet à détecter qu'une invite a été ANNULÉE (Ctrl+Z) sans nouvelle
  // commande → la pièce, jamais nommée ni citée, doit alors être nettoyée. En cas de doute
  // (erreur, recherche impossible), renvoie true → on NE supprime jamais par erreur.
  async function namingPromptPresent(rawNum, side) {
    const numStr = String(rawNum).replace(/(\d)([a-zA-Z])/g, "$1 $2");
    const needle = sideCfg(side).one + " n°" + numStr;
    let present = true;
    try {
      await Word.run(async (context) => {
        const found = context.document.body.search(needle, { matchCase: false });
        found.load("items");
        await context.sync();
        present = found.items.length > 0;
      });
    } catch (e) { present = true; }
    return present;
  }

  // Retire simplement le token tapé (« /p », « /pbail »…) — quand la palette est annulée.
  async function stripToken(rawToken) {
    await runUntracked(async (context) => {
      const para = context.document.getSelection().paragraphs.getFirst();
      const found = para.search(rawToken + " ", { matchCase: true });
      found.load("items");
      await context.sync();
      if (found.items.length) { found.items[found.items.length - 1].insertText("", "Replace"); await context.sync(); }
    });
  }

  // « /pA-B » (bornes ENTIÈRES). Trois cas :
  //  - toutes les pièces existent + style "inline"  → 1 citation « Pièces n°A à B » (suivie par id) ;
  //  - toutes les pièces existent + style "stacked" → 1 citation par pièce, empilées ;
  //  - au moins une pièce manque → création + NOMMAGE EN CHAÎNE (renvoie une chaîne à nommer).
  async function startSlashRange(rawStart, rawEnd, side) {
    const c = sideCfg(side);
    let a = parseInt(rawStart, 10), b = parseInt(rawEnd, 10);
    if (isNaN(a) || isNaN(b)) return { error: rawStart + "-" + rawEnd };
    if (a > b) { const t = a; a = b; b = t; }
    const nums = [];
    for (let n = a; n <= b; n++) nums.push(String(n));
    const token = c.token + rawStart + "-" + rawEnd + " ";
    const style = model.settings.rangeStyle || "inline";
    // Existence par NUMÉRO CALCULÉ (structuré) — le champ legacy `number` a disparu du modèle.
    const ids = onSide(side, () => nums.map((n) => pieceIdByNumber(n)));
    const allExist = ids.every(Boolean);

    // ---- CAS 1 : toutes les pièces existent → on CITE (compacte ou empilée) ----
    if (allExist) {
      const rawToken = c.token + rawStart + "-" + rawEnd;
      const cites = style === "inline"
        ? [{ tag: c.range + ids[0] + "~" + ids[ids.length - 1], title: "Citation", text: formatRange(nums[0], nums[nums.length - 1], side) }]
        : ids.map((id, i) => ({ tag: c.cite + id, title: c.one, text: formatCitation(nums[i], findPiece(id).name, side) }));
      await insertCitationsAtToken(rawToken, cites, side);
      return { mode: "cite" };
    }

    // ---- CAS 2 : au moins une pièce manque → création + nommage en chaîne ----
    const chain = [];
    onSide(side, () => {
      for (const n of nums) {
        const existId = pieceIdByNumber(n);
        let p = existId ? findPiece(existId) : null;
        // Pièces d'une plage /pA-B créées à des numéros explicites → VERROUILLÉES à ces numéros.
        if (!p) p = createPieceAt(n);
        chain.push({ id: p.id, num: n });
      }
    });
    await save();

    const applyFmt = citeFmtApplier(side);
    let baseAlign = "Left";
    await runUntracked(async (context) => {
      const sel = context.document.getSelection();
      const para = sel.paragraphs.getFirst();
      para.load("text,alignment,style,leftIndent,firstLineIndent,rightIndent,lineSpacing,spaceBefore,spaceAfter");
      para.font.load("name,size,bold,italic,underline,color,highlightColor");
      const found = para.search(token, { matchCase: false });
      found.load("items");
      await context.sync();
      baseAlign = para.alignment || "Left";
      namingRestore = captureParaFormat(para);
      if (!found.items.length) return;
      const range = found.items[found.items.length - 1];
      const tokenAlone = (para.text || "").trim().toLowerCase() === token.trim().toLowerCase();
      const first = chain[0];
      const firstName = (findPiece(first.id) || {}).name || "";
      let promptPara, promptRange;
      if (tokenAlone) {
        promptRange = range.insertText(c.one + " n°" + first.num + " : " + firstName, "Replace");
        promptPara = para;
      } else {
        range.insertText("", "Replace");
        promptPara = para.insertParagraph(c.one + " n°" + first.num + " : " + firstName, "After");
        promptRange = promptPara.getRange("Content");
      }
      applyFmt(promptPara);
      await context.sync();
      await ensureBlankBefore(context, promptPara, baseAlign);
      promptRange.select("End");
      await context.sync();
    });
    return { mode: "create", chain, baseAlign };
  }

  // Étape 2 : quand le curseur a quitté la ligne du prompt, lit le nom et finalise la citation.
  // Si nextNum est fourni (chaîne /pA-B en création), enchaîne AUSSITÔT le prompt suivant
  // sur la ligne du curseur ; sinon reprend le formatage normal du corps.
  async function finalizeSlashNamingChain(numberStr, pieceId, nextNum, nextName, baseAlign) {
    let name = null;
    const side = sideOf(pieceId);
    const c = sideCfg(side);
    const label = c.one + " n°";
    const prefix = label + numberStr + " : ";
    const applyFmt = citeFmtApplier(side);
    await runUntracked(async (context) => {
      const sel = context.document.getSelection();
      const cur = sel.paragraphs.getFirst(); // ligne où est le curseur (nouveau paragraphe après Entrée)
      const prev = cur.getPreviousOrNullObject();
      prev.load("text,isNullObject");
      await context.sync();
      if (prev.isNullObject) return;
      const ptext = prev.text || "";
      if (ptext.trim().indexOf(label + numberStr) !== 0) return; // pas encore la ligne du prompt
      const idx = ptext.indexOf(prefix);
      name = (idx >= 0 ? ptext.slice(idx + prefix.length) : ptext.replace(/^.*?:\s*/, "")).trim();
      const p = findPiece(pieceId);
      if (p) p.name = name;
      // Reconstruit la ligne du prompt en citation propre (formatée à la synchro).
      prev.clear();
      const cc = prev.getRange("Start").insertContentControl();
      cc.tag = c.cite + pieceId;
      cc.title = c.one;
      cc.appearance = "Hidden";
      cc.insertText(formatCitation(numberStr, p ? p.name : name, side), "Replace");
      // Format citation réappliqué explicitement : Word pour le web remet le texte à « normal » au
      // clear() (Windows gardait la mise en forme du paragraphe) → gras/italique perdus sinon.
      applyFmt(prev);
      if (nextNum != null) {
        // Enchaîne : la ligne du curseur devient le prompt de la pièce suivante.
        cur.insertText(label + nextNum + " : " + (nextName || ""), "Replace");
        applyFmt(cur);
        cur.getRange("End").select();
      } else {
        // Fin de chaîne : la ligne suivante reprend le format du paragraphe d'AVANT la commande…
        const restore = namingRestore;
        namingRestore = null;
        applyParaFormat(cur, restore, baseAlign);
        await context.sync();
        // …puis exactement UNE ligne vide sous la citation (réutilise `cur` si vide, retire les lignes
        // vides en trop — ex. commande tapée entre deux paragraphes déjà séparés d'une ligne vide).
        const after = await ensureBlankAfter(context, prev, baseAlign);
        // ensureBlankAfter ne remet que gras/italique/alignement : on réapplique le format complet, et on
        // RE-SÉLECTIONNE le début de la ligne — sinon Word pour le web garde le gras/italique de la frappe
        // précédente pour les caractères suivants.
        applyParaFormat(after, restore, baseAlign);
        after.getRange("Start").select();
        await context.sync();
        await ensureBlankBefore(context, prev, baseAlign);
      }
      await context.sync();
    });
    if (name != null) await save();
    return name;
  }

  // FILET ROBUSTE : cherche DANS TOUT LE DOCUMENT le paragraphe « Pièce n°<num> : <nom> »
  // (nom non vide, pas déjà géré, et PAS la ligne où est le curseur → on ne capture pas
  // pendant la frappe), l'enveloppe en citation gérée et affecte le nom à la pièce en attente.
  // Sert quand l'événement Word ne se déclenche pas après « Entrée » (poste peu fiable / 1er
  // chargement dans un document existant). Renvoie le nom capté, ou null.
  async function finalizePendingBySearch(numberStr, pieceId, baseAlign) {
    let name = null;
    let target = null;
    const c = sideCfg(sideOf(pieceId));
    const prefix = c.one + " n°" + numberStr + " : ";
    await runUntracked(async (context) => {
      const sel = context.document.getSelection();
      const cursorPara = sel.paragraphs.getFirst();
      cursorPara.load("text");
      const paras = context.document.body.paragraphs;
      paras.load("items/text");
      await context.sync();
      const cursorText = cursorPara.text || "";
      const items = paras.items;
      // Détection fiable du « déjà géré » : contrôles CONTENUS dans le paragraphe (évite l'imbrication).
      const innerCCs = items.map((p) => p.contentControls);
      innerCCs.forEach((cl) => cl.load("items/tag"));
      await context.sync();
      for (let i = 0; i < items.length; i++) {
        if (innerCCs[i].items.some((c) => c.tag && c.tag.indexOf("wp:") === 0)) continue;
        const t = items[i].text || "";
        if (t.trim().indexOf(c.one + " n°" + numberStr) !== 0) continue;
        if (t === cursorText) continue; // ligne du curseur : l'utilisateur tape peut-être encore
        const idx = t.indexOf(prefix);
        const nm = (idx >= 0 ? t.slice(idx + prefix.length) : t.replace(/^.*?:\s*/, "")).trim();
        if (!nm) continue; // pas encore de nom saisi
        name = nm;
        const p = findPiece(pieceId);
        if (p) p.name = name;
        const cc = items[i].getRange("Content").insertContentControl();
        cc.tag = c.cite + pieceId;
        cc.title = c.one;
        cc.appearance = "Hidden";
        target = items[i];
        break;
      }
      await context.sync();
      // Garantir UNE ligne vide sous la citation ainsi capturée (comme le chemin événement).
      if (target) {
        await ensureBlankAfter(context, target, baseAlign);
        await ensureBlankBefore(context, target, baseAlign);
      }
    });
    if (name != null) await save();
    return name;
  }

  // « /pf » : met la ligne au FORMAT citation (gras/italique/aligné…), sans rien capter.
  // La saisie reste du texte formaté ordinaire (aucun content control) → invisible pour
  // la renumérotation/le renommage. Renvoie l'alignement du corps pour la reprise après Entrée.
  // Mise en forme d'un paragraphe (style, retraits, interligne, police), pour la restituer plus tard.
  function captureParaFormat(para) {
    const f = para.font;
    return {
      style: para.style, alignment: para.alignment,
      leftIndent: para.leftIndent, firstLineIndent: para.firstLineIndent, rightIndent: para.rightIndent,
      lineSpacing: para.lineSpacing, spaceBefore: para.spaceBefore, spaceAfter: para.spaceAfter,
      font: { name: f.name, size: f.size, bold: f.bold, italic: f.italic, underline: f.underline, color: f.color, highlightColor: f.highlightColor },
    };
  }
  // Réapplique une mise en forme capturée. Les valeurs indéterminées (mise en forme mixte → null) sont
  // ignorées ; gras/italique/souligné retombent sur « normal » pour ne jamais garder l'aspect citation.
  function applyParaFormat(p, fmt, baseAlign) {
    const f = fmt || {};
    if (f.style) p.style = f.style;
    for (const k of ["leftIndent", "firstLineIndent", "rightIndent", "lineSpacing", "spaceBefore", "spaceAfter"]) {
      if (typeof f[k] === "number") p[k] = f[k];
    }
    p.alignment = f.alignment || baseAlign || "Left";
    const ft = f.font || {};
    p.font.bold = ft.bold === true;
    p.font.italic = ft.italic === true;
    p.font.underline = typeof ft.underline === "string" && ft.underline ? ft.underline : "None";
    if (ft.name) p.font.name = ft.name;
    if (typeof ft.size === "number") p.font.size = ft.size;
    if (ft.color) p.font.color = ft.color;
    p.font.highlightColor = ft.highlightColor || null;
  }

  // Renvoie { baseAlign, restore, anchor, endAnchor } : `restore` = mise en forme du texte AVANT le
  // « /pf » (rendue à la ligne suivante après Entrée) ; `anchor`/`endAnchor` = signets CACHÉS posés sur
  // la ligne /pf et sur le paragraphe qui la suivait, pour reconnaître à coup sûr les lignes créées
  // par « Entrée » entre les deux — même si l'utilisateur tape déjà dedans.
  async function startSlashFormat() {
    const token = "/pf ";
    const applyFmt = citeFmtApplier();
    const out = { baseAlign: "Left", restore: null, anchor: null, endAnchor: null };
    await runUntracked(async (context) => {
      const sel = context.document.getSelection();
      const para = sel.paragraphs.getFirst();
      para.load("text,alignment,style,leftIndent,firstLineIndent,rightIndent,lineSpacing,spaceBefore,spaceAfter");
      para.font.load("name,size,bold,italic,underline,color,highlightColor");
      const found = para.search(token, { matchCase: false });
      found.load("items");
      await context.sync();
      out.baseAlign = para.alignment || "Left";
      out.restore = captureParaFormat(para);
      if (!found.items.length) return;
      const range = found.items[found.items.length - 1];
      const tokenAlone = (para.text || "").trim().toLowerCase() === token.trim().toLowerCase();
      let fmtPara, fmtRange;
      if (tokenAlone) {
        fmtRange = range.insertText("", "Replace"); // ligne vidée, prête à la saisie libre
        fmtPara = para;
      } else {
        range.insertText("", "Replace"); // retire le token de la phrase (phrase intacte)
        fmtPara = para.insertParagraph("", "After");
        fmtRange = fmtPara.getRange("Content");
      }
      applyFmt(fmtPara);
      await context.sync();
      // Même espacement qu'une vraie citation : exactement une ligne vide au-dessus.
      await ensureBlankBefore(context, fmtPara, out.baseAlign);
      fmtRange.select("End");
      await context.sync();
      if (apiOK("1.4")) {
        try {
          const stamp = Date.now().toString(36);
          fmtPara.getRange("Whole").insertBookmark("_wpfmt" + stamp);
          const nx = fmtPara.getNextOrNullObject();
          nx.load("isNullObject");
          await context.sync();
          out.anchor = "_wpfmt" + stamp;
          if (!nx.isNullObject) {
            nx.getRange("Whole").insertBookmark("_wpfme" + stamp);
            await context.sync();
            out.endAnchor = "_wpfme" + stamp;
          }
        } catch (e) { /* signets indisponibles → repli sur l'heuristique */ }
      }
    });
    return out;
  }

  // Après « /pf » + Entrée : les lignes créées sous la ligne /pf (jusqu'au curseur) reprennent la mise
  // en forme du texte d'AVANT le /pf. Renvoie true quand c'est terminé (appliqué, ou abandonné parce que
  // le curseur est parti ailleurs / la ligne a été effacée) ; false tant qu'on est sur la ligne /pf.
  async function finalizeFormat(pending) {
    const pf = pending && typeof pending === "object" ? pending : { baseAlign: pending };
    if (!pf.anchor || !apiOK("1.4")) return finalizeFormatHeuristic(pf);
    let done = false;
    await runUntracked(async (context) => {
      const doc = context.document;
      const bm = doc.getBookmarkRangeOrNullObject(pf.anchor);
      const em = pf.endAnchor ? doc.getBookmarkRangeOrNullObject(pf.endAnchor) : null;
      bm.load("isNullObject");
      if (em) em.load("isNullObject");
      const cur = doc.getSelection().paragraphs.getFirst();
      await context.sync();
      const finish = async (pfPara) => {
        // …et une ligne vide sous la ligne /pf, comme sous une citation.
        if (pfPara) { try { await ensureBlankAfter(context, pfPara, pf.baseAlign); } catch (e) {} }
        doc.deleteBookmark(pf.anchor);
        if (pf.endAnchor) doc.deleteBookmark(pf.endAnchor);
        done = true;
        await context.sync();
      };
      if (bm.isNullObject) { await finish(null); return; } // ligne /pf effacée : rien à reprendre
      const pfPara = bm.paragraphs.getFirst();
      const rel = cur.getRange("Whole").compareLocationWith(pfPara.getRange("Whole"));
      const endOk = em && !em.isNullObject;
      const relEnd = endOk ? cur.getRange("Whole").compareLocationWith(em.paragraphs.getFirst().getRange("Whole")) : null;
      await context.sync();
      const after = rel.value === "After" || rel.value === "AdjacentAfter";
      if (!after) {
        if (rel.value === "Before" || rel.value === "AdjacentBefore") await finish(pfPara); // curseur remonté
        return; // toujours sur la ligne /pf : saisie en cours
      }
      const beforeEnd = !relEnd || relEnd.value === "Before" || relEnd.value === "AdjacentBefore";
      if (beforeEnd) {
        const zone = pfPara.getNext().getRange("Start").expandTo(cur.getRange("End"));
        const ps = zone.paragraphs;
        ps.load("items");
        await context.sync();
        // Sans repère de fin, prudence : on ne reformate que les quelques lignes juste créées.
        if (endOk || ps.items.length <= 3) for (const p of ps.items) applyParaFormat(p, pf.restore, pf.baseAlign);
      }
      await finish(pfPara);
    });
    return done;
  }
  // Retire les signets d'une reprise /pf abandonnée (nouvelle commande tapée entre-temps).
  async function dropFormatAnchor(pf) {
    if (!pf || !pf.anchor || !apiOK("1.4")) return;
    try {
      await Word.run(async (context) => {
        context.document.deleteBookmark(pf.anchor);
        if (pf.endAnchor) context.document.deleteBookmark(pf.endAnchor);
        await context.sync();
      });
    } catch (e) { /* signet déjà absent */ }
  }

  // Repli (API signets indisponible) : quand le curseur est passé sur une ligne VIDE, sous une ligne
  // au format citation, on lui rend la mise en forme d'avant le /pf.
  async function finalizeFormatHeuristic(pf) {
    const cite = model.settings.citation;
    let done = false;
    await runUntracked(async (context) => {
      const sel = context.document.getSelection();
      const cur = sel.paragraphs.getFirst();
      const prev = cur.getPreviousOrNullObject();
      cur.load("text");
      prev.load("text,isNullObject,alignment");
      prev.font.load("bold,italic");
      await context.sync();
      if (prev.isNullObject) return;
      if ((cur.text || "").length !== 0) return; // pas encore sur une ligne vide (saisie en cours)
      if (!(prev.text || "").trim()) return; // la ligne précédente doit contenir la citation saisie
      // La ligne précédente doit ressembler à une ligne AU FORMAT citation (sinon c'est du corps).
      const okBold = !cite.bold || prev.font.bold === true;
      const okItalic = !cite.italic || prev.font.italic === true;
      const okAlign = !cite.alignment || cite.alignment === "none" || prev.alignment === mapAlignment(cite.alignment);
      if (!(okBold && okItalic && okAlign)) return;
      applyParaFormat(cur, pf.restore, pf.baseAlign);
      done = true;
      await context.sync();
      try { await ensureBlankAfter(context, prev, pf.baseAlign); } catch (e) {}
    });
    return done;
  }

  // MODÈLE : « épinglé » = fixedNumber != null (numéro figé, ne coule pas) ; « protégé » = locked
  // (cadenas 🔒, exclu du verrou général « tout déverrouiller »). Les deux sont INDÉPENDANTS.

  // Numéro figé courant d'une pièce → chaîne à stocker dans fixedNumber (sous-pièce : son INDICE).
  function pinValueOf(p) {
    const cur = stats.numbers && stats.numbers.get(p.id);
    const curStr = cur != null ? String(cur) : "";
    if (p.parentId) { const m = curStr.match(/(\d+)\s*$/); return m ? m[1] : curStr; }
    return curStr;
  }

  // Éditer le badge = VERROUILLER au numéro saisi (badge gris). Vider = rendre COULANT (badge bleu).
  // Le numéro saisi peut changer le NIVEAU de la pièce (voir placeAt) : « 5.2 » la range dans le
  // groupe 5, « 3 » sur une sous-pièce la promeut pièce, etc.
  // Renvoie { ok, num } ou { error: "locked" | "hasSubs" | "groupTaken" | "notfound" }.
  async function setNumber(id, value) {
    const p = findPiece(id);
    if (!p) return { error: "notfound" };
    const side = sideOf(id);
    // Pièce adverse : son numéro est celui de l'adversaire → toujours explicite, jamais « verrouillé ».
    if (p.locked) return { error: "locked" }; // pièce VERROUILLÉE : il faut d'abord la déverrouiller
    const v = value == null ? "" : String(value).trim().replace(/(\d)([a-zA-Z])/g, "$1 $2");
    if (v === "") {
      if (side === "adv") return { error: "advEmpty" };
      p.fixedNumber = null; p.locked = false; // coulant
    } else {
      const wasLocked = p.locked;
      const r = onSide(side, () => placeAt(p, v));
      if (r && r.error) return r;
      if (side === "adv") p.locked = wasLocked; // adverse : modifier le numéro ne verrouille pas
    }
    await save();
    return { ok: true, num: onSide(side, computeStructuredNumbers).get(p.id) };
  }

  // Cadenas par pièce. VERROUILLER = épingle au numéro courant + protège (badge gris). DÉVERROUILLER :
  // une pièce NORMALE redevient coulante (badge bleu) ; une pièce SCANNÉE garde son numéro (pas de saut
  // automatique — on la déverrouille pour la modifier volontairement).
  async function toggleLock(id) {
    const p = findPiece(id);
    if (!p) return;
    if (sideOf(id) === "adv") {
      // Pièce adverse : le numéro est TOUJOURS celui de l'adversaire (jamais recalculé). Le verrou
      // protège seulement ce numéro d'une modification (badge hachuré, non éditable).
      p.locked = !p.locked;
      await save();
      return;
    }
    if (p.locked) {
      p.locked = false;
      if (!p.scanned) p.fixedNumber = null; // pièce normale → recoule (les scannées gardent leur n°)
    } else {
      if (p.fixedNumber == null) p.fixedNumber = pinValueOf(p);
      p.locked = true;
    }
    await save();
  }

  // Verrou GÉNÉRAL. shouldLock=true : épingle+protège les pièces visées. false : les rend COULANTES.
  // scope : "manual" (pièces créées dans WordPiece, et leurs groupes), "scanned" (pièces détectées par
  // « Rechercher les pièces existantes ») ou "all". Comme le cadenas individuel, déverrouiller une
  // pièce scannée lui laisse son numéro (aucun saut de numérotation).
  async function setAllLocks(shouldLock, scope) {
    const s = scope || "manual";
    for (const p of model.pieces) {
      const inScope = s === "all" || (s === "scanned" ? !!p.scanned : !p.scanned);
      if (!inScope) continue;
      if (shouldLock) {
        if (p.fixedNumber == null) p.fixedNumber = pinValueOf(p);
        p.locked = true;
      } else {
        p.locked = false;
        if (!p.scanned) p.fixedNumber = null;
      }
    }
    await save();
  }

  // ACTION « Numéroter par ordre d'apparition » : réordonne le volet selon l'ordre de 1re citation
  // dans le texte, et rend COULANTES les pièces non scannées (→ numéros 1,2,3… par apparition ;
  // les pièces scannées gardent leur numéro épinglé). Les groupes suivent l'apparition de leur 1re
  // sous-pièce ; chaque conteneur reste suivi de ses sous-pièces (triées elles aussi par apparition).
  async function renumberByAppearance() {
    const order = [];
    await Word.run(async (context) => {
      const ccs = context.document.body.contentControls;
      ccs.load("items/tag");
      await context.sync();
      for (const cc of ccs.items) {
        const tag = cc.tag || "";
        let ids = [];
        if (tag.indexOf(TAG_PREFIX) === 0) ids = [pieceIdFromTag(tag)];
        else if (tag.indexOf(ETSEQ_PREFIX) === 0) ids = [tag.slice(ETSEQ_PREFIX.length)];
        else if (tag.indexOf(RANGE_PREFIX) === 0) ids = tag.slice(RANGE_PREFIX.length).split("~");
        else if (tag.indexOf(LIST_PREFIX) === 0) ids = tag.slice(LIST_PREFIX.length).split("~");
        for (const id of ids) if (findPiece(id) && !order.includes(id)) order.push(id);
      }
    });
    const rank = new Map();
    order.forEach((id, i) => rank.set(id, i));
    const BIG = order.length + model.pieces.length + 1; // non citées → à la fin (ordre actuel préservé)
    const rankOf = (p) => (rank.has(p.id) ? rank.get(p.id) : BIG);
    const groupRank = (p) => {
      if (!p.container) return rankOf(p);
      let r = BIG;
      for (const k of childrenOf(p.id)) r = Math.min(r, rankOf(k));
      return r;
    };
    const stableSort = (arr, keyFn) => arr.map((p, i) => ({ p, i })).sort((a, b) => (keyFn(a.p) - keyFn(b.p)) || (a.i - b.i)).map((x) => x.p);
    // On préserve les pièces VERROUILLÉES (🔒) : elles gardent leur numéro figé. Les pièces scannées
    // sont verrouillées par défaut → protégées ; les déverrouiller permet de les renuméroter aussi.
    const keep = (p) => p.locked;
    const tops = stableSort(topPieces(), groupRank);
    const newArr = [];
    for (const top of tops) {
      if (!keep(top)) { top.fixedNumber = null; }
      newArr.push(top);
      for (const k of stableSort(childrenOf(top.id), rankOf)) {
        if (!keep(k)) { k.fixedNumber = null; }
        newArr.push(k);
      }
    }
    model.pieces = newArr;
    await save();
  }

  // Réordonne une pièce parmi ses FRÈRES, en ORDRE AFFICHÉ (par numéro). La voisine cible se cherche
  // parmi les frères DÉPLAÇABLES (non verrouillés) : on SAUTE les pièces verrouillées, qui sont des
  // murs à numéro figé (ex. 1,2,3🔒,4 → descendre la 2 l'échange avec la 4, pas avec la 3). Met à jour
  // les numéros EN MÉMOIRE aussitôt (le volet reflète le mouvement même si le sync du doc traîne).
  // Renvoie true seulement si un numéro affiché a réellement changé.
  async function movePieceStructured(id, dir) {
    const p = findPiece(id);
    if (!p || p.locked) return false; // une pièce verrouillée ne se déplace pas
    const nums = computeStructuredNumbers();
    const byNum = (a, b) => naturalCompare(nums.get(a.id), nums.get(b.id));
    const group = p.parentId ? childrenOf(p.parentId) : topPieces();
    const movable = group.filter((q) => !q.locked).sort(byNum); // frères déplaçables, ordre affiché
    const j = movable.indexOf(p) + dir;
    if (j < 0 || j >= movable.length) return false; // plus de pièce déplaçable dans cette direction
    const neighbor = movable[j];
    const snapshot = model.pieces.slice();
    // Déplace le bloc [p (+ ses sous-pièces)] juste avant/après le bloc [neighbor (+ les siennes)].
    const startA = model.pieces.indexOf(p);
    let endA = startA + 1;
    while (endA < model.pieces.length && model.pieces[endA].parentId === p.id) endA++;
    const block = model.pieces.splice(startA, endA - startA);
    const ni = model.pieces.indexOf(neighbor);
    if (dir > 0) {
      let ne = ni + 1;
      while (ne < model.pieces.length && model.pieces[ne].parentId === neighbor.id) ne++;
      model.pieces.splice(ne, 0, ...block);
    } else {
      model.pieces.splice(ni, 0, ...block);
    }
    const after = computeStructuredNumbers();
    let changed = false;
    for (const [pid, n] of after) { if (String(nums.get(pid)) !== String(n)) { changed = true; break; } }
    if (!changed) { model.pieces = snapshot; return false; }
    if (stats) stats.numbers = after; // MAJ immédiate de l'affichage
    await save();
    return true;
  }

  // Fusionne des pièces DOUBLONS dans keepId : redirige toutes les citations (piece/plage/etseq)
  // vers keepId par simple RE-TAG (sans réécrire le texte → faible risque), puis supprime les autres.
  async function mergePieces(keepId, otherIds) {
    const others = new Set((otherIds || []).filter((x) => x && x !== keepId));
    const keep = findPiece(keepId);
    if (!others.size || !keep) return;
    // Une variante « nom de la pièce gardée + suite » (« Contrat de bail, page 8 », « Contrat de bail
    // commercial ») garde son texte : la suite devient une précision propre à la citation.
    const extraOf = (otherId) => {
      const o = findPiece(otherId);
      return o ? (nameRemainder(keep.name, o.name) || "") : "";
    };
    await runUntracked(async (context) => {
      const ccs = context.document.body.contentControls;
      ccs.load("items/tag");
      await context.sync();
      const parents = ccs.items.map((cc) => cc.parentContentControlOrNullObject);
      parents.forEach((par) => par.load("tag,isNullObject"));
      await context.sync();
      for (let i = 0; i < ccs.items.length; i++) {
        const cc = ccs.items[i];
        if (!parents[i].isNullObject && isCiteTag(parents[i].tag)) continue; // citation imbriquée → ignorer
        const tag = cc.tag || "";
        if (tag.indexOf(TAG_PREFIX) === 0) {
          const sp = splitPieceTag(tag);
          if (others.has(sp.id)) cc.tag = pieceTagFor(keepId, sp.extra || extraOf(sp.id)); // le complément local suit
        } else if (tag.indexOf(RANGE_PREFIX) === 0) {
          const [a, b] = tag.slice(RANGE_PREFIX.length).split("~");
          const na = others.has(a) ? keepId : a, nb = others.has(b) ? keepId : b;
          if (na !== a || nb !== b) cc.tag = RANGE_PREFIX + na + "~" + nb;
        } else if (tag.indexOf(ETSEQ_PREFIX) === 0) {
          if (others.has(tag.slice(ETSEQ_PREFIX.length))) cc.tag = ETSEQ_PREFIX + keepId;
        } else if (tag.indexOf(LIST_PREFIX) === 0) {
          const parts = tag.slice(LIST_PREFIX.length).split("~");
          let changed = false;
          const np = parts.map((p) => { if (others.has(p)) { changed = true; return keepId; } return p; });
          if (changed) cc.tag = LIST_PREFIX + np.join("~");
        }
      }
      await context.sync();
    });
    removeFromModel([...others]);
    await save();
  }

  // ---------------- Compléments de citation (API) ----------------
  // Fixe (ou efface) la précision locale d'UNE occurrence, désignée par l'id de son contrôle.
  // Le nom de la pièce n'est pas touché : les autres citations et le bordereau restent intacts.
  async function setCitationExtra(ccId, extra) {
    await runUntracked(async (context) => {
      const cc = context.document.contentControls.getByIdOrNullObject(ccId);
      cc.load("isNullObject,tag");
      await context.sync();
      if (cc.isNullObject || !isSingleCiteTag(cc.tag)) return;
      const sp = splitPieceTag(cc.tag);
      cc.tag = pieceTagFor(sp.id, extra || "", sp.side);
      await context.sync();
    });
  }

  // ---------------- Avertissements ignorés ----------------
  // L'utilisateur peut masquer un point de diagnostic qu'il a arbitré (ex. un trou de numérotation
  // volontaire). La liste est stockée DANS LE DOCUMENT (elle voyage avec l'acte) et reste
  // réversible à tout moment via restoreWarnings().
  function ignoredList() {
    if (!Array.isArray(model.ignored)) model.ignored = [];
    return model.ignored;
  }
  function isWarningIgnored(key) { return ignoredList().indexOf(key) >= 0; }
  function ignoredCount() { return ignoredList().length; }
  async function ignoreWarning(key) {
    const a = ignoredList();
    if (a.indexOf(key) < 0) { a.push(key); await save(); }
  }
  // ---------------- Mise à niveau d'un document (une fois par format) ----------------
  // Contrôles de contenu « wp:… » laissés par d'anciennes versions et que celle-ci ne reconnaît
  // plus (ex. invites de nommage « wp:pnew: ») : on retire le contrôle en GARDANT son texte.
  // N'écrit (donc ne marque le document « modifié ») que s'il y en a.
  const KNOWN_TAG_PREFIXES = [...CITE_PREFIXES, "wp:name:"];
  const isKnownTag = (t) => t === BORDEREAU_TAG || KNOWN_TAG_PREFIXES.some((p) => t.indexOf(p) === 0);
  async function unwrapLegacyControls() {
    let n = 0;
    await runUntracked(async (context) => {
      const ccs = context.document.contentControls;
      ccs.load("items/tag");
      await context.sync();
      for (const cc of ccs.items) {
        const t = cc.tag || "";
        if (t.indexOf("wp:") !== 0 || isKnownTag(t)) continue;
        context.wpUntracked();
        cc.delete(true); // true = conserver le contenu
        n++;
      }
      if (n) await context.sync();
    }, { lazy: true });
    return n;
  }
  const DOC_FORMAT_KEY = "wordpiece.docFormat";
  function docFormat() {
    try { return Number(Office.context.document.settings.get(DOC_FORMAT_KEY)) || 0; } catch (e) { return 0; }
  }
  // Enregistre le modèle (migré) et le repère de format en un seul enregistrement.
  async function markDocFormat(v) {
    Office.context.document.settings.set(DOC_FORMAT_KEY, v);
    await save();
  }

  async function unignoreWarning(key) {
    const a = ignoredList();
    const i = a.indexOf(key);
    if (i >= 0) { a.splice(i, 1); await save(); }
  }
  async function restoreWarnings() { model.ignored = []; await save(); }

  // ---------------- Export ----------------
  window.WP = {
    TAG_PREFIX, RANGE_PREFIX, ETSEQ_PREFIX, LIST_PREFIX, BORDEREAU_TAG,
    ADV_PREFIX, ADV_RANGE_PREFIX, ADV_ETSEQ_PREFIX, ADV_LIST_PREFIX,
    sideOf, sidePieces, isSingleCiteTag, numberOf, nextAdverseNumber,
    parseAdverseList, planAdverseImport, importAdversePieces,
    get model() { return model; },
    set model(v) { model = v; },
    get stats() { return stats; },
    // persistance
    load, reload, save, saveGlobal, isDisabled, setDisabled,
    defaultSettings, defaultModel,
    // helpers
    findPiece, formatCitation, fillTemplate, buildBordereauHtml, escapeHtml,
    pieceIdFromTag, splitPieceTag, pieceTagFor, looksLikeExtra, extraAfter, setCitationExtra,
    isWarningIgnored, ignoreWarning, unignoreWarning, restoreWarnings, ignoredCount,
    extractNameFromCitation, extractNumberFromCitation, naturalCompare, sameDocName, normalizeStructure,
    // document
    sync, refreshStats, insertCitation, gotoPiece, focusSelection, focusApiSupported, selectOccurrence, generateBordereau, deleteBordereau, scanExistingPieces, unwrapCitations,
    getPieceIdAtSelection, removeCitationAtSelection,
    getSelectionContext, listOccurrences, selectCcById, removeCcById,
    unwrapLegacyControls, docFormat, markDocFormat,
    // mutations
    addPiece, renamePiece, deletePiece, deletePieceEverywhere, movePieceStructured, setNumber,
    toggleLock, setAllLocks, renumberByAppearance, computeStructuredNumbers, mergePieces,
    detectSlashAny, startSlashPrompt, startSlashEtSeq, startSlashRange, finalizeSlashNamingChain,
    startSlashNew, startSlashInsert, startSlashInsertShift, analyzeInsert, startSlashSubNew, startSlashSubInsert,
    _spacing: { ensureBlankBefore, ensureBlankAfter, deleteCitationLine }, // tests unitaires
    finalizePendingBySearch, startSlashFormat, finalizeFormat, dropFormatAnchor, startSlashList, citeAtToken, stripToken, namingPromptPresent, pieceIdByNumber, formatRange, formatEtSeq, formatList,
  };
})();
