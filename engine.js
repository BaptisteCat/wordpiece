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
    };
  }
  function defaultModel() {
    return { version: 3, settings: defaultSettings(), pieces: [] };
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
    // Réglages à DEUX couches : ceux du document s'ils existent, sinon les prefs globales (localStorage).
    const source = raw && raw.settings ? raw.settings : loadGlobal();
    model.settings = mergeSettings(defaultSettings(), normalizeLegacy(source));
    migratePieces(); // reprend les vieux docs au modèle structuré (pièces figées, non destructif)
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

  function save() {
    Office.context.document.settings.set(MODEL_KEY, model);
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

  // ---------------- Utilitaires ----------------
  function uid() { return "p_" + Date.now().toString(36) + Math.random().toString(36).slice(2, 7); }
  function findPiece(id) { return model.pieces.find((p) => p.id === id); }
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
  function runUntracked(body) {
    return Word.run(async (context) => {
      let restore = null;
      try {
        context.document.load("changeTrackingMode");
        await context.sync();
        trackingMode = context.document.changeTrackingMode || "Off";
        if (trackingMode !== "Off") {
          restore = trackingMode;
          context.document.changeTrackingMode = "Off";
          await context.sync();
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
  function fillTemplate(tpl, num, name) {
    return (tpl || "")
      .replace(/\{num\}/gi, String(num))
      .replace(/\{nom\}/gi, name || "")
      .replace(/\{name\}/gi, name || "")
      .replace(/\s+/g, " ")
      .trim();
  }
  function formatCitation(num, name) {
    return fillTemplate(model.settings.citationTemplate || "Pièce n°{num}", num, name);
  }
  // Citations multiples (plages / « et suivantes ») — phrasé juridique français fixe.
  function formatRange(startNum, endNum) {
    if (String(startNum) === String(endNum)) return "Pièce n°" + startNum;
    return "Pièces n°" + startNum + " à " + endNum;
  }
  function formatEtSeq(num) {
    return "Pièces n°" + num + " et suivantes";
  }
  // Citation multiple « Pièces n°1, 2 et 3 » (numéros triés ; « Pièces n°1 et 2 » à deux).
  function formatList(nums) {
    const s = [...nums].sort(naturalCompare);
    if (s.length === 1) return "Pièce n°" + s[0];
    return "Pièces n°" + s.slice(0, -1).join(", ") + " et " + s[s.length - 1];
  }
  // Retrouve l'id de la pièce portant ce numéro (chaîne), via les numéros calculés.
  // Si les stats ne sont pas encore peuplées (tout début de session), on calcule à la volée.
  function pieceIdByNumber(numStr) {
    const target = String(numStr).trim();
    const src = stats.numbers && stats.numbers.size ? stats.numbers : computeStructuredNumbers();
    for (const [id, n] of src) { if (String(n) === target) return id; }
    return null;
  }
  // Numéro AFFICHÉ d'une pièce (stats si disponibles, sinon calcul à la volée).
  function numberOf(id) {
    const n = stats.numbers && stats.numbers.get(id);
    return n != null ? n : computeStructuredNumbers().get(id);
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

  // Bordereau « liste » avec chaque NOM dans son propre content control (wp:name:<id>),
  // ce qui le rend éditable et détectable pour le renommage inline.
  function fillBordereauListEditable(bord, numbers) {
    const s = model.settings.bordereau;
    const ordered = orderedPieces(numbers);
    bord.clear();
    const title = bord.insertParagraph(s.title || "Bordereau de pièces", "Start");
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
      applyFont(labelRange, labelFmt);
      labelRange.font.size = listSize;
      // L'espace qui sépare du nom n'est ni souligné ni gras/italique.
      if (trail) {
        const tRange = para.insertText(trail, "End");
        tRange.font.underline = "None";
        tRange.font.bold = false;
        tRange.font.italic = false;
        tRange.font.size = listSize;
      }
      if (p.name) {
        const nameRange = para.insertText(p.name, "End");
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
  function extractNumberFromCitation(text) {
    const tpl = model.settings.citationTemplate || "Pièce n°{num}";
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
  function extractNameFromCitation(text, pieceId) {
    const tpl = model.settings.citationTemplate || "";
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

  const isCiteTag = (t) => !!t && (t.startsWith(TAG_PREFIX) || t.startsWith(RANGE_PREFIX) || t.startsWith(ETSEQ_PREFIX) || t.startsWith(LIST_PREFIX));
  // Signature du bordereau tel qu'il DEVRAIT être rendu : tant qu'elle ne change pas, on ne le
  // reconstruit pas (synchro différentielle). En mémoire seulement → 1 reconstruction par session max.
  let lastBordSig = null;
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
        const id = cc.tag.slice(TAG_PREFIX.length);
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
          const id = cc.tag.slice(TAG_PREFIX.length);
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
        if (healed) save();
      } catch (e) { /* la réparation est au mieux : ne jamais casser la synchro */ }

      const numbers = computeStructuredNumbers();
      const cite = model.settings.citation;

      const counts = {};
      let orphanCCs = 0;
      // SYNCHRO DIFFÉRENTIELLE : on ne réécrit une citation QUE si son texte doit changer (ou en
      // mode force, après un changement de réglages de mise en forme). Un document déjà à jour
      // n'est donc PAS modifié : pas de révisions parasites, pas de « document modifié », et la
      // pile Ctrl+Z de l'utilisateur reste intacte.
      const writeCite = (cc, text, title) => {
        if (!force && (cc.text || "") === text) return;
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
        cc.title = "Pièce supprimée";
        cc.appearance = "BoundingBox";
        cc.color = "#dc2626";
      };
      const countCite = (id) => { counts[id] = (counts[id] || 0) + 1; };

      for (const cc of pieceCCs) {
        const id = cc.tag.slice(TAG_PREFIX.length);
        const piece = findPiece(id);
        if (piece) {
          countCite(id);
          const num = numbers.get(id);
          writeCite(cc, formatCitation(num, piece.name), "Pièce " + num);
        } else markOrphan(cc);
      }
      // Citations multiples : plages « Pièces n°X à Y », « et s. », listes.
      for (const cc of ccs.items.filter((c) => c.tag && c.tag.startsWith(RANGE_PREFIX) && !nested.has(c))) {
        const [aId, bId] = cc.tag.slice(RANGE_PREFIX.length).split("~");
        const a = findPiece(aId), b = findPiece(bId);
        if (a && b) {
          countCite(aId); countCite(bId);
          let an = numbers.get(aId), bn = numbers.get(bId);
          if (naturalCompare(an, bn) > 0) { const t = an; an = bn; bn = t; }
          writeCite(cc, formatRange(String(an), String(bn)), "Pièces n°" + an + " à " + bn);
        } else markOrphan(cc);
      }
      for (const cc of ccs.items.filter((c) => c.tag && c.tag.startsWith(ETSEQ_PREFIX) && !nested.has(c))) {
        const aId = cc.tag.slice(ETSEQ_PREFIX.length);
        const a = findPiece(aId);
        if (a) {
          countCite(aId);
          const an = numbers.get(aId);
          writeCite(cc, formatEtSeq(String(an)), "Pièces n°" + an + " et suivantes");
        } else markOrphan(cc);
      }
      for (const cc of ccs.items.filter((c) => c.tag && c.tag.startsWith(LIST_PREFIX) && !nested.has(c))) {
        const ids = cc.tag.slice(LIST_PREFIX.length).split("~");
        if (ids.every((id) => findPiece(id))) {
          ids.forEach(countCite);
          const nums = ids.map((id) => String(numbers.get(id)));
          writeCite(cc, formatList(nums), "Citation multiple");
        } else markOrphan(cc);
      }

      // Bordereau DIFFÉRENTIEL : reconstruit seulement si son contenu attendu (pièces, numéros,
      // noms, réglages de présentation) a changé depuis la dernière reconstruction — ou en force.
      const bord = ccs.items.find((cc) => cc.tag === BORDEREAU_TAG);
      let bordSig = null;
      if (bord) {
        bordSig = JSON.stringify({
          s: model.settings.bordereau,
          rows: orderedPieces(numbers).filter((p) => !p.container)
            .map((p) => [String(numbers.get(p.id) ?? ""), p.name || "", p.parentId ? 1 : 0]),
        });
        if (force || bordSig !== lastBordSig) {
          if (model.settings.bordereau.layout === "table") bord.insertHtml(buildBordereauHtml(numbers), "Replace");
          else fillBordereauListEditable(bord, numbers);
        }
      }

      await context.sync();
      lastBordSig = bordSig; // mémorisé APRÈS un sync réussi (sinon le 2e essai sauterait le rendu)

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

      stats = { counts, orphanCCs, numbers, hasBordereau: !!bord, orderRank, tracking: trackingMode };
    });
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
          const id = tag.slice(TAG_PREFIX.length);
          if (findPiece(id) && !appearanceOrderIds.includes(id)) appearanceOrderIds.push(id);
          bump(id);
        } else if (tag.indexOf(ETSEQ_PREFIX) === 0) {
          bump(tag.slice(ETSEQ_PREFIX.length));
        } else if (tag.indexOf(RANGE_PREFIX) === 0) {
          const [a, b] = tag.slice(RANGE_PREFIX.length).split("~");
          bump(a); if (b && b !== a) bump(b);
        } else if (tag.indexOf(LIST_PREFIX) === 0) {
          tag.slice(LIST_PREFIX.length).split("~").forEach(bump);
        }
      }
    });
    const numbers = computeStructuredNumbers();
    const orderRank = new Map();
    let rank = 0;
    for (const id of appearanceOrderIds) if (findPiece(id) && !orderRank.has(id)) orderRank.set(id, rank++);
    for (const p of model.pieces) if (!orderRank.has(p.id)) orderRank.set(p.id, rank++);
    stats = { counts, orphanCCs, numbers, hasBordereau, orderRank, tracking: trackingMode };
    return stats;
  }

  // ---------------- Actions document ----------------
  async function insertCitation(id) {
    const piece = findPiece(id);
    if (!piece) return;
    const onNewLine = model.settings.citation.newLine !== false;
    await runUntracked(async (context) => {
      if (onNewLine) {
        // La citation va dans SON PROPRE paragraphe (l'alignement ne touche donc pas la phrase),
        // et le curseur repart sur une ligne normale en dessous.
        const para = context.document.getSelection().paragraphs.getLast();
        para.load("alignment");
        await context.sync();
        const baseAlign = para.alignment;

        const citePara = para.insertParagraph(formatCitation("?", piece.name), "After");
        const cc = citePara.getRange("Content").insertContentControl();
        cc.tag = TAG_PREFIX + id;
        cc.title = "Pièce";
        cc.appearance = "Hidden";

        const afterPara = citePara.insertParagraph("", "After");
        afterPara.font.bold = false;
        afterPara.font.italic = false;
        afterPara.font.underline = "None";
        afterPara.alignment = baseAlign || "Left";
        afterPara.getRange("Start").select();
        await context.sync();
      } else {
        const cc = context.document.getSelection().insertContentControl();
        cc.tag = TAG_PREFIX + id;
        cc.title = "Pièce";
        cc.appearance = "Hidden";
        cc.insertText(formatCitation("?", piece.name), "Replace");
        await context.sync();
      }
    });
  }

  async function gotoPiece(id) {
    await Word.run(async (context) => {
      const ccs = context.document.body.contentControls;
      ccs.load("items/tag");
      await context.sync();
      const cc = ccs.items.find((c) => c.tag === TAG_PREFIX + id);
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
      const matching = ccs.items.filter((c) => c.tag === TAG_PREFIX + pieceId);
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
  const SCAN_RE = /^\s*pi[eè]ces?\s+(?:n\s*[°ºo]?\s*)?(\d+(?:\.\d+)*(?:\s+(?:bis|ter|quater))?)\s*[)\].:_·•\-–—]+\s*(.+?)\s*$/i;
  // Construit une regex de scan à partir d'un GABARIT personnalisé « … {num} … {nom} … » (repli manuel
  // pour les formats exotiques). {num} avant {nom}. Renvoie null si le gabarit est inutilisable.
  function buildScanRegex(tpl) {
    if (!/\{num\}/i.test(tpl) || !/\{nom\}|\{name\}/i.test(tpl)) return null;
    let pattern = "";
    for (const tk of String(tpl).split(/(\{num\}|\{nom\}|\{name\})/i)) {
      if (/^\{num\}$/i.test(tk)) pattern += "(\\d+(?:\\.\\d+)*(?:\\s+(?:bis|ter|quater))?)";
      else if (/^\{nom\}$/i.test(tk) || /^\{name\}$/i.test(tk)) pattern += "(.+?)";
      else pattern += escapeRegex(tk).replace(/\s+/g, "\\s+");
    }
    try { return new RegExp("^\\s*" + pattern + "\\s*$", "i"); } catch (e) { return null; }
  }
  async function scanExistingPieces(customTemplate) {
    const re = (customTemplate && customTemplate.trim()) ? (buildScanRegex(customTemplate) || SCAN_RE) : SCAN_RE;
    const result = { wrapped: 0, newPieces: 0, conflicts: 0 };
    const wrappedCCs = []; // contrôles créés par ce scan (pour l'annulation)
    const newIds = [];     // pièces créées par ce scan (pour l'annulation)
    const byNum = new Map(); // numéro -> id (pièces existantes + créées pendant le scan)
    for (const p of model.pieces) { const n = pieceNumberOf(p); if (n) byNum.set(n, p.id); }

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

      for (let i = 0; i < items.length; i++) {
        // Déjà géré si le paragraphe contient un contrôle wp: (citation ou nom de bordereau).
        if (innerCCs[i].items.some((c) => c.tag && c.tag.indexOf("wp:") === 0)) continue;
        const m = (items[i].text || "").match(re);
        if (!m) continue;
        const rawNum = m[1].replace(/\s+/g, " ").trim();
        const name = m[2].replace(/\s+/g, " ").trim();
        let id = byNum.get(rawNum);
        const pc = id ? findPiece(id) : null;
        // CONFLIT : une pièce N existe DÉJÀ avec un nom DIFFÉRENT → on ne l'absorbe pas ; on crée
        // une pièce séparée (doublon de numéro signalé au diagnostic, à arbitrer par l'utilisateur).
        const conflict = pc && pc.name && pc.name.trim() && name && normScan(pc.name) !== normScan(name);
        if (!id || conflict) {
          // Pièce détectée avec un numéro explicite → reprise ÉPINGLÉE + PROTÉGÉE à ce numéro.
          // `scanned` : exclue du verrou général (elle reste figée à son numéro d'origine).
          const piece = makePiece({ name, locked: true, fixedNumber: rawNum, scanned: true });
          model.pieces.push(piece);
          if (!id) byNum.set(rawNum, piece.id); // le 1er occupant garde la clé ; un conflit ne l'écrase pas
          id = piece.id;
          newIds.push(piece.id);
          result.newPieces++;
          if (conflict) result.conflicts = (result.conflicts || 0) + 1;
        } else if (pc && !(pc.name && pc.name.trim()) && name) {
          pc.name = name; // même pièce : complète un nom vide
        }
        const cc = items[i].getRange("Content").insertContentControl();
        cc.tag = TAG_PREFIX + id;
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
      ccs.forEach((cc) => { if (!cc.isNullObject) cc.delete(false); }); // delete(false) = retire le contrôle, garde le texte
      await context.sync();
    });
  }

  async function getPieceIdAtSelection() {
    let id = null;
    await Word.run(async (context) => {
      const cc = context.document.getSelection().parentContentControlOrNullObject;
      cc.load("tag,isNullObject");
      await context.sync();
      if (!cc.isNullObject && cc.tag && cc.tag.startsWith(TAG_PREFIX)) id = cc.tag.slice(TAG_PREFIX.length);
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
        cc.delete(false); // supprime le contrôle ET son texte
        removed = true;
        await context.sync();
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
      if (!cc.isNullObject && cc.tag && cc.tag.startsWith(TAG_PREFIX)) {
        out.pieceId = cc.tag.slice(TAG_PREFIX.length);
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
      const matching = ccs.items.filter((cc) => cc.tag === TAG_PREFIX + pieceId);
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

  // ---------------- Mutations du modèle (persistées) ----------------
  async function addPiece(name) {
    const clean = (name || "").trim();
    if (!clean) return null;
    const piece = makePiece({ name: clean });
    model.pieces.push(piece);
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
    model.pieces = model.pieces.filter((p) => p.id !== id);
    await save();
  }

  // Supprime la pièce PARTOUT : retire ses citations du corps du texte, puis du modèle
  // (le bordereau et le volet se mettent à jour à la synchro suivante).
  // Un GROUPE (conteneur) emporte ses sous-pièces ; supprimer la dernière sous-pièce d'un groupe
  // supprime aussi le conteneur devenu vide.
  async function deletePieceEverywhere(id) {
    const target = findPiece(id);
    // Groupe → supprime d'abord toutes ses sous-pièces (avec leurs citations).
    if (target && target.container) {
      for (const kid of childrenOf(id)) await deletePieceEverywhere(kid.id);
    }
    const removeIds = [id];
    // Sous-pièce dont le groupe deviendra vide → on supprime aussi le conteneur.
    if (target && target.parentId) {
      const cont = findPiece(target.parentId);
      if (cont && cont.container && childrenOf(cont.id).filter((k) => k.id !== id).length === 0) removeIds.push(cont.id);
    }
    await runUntracked(async (context) => {
      const ccs = context.document.body.contentControls;
      ccs.load("items/tag,items/id");
      await context.sync();
      const matching = ccs.items.filter((cc) => removeIds.some((rid) => cc.tag === TAG_PREFIX + rid));
      const paras = matching.map((cc) => cc.getRange("Whole").paragraphs.getFirst());
      matching.forEach((cc) => cc.load("text"));
      paras.forEach((p) => p.load("text"));
      await context.sync();
      for (let i = 0; i < matching.length; i++) {
        const ccText = (matching[i].text || "").trim();
        const paraText = (paras[i].text || "").trim();
        // Citation seule sur sa ligne → on supprime la ligne ; sinon juste la citation.
        if (ccText.length > 0 && paraText === ccText) paras[i].delete();
        else matching[i].delete(false);
      }
      await context.sync();
    });
    model.pieces = model.pieces.filter((p) => !removeIds.includes(p.id));
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
      if (out) out._sig = t;
    });
    return out;
  }

  // Garantit UNE ligne vide juste AVANT le paragraphe `p`. N'ajoute rien si la ligne
  // précédente est déjà vide (pas de doublon) ni si `p` est en tout début de document.
  // `context` doit être déjà synchronisé (p en place). Fait sa propre synchro.
  async function ensureBlankBefore(context, p, baseAlign) {
    const prev = p.getPreviousOrNullObject();
    prev.load("text,isNullObject");
    await context.sync();
    if (prev.isNullObject || (prev.text || "").trim() === "") return; // début de doc ou déjà vide
    const b = p.insertParagraph("", "Before");
    b.font.bold = false; b.font.italic = false; b.font.underline = "None";
    if (baseAlign) b.alignment = baseAlign;
  }

  // Renvoie UNE ligne vide juste APRÈS le paragraphe `p` : réutilise la ligne suivante
  // si elle est DÉJÀ vide (pas de doublon), sinon en crée une. `context` déjà synchronisé.
  async function ensureBlankAfter(context, p, baseAlign) {
    const next = p.getNextOrNullObject();
    next.load("text,isNullObject");
    await context.sync();
    let after;
    if (!next.isNullObject && (next.text || "").trim() === "") {
      after = next; // ligne vide déjà présente → on la réutilise (pas de 2e ligne vide)
    } else {
      after = p.insertParagraph("", "After");
    }
    after.font.bold = false; after.font.italic = false; after.font.underline = "None";
    if (baseAlign) after.alignment = baseAlign;
    return after;
  }

  // Étape 1 : remplace « /pN » par « Pièce n°N :  » DÉJÀ FORMATÉE, curseur prêt pour le nom.
  // Si la pièce existe déjà, insère simplement une citation normale (pas de nommage).
  async function startSlashPrompt(rawNum) {
    // Numéro affiché : « 1bis » → « 1 bis » (espace entre chiffre et lettres). « 1.1 » inchangé.
    const numberStr = String(rawNum).replace(/(\d)([a-zA-Z])/g, "$1 $2");
    // Résolution par NUMÉRO CALCULÉ (structuré) : cite la pièce affichant ce numéro si elle existe.
    const existingId = pieceIdByNumber(numberStr);
    let piece = existingId ? findPiece(existingId) : null;
    // On (re)propose le nom si la pièce n'existe pas OU existe mais n'a pas encore de nom.
    const naming = !piece || !piece.name;
    if (!piece) {
      // Création AU numéro demandé → pièce VERROUILLÉE (n'entraîne aucun décalage).
      const subMatch = numberStr.match(/^(\d+)\.(\d+)$/);
      const parentId = subMatch ? pieceIdByNumber(subMatch[1]) : null;
      if (subMatch && parentId) {
        piece = makePiece({ parentId, locked: true, fixedNumber: subMatch[2] });
        insertPieceBefore(piece, afterChildrenBlockId(parentId));
      } else {
        piece = makePiece({ locked: true, fixedNumber: numberStr });
        model.pieces.push(piece);
      }
      await save();
    }
    if (naming) {
      const baseAlign = await insertNamingPromptFor("/p" + rawNum, numberStr);
      return { id: piece.id, num: numberStr, naming: true, baseAlign };
    }
    // Pièce existante déjà nommée → citation directe, au texte FINAL (la synchro n'aura rien à réécrire).
    await insertCitationsAtToken("/p" + rawNum, [
      { tag: TAG_PREFIX + piece.id, title: "Pièce", text: formatCitation(numberStr, piece.name) },
    ]);
    return { id: piece.id, num: numberStr, naming: false, baseAlign: "Left" };
  }

  // Écrit le prompt de nommage « Pièce n°<disp> :  » à la place du token, formaté, curseur prêt.
  // Factorisé pour /pn, /pnN, /pn4., /pn4.2 (le token complet est passé tel quel).
  async function insertNamingPromptFor(rawToken, dispNum) {
    const applyFmt = citeFmtApplier();
    let baseAlign = "Left";
    await runUntracked(async (context) => {
      const sel = context.document.getSelection();
      const para = sel.paragraphs.getFirst();
      para.load("text,alignment");
      const found = para.search(rawToken + " ", { matchCase: true });
      found.load("items");
      await context.sync();
      baseAlign = para.alignment || "Left";
      if (!found.items.length) return;
      const range = found.items[found.items.length - 1];
      const tokenAlone = (para.text || "").trim() === rawToken;
      let promptPara, promptRange;
      if (tokenAlone) {
        promptRange = range.insertText("Pièce n°" + dispNum + " : ", "Replace");
        promptPara = para;
      } else {
        range.insertText("", "Replace");
        promptPara = para.insertParagraph("Pièce n°" + dispNum + " : ", "After");
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
  async function startSlashNew(token) {
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
    const id = pieceIdByNumber(N);
    const p = id ? findPiece(id) : null;
    if (p && p.parentId) return null;      // N est une sous-pièce → cible invalide
    if (p && p.container) return p;         // groupe déjà existant → on ajoute simplement dedans
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
    const container = ensureGroup(rawParent);
    if (!container) return { error: "noparent", num: rawParent + "." + rawSub };
    const S = parseInt(rawSub, 10);
    const kids = childrenOf(container.id);
    const nums = computeStructuredNumbers();
    const parentDisp = nums.get(container.id);
    if (lockedIntSet(kids).has(S)) return { error: "locked", num: parentDisp + "." + S };
    let beforeId;
    const target = kids.find((k) => nums.get(k.id) === parentDisp + "." + S);
    if (target) {
      if (target.fixedNumber != null) return { error: "locked", num: parentDisp + "." + S };
      beforeId = target.id;
    } else {
      if (S !== nextSubIndex(container.id)) return { error: "range", num: parentDisp + "." + S }; // hors suite
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
  function citeFmtApplier() {
    const cite = model.settings.citation;
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
  async function insertCitationsAtToken(rawToken, cites) {
    const token = rawToken + " ";
    const applyFmt = citeFmtApplier();
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
  async function startSlashEtSeq(rawStart) {
    const dispStart = String(rawStart).replace(/(\d)([a-zA-Z])/g, "$1 $2");
    const startId = pieceIdByNumber(dispStart);
    if (!startId) return { error: dispStart };
    const num = String(numberOf(startId));
    await insertCitationsAtToken("/p" + rawStart + "+", [
      { tag: ETSEQ_PREFIX + startId, title: "Pièces n°" + num + " et suivantes", text: formatEtSeq(num) },
    ]);
    return { ok: true };
  }

  // « /p1,2,5 » : citation MULTIPLE de plusieurs pièces EXISTANTES → « Pièces n°1, 2 et 5 » (sans les
  // noms). Si une pièce n'existe pas → erreur, AUCUNE création. Suit la renumérotation (id-tracké).
  async function startSlashList(rawList) {
    const toks = String(rawList).split(",").map((s) => s.trim().replace(/(\d)([a-zA-Z])/g, "$1 $2")).filter(Boolean);
    const ids = [];
    for (const t of toks) {
      const id = pieceIdByNumber(t);
      if (!id) return { error: t }; // pièce inexistante → on n'insère rien et on ne crée rien
      if (!ids.includes(id)) ids.push(id);
    }
    if (ids.length < 2) return { error: rawList };
    await insertCitationsAtToken("/p" + rawList, [
      { tag: LIST_PREFIX + ids.join("~"), title: "Citation multiple", text: formatList(ids.map((id) => String(numberOf(id)))) },
    ]);
    return { ok: true };
  }

  // Insère une citation d'une pièce EXISTANTE à la place du token tapé (« /p », « /pbail »…) — sert à
  // la citation par NOM depuis la palette.
  async function citeAtToken(rawToken, pieceId) {
    const piece = findPiece(pieceId);
    if (!piece) return { error: "notfound" };
    const num = numberOf(pieceId);
    await insertCitationsAtToken(rawToken, [
      { tag: TAG_PREFIX + pieceId, title: "Pièce", text: formatCitation(num != null ? num : "?", piece.name) },
    ]);
    return { ok: true };
  }
  // Vérifie si l'invite de nommage « Pièce n°<num> » d'une pièce en attente est ENCORE présente
  // dans le document. Sert au volet à détecter qu'une invite a été ANNULÉE (Ctrl+Z) sans nouvelle
  // commande → la pièce, jamais nommée ni citée, doit alors être nettoyée. En cas de doute
  // (erreur, recherche impossible), renvoie true → on NE supprime jamais par erreur.
  async function namingPromptPresent(rawNum) {
    const numStr = String(rawNum).replace(/(\d)([a-zA-Z])/g, "$1 $2");
    const needle = "Pièce n°" + numStr;
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
  async function startSlashRange(rawStart, rawEnd) {
    let a = parseInt(rawStart, 10), b = parseInt(rawEnd, 10);
    if (isNaN(a) || isNaN(b)) return { error: rawStart + "-" + rawEnd };
    if (a > b) { const t = a; a = b; b = t; }
    const nums = [];
    for (let n = a; n <= b; n++) nums.push(String(n));
    const token = "/p" + rawStart + "-" + rawEnd + " ";
    const style = model.settings.rangeStyle || "inline";
    // Existence par NUMÉRO CALCULÉ (structuré) — le champ legacy `number` a disparu du modèle.
    const ids = nums.map((n) => pieceIdByNumber(n));
    const allExist = ids.every(Boolean);

    // ---- CAS 1 : toutes les pièces existent → on CITE (compacte ou empilée) ----
    if (allExist) {
      const rawToken = "/p" + rawStart + "-" + rawEnd;
      const cites = style === "inline"
        ? [{ tag: RANGE_PREFIX + ids[0] + "~" + ids[ids.length - 1], title: "Citation", text: formatRange(nums[0], nums[nums.length - 1]) }]
        : ids.map((id, i) => ({ tag: TAG_PREFIX + id, title: "Pièce", text: formatCitation(nums[i], findPiece(id).name) }));
      await insertCitationsAtToken(rawToken, cites);
      return { mode: "cite" };
    }

    // ---- CAS 2 : au moins une pièce manque → création + nommage en chaîne ----
    const chain = [];
    for (const n of nums) {
      const existId = pieceIdByNumber(n);
      let p = existId ? findPiece(existId) : null;
      // Pièces d'une plage /pA-B créées à des numéros explicites → VERROUILLÉES à ces numéros.
      if (!p) { p = makePiece({ locked: true, fixedNumber: n }); model.pieces.push(p); }
      chain.push({ id: p.id, num: n });
    }
    await save();

    const applyFmt = citeFmtApplier();
    let baseAlign = "Left";
    await runUntracked(async (context) => {
      const sel = context.document.getSelection();
      const para = sel.paragraphs.getFirst();
      para.load("text,alignment");
      const found = para.search(token, { matchCase: false });
      found.load("items");
      await context.sync();
      baseAlign = para.alignment || "Left";
      if (!found.items.length) return;
      const range = found.items[found.items.length - 1];
      const tokenAlone = (para.text || "").trim().toLowerCase() === token.trim().toLowerCase();
      const first = chain[0];
      const firstName = (findPiece(first.id) || {}).name || "";
      let promptPara, promptRange;
      if (tokenAlone) {
        promptRange = range.insertText("Pièce n°" + first.num + " : " + firstName, "Replace");
        promptPara = para;
      } else {
        range.insertText("", "Replace");
        promptPara = para.insertParagraph("Pièce n°" + first.num + " : " + firstName, "After");
        promptRange = promptPara.getRange("Content");
      }
      applyFmt(promptPara);
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
    const prefix = "Pièce n°" + numberStr + " : ";
    const applyFmt = citeFmtApplier();
    await runUntracked(async (context) => {
      const sel = context.document.getSelection();
      const cur = sel.paragraphs.getFirst(); // ligne où est le curseur (nouveau paragraphe après Entrée)
      const prev = cur.getPreviousOrNullObject();
      prev.load("text,isNullObject");
      await context.sync();
      if (prev.isNullObject) return;
      const ptext = prev.text || "";
      if (ptext.trim().indexOf("Pièce n°" + numberStr) !== 0) return; // pas encore la ligne du prompt
      const idx = ptext.indexOf(prefix);
      name = (idx >= 0 ? ptext.slice(idx + prefix.length) : ptext.replace(/^.*?:\s*/, "")).trim();
      const p = findPiece(pieceId);
      if (p) p.name = name;
      // Reconstruit la ligne du prompt en citation propre (formatée à la synchro).
      prev.clear();
      const cc = prev.getRange("Start").insertContentControl();
      cc.tag = TAG_PREFIX + pieceId;
      cc.title = "Pièce";
      cc.appearance = "Hidden";
      cc.insertText(formatCitation(numberStr, p ? p.name : name), "Replace");
      if (nextNum != null) {
        // Enchaîne : la ligne du curseur devient le prompt de la pièce suivante.
        cur.insertText("Pièce n°" + nextNum + " : " + (nextName || ""), "Replace");
        applyFmt(cur);
        cur.getRange("End").select();
      } else {
        // Fin de chaîne : retour au formatage normal du document…
        cur.font.bold = false;
        cur.font.italic = false;
        cur.font.underline = "None";
        if (baseAlign) cur.alignment = baseAlign;
        await context.sync();
        // …puis garantir UNE ligne vide sous la citation (réutilise `cur` si déjà vide → pas de doublon).
        await ensureBlankAfter(context, prev, baseAlign);
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
    const prefix = "Pièce n°" + numberStr + " : ";
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
        if (t.trim().indexOf("Pièce n°" + numberStr) !== 0) continue;
        if (t === cursorText) continue; // ligne du curseur : l'utilisateur tape peut-être encore
        const idx = t.indexOf(prefix);
        const nm = (idx >= 0 ? t.slice(idx + prefix.length) : t.replace(/^.*?:\s*/, "")).trim();
        if (!nm) continue; // pas encore de nom saisi
        name = nm;
        const p = findPiece(pieceId);
        if (p) p.name = name;
        const cc = items[i].getRange("Content").insertContentControl();
        cc.tag = TAG_PREFIX + pieceId;
        cc.title = "Pièce";
        cc.appearance = "Hidden";
        target = items[i];
        break;
      }
      await context.sync();
      // Garantir UNE ligne vide sous la citation ainsi capturée (comme le chemin événement).
      if (target) await ensureBlankAfter(context, target, baseAlign);
    });
    if (name != null) await save();
    return name;
  }

  // « /pf » : met la ligne au FORMAT citation (gras/italique/aligné…), sans rien capter.
  // La saisie reste du texte formaté ordinaire (aucun content control) → invisible pour
  // la renumérotation/le renommage. Renvoie l'alignement du corps pour la reprise après Entrée.
  async function startSlashFormat() {
    const token = "/pf ";
    const applyFmt = citeFmtApplier();
    let baseAlign = "Left";
    await runUntracked(async (context) => {
      const sel = context.document.getSelection();
      const para = sel.paragraphs.getFirst();
      para.load("text,alignment");
      const found = para.search(token, { matchCase: false });
      found.load("items");
      await context.sync();
      baseAlign = para.alignment || "Left";
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
      fmtRange.select("End");
      await context.sync();
    });
    return { baseAlign };
  }

  // Après « /pf » : quand le curseur est passé sur une ligne VIDE, sous une ligne mise
  // au format citation, on reprend le formatage du corps (comme après le nommage d'une pièce).
  async function finalizeFormat(baseAlign) {
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
      cur.font.bold = false;
      cur.font.italic = false;
      cur.font.underline = "None";
      if (baseAlign) cur.alignment = baseAlign;
      done = true;
      await context.sync();
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
  async function setNumber(id, value) {
    const p = findPiece(id);
    if (!p) return;
    if (p.locked) return; // pièce VERROUILLÉE : numéro non modifiable (il faut d'abord la déverrouiller)
    const v = value == null ? "" : String(value).trim();
    if (v === "") { p.fixedNumber = null; p.locked = false; await save(); return; } // coulant
    if (p.parentId) { const m = v.match(/(\d+)\s*$/); p.fixedNumber = m ? m[1] : v; }
    else { p.fixedNumber = v; }
    p.locked = true; // un numéro saisi à la main est figé (protégé)
    await save();
  }

  // Cadenas par pièce. VERROUILLER = épingle au numéro courant + protège (badge gris). DÉVERROUILLER :
  // une pièce NORMALE redevient coulante (badge bleu) ; une pièce SCANNÉE garde son numéro (pas de saut
  // automatique — on la déverrouille pour la modifier volontairement).
  async function toggleLock(id) {
    const p = findPiece(id);
    if (!p) return;
    if (p.locked) {
      p.locked = false;
      if (!p.scanned) p.fixedNumber = null; // pièce normale → recoule (les scannées gardent leur n°)
    } else {
      if (p.fixedNumber == null) p.fixedNumber = pinValueOf(p);
      p.locked = true;
    }
    await save();
  }

  // Verrou GÉNÉRAL. shouldLock=true : épingle+protège toutes les pièces. false : rend toutes les
  // pièces COULANTES (auto). Dans les deux cas, les pièces SCANNÉES sont épargnées (jamais touchées).
  async function setAllLocks(shouldLock) {
    for (const p of model.pieces) {
      if (p.scanned) continue;
      if (shouldLock) {
        if (p.fixedNumber == null) p.fixedNumber = pinValueOf(p);
        p.locked = true;
      } else {
        p.locked = false;
        p.fixedNumber = null;
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
        if (tag.indexOf(TAG_PREFIX) === 0) ids = [tag.slice(TAG_PREFIX.length)];
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
    if (!others.size || !findPiece(keepId)) return;
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
          if (others.has(tag.slice(TAG_PREFIX.length))) cc.tag = TAG_PREFIX + keepId;
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
    model.pieces = model.pieces.filter((p) => !others.has(p.id));
    await save();
  }

  // ---------------- Export ----------------
  window.WP = {
    TAG_PREFIX, RANGE_PREFIX, ETSEQ_PREFIX, LIST_PREFIX, BORDEREAU_TAG,
    get model() { return model; },
    set model(v) { model = v; },
    get stats() { return stats; },
    // persistance
    load, reload, save, saveGlobal, isDisabled, setDisabled,
    defaultSettings, defaultModel,
    // helpers
    findPiece, formatCitation, fillTemplate, buildBordereauHtml, escapeHtml,
    extractNameFromCitation, extractNumberFromCitation, naturalCompare,
    // document
    sync, refreshStats, insertCitation, gotoPiece, selectOccurrence, generateBordereau, deleteBordereau, scanExistingPieces, unwrapCitations,
    getPieceIdAtSelection, removeCitationAtSelection,
    getSelectionContext, listOccurrences, selectCcById, removeCcById,
    // mutations
    addPiece, renamePiece, deletePiece, deletePieceEverywhere, movePieceStructured, setNumber,
    toggleLock, setAllLocks, renumberByAppearance, computeStructuredNumbers, mergePieces,
    detectSlashAny, startSlashPrompt, startSlashEtSeq, startSlashRange, finalizeSlashNamingChain,
    startSlashNew, startSlashInsert, startSlashInsertShift, analyzeInsert, startSlashSubNew, startSlashSubInsert,
    finalizePendingBySearch, startSlashFormat, finalizeFormat, startSlashList, citeAtToken, stripToken, namingPromptPresent, pieceIdByNumber, formatRange, formatEtSeq, formatList,
  };
})();
