/* global window */
// ============================================================
// LexRef — lecture et nettoyage d'un document Word (.docx), sans Word.
// S'appuie sur window.WPZip. Expose window.WPDocx :
//   readModel(entries)          → modèle LexRef enregistré dans le document, ou null
//   readTrailingList(entries)   → texte de la liste de pièces qui termine le document
//                                 (« Bordereau… » puis une ligne par pièce), ou ""
//   cleanPackage(entries, id)   → entrées d'une COPIE PROPRE : sans les contrôles LexRef
//                                 (le texte reste), sans les données ni le volet de l'extension
// Analyse par expressions régulières (et non DOMParser) : testable hors navigateur, et le XML
// de Word est assez régulier pour cela.
// ============================================================
(function (root) {
  const MODEL_KEY = "wordpiece.model.v1";
  const dec = (u8) => new TextDecoder("utf-8").decode(u8);
  const enc = (s) => new TextEncoder().encode(s);
  const find = (entries, name) => entries.find((e) => e.name === name);
  function unescapeXml(s) {
    return String(s)
      .replace(/&#x([0-9a-f]+);/gi, (m, h) => String.fromCodePoint(parseInt(h, 16)))
      .replace(/&#(\d+);/g, (m, d) => String.fromCodePoint(parseInt(d, 10)))
      .replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'")
      .replace(/&amp;/g, "&");
  }

  // ---------- Modèle LexRef (réglages Office du document : parties « webextension ») ----------
  function readModel(entries) {
    for (const e of entries) {
      if (!/^word\/webextensions\/webextension\d*\.xml$/i.test(e.name)) continue;
      const xml = dec(e.data);
      const re = /<we:property\b[^>]*\bname="([^"]*)"[^>]*\bvalue="([^"]*)"/g;
      let m;
      while ((m = re.exec(xml))) {
        if (unescapeXml(m[1]) !== MODEL_KEY) continue;
        try {
          let v = JSON.parse(unescapeXml(m[2]));
          if (typeof v === "string") v = JSON.parse(v); // valeur parfois sérialisée deux fois
          if (v && typeof v === "object" && Array.isArray(v.pieces)) return v;
        } catch (err) { /* propriété illisible : on tentera la lecture du texte */ }
      }
    }
    return null;
  }

  // ---------- Texte du document, ligne par ligne ----------
  // Un paragraphe = une ligne ; une ligne de tableau = ses cellules séparées par des tabulations
  // (le format que l'import sait déjà lire). Le texte supprimé en suivi des modifications
  // (w:delText) n'est pas lu.
  function documentLines(xml) {
    const lines = [];
    const paras = [];          // pile : paragraphes imbriqués (zones de texte)
    let row = null, cell = null;
    const tok = /<(\/?)(w:p|w:tr|w:tc|w:t|w:tab|w:br|w:cr)\b([^>]*?)(\/?)>|([^<]+)/g;
    let m, inText = false;
    while ((m = tok.exec(xml))) {
      if (m[5] !== undefined) { if (inText && paras.length) paras[paras.length - 1] += unescapeXml(m[5]); continue; }
      const close = m[1] === "/", tag = m[2], selfClose = m[4] === "/";
      if (tag === "w:t") { inText = !close && !selfClose; continue; }
      if (tag === "w:tab") { if (!close && paras.length) paras[paras.length - 1] += "\t"; continue; }
      if (tag === "w:br" || tag === "w:cr") { if (!close && paras.length) paras[paras.length - 1] += " "; continue; }
      if (tag === "w:p") {
        if (!close && !selfClose) paras.push("");
        else if (selfClose) { if (cell) cell.push(""); else if (!row) lines.push(""); }
        else {
          const text = paras.pop() || "";
          if (paras.length) continue; // paragraphe de zone de texte : rattaché au paragraphe parent
          if (cell) cell.push(text); else lines.push(text);
        }
        continue;
      }
      if (tag === "w:tr") { if (!close) row = []; else { if (row) lines.push(row.join("\t")); row = null; } continue; }
      if (tag === "w:tc") {
        if (!close && !selfClose) cell = [];
        else if (close) { if (row && cell) row.push(cell.join(" ").replace(/\s+/g, " ").trim()); cell = null; }
      }
    }
    return lines;
  }

  // Liste de pièces de FIN de document. Priorité au DERNIER titre de bordereau (« Bordereau de
  // pièces », « Liste des pièces », « Pièces communiquées »…) : tout ce qui le suit. À défaut, le
  // dernier bloc de lignes numérotées. L'analyse ligne à ligne (côté moteur) écarte le reste.
  const TITLE_RE = /^\s*(?:bordereau|liste\s+des\s+pi[eè]ces|inventaire\s+des\s+pi[eè]ces|pi[eè]ces\s+(?:communiqu|produites|vis[ée]es|jointes|vers[ée]es)|productions?\b)/i;
  const NUMBERED_RE = /^\s*(?:[-–—•·*▪►]\s*)?(?:pi[eè]ces?\s*)?(?:n\s*[°ºo]\.?\s*)?\d{1,4}(?:\.\d{1,3})*(?:\s*(?:bis|ter|quater))?(?![\d/])\s*(?:[)\].:_·•\-–—\t|]+|\s)/i;
  function readTrailingList(entries) {
    const doc = find(entries, "word/document.xml");
    if (!doc) return "";
    const lines = documentLines(dec(doc.data));
    let start = -1;
    for (let i = lines.length - 1; i >= 0; i--) {
      const t = lines[i].trim();
      if (t.length <= 90 && TITLE_RE.test(t)) { start = i + 1; break; }
    }
    if (start < 0) {
      let last = -1;
      for (let i = lines.length - 1; i >= 0; i--) if (NUMBERED_RE.test(lines[i])) { last = i; break; }
      if (last < 0) return "";
      start = last;
      // On remonte tant qu'on reste dans la liste : lignes numérotées, vides, ou suites d'intitulé.
      for (let i = last - 1; i >= 0; i--) {
        const t = lines[i];
        if (NUMBERED_RE.test(t) || !t.trim() || /^\s*[a-zà-ÿ(«"“]/.test(t)) start = i; else break;
      }
    }
    return lines.slice(start).join("\n").trim();
  }

  // ---------- Copie propre ----------
  // 1) Contrôles de contenu LexRef (balise « wp:… ») : on garde leur contenu, on retire l'enveloppe.
  //    Traitement du plus INTERNE au plus externe : un contrôle sans autre contrôle dedans est
  //    déballé (s'il est à nous) ou mis de côté (sinon), et on recommence jusqu'à épuisement.
  const HOLD = "\u0000SDT";
  function unwrapLexRefControls(xml) {
    const kept = [];
    const innermost = /<w:sdt\b(?:\s[^>]*)?>((?:(?!<w:sdt\b)[\s\S])*?)<\/w:sdt>/g;
    let prev;
    do {
      prev = xml;
      xml = xml.replace(innermost, (whole, inner) => {
        const pr = inner.match(/<w:sdtPr\b[\s\S]*?<\/w:sdtPr>/);
        const tag = pr && pr[0].match(/<w:tag\b[^>]*\bw:val="([^"]*)"/);
        if (tag && unescapeXml(tag[1]).indexOf("wp:") === 0) {
          const content = inner.match(/<w:sdtContent\b[^>]*>([\s\S]*)<\/w:sdtContent>/);
          return content ? content[1] : "";
        }
        kept.push(whole);
        return HOLD + (kept.length - 1) + "\u0000";
      });
    } while (xml !== prev);
    // Les contrôles qui ne sont pas à LexRef reprennent leur place ; ils peuvent en contenir
    // d'autres mis de côté avant eux, d'où la boucle.
    const back = /\u0000SDT(\d+)\u0000/g;
    while (back.test(xml)) { back.lastIndex = 0; xml = xml.replace(back, (m, i) => kept[+i]); }
    return xml;
  }

  // 2) Données et volet de l'extension (parties « webextension » qui la désignent).
  function cleanPackage(entries, addinId) {
    const ours = new Set();
    const idRe = addinId ? new RegExp('\\bid="\\{?' + String(addinId).replace(/[^0-9a-f-]/gi, "") + '\\}?"', "i") : null;
    for (const e of entries) {
      if (!/^word\/webextensions\/webextension\d*\.xml$/i.test(e.name)) continue;
      const xml = dec(e.data);
      if ((idRe && idRe.test(xml)) || /\bname="wordpiece\./.test(xml)) ours.add(e.name);
    }
    let out = entries.filter((e) => !ours.has(e.name));
    const short = (n) => n.replace(/^word\/webextensions\//, "");
    // Relations du volet vers les parties retirées, puis volets qui les utilisaient.
    const removedRels = new Set();
    const relsName = "word/webextensions/_rels/taskpanes.xml.rels";
    const rels = find(out, relsName);
    if (rels) {
      rels.data = enc(dec(rels.data).replace(/<Relationship\b[^>]*\/>/g, (r) => {
        const target = ((r.match(/\bTarget="([^"]*)"/) || [])[1] || "").replace(/^\.?\//, "");
        const id = (r.match(/\bId="([^"]*)"/) || [])[1];
        if ([...ours].some((n) => short(n) === target || n === target)) { removedRels.add(id); return ""; }
        return r;
      }));
    }
    const tpName = "word/webextensions/taskpanes.xml";
    const tp = find(out, tpName);
    const dropped = new Set(ours);
    if (tp) {
      const xml = dec(tp.data).replace(/<wetp:taskpane\b[\s\S]*?<\/wetp:taskpane>/g, (t) => {
        const rid = (t.match(/\br:id="([^"]*)"/) || [])[1];
        return removedRels.has(rid) ? "" : t;
      });
      if (/<wetp:taskpane\b/.test(xml)) tp.data = enc(xml);
      else {
        // Plus aucun volet : la liste des volets et ses relations disparaissent aussi.
        dropped.add(tpName); dropped.add(relsName);
        out = out.filter((e) => e.name !== tpName && e.name !== relsName);
        const rootRels = find(out, "_rels/.rels");
        if (rootRels) rootRels.data = enc(dec(rootRels.data).replace(/<Relationship\b[^>]*\bTarget="\/?word\/webextensions\/taskpanes\.xml"[^>]*\/>/g, ""));
      }
    }
    const ct = find(out, "[Content_Types].xml");
    if (ct) ct.data = enc(dec(ct.data).replace(/<Override\b[^>]*\bPartName="\/([^"]*)"[^>]*\/>/g, (o, part) => (dropped.has(part) ? "" : o)));
    // 3) Corps, en-têtes, pieds de page, notes : contrôles LexRef déballés.
    for (const e of out) {
      if (!/^word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$/i.test(e.name)) continue;
      const xml = dec(e.data);
      if (xml.indexOf('w:val="wp:') >= 0) e.data = enc(unwrapLexRefControls(xml));
    }
    return out;
  }

  root.WPDocx = { readModel, readTrailingList, cleanPackage, documentLines, unwrapLexRefControls };
})(typeof window !== "undefined" ? window : globalThis);
