/* global WP, el, render, withBusy, syncDoc, snapshotModel, commitUndo, setPaneSide, toast, reportError, isLetterDoc, flashIds:writable */
// ============================================================
// LexRef — volet : import de pièces (menu ☰, ou bouton de l'onglet adverse).
// Source : une liste COLLÉE, ou un DOCUMENT WORD (acte précédent, conclusions adverses…) dont on
// reprend la liste LexRef ou, à défaut, le bordereau qui le termine. Destination au choix : nos
// pièces ou les pièces adverses. Aperçu cochable ; « Importer » crée les pièces (verrouillées à leur
// numéro) et renomme celles cochées dont le nom diffère.
// Chargé AVANT taskpane.js (qui appelle bindImportUI au démarrage).
// ============================================================
let importDest = "own";     // "own" (nos pièces / pièces jointes) | "adv"
let importPlan = [];
let importFileInfo = null;  // { name, lexref, own, adv, list } — document Word choisi
let importFilledText = "";  // texte posé depuis le document (une saisie de l'utilisateur le remplace)

function openImport(dest) {
  closeAdvImportToggle(true);
  el("importPanel").classList.remove("hidden");
  setImportDest(dest || "own");
  el("importPanel").scrollIntoView({ block: "nearest" });
  el("importText").focus();
}
function closeImport() {
  const box = el("importPanel");
  if (!box) return;
  box.classList.add("hidden");
  closeAdvImportToggle(false);
  el("importText").value = "";
  el("importFile").value = "";
  importFileInfo = null; importFilledText = ""; importPlan = [];
  renderImportSource();
  renderImportPreview();
}
// Le bouton « Coller la liste des pièces adverses » (onglet adverse) s'efface tant que le panneau est ouvert.
function closeAdvImportToggle(hide) {
  const t = el("advImportToggle");
  if (t) t.classList.toggle("hidden", hide);
}
function setImportDest(dest) {
  importDest = dest === "adv" ? "adv" : "own";
  const panel = el("importPanel");
  panel.classList.toggle("dest-adv", importDest === "adv");
  panel.querySelectorAll(".seg-btn[data-dest]").forEach((b) => b.setAttribute("aria-pressed", String(b.dataset.dest === importDest)));
  panel.querySelector('.seg-btn[data-dest="own"]').textContent = isLetterDoc() ? "Pièces jointes" : "Nos pièces";
  // « Déjà communiquées » : nos pièces d'un ACTE seulement.
  el("importCommRow").classList.toggle("hidden", importDest === "adv" || isLetterDoc());
  // Document Word déjà lu : sa liste pour la nouvelle destination (sauf si l'utilisateur a retouché le texte).
  if (importFileInfo && el("importText").value === importFilledText) fillFromFile();
  onImportInput();
}

// Liste du document Word pour la destination : liste LexRef du même côté (à défaut, l'autre),
// ou bordereau de fin de document.
function fileListFor(info, dest) {
  if (info.lexref) {
    const same = info[dest], other = info[dest === "adv" ? "own" : "adv"];
    return { text: same || other || "", kind: same ? dest : other ? (dest === "adv" ? "own" : "adv") : "" };
  }
  return { text: info.list || "", kind: info.list ? "list" : "" };
}
function fillFromFile() {
  const r = fileListFor(importFileInfo, importDest);
  importFilledText = r.text;
  el("importText").value = r.text;
  importFileInfo.kind = r.kind;
  renderImportSource();
}
function renderImportSource() {
  const box = el("importSource");
  if (!importFileInfo) { box.classList.add("hidden"); box.textContent = ""; return; }
  const esc = WP.escapeHtml, info = importFileInfo;
  const what = {
    own: "sa liste LexRef « " + (isLetterDoc() ? "Pièces jointes" : "Nos pièces") + " »",
    adv: "sa liste LexRef « Pièces adverses »",
    list: "la liste de pièces qui le termine",
  }[info.kind];
  box.innerHTML = what
    ? `Lu dans <b>${esc(info.name)}</b> : ${what}. Vérifiez puis importez.`
    : `<span class="warn">Aucune liste de pièces trouvée dans <b>${esc(info.name)}</b>.</span> Collez-la ci-dessous.`;
  box.classList.remove("hidden");
}
async function onImportFile(file) {
  if (!file) return;
  try {
    const info = await withBusy(async () => WP.readImportFile(await file.arrayBuffer()));
    if (!info) return; // erreur déjà signalée par withBusy
    importFileInfo = { ...info, name: file.name };
    fillFromFile();
    onImportInput();
  } catch (e) { reportError(e); }
}

function onImportInput() {
  importPlan = WP.planImport(WP.parseAdverseList(el("importText").value), importDest);
  renderImportPreview();
}
function renderImportPreview() {
  const box = el("importPreview");
  const btn = el("importBtn");
  if (!box) return;
  const text = el("importText").value.trim();
  if (!text) { box.innerHTML = ""; btn.disabled = true; btn.textContent = "Importer"; return; }
  if (!importPlan.length) {
    box.innerHTML = `<div class="imp-summary none">Aucune pièce détectée — chaque ligne doit commencer par un numéro (« 1. », « Pièce n°1 : », « 1 - »…).</div>`;
    btn.disabled = true; btn.textContent = "Importer";
    return;
  }
  const n = importPlan.length;
  const esc = WP.escapeHtml;
  const note = (it) => {
    if (it.status === "same") return `<span class="imp-note">déjà présente</span>`;
    if (it.status === "dup") return `<span class="imp-note warn">n°${esc(it.num)} en double dans la liste — ignorée</span>`;
    if (it.status === "rename") return it.oldName
      ? `<span class="imp-note warn">n°${esc(it.num)} existe déjà : « ${esc(it.oldName)} » — cocher pour la renommer</span>`
      : `<span class="imp-note">pièce existante sans nom — sera nommée</span>`;
    return "";
  };
  box.innerHTML =
    `<div class="imp-summary">${n} pièce${n > 1 ? "s" : ""} détectée${n > 1 ? "s" : ""}</div>` +
    `<div class="imp-list">` + importPlan.map((it, i) => {
      const fixed = it.status === "same" || it.status === "dup";
      return `<label class="imp-row ${it.checked ? "" : "off"}">` +
        `<input type="checkbox" data-imp="${i}" ${it.checked ? "checked" : ""} ${fixed ? "disabled" : ""} />` +
        `<span class="imp-num">${esc(it.num)}</span>` +
        `<span class="imp-body"><span class="imp-name">${esc(it.name)}</span>${note(it)}</span></label>`;
    }).join("") + `</div>`;
  const k = importPlan.filter((it) => it.checked).length;
  btn.disabled = k === 0;
  btn.textContent = k ? `Importer ${k} pièce${k > 1 ? "s" : ""}` : "Importer";
}

async function runImport() {
  if (!importPlan.some((it) => it.checked)) return;
  const dest = importDest;
  const communicated = dest === "own" && !isLetterDoc() && el("importCommunicated").checked;
  const snap = snapshotModel();
  let r = null;
  await withBusy(async () => {
    r = await WP.importPieces(importPlan, dest, { communicated });
    await syncDoc();
  });
  if (!r) return;
  const what = dest === "adv" ? "pièces adverses" : isLetterDoc() ? "pièces jointes" : "pièces";
  commitUndo(snap, "import des " + what);
  closeImport();
  setPaneSide(dest); // la liste qui vient d'être complétée
  flashIds = new Set(r.ids);
  render();
  const s = (k) => (k > 1 ? "s" : "");
  const parts = [];
  if (r.created) parts.push(`${r.created} pièce${s(r.created)} importée${s(r.created)}`);
  if (r.renamed) parts.push(`${r.renamed} renommée${s(r.renamed)}`);
  toast((parts.join(" · ") || "Rien à importer") + " (flèche ↶ du volet pour annuler)");
}

function bindImportUI() {
  el("advImportToggle").addEventListener("click", () => openImport("adv"));
  el("importClose").addEventListener("click", closeImport);
  el("importCancel").addEventListener("click", closeImport);
  el("importBtn").addEventListener("click", () => runImport().catch(reportError));
  el("importFileBtn").addEventListener("click", () => el("importFile").click());
  el("importFile").addEventListener("change", (e) => onImportFile(e.target.files && e.target.files[0]));
  el("importPanel").addEventListener("click", (e) => {
    const b = e.target.closest(".seg-btn[data-dest]");
    if (b) setImportDest(b.dataset.dest);
  });
  el("importText").addEventListener("input", onImportInput);
  el("importText").addEventListener("keydown", (e) => {
    if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); closeImport(); }
    else if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) { e.preventDefault(); runImport().catch(reportError); }
  });
  el("importPreview").addEventListener("change", (e) => {
    const cb = e.target.closest("input[data-imp]");
    if (!cb) return;
    const it = importPlan[+cb.dataset.imp];
    if (it) it.checked = cb.checked;
    renderImportPreview();
  });
}
