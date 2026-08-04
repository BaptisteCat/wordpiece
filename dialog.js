/* global Office */
// Petite fenêtre de saisie. Renvoie le résultat au runtime des commandes via messageParent.

Office.onReady(() => {
  const params = new URLSearchParams(location.search);
  const mode = params.get("mode") || "rename";
  const label = params.get("label") || "";
  const title = params.get("title") || "WordPiece";

  const $ = (id) => document.getElementById(id);
  $("title").textContent = title;
  const input = $("input");
  const okBtn = $("ok");

  function send(obj) { Office.context.ui.messageParent(JSON.stringify(obj)); }

  if (mode === "confirm") {
    input.style.display = "none";
    const customMsg = params.get("msg");
    $("msg").textContent = customMsg || `Supprimer « ${label} » ? Elle disparaîtra du corps du texte, du bordereau et de la liste.`;
    okBtn.textContent = params.get("ok") || "Supprimer";
    if (params.get("danger") !== "0") okBtn.classList.add("danger"); // rouge par défaut (suppression)
    okBtn.onclick = () => send({ confirmed: true });
  } else if (mode === "renumber") {
    $("msg").textContent = "Numéro de la pièce (libre : 2, 19, 2 bis, 1.2.4…) :";
    input.type = "text";
    input.value = label;
    okBtn.onclick = () => send({ value: input.value });
  } else {
    $("msg").textContent = "Nouveau nom de la pièce :";
    input.type = "text"; input.value = label;
    okBtn.onclick = () => send({ value: input.value });
  }

  const cancelLabel = params.get("cancel");
  if (cancelLabel) $("cancel").textContent = cancelLabel;
  $("cancel").onclick = () => send({ cancelled: true });
  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter") okBtn.click();
    else if (e.key === "Escape") $("cancel").click();
  });
  input.focus();
  if (input.select) input.select();
});
