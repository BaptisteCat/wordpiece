/* global Office, WPMsg */
// Fenêtre CENTRALE du mode d'emploi : contenu statique ; seule la fermeture passe par le volet (parent).

WPMsg.ready(() => {
  const close = () => WPMsg.send(JSON.stringify({ close: true }));
  document.getElementById("close").addEventListener("click", close);
  document.addEventListener("keydown", (e) => { if (e.key === "Escape") close(); });
});
