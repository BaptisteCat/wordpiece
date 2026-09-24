/* global Office */
// Transport commun aux pages « fenêtre » (dialog, search-piece, settings-dialog, help-dialog).
//
// Deux contextes possibles :
//   - VRAIE FENÊTRE Office (Windows, Mac) : messagerie Office (messageParent / DialogParentMessageReceived).
//   - INTÉGRÉE DANS LE VOLET (Word pour le web) : simple <iframe> du volet → postMessage entre
//     même origine. Word pour le web refuse d'ouvrir les fenêtres Office quand le navigateur bloque
//     le stockage tiers ; le volet ouvre alors la page ici plutôt que dans une fenêtre.
//
// Le volet ajoute « ?embed=1 » à l'URL pour demander ce mode.
(function () {
  const embed = new URLSearchParams(location.search).get("embed") === "1";

  // office.js n'est chargé QUE pour une vraie fenêtre Office : intégré dans le volet, il constate
  // qu'il n'est pas dans un hôte Office (« not hosted in plain browser top window ») et VIDE la page.
  // document.write : chargement synchrone pendant l'analyse du document, comme la balise <script>
  // d'origine, pour que theme.js et le script de la page trouvent Office déjà défini.
  if (!embed) {
    document.write('<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"><\/script>');
  }

  function send(obj) {
    const s = typeof obj === "string" ? obj : JSON.stringify(obj);
    if (embed) {
      try { parent.postMessage(s, location.origin); } catch (e) { /* volet fermé */ }
      return;
    }
    try { Office.context.ui.messageParent(s); } catch (e) { /* fenêtre déjà fermée */ }
  }

  // cb reçoit { message } comme le gestionnaire Office ; `done` est appelé une fois l'écoute en place.
  function onParent(cb, done) {
    if (embed) {
      window.addEventListener("message", (e) => {
        if (e.source !== parent || e.origin !== location.origin) return;
        cb({ message: typeof e.data === "string" ? e.data : JSON.stringify(e.data) });
      });
      if (done) done();
      return;
    }
    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, cb, () => { if (done) done(); });
  }

  // Intégrée, la page n'est PAS une fenêtre Office : Office.onReady ne se résout pas → on démarre seul.
  function ready(cb) {
    if (embed) {
      if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", cb);
      else cb();
      return;
    }
    Office.onReady().then(cb);
  }

  window.WPMsg = { embed, send, onParent, ready };
})();
