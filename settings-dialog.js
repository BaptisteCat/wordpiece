/* global Office */
// Fenêtre CENTRALE des réglages. Reçoit du volet (parent) l'objet settings + le schéma de liaison
// (bindings), affiche le formulaire, et renvoie CHAQUE changement au parent sous forme { path, value }.
// Le parent applique et enregistre (source de vérité unique : WP.model.settings dans le volet).

Office.onReady(() => {
  let bindings = [];
  const byId = {};
  const el = (id) => document.getElementById(id);
  const getPath = (obj, path) => path.reduce((o, k) => (o ? o[k] : undefined), obj);

  function applySettings(settings) {
    for (const b of bindings) {
      const elm = el(b.id);
      if (!elm) continue;
      const v = getPath(settings, b.path);
      if (b.type === "checked") elm.checked = !!v;
      else elm.value = v == null ? "" : v;
    }
    updateLayoutVisibility(settings);
  }

  function updateLayoutVisibility(settings) {
    const isList = !settings || !settings.bordereau || settings.bordereau.layout !== "table";
    document.querySelectorAll('[data-layout="list"]').forEach((n) => n.classList.toggle("hidden", !isList));
  }

  function readValue(b, elm) {
    if (b.type === "checked") return elm.checked;
    if (b.type === "number") return parseInt(elm.value, 10) || 14;
    return elm.value; // "value" et "raw"
  }

  function bindChangeHandlers() {
    for (const b of bindings) {
      const elm = el(b.id);
      if (!elm || byId[b.id]) continue;
      byId[b.id] = b;
      elm.addEventListener("change", () => {
        const value = readValue(b, elm);
        // Si la présentation du bordereau change, on ajuste tout de suite la visibilité locale.
        if (b.id === "bordLayout") {
          document.querySelectorAll('[data-layout="list"]').forEach((n) => n.classList.toggle("hidden", value === "table"));
        }
        Office.context.ui.messageParent(JSON.stringify({ path: b.path, value }));
      });
    }
  }

  el("close").addEventListener("click", () => Office.context.ui.messageParent(JSON.stringify({ close: true })));
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") Office.context.ui.messageParent(JSON.stringify({ close: true }));
  });

  // Handshake : on enregistre le récepteur AVANT de signaler « ready » (sinon on raterait le message).
  Office.context.ui.addHandlerAsync(
    Office.EventType.DialogParentMessageReceived,
    (arg) => {
      let data = null;
      try { data = JSON.parse(arg.message); } catch (e) { data = null; }
      if (!data) return;
      bindings = Array.isArray(data.bindings) ? data.bindings : [];
      applySettings(data.settings || {});
      bindChangeHandlers();
    },
    () => { Office.context.ui.messageParent(JSON.stringify({ ready: true })); }
  );
});
