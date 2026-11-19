/* global Office */
// KIT JURITEL — apparence : thème clair / sombre et densité du volet.
// À charger dans CHAQUE page (volet et fenêtres), après office.js :
//     <script src="theme.js" data-key="monextension.theme"></script>
// data-key : clé de stockage propre à l'extension (facultatif). Deux extensions
// ouvertes en même temps ne se marchent pas dessus si leurs clés diffèrent.
//
// THÈME
// 1) Tout de suite : dernier thème connu (mémorisé) → pas de flash de couleur.
// 2) Dès qu'Office est prêt : lit le fond du thème Office (Office.context.officeTheme) et en
//    déduit clair ou sombre, puis le RE-LIT régulièrement : le volet vit aussi longtemps que Word
//    (runtime partagé) et Word ne signale pas toujours un changement de thème.
// 3) L'utilisateur peut IMPOSER son choix : JT.setTheme("light" | "dark"), ou "auto" pour
//    revenir au thème de Word. Un choix imposé l'emporte sur la lecture d'Office.
// 4) Les fenêtres (réglages, recherche…) suivent le volet via le stockage partagé (événement storage).
// Sans information (navigateur, ancienne version), le kit suit le thème système.
//
// DENSITÉ
//   JT.setDensity("compact" | "standard" | "comfortable") — pose data-density sur <html>,
//   que juritel-design.css traduit en facteur d'échelle. Mémorisée et propagée comme le thème.
//
// Ces deux préférences appartiennent au POSTE : elles vivent dans localStorage et ne
// doivent jamais être écrites dans le document.
(function () {
  var me = document.currentScript;
  var KEY = (me && me.dataset && me.dataset.key) || window.JT_THEME_KEY || "juritel.theme";
  var BASE = KEY.replace(/\.theme$/, "");
  var PREF = BASE + ".theme.pref"; // choix de l'utilisateur : auto | light | dark
  var DENS = BASE + ".density";    // compact | standard | comfortable

  var read = function (k) { try { return localStorage.getItem(k); } catch (e) { return null; } };
  var write = function (k, v) { try { localStorage.setItem(k, v); } catch (e) { /* stockage bloqué */ } };

  // ---------------- Thème ----------------
  function setAttr(t) {
    var d = document.documentElement;
    if (t === "dark" || t === "light") { if (d.dataset.theme !== t) d.dataset.theme = t; }
    else if (d.dataset.theme) { delete d.dataset.theme; } // ni choix ni thème connu → on suit le système
  }
  function themePref() {
    var p = read(PREF);
    return p === "light" || p === "dark" ? p : "auto";
  }
  // Le choix de l'utilisateur d'abord ; à défaut, le dernier thème de Word connu.
  function applyTheme() {
    var p = themePref();
    setAttr(p === "auto" ? read(KEY) : p);
  }

  function fromOffice(ev) {
    try {
      var th = (ev && ev.officeTheme) || (Office.context && Office.context.officeTheme);
      var bg = th && th.bodyBackgroundColor;
      if (!bg) return;
      var hex = String(bg).replace("#", "");
      if (hex.length === 3) hex = hex.replace(/./g, "$&$&");
      var r = parseInt(hex.substr(0, 2), 16), g = parseInt(hex.substr(2, 2), 16), b = parseInt(hex.substr(4, 2), 16);
      if (isNaN(r + g + b)) return;
      var t = 0.299 * r + 0.587 * g + 0.114 * b < 128 ? "dark" : "light";
      // On mémorise TOUJOURS le thème de Word (il redevient la référence si l'utilisateur
      // repasse en « auto »), mais on ne l'applique que si aucun choix n'est imposé.
      if (read(KEY) !== t) write(KEY, t);
      applyTheme();
    } catch (e) { /* ignoré */ }
  }

  // ---------------- Densité ----------------
  var DENSITIES = { compact: 1, standard: 1, comfortable: 1 };
  function densityPref() {
    var d = read(DENS);
    return DENSITIES[d] ? d : "standard";
  }
  function applyDensity() {
    document.documentElement.dataset.density = densityPref();
  }

  // ---------------- Mise en place ----------------
  applyTheme();
  applyDensity();

  // Une fenêtre suit le volet : même origine, donc même localStorage.
  window.addEventListener("storage", function (e) {
    if (e.key === KEY || e.key === PREF) applyTheme();
    else if (e.key === DENS) applyDensity();
  });

  var JT = window.JT || (window.JT = {});
  JT.theme = themePref;
  JT.setTheme = function (p) {
    write(PREF, p === "light" || p === "dark" ? p : "auto");
    applyTheme();
  };
  JT.density = densityPref;
  JT.setDensity = function (d) {
    write(DENS, DENSITIES[d] ? d : "standard");
    applyDensity();
  };

  if (window.Office && Office.onReady) {
    Office.onReady(function () {
      fromOffice();
      // Événement officiel (pris en charge selon l'application/la version : sinon, sans effet).
      try {
        Office.context.document.addHandlerAsync(Office.EventType.OfficeThemeChanged, fromOffice, function () {});
      } catch (e) { /* non pris en charge */ }
      setInterval(fromOffice, 3000);
      document.addEventListener("visibilitychange", function () { if (!document.hidden) fromOffice(); });
    });
  }
})();
