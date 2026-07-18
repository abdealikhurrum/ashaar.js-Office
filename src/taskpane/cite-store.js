(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) { module.exports = factory(); }
  else { root.CiteStore = factory(); }
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  var REFS_KEY = "AshaarCiteRefs";

  function serializeRefs(items) { return JSON.stringify({ v: 1, items: items || {} }); }

  function parseRefs(str) {
    if (!str || typeof str !== "string") { return {}; }
    try {
      var obj = JSON.parse(str);
      if (!obj || obj.v !== 1 || !obj.items || typeof obj.items !== "object") { return {}; }
      return obj.items;
    } catch (e) { return {}; }
  }

  // The document settings bag, or null when Office isn't present (bare browser).
  function resolveSettings() {
    try {
      if (typeof Office !== "undefined" && Office.context && Office.context.document &&
          Office.context.document.settings) { return Office.context.document.settings; }
    } catch (e) { /* fall through */ }
    return null;
  }

  function saveRefs(items, settingsImpl) {
    var s = settingsImpl || resolveSettings();
    if (!s) { return Promise.resolve(); } // browser preview: no-op
    return new Promise(function (resolve, reject) {
      s.set(REFS_KEY, serializeRefs(items));
      s.saveAsync(function (res) {
        var status = res && res.status;
        // Office.AsyncResultStatus.Succeeded === "succeeded" (compare by value).
        if (status === "succeeded" || status === 0) { resolve(); }
        else { reject(new Error((res && res.error && res.error.message) || "settings saveAsync failed")); }
      });
    });
  }

  function loadRefs(settingsImpl) {
    var s = settingsImpl || resolveSettings();
    if (!s) { return Promise.resolve({}); }
    return Promise.resolve().then(function () { return parseRefs(s.get(REFS_KEY)); })
      .catch(function () { return {}; });
  }

  return {
    REFS_KEY: REFS_KEY,
    serializeRefs: serializeRefs,
    parseRefs: parseRefs,
    resolveSettings: resolveSettings,
    saveRefs: saveRefs,
    loadRefs: loadRefs
  };
}));
