(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) {
    module.exports = factory(require("../vendor/citeproc.js"));
  } else {
    root.CiteEngine = factory(root.CSL);
  }
}(typeof globalThis !== "undefined" ? globalThis : this, function (CSL) {
  "use strict";

  // Join citeproc's bibliography params + entry strings into one HTML string.
  function joinBibliography(result) {
    if (!result) { return ""; }
    var params = result[0] || {};
    var entries = result[1] || [];
    var open = (params.bibstart || "");
    var close = (params.bibend || "");
    return open + entries.join("") + close;
  }

  function build(opts) {
    var items = opts.items || {};
    var locales = opts.locales || {};
    var sys = {
      retrieveItem: function (id) { return items[id]; },
      retrieveLocale: function (lang) {
        return locales[lang] || locales["en-US"] || locales["us"];
      }
    };
    var engine = new CSL.Engine(sys, opts.styleXml, opts.lang || "en-US", !!opts.lang);
    var allIds = Object.keys(items);
    engine.updateItems(allIds);

    // CSL-M multilingual variant rendering (Task 4).
    // langPrefs selects which language *slot* to render per segment
    // (persons/titles/... -> "orig" | "translit" | "translat"), while
    // langPrefs.translit/translat register the concrete language tags
    // (e.g. "ar-Latn") that the engine treats as transliteration/translation
    // when matching an item's multi._keys[field][tag] / creator.multi._key[tag].
    if (opts.langPrefs) {
      var lp = opts.langPrefs;
      if (lp.translit && typeof engine.setLangTagsForCslTransliteration === "function") {
        engine.setLangTagsForCslTransliteration(lp.translit);
      }
      if (lp.translat && typeof engine.setLangTagsForCslTranslation === "function") {
        engine.setLangTagsForCslTranslation(lp.translat);
      }
      if (typeof engine.setLangPrefsForCites === "function") {
        engine.setLangPrefsForCites(lp);
      }
    }

    return {
      raw: engine,
      cite: function (citationItems) {
        var items = (citationItems || []).map(function (c) {
          if (typeof c === "string") { return { id: c }; }
          var out = { id: c.id };
          if (c.locator !== undefined && c.locator !== null && String(c.locator) !== "") {
            out.locator = String(c.locator);
          }
          if (c.label) { out.label = c.label; }
          return out;
        });
        return engine.makeCitationCluster(items);
      },
      bibliography: function () {
        return joinBibliography(engine.makeBibliography());
      },
      isRTL: function () {
        return /^ar\b/i.test(opts.lang || "");
      }
    };
  }

  return { build: build, joinBibliography: joinBibliography };
}));
