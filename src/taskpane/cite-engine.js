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

    return {
      raw: engine,
      cite: function (itemKeys) {
        var citationItems = (itemKeys || []).map(function (id) { return { id: id }; });
        return engine.makeCitationCluster(citationItems);
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
