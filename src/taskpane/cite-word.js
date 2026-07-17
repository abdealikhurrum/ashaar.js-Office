(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) {
    module.exports = factory();
  } else {
    root.CiteWord = factory();
  }
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  var KEEP = { i: 1, b: 1, em: 1, strong: 1, span: 1, sup: 1, sub: 1, br: 1 };

  // Strip all attributes from kept tags; unwrap (keep inner content of) all other tags.
  function sanitize(html) {
    if (!html) { return ""; }
    return String(html).replace(/<(\/?)([a-zA-Z0-9]+)([^>]*)>/g, function (m, slash, name) {
      var tag = name.toLowerCase();
      if (KEEP[tag]) { return "<" + slash + tag + ">"; }
      return ""; // unwrap: remove the tag, keep its text content
    });
  }

  function direction(rtl) { return rtl ? "Rtl" : "Ltr"; }

  function buildNotePayload(o) {
    return { html: sanitize(o.html), direction: direction(o.rtl) };
  }

  function buildBibliographyPayload(o) {
    return { html: sanitize(o.html), direction: direction(o.rtl), tag: o.tag || "AshaarBibliography" };
  }

  function citationTag(itemKeys, style) {
    return "AshaarCite:" + style + ":" + (itemKeys || []).join(",");
  }

  function parseCitationTag(tag) {
    var m = /^AshaarCite:([^:]*):(.*)$/.exec(tag || "");
    if (!m) { return null; }
    return { style: m[1], itemKeys: m[2] ? m[2].split(",") : [] };
  }

  return {
    sanitize: sanitize,
    buildNotePayload: buildNotePayload,
    buildBibliographyPayload: buildBibliographyPayload,
    citationTag: citationTag,
    parseCitationTag: parseCitationTag
  };
}));
