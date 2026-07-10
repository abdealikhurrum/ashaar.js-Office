/**
 * AshaarFonts — single source of truth for the fonts the add-in offers.
 * Each descriptor carries the CSS family (preview), the Word cs font name
 * (OOXML <w:rFonts w:cs>), and the kashida `mechanism` that selects the
 * justify strategy. Pure (no DOM); safe to require in Node tests.
 */
(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) module.exports = factory();
  else root.AshaarFonts = factory();
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  var LIST = {
    document: { id: "document", label: "Document default",
      css: null, wordName: null, mechanism: "whitespace", bundled: false },
    "arabic-serif": { id: "arabic-serif", label: "Arabic serif",
      css: "'Scheherazade New','Amiri','Times New Roman',serif",
      wordName: "Scheherazade New", mechanism: "whitespace", bundled: false },
    noto: { id: "noto", label: "Noto Nastaliq Urdu",
      css: "'Noto Nastaliq Urdu',serif", wordName: "Noto Nastaliq Urdu",
      mechanism: "whitespace", bundled: false },
    mehr: { id: "mehr", label: "Mehr Nastaliq",
      css: "'Mehr Nastaliq Web','Noto Nastaliq Urdu',serif", wordName: "Mehr Nastaliq Web",
      mechanism: "tatweel", bundled: true, file: "MehrNastaliqWeb.woff2",
      readerNote: true,
      tatweelRules: {
        version: "beta-2.0",
        medialInto: ["ب","پ","ت","ٹ","ث","س","ش","ف","ک","گ"],
        finalInto:  ["ب","پ","ت","ٹ","ث","ف","ک","گ"]
      } },
    // Gate G2 (2026-07-10, manual Word test) confirmed Jameel Noori Kasheeda's
    // elongated forms are applied by FONT (the "Kasheeda" named style in the
    // family), not by an italic-run or whitespace mechanism. Jameel fills a
    // line by swapping whole fasls to the wider Kasheeda face; see
    // kashida-fontswap.js for the pure selection logic.
    jameel: { id: "jameel", label: "Jameel Noori Kasheeda",
      css: "'Jameel Noori Nastaleeq Kasheeda','Jameel Noori Nastaleeq',serif",
      wordName: "Jameel Noori Nastaleeq",                 // base face
      kasheedaName: "Jameel Noori Nastaleeq Kasheeda",    // wider face (font-swap target)
      mechanism: "font-swap", bundled: true, private: true, readerNote: true,
      file: "JameelNooriNastaleeq-Regular.ttf", kasheedaFile: "JameelNooriNastaleeqKasheeda.ttf" },
    gulzar: { id: "gulzar", label: "Gulzar",
      css: "'Gulzar',serif", wordName: "Gulzar",
      mechanism: "whitespace", bundled: true, file: "Gulzar-Regular.woff2" }
  };

  function get(id) { return LIST[id] || null; }
  function mechanismOf(id) { var d = get(id); return d ? d.mechanism : "whitespace"; }
  function wordNameOf(id) { var d = get(id); return d && d.wordName ? d.wordName : null; }
  function kasheedaNameOf(id) { var d = get(id); return d && d.kasheedaName ? d.kasheedaName : null; }
  function cssFamilyOf(id) { var d = get(id); return d && d.css ? d.css : null; }
  function tatweelRulesOf(id) { var d = get(id); return d && d.tatweelRules ? d.tatweelRules : null; }

  return { LIST: LIST, get: get, mechanismOf: mechanismOf, wordNameOf: wordNameOf,
    kasheedaNameOf: kasheedaNameOf, cssFamilyOf: cssFamilyOf, tatweelRulesOf: tatweelRulesOf };
}));
