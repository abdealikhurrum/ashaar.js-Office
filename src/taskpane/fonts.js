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
        // Letters that take a clean trailing tatweel, by shaping form. Isolated
        // adds seen/sheen (س ش) which final does not; medial takes no tatweel.
        isolatedInto: ["ب","پ","ت","ٹ","ث","س","ش","ف","ک","گ"],
        finalInto:    ["ب","پ","ت","ٹ","ث","ف","ک","گ"]
      } },
    // Gate G2 (2026-07-10, manual Word test) confirmed Jameel Noori Kasheeda's
    // elongated forms are applied by FONT (the "Kasheeda" named style in the
    // family), not by an italic-run or whitespace mechanism. Jameel fills a
    // line by swapping whole fasls to the wider Kasheeda face; see
    // kashida-fontswap.js for the pure selection logic.
    jameel: { id: "jameel", label: "Jameel Noori Kasheeda",
      // MarkSafe build: vocalized Arabic keeps the wide kashida ligature forms
      // with harakat anchored (the stock Kasheeda drops vocalized words to
      // narrow forms). Family name MUST stay ≤31 chars: Word truncates longer
      // font names at 31, and "Jameel Noori Nastaleeq Kasheeda MarkSafe"
      // truncated to exactly the stock Kasheeda's name — swaps silently
      // rendered the old font (burned 2026-07-13). Old plain-Kasheeda runs
      // migrate via LEGACY_NAMES below.
      css: "'Jameel Kasheeda MarkSafe','Jameel Noori Nastaleeq',serif",
      wordName: "Jameel Noori Nastaleeq",                 // base face
      kasheedaName: "Jameel Kasheeda MarkSafe", // wider face (font-swap target)
      mechanism: "font-swap", bundled: true, private: true, readerNote: true,
      file: "JameelNooriNastaleeq-Regular.ttf", kasheedaFile: "JameelKasheedaMarkSafe.ttf" },
    gulzar: { id: "gulzar", label: "Gulzar",
      css: "'Gulzar',serif", wordName: "Gulzar",
      mechanism: "whitespace", bundled: true, file: "Gulzar-Regular.woff2" }
  };

  function get(id) { return LIST[id] || null; }
  function mechanismOf(id) { var d = get(id); return d ? d.mechanism : "whitespace"; }

  // Resolve the full registry descriptor from a run's ACTUAL Word font name
  // (used by the Justify path's per-cell dispatch to read wordName/kasheedaName/
  // tatweelRules/mechanism without knowing the dropdown id). Registry fonts
  // match by their wordName / kasheedaName; anything unrecognised — arbitrary
  // Arabic fonts (e.g. Fatemi Maqala), Latin defaults, empty/absent — returns a
  // synthetic "generic" descriptor so those runs run the tatweel engine
  // (AshaarJustify.justifyLine/justifyRuns) instead of being forced to spacing.
  // Retired face names → the base face they migrate to. The plain Kasheeda
  // face was replaced by the MarkSafe build (emit-only rename, 2026-07-13);
  // runs/tag-packs from older sessions still carry the plain name. Mapping it
  // to the BASE face lets the swap engine re-decide and re-emit as MarkSafe —
  // documents self-heal on their next justify instead of falling to generic.
  var LEGACY_NAMES = { "Jameel Noori Nastaleeq Kasheeda": "Jameel Noori Nastaleeq" };
  function normalizeLegacyFontName(name) {
    var n = String(name == null ? "" : name).trim();
    return LEGACY_NAMES[n] || n;
  }

  function descriptorForFontName(name) {
    var n = normalizeLegacyFontName(name);
    if (n) {
      for (var id in LIST) {
        if (!LIST.hasOwnProperty(id)) continue;
        var d = LIST[id];
        if ((d.wordName && d.wordName === n) || (d.kasheedaName && d.kasheedaName === n)) {
          return d;
        }
      }
    }
    return { id: "generic", mechanism: "generic", wordName: null, kasheedaName: null, tatweelRules: null };
  }

  // Thin wrapper: just the mechanism for a run's real font.
  function mechanismForFontName(name) { return descriptorForFontName(name).mechanism; }
  function wordNameOf(id) { var d = get(id); return d && d.wordName ? d.wordName : null; }
  function kasheedaNameOf(id) { var d = get(id); return d && d.kasheedaName ? d.kasheedaName : null; }
  function cssFamilyOf(id) { var d = get(id); return d && d.css ? d.css : null; }
  function tatweelRulesOf(id) { var d = get(id); return d && d.tatweelRules ? d.tatweelRules : null; }

  return { LIST: LIST, get: get, mechanismOf: mechanismOf,
    descriptorForFontName: descriptorForFontName,
    normalizeLegacyFontName: normalizeLegacyFontName,
    mechanismForFontName: mechanismForFontName, wordNameOf: wordNameOf,
    kasheedaNameOf: kasheedaNameOf, cssFamilyOf: cssFamilyOf, tatweelRulesOf: tatweelRulesOf };
}));
