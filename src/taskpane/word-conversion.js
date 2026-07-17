/**
 * AshaarConversion — pure data + engine for the "Convert" tab. Converts
 * Arabic-script text between the legacy LD double-press/AL-KANZ encoding and a
 * modern encoding (Unicode where it exists, Fatemi !keyword! otherwise).
 * No Office.js/DOM; the Word.run() orchestration lives in conversion-pane.js.
 * See docs/superpowers/specs/2026-07-16-text-conversion-design.md.
 */
(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) module.exports = factory();
  else root.AshaarConversion = factory();
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  var TATWEEL = "ـ";
  var DIRECTIONS = { TO_MODERN: "toModern", TO_LEGACY: "toLegacy" };

  // legacy = old double-press/placeholder form; modern = Unicode-or-!keyword!.
  // wholeWord: only collapse when standalone (Modern→Legacy). lossy: not safely
  // round-trippable (the source character has a legitimate independent meaning).
  var MAPPINGS = [
    { id: "seen-baariye",   category: "letter", legacy: "سس", modern: "ے",    label: "ے baari ye",            wholeWord: false, lossy: false },
    { id: "zah-heh",        category: "letter", legacy: "ظظ", modern: "ہ",    label: "ہ gol he",              wholeWord: false, lossy: false },
    { id: "tah-noonghunna", category: "letter", legacy: "طط", modern: "ں",    label: "ں noon ghunna",         wholeWord: false, lossy: false },
    { id: "kaf-gaf",        category: "letter", legacy: "كك", modern: "گ",    label: "گ gaf",                 wholeWord: false, lossy: false },
    { id: "cheh-hah",       category: "letter", legacy: "حح", modern: "چ",    label: "چ cheh",                wholeWord: false, lossy: false },
    { id: "tteh-dad",       category: "letter", legacy: "ضض", modern: "ٹ",    label: "ٹ tteh",                wholeWord: false, lossy: false },
    { id: "rreh-re",        category: "letter", legacy: "رٌ", modern: "ڑ",    label: "ڑ rreh (rā+dammatan)",  wholeWord: false, lossy: false },
    { id: "ddal-dal",       category: "letter", legacy: "دٌ", modern: "ڈ",    label: "ڈ ddal (dāl+dammatan)", wholeWord: false, lossy: false },
    { id: "peh-theh",       category: "letter", legacy: "ثث", modern: "پ",    label: "پ peh",                 wholeWord: false, lossy: false },
    { id: "chhay-semicolon",category: "letter", legacy: "؛",  modern: "چھے", label: "چھے ⇄ ؛ (semicolon)",   wholeWord: true,  lossy: false },

    // ── mark tier: legacy Arabic-101 keyboard repurposings (confirmed rows) ──
    // Shift+X emitted sukun (U+0652) but the old font drew it as khari zabar /
    // dagger alef (U+0670). LOSSY: a genuine sukun would be reinterpreted.
    { id: "sukun-kharizabar", category: "mark", legacy: "ْ", modern: "ٰ", label: "khari zabar / dagger alef (was sukun)", wholeWord: false, lossy: true }
    // >>> GENERATED symbol rows (scripts/generate-conversion-table.mjs) — do not edit by hand.
    // Also pending: high jeem (modern U+06DA) and high noon (modern U+06E8) mark
    // rows — modern targets confirmed from features.fea, legacy Shift+C/Shift+V
    // source characters await confirmation.
    // <<< GENERATED
  ];

  function isDoubledConsonant(m) {
    return m.category === "letter" && m.legacy.length === 2 && m.legacy[0] === m.legacy[1];
  }

  function enabledSet(ids) {
    if (!ids) return null;              // null/undefined = all rows enabled
    var s = {};
    ids.forEach(function (i) { s[i] = true; });
    return s;
  }

  function byFindLenDesc(key) {
    return function (a, b) { return b[key].length - a[key].length; };
  }

  // Ordered literal ops. Ordering (not runtime context) enforces the escape rule.
  function buildOperations(direction, enabledIds) {
    var on = enabledSet(enabledIds);
    var rows = MAPPINGS.filter(function (m) { return !on || on[m.id]; });
    var ops = [];

    if (direction === DIRECTIONS.TO_LEGACY) {
      // 1) protect genuine doubles: سس → سـس (before ے→سس creates new doubles)
      rows.filter(isDoubledConsonant).forEach(function (m) {
        var b = m.legacy[0];
        ops.push({ find: b + b, replaceWith: b + TATWEEL + b, wholeWord: false, category: m.category });
      });
      // 2) modern→legacy (longest find first; whole-word where flagged)
      rows.slice().sort(byFindLenDesc("modern")).forEach(function (m) {
        ops.push({ find: m.modern, replaceWith: m.legacy, wholeWord: !!m.wholeWord, category: m.category });
      });
    } else {
      // 1) contiguous / direct legacy→modern (longest find first)
      rows.slice().sort(byFindLenDesc("legacy")).forEach(function (m) {
        ops.push({ find: m.legacy, replaceWith: m.modern, wholeWord: false, category: m.category });
      });
      // 2) escape-drop: سـس → سس (after step 1, so the separated form survived it)
      rows.filter(isDoubledConsonant).forEach(function (m) {
        var b = m.legacy[0];
        ops.push({ find: b + TATWEEL + b, replaceWith: b + b, wholeWord: false, category: m.category });
      });
    }
    return ops;
  }

  // Arabic-script "letter or mark" ranges, for whole-word boundary detection.
  var WORDCHAR = /[؀-ۿݐ-ݿࢠ-ࣿﭐ-﷿ﹰ-﻿]/;
  function isWordChar(ch) { return !!ch && WORDCHAR.test(ch); }

  function replaceAll(text, find, replaceWith, wholeWord) {
    if (!find) return text;
    var out = "", i = 0;
    while (i < text.length) {
      if (text.substr(i, find.length) === find) {
        if (wholeWord) {
          var before = i > 0 ? text.charAt(i - 1) : "";
          var after = text.charAt(i + find.length);
          if (isWordChar(before) || isWordChar(after)) { out += text.charAt(i); i += 1; continue; }
        }
        out += replaceWith; i += find.length;
      } else {
        out += text.charAt(i); i += 1;
      }
    }
    return out;
  }

  function convert(text, direction, enabledIds) {
    var ops = buildOperations(direction, enabledIds);
    var s = String(text == null ? "" : text);
    ops.forEach(function (op) { s = replaceAll(s, op.find, op.replaceWith, op.wholeWord); });
    return s;
  }

  function groupsForUi() {
    var order = ["letter", "mark", "symbol"], byCat = {};
    MAPPINGS.forEach(function (m) { (byCat[m.category] = byCat[m.category] || []).push(m); });
    return order.filter(function (c) { return byCat[c]; })
      .map(function (c) { return { category: c, rows: byCat[c] }; });
  }

  return {
    TATWEEL: TATWEEL,
    DIRECTIONS: DIRECTIONS,
    MAPPINGS: MAPPINGS,
    buildOperations: buildOperations,
    convert: convert,
    groupsForUi: groupsForUi
  };
}));
