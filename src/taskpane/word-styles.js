/**
 * AshaarStyles — style-group data model for the prose Styles tab (headings,
 * emphasis, block quotes, Quran quotes) and RTL document setup. Pure (no
 * Office.js/DOM); the Word.run() orchestration lives in styles-pane.js.
 *
 * See docs/superpowers/specs/2026-07-16-ashaar-styles-design.md.
 */
(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) module.exports = factory();
  else root.AshaarStyles = factory();
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  // Order matters: "quote" must be created/configured before "quranQuote",
  // since Ashaar Quran Quote is basedOn Ashaar Quote (not a Word built-in).
  var ROLES = ["heading1", "heading2", "heading3", "emphasis", "quote", "quranQuote"];

  var STYLE_NAME = {
    heading1: "Ashaar Heading 1",
    heading2: "Ashaar Heading 2",
    heading3: "Ashaar Heading 3",
    emphasis: "Ashaar Emphasis",
    quote: "Ashaar Quote",
    quranQuote: "Ashaar Quran Quote"
  };

  var BASE_STYLE = {
    heading1: "Heading 1",
    heading2: "Heading 2",
    heading3: "Heading 3",
    emphasis: "Emphasis",
    quote: "Quote",
    quranQuote: "Ashaar Quote"
  };

  var STYLE_TYPE = {
    heading1: "Paragraph",
    heading2: "Paragraph",
    heading3: "Paragraph",
    emphasis: "Character",
    quote: "Paragraph",
    quranQuote: "Paragraph"
  };

  function isObj(v) { return v && typeof v === "object" && !Array.isArray(v); }

  // The authoritative default group recipe. Headings are bold+centered
  // (fixed, not user-exposed); font/size are the adjustable style-level
  // fields. indentPt/lineHeightPt double as both the style-level default AND
  // the seed for a per-instance override (applied as direct formatting on
  // top of the named style).
  function defaultGroup(name) {
    return {
      name: typeof name === "string" ? name : "",
      heading1: { font: "Marjaan", sizePt: 18 },
      heading2: { font: "Marjaan", sizePt: 16 },
      heading3: { font: "Marjaan", sizePt: 14 },
      emphasis: { color: "#FF0000", bumpPt: 3 },
      quote: { borderColor: "#000000", borderWidth: "Pt050", indentPt: 0 },
      quranQuote: { font: "Amiri Quran", lineHeightPt: null } // null = Word auto
    };
  }

  // Shallow-merge `partial` onto `base` one level deep, per role. Returns a
  // new object; never mutates `base` or `partial`.
  function mergeGroup(base, partial) {
    var out = {};
    var b = base || {};
    var p = partial || {};
    out.name = ("name" in p) ? p.name : b.name;
    ROLES.forEach(function (role) {
      var br = isObj(b[role]) ? b[role] : {};
      var pr = isObj(p[role]) ? p[role] : {};
      var merged = {};
      Object.keys(br).forEach(function (k) { merged[k] = br[k]; });
      Object.keys(pr).forEach(function (k) { merged[k] = pr[k]; });
      out[role] = merged;
    });
    return out;
  }

  // Fill any missing roles/fields of `g` from the defaults (deep, via mergeGroup).
  function normalizeGroup(g) {
    return mergeGroup(defaultGroup((g && g.name) || ""), g || {});
  }

  var BUILTIN_GROUPS = {
    General: defaultGroup("General"),
    Petition: mergeGroup(defaultGroup("Petition"), {
      heading1: { sizePt: 16 }, heading2: { sizePt: 14 }, heading3: { sizePt: 12 },
      quote: { indentPt: 18 }
    }),
    Maqala: mergeGroup(defaultGroup("Maqala"), {
      heading1: { sizePt: 16 }, heading2: { sizePt: 14 }, heading3: { sizePt: 12 }
    }),
    Waaz: mergeGroup(defaultGroup("Waaz"), {
      heading1: { font: "Fatemi", sizePt: 20 }, heading2: { font: "Fatemi", sizePt: 17 }
    })
  };

  return {
    ROLES: ROLES,
    STYLE_NAME: STYLE_NAME,
    BASE_STYLE: BASE_STYLE,
    STYLE_TYPE: STYLE_TYPE,
    defaultGroup: defaultGroup,
    mergeGroup: mergeGroup,
    normalizeGroup: normalizeGroup,
    BUILTIN_GROUPS: BUILTIN_GROUPS
  };
}));
