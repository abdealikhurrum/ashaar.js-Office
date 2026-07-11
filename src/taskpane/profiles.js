/**
 * profiles.js — Qaseeda profile data model and shared-width math (pure, UMD).
 *
 * A "qaseeda" is a named link across Ashaar Poem blocks that share one profile
 * (width, justification, gap, symbol, colours). This module owns the profile
 * shape, defaults/merge, the per-font fill correction, and the derive-shared-
 * widths math. It has NO Office.js dependency — orchestration (the document
 * store, tagging, in-place resize) lives in taskpane.js.
 *
 * See docs/superpowers/specs/2026-05-31-qaseeda-profiles-design.md.
 */
(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) module.exports = factory();
  else root.AshaarProfiles = factory();
}(this, function () {
  "use strict";

  // The authoritative default profile. `derived` is computed from the qaseeda's
  // content (see deriveSharedWidths) and refreshed on the hybrid trigger.
  function defaultProfile(name) {
    return {
      name: typeof name === "string" ? name : "",
      width: { mode: "auto-fit", pct: 50 },        // "auto-fit" | "fixed"
      justify: { mode: "kashida", strength: 6, fillMode: "natural-fit" },   // "kashida" | "spacing" | "css" | "none"
      gap: 4,
      misraSymbol: "",
      symbolColor: "",
      debugColors: { tatweel: "", space: "" },
      font: "",
      fontCorrections: {},                         // { <fontName>: factor }
      derived: { colWidthVector: null, calibrationRecipe: null },
    };
  }

  function isObj(v) { return v && typeof v === "object" && !Array.isArray(v); }

  // Shallow-merge `partial` onto `base` one level deep for the known nested
  // buckets (width, justify, debugColors, fontCorrections, derived). Returns a
  // new object; never mutates `base` or `partial`.
  function mergeProfile(base, partial) {
    var out = {};
    var b = base || {};
    var p = partial || {};
    var keys = ["name", "gap", "misraSymbol", "symbolColor", "font"];
    keys.forEach(function (k) { out[k] = (k in p) ? p[k] : b[k]; });
    var nested = ["width", "justify", "debugColors", "fontCorrections", "derived"];
    nested.forEach(function (k) {
      var bo = isObj(b[k]) ? b[k] : {};
      var po = isObj(p[k]) ? p[k] : {};
      var merged = {};
      Object.keys(bo).forEach(function (kk) { merged[kk] = bo[kk]; });
      Object.keys(po).forEach(function (kk) { merged[kk] = po[kk]; });
      out[k] = merged;
    });
    return out;
  }

  // Fill any missing fields of `p` from the defaults (deep, via mergeProfile).
  function normalizeProfile(p) {
    return mergeProfile(defaultProfile((p && p.name) || ""), p || {});
  }

  // px × the per-font correction factor (1 when the font has no entry).
  function applyFontCorrection(px, font, corrections) {
    var f = (corrections && font && corrections[font]) || 1;
    return px * f;
  }

  // Given per-column measurement lists, return the target px each column needs:
  // the widest corrected measurement in that column, scaled by kashida headroom.
  //   columns: [ [ {px, font}, ... ] (col 0), [ ... ] (col 1), ... ]
  //   opts:    { headroom = 1, corrections = {} }
  function deriveSharedWidths(columns, opts) {
    opts = opts || {};
    var headroom = typeof opts.headroom === "number" ? opts.headroom : 1;
    var corrections = opts.corrections || {};
    return (columns || []).map(function (col) {
      var max = 0;
      (col || []).forEach(function (m) {
        var w = applyFontCorrection(m.px || 0, m.font, corrections);
        if (w > max) max = w;
      });
      return max * headroom;
    });
  }

  // Content-width px (at 96 dpi) → a Word column width in points, adding back
  // both reserved cell side margins. Mirror of taskpane's contentPx() inverse.
  function columnPointsFromContentPx(px, marginPt) {
    return (px || 0) * 72 / 96 + 2 * (marginPt || 0);
  }

  // Kashida-strength slider (1..10) → justify targetFill (0.90..1.0), matching
  // word-html's sliderToFill. Values outside [1,10] clamp.
  function strengthToTargetFill(strength) {
    var s = Math.max(1, Math.min(10, Number(strength) || 1));
    return 0.90 + ((s - 1) / 9) * 0.10;
  }

  // Sanitise a profile's stored strength to the 1–10 domain. Profiles are not yet
  // in real use, so this is a plain clamp — no proportional legacy 0–24 remap.
  function normalizeStrength(strength) {
    return Math.max(1, Math.min(10, Number(strength) || 1));
  }

  // Sanitise a profile's stored fill mode. Cell-fit = fill to the true cell edge
  // (Word distribute residual); Natural-fit = fill to the per-position matrix
  // width (capped micro-space residual). Anything else defaults to natural-fit.
  function normalizeFillMode(mode) {
    return mode === "cell-fit" ? "cell-fit" : "natural-fit";
  }

  return {
    defaultProfile: defaultProfile,
    normalizeProfile: normalizeProfile,
    mergeProfile: mergeProfile,
    applyFontCorrection: applyFontCorrection,
    deriveSharedWidths: deriveSharedWidths,
    columnPointsFromContentPx: columnPointsFromContentPx,
    strengthToTargetFill: strengthToTargetFill,
    normalizeStrength: normalizeStrength,
    normalizeFillMode: normalizeFillMode,
  };
}));
