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
      // mode: "kashida" | "spacing" | "css" | "none". widthPt: explicit misra
      // fill-target width for every content cell (null = computed harmony /
      // cell-fit target); overridable per bandh (tag payload.widthPt) and per
      // cell — precedence cell > bandh > qaseeda > computed.
      justify: { mode: "kashida", strength: 6, fillMode: "natural-fit", widthPt: null },
      gap: 4,
      misraSymbol: "",
      symbolColor: "",
      debugColors: { tatweel: "", space: "" },
      spacingDecor: {},                            // { "<slot>": { symbol, fill, color } }
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
    var keys = ["name", "gap", "misraSymbol", "symbolColor"];
    keys.forEach(function (k) { out[k] = (k in p) ? p[k] : b[k]; });
    var nested = ["width", "justify", "debugColors", "fontCorrections", "derived", "spacingDecor"];
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
  // width (capped micro-space residual); Adaptive = one shared target every
  // misra can REACH with its own kashida plus at most ~0.25em/gap of spacing —
  // equal line widths and small gaps, at the cost of stopping short of the cell
  // edge when the font can't stretch (see the Jameel-harakat ceiling).
  // Anything else defaults to natural-fit.
  function normalizeFillMode(mode) {
    return (mode === "cell-fit" || mode === "adaptive") ? mode : "natural-fit";
  }

  // ── Canonical settings (unified panel) ────────────────────────────────────
  // One flat shape shared by the resolver, the panel, and tag `local` maps.
  // Layering: defaults → profile → block local → bandh widthPt → cell override.

  function defaultSettings() {
    return {
      justifyMode: "kashida",     // "kashida" | "css" | "spacing" | "none"
      fillMode: "natural-fit",    // "natural-fit" | "cell-fit" | "adaptive"
      strength: 6,                // 1..10
      gap: 4,                     // middle-gap grid columns, 0..20
      widthMode: "auto-fit",      // "auto-fit" | "fixed"
      widthPct: 50,               // 25..100 (only meaningful when fixed)
      misraWidthPt: null,         // explicit fill target; null = computed
      layoutMode: "balanced",     // "balanced"|"equal"|"compact"|"stacked"|"auto"
      colWidthMode: "optimized",  // "optimized" | "fixed" (column-width strategy)
      capEm: 0.28,                // residual spacing cap (cell scope)
      fontCorrections: {},
      debugColors: { tatweel: "", space: "" },
    };
  }

  // Keys a profile owns. layoutMode/colWidthMode/capEm are block/cell-level
  // preferences with no profile layer.
  function settingsFromProfile(profile) {
    var p = normalizeProfile(profile || {});
    var out = {
      justifyMode: p.justify.mode,
      fillMode: normalizeFillMode(p.justify.fillMode),
      strength: normalizeStrength(p.justify.strength),
      gap: Number(p.gap),
      widthMode: p.width.mode === "fixed" ? "fixed" : "auto-fit",
      widthPct: Number(p.width.pct),
    };
    if (p.justify.widthPt != null) out.misraWidthPt = p.justify.widthPt;
    if (p.fontCorrections && Object.keys(p.fontCorrections).length) out.fontCorrections = p.fontCorrections;
    if (p.debugColors && (p.debugColors.tatweel || p.debugColors.space)) out.debugColors = p.debugColors;
    return out;
  }

  // Inverse of settingsFromProfile: canonical values → profile-schema object.
  function profileFromSettings(name, values) {
    var v = values || {};
    var p = defaultProfile(name);
    if (v.justifyMode != null) p.justify.mode = v.justifyMode;
    if (v.fillMode != null) p.justify.fillMode = normalizeFillMode(v.fillMode);
    if (v.strength != null) p.justify.strength = normalizeStrength(v.strength);
    if (v.misraWidthPt !== undefined) p.justify.widthPt = v.misraWidthPt;
    if (v.gap != null) p.gap = Number(v.gap);
    if (v.widthMode != null) p.width.mode = v.widthMode === "fixed" ? "fixed" : "auto-fit";
    if (v.widthPct != null) p.width.pct = Number(v.widthPct);
    if (v.fontCorrections) p.fontCorrections = v.fontCorrections;
    if (v.debugColors) p.debugColors = v.debugColors;
    return p;
  }

  // Resolve the effective settings for a target.
  //   payload:      parsed v3 tag payload (or null for a plain selection)
  //   profileStore: { name: profile } (localStorage contents)
  //   scope:        { level: "poem"|"bandh"|"cell"|"gap", key?: "A2:3" }
  // Returns { values, source, profileName, profileMissing, usedCache }.
  function resolveSettings(args) {
    args = args || {};
    var payload = args.payload || null;
    var store = args.profileStore || {};
    var scope = args.scope || { level: "poem" };
    var values = defaultSettings();
    var source = {};
    Object.keys(values).forEach(function (k) { source[k] = "default"; });

    var profileName = (payload && typeof payload.profile === "string") ? payload.profile : "";
    var prof = profileName ? store[profileName] : null;
    var profileMissing = !!(profileName && !prof);
    var usedCache = false;
    var layer = null;
    var defs = defaultSettings();
    if (prof) layer = settingsFromProfile(prof);
    else if (profileMissing && payload && isObj(payload.profileCache)) { layer = payload.profileCache; usedCache = true; }
    if (layer) {
      Object.keys(layer).forEach(function (k) {
        if (k in values && layer[k] !== defs[k]) { values[k] = layer[k]; source[k] = "profile"; }
      });
    }

    var local = (payload && isObj(payload.local)) ? payload.local : {};
    Object.keys(local).forEach(function (k) {
      if (k in values) { values[k] = local[k]; source[k] = "local"; }
    });

    if (scope.level === "bandh" || scope.level === "cell") {
      if (payload && typeof payload.widthPt === "number" && payload.widthPt > 0) {
        values.misraWidthPt = payload.widthPt; source.misraWidthPt = "bandh";
      }
    }
    if (scope.level === "cell" && scope.key && payload && isObj(payload.overrides) && isObj(payload.overrides[scope.key])) {
      var ov = payload.overrides[scope.key];
      if (ov.strength != null) { values.strength = ov.strength; source.strength = "cell"; }
      if (ov.widthPt != null) { values.misraWidthPt = ov.widthPt; source.misraWidthPt = "cell"; }
      if (ov.capEm != null) { values.capEm = ov.capEm; source.capEm = "cell"; }
    }

    return { values: values, source: source, profileName: profileName, profileMissing: profileMissing, usedCache: usedCache };
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
    defaultSettings: defaultSettings,
    settingsFromProfile: settingsFromProfile,
    profileFromSettings: profileFromSettings,
    resolveSettings: resolveSettings,
  };
}));
