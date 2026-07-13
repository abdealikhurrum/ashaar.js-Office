/**
 * AshaarOverrides — per-cell justify override key + merge (pure, UMD;
 * node-testable, no DOM/Office). An override deviates one cell from its block's
 * justify defaults on strength / target width / cap-lift. fillMode is NOT
 * overridable (block-wide choice). See
 * docs/superpowers/specs/2026-07-11-per-cell-overrides-design.md.
 */
(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) module.exports = factory();
  else root.AshaarOverrides = factory();
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  // Per-cell key: the SP1 label prefixed by the cell's table (bandh) index in
  // the block, so an override targets exactly one cell.
  function overrideKey(tableIndex, label) {
    return (tableIndex || 0) + ":" + label;
  }

  // Merge a cell override onto the block's base justify settings. Absent fields
  // inherit. fillMode always comes from base. widthPt inherits base.widthPt
  // (the bandh- or qaseeda-level misra width) when the cell doesn't set one;
  // null when unset everywhere (the caller then uses the matrix target).
  // capEm is null when unset (the caller applies its own default).
  function resolveCellOverride(base, override) {
    base = base || {};
    override = override || {};
    var s = (override.strength != null) ? override.strength : base.strength;
    return {
      strength: s,
      fillMode: base.fillMode,
      widthPt: (override.widthPt != null) ? override.widthPt
        : (base.widthPt != null) ? base.widthPt : null,
      capEm: (override.capEm != null) ? override.capEm : null
    };
  }

  // Merge a per-slot decoration override onto the profile default. Per field: an
  // override key present wins (empty string = explicit none); else inherit the
  // profile; else "". Fields: symbol / fill / color.
  function resolveSlotDecor(profileDecor, override) {
    profileDecor = profileDecor || {};
    override = override || {};
    function pick(k) { return (k in override) ? (override[k] || "") : (profileDecor[k] || ""); }
    return { symbol: pick("symbol"), fill: pick("fill"), color: pick("color") };
  }

  return {
    overrideKey: overrideKey,
    resolveCellOverride: resolveCellOverride,
    resolveSlotDecor: resolveSlotDecor
  };
}));
