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
  // inherit. fillMode always comes from base. widthPt/capEm are null when unset
  // (the caller then uses the matrix target / the 0.28em default).
  function resolveCellOverride(base, override) {
    base = base || {};
    override = override || {};
    var s = (override.strength != null) ? override.strength : base.strength;
    return {
      strength: s,
      fillMode: base.fillMode,
      widthPt: (override.widthPt != null) ? override.widthPt : null,
      capEm: (override.capEm != null) ? override.capEm : null
    };
  }

  return { overrideKey: overrideKey, resolveCellOverride: resolveCellOverride };
}));
