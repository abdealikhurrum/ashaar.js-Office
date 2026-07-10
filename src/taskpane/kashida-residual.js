/**
 * AshaarResidual — hybrid-fill residual spacing. After calligraphic elongation
 * (Mehr tatweel / Jameel font-swap) undershoots the column, close the gap with
 * a CAPPED number of inter-word micro-spaces. Pure (no DOM); node-testable.
 */
(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) module.exports = factory();
  else root.AshaarResidual = factory();
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  // Hair space (U+200A) — same glyph the spacing path uses to realize word gaps.
  var HAIR_SPACE = " ";

  // How many micro-space glyphs to add to close `residualPx`, never exceeding
  // capEm*sizePx per gap (default 0.28em/gap). Returns 0 for non-positive
  // residual / no gaps / unmeasurable glyph.
  function capMicroSpaces(residualPx, gaps, spaceGlyphPx, sizePx, capEm) {
    if (capEm == null) capEm = 0.28;
    if (!(residualPx > 0) || !(gaps > 0) || !(spaceGlyphPx > 0)) return 0;
    var capPx = capEm * sizePx * gaps;
    var wantPx = Math.min(residualPx, capPx);
    return Math.max(0, Math.round(wantPx / spaceGlyphPx));
  }

  return { HAIR_SPACE: HAIR_SPACE, capMicroSpaces: capMicroSpaces };
}));
