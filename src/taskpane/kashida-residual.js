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
  var HAIR_SPACE = " ";

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

  // Spread `n` spaceChar glyphs across the inter-word (" ") runs of a font-swap
  // run-list, as evenly as possible (earlier gaps take the remainder). Returns a
  // new array of copied run objects; input is never mutated.
  function injectSpaceRuns(runs, n, spaceChar) {
    if (spaceChar == null) spaceChar = HAIR_SPACE;
    var out = (runs || []).map(function (r) {
      var c = {}; for (var k in r) { if (r.hasOwnProperty(k)) c[k] = r[k]; } return c;
    });
    if (!(n > 0)) return out;
    var gapIdx = [];
    for (var i = 0; i < out.length; i++) { if (out[i].text === " ") gapIdx.push(i); }
    if (!gapIdx.length) return out;
    var base = Math.floor(n / gapIdx.length), rem = n % gapIdx.length;
    for (var g = 0; g < gapIdx.length; g++) {
      var add = base + (g < rem ? 1 : 0);
      if (add > 0) out[gapIdx[g]].text += new Array(add + 1).join(spaceChar);
    }
    return out;
  }

  return { HAIR_SPACE: HAIR_SPACE, capMicroSpaces: capMicroSpaces, injectSpaceRuns: injectSpaceRuns };
}));
