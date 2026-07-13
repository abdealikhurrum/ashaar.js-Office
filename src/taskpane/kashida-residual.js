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
  //
  // The 0.28em/gap cap is a deliberate typographic limit: when elongation +
  // capped spacing still don't reach the column edge, the line is ACCEPTED
  // SHORT rather than over-spaced (unbounded word gaps read worse than a
  // slightly short line). Recourse for a persistently short line is to give the
  // engine more room to distribute into — INCREASE table width % or REDUCE the
  // font size — not to lift the cap. See the hybrid-fill design §5a.
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

  // 96 px/inch vs 1440 twips/inch → 1px = 15 twips.
  var PX_TO_TWIPS = 15;

  // Close `residualPx` with per-gap CHARACTER SPACING (rPr w:spacing twips on a
  // dedicated single-space run per word gap) instead of injected glyphs. Gaps
  // come out pixel-exact and even (±1 twip), and the text round-trips as the
  // clean source — nothing to strip on re-capture. Same typographic cap as
  // capMicroSpaces: at most capEm*sizePx per gap; accept-short beyond it.
  // Splits each run's text at spaces (word runs + space runs, style props
  // copied); pre-existing standalone " " runs count as gaps too. When nothing
  // applies (no positive residual / no gaps) returns flat copies unsplit.
  // Returns { runs, appliedPx, gaps }.
  function spreadResidualSpacing(runs, residualPx, sizePx, capEm) {
    if (capEm == null) capEm = 0.28;
    function copyRun(r, text) {
      var c = {}; for (var k in r) { if (r.hasOwnProperty(k)) c[k] = r[k]; }
      if (text != null) c.text = text;
      return c;
    }
    var flat = (runs || []).map(function (r) { return copyRun(r); });
    var gaps = flat.reduce(function (a, r) {
      return a + (String(r.text || "").split(" ").length - 1);
    }, 0);
    if (!(residualPx > 0) || !(gaps > 0)) return { runs: flat, appliedPx: 0, gaps: gaps };

    var perGapPx = Math.min(residualPx / gaps, (Number(capEm) || 0) * (Number(sizePx) || 0));
    var totalTwips = Math.max(0, Math.round(perGapPx * gaps * PX_TO_TWIPS));
    var base = Math.floor(totalTwips / gaps), rem = totalTwips % gaps;

    var out = [], gi = 0;
    (runs || []).forEach(function (r) {
      var parts = String(r.text || "").split(" ");
      parts.forEach(function (word, pi) {
        if (pi > 0) {
          var sp = copyRun(r, " ");
          var tw = base + (gi < rem ? 1 : 0);
          if (tw > 0) sp.spacingTwips = tw;
          gi++;
          out.push(sp);
        }
        if (word) out.push(copyRun(r, word));
      });
    });
    return { runs: out, appliedPx: totalTwips / PX_TO_TWIPS, gaps: gaps };
  }

  return { HAIR_SPACE: HAIR_SPACE, capMicroSpaces: capMicroSpaces, injectSpaceRuns: injectSpaceRuns, spreadResidualSpacing: spreadResidualSpacing };
}));
