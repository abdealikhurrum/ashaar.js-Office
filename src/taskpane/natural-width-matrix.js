/**
 * AshaarMatrix — the per-position natural-width matrix and the two justification
 * modes' target math (pure, UMD; no DOM/Office dependency, node-testable).
 *
 * Visual harmony means corresponding cells across bandhs (stanzas) share a
 * width. A content cell's POSITION is its grid signature (row within the bandh,
 * grid-column start, span). The matrix maps each position → Wpos, the longest
 * tatweel-free ("natural") width among all content cells at that position.
 *
 *   Cell-fit budget:  natural + φ·(colPx − natural)   — φ = elongation share
 *   Natural-fit target: Wpos   + φ·(reach − Wpos)     — φ = misra width dial
 *
 * See docs/superpowers/specs/2026-07-11-poetry-justification-modes-design.md.
 */
(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) module.exports = factory();
  else root.AshaarMatrix = factory();
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  // Stable signature for a grid position. Cells with the same (row, col, span)
  // across bandhs occupy the "same position" and balance to one width.
  function positionKey(sig) {
    sig = sig || {};
    return (sig.row || 0) + ":" + (sig.col || 0) + ":" + (sig.span || 0);
  }

  // A content cell holds misra text (participates in the matrix and is
  // justified); a spacing cell is a structural gap (no text) — excluded.
  function isContentCell(text) {
    return !!(text && String(text).replace(/\s+/g, "").length);
  }

  // cells: [{key, natural}] → { key → max natural }. The baseline Wpos per
  // position is the LONGEST natural width, so shorter misras elongate up to it.
  function buildMatrix(cells) {
    var out = {};
    (cells || []).forEach(function (c) {
      if (!c || !c.key) return;
      var n = Number(c.natural) || 0;
      if (!(c.key in out) || n > out[c.key]) out[c.key] = n;
    });
    return out;
  }

  // Natural-fit total fill target: harmony baseline extended toward the
  // container by strength. φ=0 → Wpos (pure harmony); φ=1 → reach.
  function naturalFitTarget(Wpos, reach, phi) {
    Wpos = Number(Wpos) || 0;
    return Wpos + (Number(phi) || 0) * Math.max(0, (Number(reach) || 0) - Wpos);
  }

  // Cell-fit elongation budget: strength is the elongation:spacing ratio, so φ
  // is the fraction of the cell gap the tatweel engine tries to cover; the rest
  // is left for Word's distribute residual. φ=0 → natural (spacing-only).
  function cellFitBudget(natural, colPx, phi) {
    natural = Number(natural) || 0;
    return natural + (Number(phi) || 0) * Math.max(0, (Number(colPx) || 0) - natural);
  }

  return {
    positionKey: positionKey,
    isContentCell: isContentCell,
    buildMatrix: buildMatrix,
    naturalFitTarget: naturalFitTarget,
    cellFitBudget: cellFitBudget,
  };
}));
