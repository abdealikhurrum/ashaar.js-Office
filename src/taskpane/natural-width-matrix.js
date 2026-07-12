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

  // Compute the shared per-grid-column width vector (px) for a set of same-shape
  // bandhs plus each bandh's per-position fill target.
  //   auto-fit: each position width = longest natural × (1 + headroom); total capped at pagePx.
  //   fixed:    scale so total = (pct/100) × pagePx, proportions from the matrix.
  // colPx[j] is the width of grid column j; a position's target = sum of its spanned colPx.
  function computeTargetGrid(bandhs, opts) {
    bandhs = bandhs || [];
    opts = opts || {};
    var pagePx = Number(opts.pagePx) || 0;
    var headroom = Number(opts.headroom) || 0;

    // sameShape: every bandh shares one GRID.
    var GRID = bandhs.length ? Number(bandhs[0].GRID) || 0 : 0;
    var sameShape = GRID > 0 && bandhs.every(function (b) { return Number(b.GRID) === GRID; });

    // Cross-bandh matrix: longest natural per position (px).
    var flat = [];
    bandhs.forEach(function (b) { (b.cells || []).forEach(function (c) { flat.push(c); }); });
    var matrix = buildMatrix(flat);

    // Per-position width need = longest natural × (1 + headroom). Spread a
    // position's need evenly across the grid columns it spans; a column's width
    // is the max need imposed by any position covering it.
    var need = {};
    Object.keys(matrix).forEach(function (k) { need[k] = matrix[k] * (1 + headroom); });

    var colPx = [];
    for (var j0 = 0; j0 < GRID; j0++) colPx.push(0);
    var layout = {};
    (bandhs[0] && bandhs[0].cells || []).forEach(function (c) { layout[c.key] = { col: c.col, span: c.span }; });
    Object.keys(need).forEach(function (k) {
      var L = layout[k]; if (!L || !L.span) return;
      var per = need[k] / L.span;
      for (var j = L.col; j < L.col + L.span && j < GRID; j++) colPx[j] = Math.max(colPx[j], per);
    });

    var total = colPx.reduce(function (a, b) { return a + b; }, 0);
    if (opts.mode === "fixed") {
      var want = (Number(opts.pct) || 100) / 100 * pagePx;
      if (total > 0 && want > 0) { var kf = want / total; colPx = colPx.map(function (w) { return w * kf; }); }
    } else if (pagePx > 0 && total > pagePx) {
      var ka = pagePx / total; colPx = colPx.map(function (w) { return w * ka; });
    }

    var bandhTargets = bandhs.map(function (b) {
      var t = {};
      (b.cells || []).forEach(function (c) {
        var sum = 0;
        for (var j = c.col; j < c.col + c.span && j < GRID; j++) sum += colPx[j];
        t[c.key] = sum;
      });
      return t;
    });

    return { sameShape: sameShape, colPx: colPx, bandhTargets: bandhTargets };
  }

  return {
    positionKey: positionKey,
    isContentCell: isContentCell,
    buildMatrix: buildMatrix,
    naturalFitTarget: naturalFitTarget,
    cellFitBudget: cellFitBudget,
    computeTargetGrid: computeTargetGrid,
  };
}));
