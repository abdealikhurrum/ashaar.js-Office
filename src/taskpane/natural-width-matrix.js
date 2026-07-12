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
    var mode = opts.mode === "fixed" ? "fixed" : "auto-fit";
    var pctWant = (Number(opts.pct) || 100) / 100 * pagePx;

    var GRID = bandhs.length ? Number(bandhs[0].GRID) || 0 : 0;
    var sameShape = GRID > 0 && bandhs.every(function (b) { return Number(b.GRID) === GRID; });

    // colPx vector for one bandh: each position's need (longest natural × (1+headroom),
    // from `matrix` when shared) spread evenly over its spanned columns; a column
    // takes the max need of any position covering it.
    function vectorFor(cells, grid, matrix) {
      var v = []; for (var j = 0; j < grid; j++) v.push(0);
      (cells || []).forEach(function (c) {
        var w = ((matrix ? matrix[c.key] : c.natural) || 0) * (1 + headroom);
        var per = w / c.span;
        for (var j = c.col; j < c.col + c.span && j < grid; j++) v[j] = Math.max(v[j], per);
      });
      return v;
    }
    function totalOf(v) { return v.reduce(function (a, b) { return a + b; }, 0); }
    function scaleTo(v, targetTotal) {
      var tot = totalOf(v);
      if (tot <= 0 || targetTotal <= 0) return v;
      var k = targetTotal / tot; return v.map(function (w) { return w * k; });
    }
    function targetsFrom(cells, v, grid) {
      var t = {};
      (cells || []).forEach(function (c) {
        var sum = 0; for (var j = c.col; j < c.col + c.span && j < grid; j++) sum += v[j];
        t[c.key] = sum;
      });
      return t;
    }

    if (sameShape) {
      var flat = []; bandhs.forEach(function (b) { (b.cells || []).forEach(function (c) { flat.push(c); }); });
      var matrix = buildMatrix(flat);
      // Union of positions across all bandhs (a position may be absent from bandh[0]).
      var layout = {};
      bandhs.forEach(function (b) { (b.cells || []).forEach(function (c) { if (!layout[c.key]) layout[c.key] = { col: c.col, span: c.span, key: c.key }; }); });
      var colPx = vectorFor(Object.keys(layout).map(function (k) { return { key: k, col: layout[k].col, span: layout[k].span }; }), GRID, matrix);
      var total = totalOf(colPx);
      if (mode === "fixed") colPx = scaleTo(colPx, pctWant);
      else if (pagePx > 0 && total > pagePx) colPx = scaleTo(colPx, pagePx);
      var bandhTargets = bandhs.map(function (b) { return targetsFrom(b.cells, colPx, GRID); });
      return { sameShape: true, colPx: colPx, perBandhColPx: bandhs.map(function () { return colPx; }), bandhTargets: bandhTargets };
    }

    // Different shapes: per-bandh vectors, all scaled to one shared total.
    var vecs = bandhs.map(function (b) { return vectorFor(b.cells, Number(b.GRID) || 0, null); });
    var totals = vecs.map(totalOf);
    var shared = mode === "fixed" ? pctWant : Math.min(pagePx || Infinity, Math.max.apply(null, totals.concat([0])));
    var scaled = vecs.map(function (v) { return scaleTo(v, shared); });
    var bt = bandhs.map(function (b, i) { return targetsFrom(b.cells, scaled[i], Number(b.GRID) || 0); });
    return { sameShape: false, colPx: null, perBandhColPx: scaled, bandhTargets: bt };
  }

  // Option A: the single grid-slot width (px) for the whole qaseeda. Non-uniform
  // CELL widths come from cells spanning different numbers of these equal slots,
  // so one slot size preserves the skinny-gap / wide-text layout and harmonizes
  // same-shape bandhs (same GRID + same slot → matching cells equal).
  //   rawSlot   = max over content cells of natural×(1+headroom) / span
  //   auto-fit  = min(rawSlot, pagePx / maxGRID)   — widest bandh fits the page
  //   fixed     = (pct/100 × pagePx) / maxGRID     — widest bandh = pct of page
  function uniformSlotPx(bandhs, opts) {
    bandhs = bandhs || [];
    opts = opts || {};
    var pagePx = Number(opts.pagePx) || 0;
    var headroom = Number(opts.headroom) || 0;
    var maxGRID = 0;
    bandhs.forEach(function (b) { maxGRID = Math.max(maxGRID, Number(b.GRID) || 0); });
    if (maxGRID <= 0) return 0;

    if (opts.mode === "fixed") {
      return (Number(opts.pct) || 100) / 100 * pagePx / maxGRID;
    }
    var rawSlot = 0;
    bandhs.forEach(function (b) {
      (b.cells || []).forEach(function (c) {
        var span = Number(c.span) || 1;
        var need = (Number(c.natural) || 0) * (1 + headroom) / span;
        if (need > rawSlot) rawSlot = need;
      });
    });
    var cap = pagePx > 0 ? pagePx / maxGRID : Infinity;
    return Math.min(rawSlot, cap);
  }

  return {
    positionKey: positionKey,
    isContentCell: isContentCell,
    buildMatrix: buildMatrix,
    naturalFitTarget: naturalFitTarget,
    cellFitBudget: cellFitBudget,
    computeTargetGrid: computeTargetGrid,
    uniformSlotPx: uniformSlotPx,
  };
}));
