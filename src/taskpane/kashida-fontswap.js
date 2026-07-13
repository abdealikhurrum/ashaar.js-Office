/**
 * AshaarKashidaFontswap — Jameel Noori Kasheeda fills a line by swapping whole
 * connected segments (fasl/PAW) from the base face to the wider "Kasheeda" face.
 * Only fasls whose Kasheeda form is actually wider contribute; the rest measure
 * equal and are never swapped. splitSpans + selectSwapRuns are pure; the width
 * measurement (base vs Kasheeda font on a canvas) lives in the browser caller.
 */
(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) module.exports = factory();
  else root.AshaarKashidaFontswap = factory();
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  // Letters that do NOT join to the following letter → a segment ends after them.
  var NONJOIN = "اأإآٱدذڈرزڑژوؤءے";

  function splitSpans(text) {
    var spans = [], cur = "";
    for (var i = 0; i < text.length; i++) {
      var ch = text.charAt(i);
      if (ch === " ") { if (cur) { spans.push(cur); cur = ""; } spans.push(" "); continue; }
      cur += ch;
      if (NONJOIN.indexOf(ch) !== -1) { spans.push(cur); cur = ""; }
    }
    if (cur) spans.push(cur);
    return spans;
  }

  // maxPx (optional): hard ceiling above targetPx (e.g. the cell text box).
  // When given, one extra step may OVERSHOOT the target if that lands closer to
  // it than staying short — a line whose every remaining fasl gain exceeds the
  // budget otherwise sits at natural width while its bandh-mates stretch
  // (alignment beats exactness). Without maxPx: strict never-overshoot.
  // spaceClosePx (optional): how much the DOWNSTREAM residual filler (capped
  // micro-spaces) can still add. Spaces land a short line on the target to the
  // pixel, so overshoot must yield to them: it fires only when the shortfall
  // that remains AFTER spaces is still worse than the overshoot.
  function selectSwapRuns(spans, widthsBase, widthsWide, targetPx, maxPx, spaceClosePx) {
    var n = spans.length, swap = new Array(n), total = 0, i;
    for (i = 0; i < n; i++) { swap[i] = false; total += widthsBase[i]; }

    var cand = [];
    for (i = 0; i < n; i++) {
      var gain = widthsWide[i] - widthsBase[i];
      if (gain > 0) cand.push({ i: i, gain: gain });
    }
    cand.sort(function (a, b) { return b.gain - a.gain; });

    var reason = null;
    if (!cand.length) reason = "no kasheeda variants";

    for (var k = 0; k < cand.length; k++) {
      var add = cand[k].gain;
      if (total + add <= targetPx) { swap[cand[k].i] = true; total += add; }
    }
    if (maxPx != null && maxPx > targetPx && total < targetPx) {
      // Where the line lands if we DON'T overshoot: spaces close up to
      // spaceClosePx of the shortfall (exactly, capped), never past the target.
      var afterSpaces = Math.min(targetPx, total + (spaceClosePx || 0));
      var shortAfter = targetPx - afterSpaces;
      if (shortAfter > 0) {
        var best = null; // smallest remaining gain = closest overshoot landing
        for (k = cand.length - 1; k >= 0; k--) { if (!swap[cand[k].i]) { best = cand[k]; break; } }
        if (best && total + best.gain <= maxPx &&
            (total + best.gain - targetPx) < shortAfter) {
          swap[best.i] = true; total += best.gain;
        }
      }
    }
    if (!reason && total < targetPx) reason = "discrete steps underfill";

    var runs = [];
    for (i = 0; i < n; i++) runs.push({ text: spans[i], swap: swap[i] });
    return { runs: runs, fill: targetPx > 0 ? total / targetPx : 0, reason: reason };
  }

  return { splitSpans: splitSpans, selectSwapRuns: selectSwapRuns };
}));
