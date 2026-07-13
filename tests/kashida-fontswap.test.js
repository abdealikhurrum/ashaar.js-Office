"use strict";
const assert = require("assert");
const K = require("../src/taskpane/kashida-fontswap");

// splitSpans: break at non-joining letters. "ستارہ" = س-ت-ا | ر | ہ
const spans = K.splitSpans("ستارہ");
assert.deepStrictEqual(spans, ["ستا", "ر", "ہ"]);
assert.strictEqual(spans.join(""), "ستارہ");

// selectSwapRuns: swap highest-gain fasls (wider face) until <= target.
// base sum 30; wider adds gain [+8,+2,+0]; target 36 → swap span1 (+2 → 32), span0 (+8 → 38>36 skip)
const r = K.selectSwapRuns(["a","b","c"], [10,10,10], [18,12,10], 36);
assert.strictEqual(r.runs.length, 3);
assert.strictEqual(r.runs[1].swap, true);
assert.strictEqual(r.runs[0].swap, false);
assert.strictEqual(r.runs[2].swap, false);
assert.ok(r.fill > 0 && r.fill <= 1);

// no fasl has a wider variant (widths equal) → reason set, nothing swapped
const r2 = K.selectSwapRuns(["a","b"], [10,10], [10,10], 40);
assert.strictEqual(r2.runs.every(function (x){return !x.swap;}), true);
assert.strictEqual(r2.reason, "no kasheeda variants");

// Bounded overshoot (maxPx): when every remaining step overshoots the target,
// staying short can leave a line far from its bandh-mates (seen live: nat 116,
// tgt 140, achieved 117 — every fasl gain > 23). Take the SMALLEST remaining
// step when it lands closer to the target than staying short, never past maxPx.
{
  // base 100; gains [30, 25]; target 124 → greedy takes nothing (both overshoot).
  // maxPx 140: smallest remaining gain 25 → 125, |125-124|=1 < |124-100|=24 → swap.
  const o = K.selectSwapRuns(["a","b"], [50,50], [80,75], 124, 140);
  assert.strictEqual(o.runs[1].swap, true, "smallest overshoot step taken");
  assert.strictEqual(o.runs[0].swap, false, "larger step not taken");
  // Overshoot must never pass maxPx: same line, maxPx 120 → no swap.
  const o2 = K.selectSwapRuns(["a","b"], [50,50], [80,75], 124, 120);
  assert.strictEqual(o2.runs.every(function (x){return !x.swap;}), true, "capped by maxPx");
  // Overshoot only when CLOSER than the undershoot: base 100, gain 30,
  // target 110 → 130 is farther (20) than 100 (10) → no swap.
  const o3 = K.selectSwapRuns(["a"], [100], [130], 110, 200);
  assert.strictEqual(o3.runs[0].swap, false, "overshoot farther than undershoot rejected");
  // Without maxPx the old strict-under behavior is unchanged.
  const o4 = K.selectSwapRuns(["a","b"], [50,50], [80,75], 124);
  assert.strictEqual(o4.runs.every(function (x){return !x.swap;}), true, "no maxPx -> never overshoot");

  // spaceClosePx: the residual micro-spaces close a shortfall EXACTLY, so an
  // overshooting swap must yield to them (seen live: a +19 swap overshot to 147
  // when spaces could land the line on 140 to the pixel — misaligning the pair).
  // base 100, gain 25, target 124, maxPx 140:
  //  - spaces can close 24 (>= shortfall) → no overshoot, spaces finish it.
  const s1 = K.selectSwapRuns(["a","b"], [50,50], [80,75], 124, 140, 30);
  assert.strictEqual(s1.runs.every(function (x){return !x.swap;}), true, "spaces can close fully -> no overshoot");
  //  - spaces can close only 10 → after-spaces lands at 110 (14 short);
  //    overshoot to 125 is |1| off → closer → swap.
  const s2 = K.selectSwapRuns(["a","b"], [50,50], [80,75], 124, 140, 10);
  assert.strictEqual(s2.runs[1].swap, true, "spaces too capped -> overshoot wins");
  //  - spaces close 20 → after-spaces 120 (4 short); overshoot 125 is 1 off → still closer → swap.
  const s3 = K.selectSwapRuns(["a","b"], [50,50], [80,75], 124, 140, 20);
  assert.strictEqual(s3.runs[1].swap, true, "overshoot closer than capped spaces -> swap");
  //  - spaces close 22 → after-spaces 122 (2 short); overshoot 125 is 1 off... closer → swap;
  //    but with gain 30 (→130, 6 off) spaces win.
  const s4 = K.selectSwapRuns(["a"], [100], [130], 124, 140, 22);
  assert.strictEqual(s4.runs[0].swap, false, "overshoot farther than capped-space landing -> spaces win");
}

// ── splitSpans keeps combining marks attached to their base letter ───────────
// Orphaned marks shattered vocalized words once MarkSafe made vocalized fasls
// swappable: a span starting with shadda/fatha/damma renders on a dotted
// circle when its run gets a different font (seen live 2026-07-13).
{
  const MARK_AT_START = /^[ؐ-ًؚ-ٰٟۖ-ۜ۟-۪ۨ-ۭ]/;
  assert.deepStrictEqual(K.splitSpans("تَدَّعيهِ"), ["تَدَّ", "عيهِ"], "shadda+fatha stay with د");
  assert.deepStrictEqual(K.splitSpans("صِغارُها"), ["صِغا", "رُ", "ها"], "damma stays with ر");
  assert.deepStrictEqual(K.splitSpans("عَجِزَت"), ["عَجِزَ", "ت"], "fatha stays with ز");
  assert.deepStrictEqual(K.splitSpans("هَمَّهُ"), ["هَمَّهُ"], "fully-joined word unchanged");
  // Property: NO span may ever start with a combining mark.
  ["تَدَّعيهِ الضَراغِم", "صِغارُها", "وَقَدْ", "أَوَّاب"].forEach(function (t) {
    K.splitSpans(t).forEach(function (s) {
      assert.strictEqual(MARK_AT_START.test(s), false, "span starts with mark: " + JSON.stringify(s) + " in " + t);
    });
  });
  // Round-trip: spans always reassemble to the input.
  assert.strictEqual(K.splitSpans("تَدَّعيهِ الضَراغِم").join(""), "تَدَّعيهِ الضَراغِم");
}

console.log("kashida-fontswap tests passed");
