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

console.log("kashida-fontswap tests passed");
