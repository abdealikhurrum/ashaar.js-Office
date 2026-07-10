const assert = require("assert");
const AshaarJustify = require("../src/vendor/ashaar-justify");

const TATWEEL = "ـ";

// ── insertIntoWords ─────────────────────────────────────────────────────────
// Insert 2 tatweels at position 1 of word 0 ("بيت" → "ب" + ــ + "يت").
{
  const out = AshaarJustify.insertIntoWords("بيت نور", { "0:1": 2 }, TATWEEL);
  assert.equal(out, "ب" + TATWEEL + TATWEEL + "يت نور");
}

// ── measureRunsNatural ──────────────────────────────────────────────────────
// Sum of per-run measures, with tatweels stripped first.
{
  const runs = [
    { text: "بيت", measure: (s) => s.length },
    { text: "ن" + TATWEEL + "ور", measure: (s) => s.length }, // stripped → "نور" (len 3)
  ];
  assert.equal(AshaarJustify.measureRunsNatural(runs), 6);
}

// measureRunsNatural throws when a run lacks measure().
assert.throws(function () {
  AshaarJustify.measureRunsNatural([{ text: "بيت" }]);
}, TypeError);

// ── applySlotsMulti ─────────────────────────────────────────────────────────
// Two runs, one slot each; n=2 round-robins one tatweel into each run.
{
  const slots = [
    { ri: 0, wi: 0, pos: 1, score: 7 },
    { ri: 1, wi: 0, pos: 1, score: 7 },
  ];
  const out = AshaarJustify.applySlotsMulti(["بت", "نر"], slots, 2);
  assert.deepEqual(out, ["ب" + TATWEEL + "ت", "ن" + TATWEEL + "ر"]);
}

// applySlotsMulti with n=0 or no slots returns the texts unchanged.
assert.deepEqual(AshaarJustify.applySlotsMulti(["بت", "نر"], [], 5), ["بت", "نر"]);
assert.deepEqual(AshaarJustify.applySlotsMulti(["بت"], [{ ri: 0, wi: 0, pos: 1, score: 7 }], 0), ["بت"]);

// ── applySlots parity (guards the insertIntoWords refactor) ──────────────────
// Existing single-string applySlots still inserts correctly via the shared helper.
{
  const slots = AshaarJustify.buildSlots("بيت", {}, null);
  const out = AshaarJustify.applySlots("بيت", slots, 1);
  assert.equal((out.match(/ـ/g) || []).length, 1);
}

// ── justifyRuns: two runs, fake measure = char length (tatweel width 1) ──────
{
  const runs = [
    { text: "بيت", measure: (s) => s.length }, // 2 slots (ب→ي, ي→ت)
    { text: "نور", measure: (s) => s.length }, // 1 slot  (ن→و; و is right-join)
  ];
  // natural = 6, target = 12 → 6 tatweels distributed across both runs.
  const out = AshaarJustify.justifyRuns(runs, 12, { targetFill: 1 });
  assert.equal(out.length, 2);
  const total = (out[0].text + out[1].text).split("").filter((c) => c === TATWEEL).length;
  assert.equal(total, 6);
  assert.ok(out[0].text.length > 3, "run 0 received tatweels");
  assert.ok(out[1].text.length > 3, "run 1 received tatweels");
}

// ── justifyLine parity: delegating through justifyRuns is transparent ────────
{
  const fakeCtx = { measureText: (s) => ({ width: s.length }) };
  const viaLine = AshaarJustify.justifyLine("بيت", 12, fakeCtx, { targetFill: 1 }, null);
  const viaRuns = AshaarJustify.justifyRuns(
    [{ text: "بيت", measure: (s) => s.length }], 12, { targetFill: 1 }
  )[0].text;
  assert.equal(viaLine, viaRuns);
}

// ── quality-first across runs: boosted run fills before the other ────────────
{
  const boosted = { getQuality: () => 0.9 };
  const runs = [
    { text: "بيت", measure: (s) => s.length, fontProfile: boosted },
    { text: "نور", measure: (s) => s.length }, // no profile
  ];
  // Just above natural: only ~1 tatweel fits → it must land in the boosted run.
  const out = AshaarJustify.justifyRuns(runs, 7, { targetFill: 1, fontQualityBoost: 100 });
  assert.ok(out[0].text.length > 3, "boosted run 0 received the tatweel");
  assert.equal(out[1].text, "نور", "run 1 unchanged");
}

// ── edge cases ───────────────────────────────────────────────────────────────
assert.deepEqual(AshaarJustify.justifyRuns([], 100, {}), []);
assert.deepEqual(
  AshaarJustify.justifyRuns([{ text: "   ", measure: (s) => s.length }], 100, {}),
  [{ text: "   " }]
); // whitespace-only: no visible text
assert.deepEqual(
  AshaarJustify.justifyRuns([{ text: "hello", measure: (s) => s.length }], 100, {}),
  [{ text: "hello" }]
); // Latin: no legal slots
assert.deepEqual(
  AshaarJustify.justifyRuns([{ text: "بيت", measure: (s) => s.length }], 2, {}),
  [{ text: "بيت" }]
); // already >= target
// reducibility: pre-tatweeled input == bare input
{
  const bare = AshaarJustify.justifyRuns([{ text: "بيت", measure: (s) => s.length }], 12, { targetFill: 1 });
  const pre = AshaarJustify.justifyRuns([{ text: "بـيـت", measure: (s) => s.length }], 12, { targetFill: 1 });
  assert.deepEqual(pre, bare);
}
// missing measure throws
assert.throws(function () { AshaarJustify.justifyRuns([{ text: "بيت" }], 12, {}); }, TypeError);

console.log("ashaar-justify-runs: Task 1 + Task 2 helpers OK");
