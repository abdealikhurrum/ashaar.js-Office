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

console.log("ashaar-justify-runs: Task 1 helpers OK");
