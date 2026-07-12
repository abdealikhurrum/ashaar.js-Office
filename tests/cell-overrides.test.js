const assert = require("assert");
const AshaarOverrides = require("../src/taskpane/cell-overrides");
const AshaarWord = require("../src/taskpane/word-html");

// ── overrideKey ──────────────────────────────────────────────────────────────
assert.strictEqual(AshaarOverrides.overrideKey(2, "A1"), "2:A1");
assert.strictEqual(AshaarOverrides.overrideKey(0, "B2"), "0:B2");

// ── resolveCellOverride ──────────────────────────────────────────────────────
{
  const base = { strength: 7, fillMode: "natural-fit" };
  // No override → base strength/fillMode, no width/cap.
  assert.deepStrictEqual(
    AshaarOverrides.resolveCellOverride(base, null),
    { strength: 7, fillMode: "natural-fit", widthPt: null, capEm: null }
  );
  // Strength override wins; fillMode still from base.
  assert.deepStrictEqual(
    AshaarOverrides.resolveCellOverride(base, { strength: 9 }),
    { strength: 9, fillMode: "natural-fit", widthPt: null, capEm: null }
  );
  // Width + cap pass through.
  assert.deepStrictEqual(
    AshaarOverrides.resolveCellOverride(base, { widthPt: 320, capEm: 0.5 }),
    { strength: 7, fillMode: "natural-fit", widthPt: 320, capEm: 0.5 }
  );
  // Empty override object → base.
  assert.deepStrictEqual(
    AshaarOverrides.resolveCellOverride(base, {}),
    { strength: 7, fillMode: "natural-fit", widthPt: null, capEm: null }
  );
}

// ── setTagOverride: add / replace / remove, round-trip, other fields intact ──
{
  const tag0 = AshaarWord.contentControlTag("poem", { qaseeda: "Q", tableWidthPct: 60 }, [[["c", "g", "c"]]]);
  const tag1 = AshaarWord.setTagOverride(tag0, "0:A1", { strength: 9 });
  const p1 = AshaarWord.parseContentControlTag(tag1);
  assert.deepStrictEqual(p1.overrides, { "0:A1": { strength: 9 } }, "override added");
  assert.equal(p1.qaseeda, "Q", "other payload fields intact");
  assert.deepStrictEqual(p1.cells, [[["c", "g", "c"]]], "cells intact");

  const tag2 = AshaarWord.setTagOverride(tag1, "0:A1", { strength: 5, widthPt: 300 });
  assert.deepStrictEqual(AshaarWord.parseContentControlTag(tag2).overrides, { "0:A1": { strength: 5, widthPt: 300 } }, "replaced");

  const tag3 = AshaarWord.setTagOverride(tag2, "0:A1", null);
  assert.deepStrictEqual(AshaarWord.parseContentControlTag(tag3).overrides, {}, "removed → empty map");

  // Empty override object also removes the key.
  const tag4 = AshaarWord.setTagOverride(tag1, "0:A1", {});
  assert.deepStrictEqual(AshaarWord.parseContentControlTag(tag4).overrides, {}, "empty override removes");

  // Non-ashaar tag returned unchanged.
  assert.strictEqual(AshaarWord.setTagOverride("not-ashaar", "0:A1", { strength: 9 }), "not-ashaar");
}

console.log("cell-overrides.test.js OK");
