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
  const tag0 = AshaarWord.contentControlTag("poem", { profile: "Q", local: { widthPct: 60 } }, [[["c", "g", "c"]]]);
  const tag1 = AshaarWord.setTagOverride(tag0, "0:A1", { strength: 9 });
  const p1 = AshaarWord.parseContentControlTag(tag1);
  assert.deepStrictEqual(p1.overrides, { "0:A1": { strength: 9 } }, "override added");
  assert.equal(p1.profile, "Q", "other payload fields intact");
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

// ── resolveSlotDecor: override wins per field; "" = explicit none; else inherit
{
  const prof = { symbol: "؎", fill: "f5f0e0", color: "a7352a" };
  assert.deepStrictEqual(AshaarOverrides.resolveSlotDecor(prof, null),
    { symbol: "؎", fill: "f5f0e0", color: "a7352a" }, "no override → profile");
  assert.deepStrictEqual(AshaarOverrides.resolveSlotDecor(prof, { symbol: "*" }),
    { symbol: "*", fill: "f5f0e0", color: "a7352a" }, "symbol overridden");
  assert.deepStrictEqual(AshaarOverrides.resolveSlotDecor(prof, { symbol: "" }),
    { symbol: "", fill: "f5f0e0", color: "a7352a" }, "empty string suppresses");
  assert.deepStrictEqual(AshaarOverrides.resolveSlotDecor(null, null),
    { symbol: "", fill: "", color: "" }, "nothing → all none");
}

// ── setTagSlotDecor: add / remove, round-trip, cells/overrides intact ────────
{
  const t0 = AshaarWord.contentControlTag("poem", { profile: "Q" }, [[["c", "g", "c"]]]);
  const t1 = AshaarWord.setTagSlotDecor(t0, "0:A#1", { symbol: "؎", fill: "eeeeee" });
  const p1 = AshaarWord.parseContentControlTag(t1);
  assert.deepStrictEqual(p1.slotDecor, { "0:A#1": { symbol: "؎", fill: "eeeeee" } });
  assert.deepStrictEqual(p1.cells, [[["c", "g", "c"]]], "cells intact");
  assert.equal(p1.profile, "Q");
  const t2 = AshaarWord.setTagSlotDecor(t1, "0:A#1", null);
  assert.deepStrictEqual(AshaarWord.parseContentControlTag(t2).slotDecor, {}, "removed");
  const t3 = AshaarWord.setTagSlotDecor(t1, "0:A#1", { symbol: "", fill: "", color: "" });
  assert.deepStrictEqual(AshaarWord.parseContentControlTag(t3).slotDecor, {}, "all-empty removes");
  assert.strictEqual(AshaarWord.setTagSlotDecor("nope", "0:A#1", { symbol: "x" }), "nope");
}

// ── resolveCellOverride: widthPt inherits base (bandh/qaseeda misra width) ───
{
  const base = { strength: 6, fillMode: "natural-fit", widthPt: 120 };
  assert.strictEqual(AshaarOverrides.resolveCellOverride(base, null).widthPt, 120,
    "no cell override → inherits base widthPt");
  assert.strictEqual(AshaarOverrides.resolveCellOverride(base, { widthPt: 90 }).widthPt, 90,
    "cell widthPt wins over base");
  assert.strictEqual(AshaarOverrides.resolveCellOverride({ strength: 6, fillMode: "natural-fit" }, null).widthPt, null,
    "unset everywhere → null (computed target)");
}

// ── colorClearKeys: keys transitioning color→none under the fan-out Apply ───
{
  const old = {
    "0:A1": { strength: 9, color: "#A7352A" },
    "0:A2": { fill: "#F5F0E0" },              // fill only — no color to clear
    "0:B1": { color: "#112233" },
    "1:A1": { color: "#445566" },             // not targeted — must not appear
  };
  const keys = ["0:A1", "0:A2", "0:B1", "0:B2"]; // B2 has no old override at all
  // Incoming override drops color → every targeted key that HAD one clears.
  assert.deepStrictEqual(
    AshaarOverrides.colorClearKeys(old, keys, { strength: 5, color: null }),
    ["0:A1", "0:B1"]
  );
  // Empty string counts as no color on both sides.
  assert.deepStrictEqual(
    AshaarOverrides.colorClearKeys({ "0:A1": { color: "" } }, ["0:A1"], { color: null }),
    [], "old empty-string color is not a real color");
  assert.deepStrictEqual(
    AshaarOverrides.colorClearKeys(old, keys, { color: "" }),
    ["0:A1", "0:B1"], "incoming empty string = no color");
  // Incoming override still carries a color → nothing clears.
  assert.deepStrictEqual(
    AshaarOverrides.colorClearKeys(old, keys, { color: "#000000" }),
    []);
  // Degenerate inputs.
  assert.deepStrictEqual(AshaarOverrides.colorClearKeys(null, keys, null), []);
  assert.deepStrictEqual(AshaarOverrides.colorClearKeys(old, null, null), []);
}

// ── mergeFanOutOverride: fan-out onto sibling keys preserves untouched fields
// (final review I2 — a bandh/poem-target Apply must not wipe unrelated
// per-cell strength/color when only e.g. fill was edited) ──────────────────
{
  const existing = { strength: 9, widthPt: 120, color: "#112233" };
  const incoming = { strength: null, widthPt: null, capEm: null, fill: "#F5F0E0", color: null };
  // Only "fill" touched → strength/widthPt/color survive from existing;
  // capEm stays null (never set anywhere); fill comes from incoming.
  assert.deepStrictEqual(
    AshaarOverrides.mergeFanOutOverride(existing, incoming, { fill: true }),
    { strength: 9, widthPt: 120, capEm: null, fill: "#F5F0E0", color: "#112233" },
    "untouched fields survive on a non-current key"
  );

  // A ⟲-cleared field (touched=true, incoming=null) clears even though the
  // sibling key had its own value — the clear is spec-mandated to fan out.
  assert.deepStrictEqual(
    AshaarOverrides.mergeFanOutOverride(existing, incoming, { strength: true }),
    { strength: null, widthPt: 120, capEm: null, fill: null, color: "#112233" },
    "⟲-cleared field clears on every targeted key, even untouched ones' own value"
  );

  // Nothing touched → the key is untouched by this Apply entirely; every
  // field falls back to its own existing value.
  assert.deepStrictEqual(
    AshaarOverrides.mergeFanOutOverride(existing, incoming, {}),
    { strength: 9, widthPt: 120, capEm: null, fill: null, color: "#112233" },
    "no touched fields → sibling key is untouched"
  );

  // A key with no prior override at all: untouched fields stay null (no
  // existing value to fall back to); touched fields still land.
  assert.deepStrictEqual(
    AshaarOverrides.mergeFanOutOverride(null, { strength: 8, widthPt: null, capEm: null, fill: null, color: null },
      { strength: true }),
    { strength: 8, widthPt: null, capEm: null, fill: null, color: null },
    "no existing override → untouched fields stay null, touched field lands"
  );
}

console.log("cell-overrides.test.js OK");
