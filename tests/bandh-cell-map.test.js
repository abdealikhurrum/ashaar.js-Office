const assert = require("assert");
const AshaarCellMap = require("../src/taskpane/bandh-cell-map");

// ── buildBandhCellMap: couplet row [c,g,c] → A1 (sadr/right), gap, A2 (ajuz) ──
{
  const map = AshaarCellMap.buildBandhCellMap([["c", "g", "c"]]);
  assert.equal(map.length, 3);
  assert.deepStrictEqual(
    map.map((e) => e.kind),
    ["content", "spacing", "content"]
  );
  assert.deepStrictEqual(
    map.map((e) => e.label),
    ["A1", null, "A2"],
    "content numbered 1,2 in emission order (RTL rightmost = A1)"
  );
  assert.equal(map[1].slot, "A#1", "gap gets a slot, no label");
  assert.deepStrictEqual(map.map((e) => e.index), [0, 1, 2]);
  assert.deepStrictEqual(map.map((e) => e.row), [0, 0, 0]);
}

// ── solo row [g,c,g] → single A1 flanked by gaps ─────────────────────────────
{
  const map = AshaarCellMap.buildBandhCellMap([["g", "c", "g"]]);
  assert.deepStrictEqual(map.map((e) => e.label), [null, "A1", null]);
  assert.deepStrictEqual(map.map((e) => e.slot), ["A#1", null, "A#2"]);
}

// ── multi-row marsiya bandh → A/B/C letters per row ──────────────────────────
{
  // Row A: [c,g,c] ; Row B: [g,c,g] ; Row C: [c,g,c]
  const map = AshaarCellMap.buildBandhCellMap([
    ["c", "g", "c"],
    ["g", "c", "g"],
    ["c", "g", "c"],
  ]);
  assert.deepStrictEqual(
    map.filter((e) => e.kind === "content").map((e) => e.label),
    ["A1", "A2", "B1", "C1", "C2"]
  );
  assert.deepStrictEqual(map.filter((e) => e.kind === "content").map((e) => e.row), [0, 0, 1, 2, 2]);
}

// ── degenerate ───────────────────────────────────────────────────────────────
assert.deepStrictEqual(AshaarCellMap.buildBandhCellMap([]), []);
assert.deepStrictEqual(AshaarCellMap.buildBandhCellMap(null), []);

// ── alignPatternToTable ──────────────────────────────────────────────────────
assert.strictEqual(AshaarCellMap.alignPatternToTable([3], [["c", "g", "c"]]), true);
assert.strictEqual(AshaarCellMap.alignPatternToTable([3, 3], [["c", "g", "c"]]), false, "row count mismatch");
assert.strictEqual(AshaarCellMap.alignPatternToTable([2], [["c", "g", "c"]]), false, "cell count mismatch");
assert.strictEqual(AshaarCellMap.alignPatternToTable([3], null), false);
assert.strictEqual(AshaarCellMap.alignPatternToTable(null, [["c"]]), false);

console.log("bandh-cell-map.test.js OK");
