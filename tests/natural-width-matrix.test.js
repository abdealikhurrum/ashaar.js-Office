const assert = require("assert");
const AshaarMatrix = require("../src/taskpane/natural-width-matrix");

// ── positionKey: stable signature, order-independent fields ──────────────────
assert.strictEqual(AshaarMatrix.positionKey({ row: 0, col: 6, span: 6 }), "0:6:6");
assert.strictEqual(AshaarMatrix.positionKey({ row: 1, col: 0 }), "1:0:0"); // missing span → 0
assert.strictEqual(
  AshaarMatrix.positionKey({ row: 0, col: 0, span: 6 }),
  AshaarMatrix.positionKey({ row: 0, col: 0, span: 6 }),
  "same signature → same key"
);
assert.notStrictEqual(
  AshaarMatrix.positionKey({ row: 0, col: 0, span: 6 }),
  AshaarMatrix.positionKey({ row: 0, col: 6, span: 6 }),
  "different column → different key"
);

// ── buildMatrix: longest natural per position across bandhs ──────────────────
{
  // Two bandhs (A, B) sharing a 3-cell template. Position "0:0:4" appears twice.
  const cells = [
    { key: "0:0:4", natural: 300 }, // bandh A
    { key: "0:0:4", natural: 280 }, // bandh B → position width is the max (300)
    { key: "0:4:4", natural: 250 },
    { key: "0:8:4", natural: 260 },
  ];
  const m = AshaarMatrix.buildMatrix(cells);
  assert.strictEqual(m["0:0:4"], 300, "position takes the longest natural");
  assert.strictEqual(m["0:4:4"], 250);
  assert.strictEqual(m["0:8:4"], 260);
}
{
  // A position with a single cell → its own natural (trivial single-bandh case).
  const m = AshaarMatrix.buildMatrix([{ key: "0:0:6", natural: 325 }]);
  assert.strictEqual(m["0:0:6"], 325);
}
{
  // Empty / degenerate.
  assert.deepStrictEqual(AshaarMatrix.buildMatrix([]), {});
  assert.deepStrictEqual(AshaarMatrix.buildMatrix(null), {});
}

// ── isContentCell ────────────────────────────────────────────────────────────
assert.strictEqual(AshaarMatrix.isContentCell("دل ناداں"), true);
assert.strictEqual(AshaarMatrix.isContentCell(""), false);
assert.strictEqual(AshaarMatrix.isContentCell("   "), false);
assert.strictEqual(AshaarMatrix.isContentCell(null), false);
assert.strictEqual(AshaarMatrix.isContentCell(undefined), false);

// ── naturalFitTarget: Wpos at φ=0, reach at φ=1, linear between ──────────────
assert.strictEqual(AshaarMatrix.naturalFitTarget(300, 400, 0), 300, "φ=0 → harmony baseline Wpos");
assert.strictEqual(AshaarMatrix.naturalFitTarget(300, 400, 1), 400, "φ=1 → reach (container)");
assert.strictEqual(AshaarMatrix.naturalFitTarget(300, 400, 0.5), 350, "φ=0.5 → halfway");
assert.strictEqual(AshaarMatrix.naturalFitTarget(400, 300, 0.5), 400, "Wpos>reach → no shrink (max 0)");

// ── cellFitBudget: natural at φ=0, colPx at φ=1 ──────────────────────────────
assert.strictEqual(AshaarMatrix.cellFitBudget(200, 400, 0), 200, "φ=0 → all spacing (no elongation)");
assert.strictEqual(AshaarMatrix.cellFitBudget(200, 400, 1), 400, "φ=1 → elongation-dominant to edge");
assert.strictEqual(AshaarMatrix.cellFitBudget(200, 400, 0.5), 300);
assert.strictEqual(AshaarMatrix.cellFitBudget(500, 400, 0.5), 500, "natural>colPx → no negative budget");

console.log("natural-width-matrix.test.js OK");
