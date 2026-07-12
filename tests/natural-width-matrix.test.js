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

// ── computeTargetGrid: harmony + auto-fit + fixed-% ──────────────────────────
{
  // Two same-shape bandhs, GRID=6: cols 0..5. Position 0:0:3 and 0:3:3.
  const bandhs = [
    { GRID: 6, cells: [
      { key: "0:0:3", natural: 180, col: 0, span: 3 },
      { key: "0:3:3", natural: 120, col: 3, span: 3 },
    ]},
    { GRID: 6, cells: [
      { key: "0:0:3", natural: 150, col: 0, span: 3 }, // shorter → matrix keeps 180
      { key: "0:3:3", natural: 140, col: 3, span: 3 }, // longer  → matrix 140
    ]},
  ];
  // auto-fit: each position = longest natural × (1+headroom); columns split evenly within a span.
  const g = AshaarMatrix.computeTargetGrid(bandhs, { mode: "auto-fit", pagePx: 1000, headroom: 0 });
  assert.strictEqual(g.sameShape, true);
  // position 0:0:3 longest=180 → 60/col; 0:3:3 longest=140 → ~46.67/col
  assert.strictEqual(Math.round(g.colPx[0]), 60);
  assert.strictEqual(Math.round(g.colPx[3]), 47);
  // both bandhs get the SAME colPx vector (harmony)
  assert.strictEqual(Math.round(g.bandhTargets[0]["0:0:3"]), 180);
  assert.strictEqual(Math.round(g.bandhTargets[1]["0:0:3"]), 180, "shorter cell targets the shared 180");
}
{
  // fixed-% scales the whole vector so the total = pct × pagePx, proportions from the matrix.
  const bandhs = [{ GRID: 2, cells: [
    { key: "0:0:1", natural: 100, col: 0, span: 1 },
    { key: "0:1:1", natural: 300, col: 1, span: 1 },
  ]}];
  const g = AshaarMatrix.computeTargetGrid(bandhs, { mode: "fixed", pct: 50, pagePx: 1000 });
  // total target = 500; proportions 100:300 → 125 and 375
  assert.strictEqual(Math.round(g.colPx[0]), 125);
  assert.strictEqual(Math.round(g.colPx[1]), 375);
}
{
  // auto-fit total capped at pagePx.
  const bandhs = [{ GRID: 2, cells: [
    { key: "0:0:1", natural: 800, col: 0, span: 1 },
    { key: "0:1:1", natural: 800, col: 1, span: 1 },
  ]}];
  const g = AshaarMatrix.computeTargetGrid(bandhs, { mode: "auto-fit", pagePx: 1000, headroom: 0.2 });
  const total = g.colPx.reduce((a, b) => a + b, 0);
  assert.ok(Math.abs(total - 1000) < 1, "capped at page width, got " + total);
}
console.log("computeTargetGrid OK");

{
  // Different GRID counts → not same shape; each bandh keeps its own layout but
  // shares one total width (max of the per-bandh needs, capped at page).
  const bandhs = [
    { GRID: 2, cells: [ { key: "0:0:1", natural: 100, col: 0, span: 1 }, { key: "0:1:1", natural: 100, col: 1, span: 1 } ] },
    { GRID: 3, cells: [ { key: "0:0:1", natural: 90, col: 0, span: 1 }, { key: "0:1:2", natural: 210, col: 1, span: 2 } ] },
  ];
  const g = AshaarMatrix.computeTargetGrid(bandhs, { mode: "auto-fit", pagePx: 1000, headroom: 0 });
  assert.strictEqual(g.sameShape, false);
  assert.strictEqual(g.colPx, null, "no shared column vector when shapes differ");
  // per-bandh vectors present, one per bandh, with each bandh's own grid length
  assert.strictEqual(g.perBandhColPx.length, 2);
  assert.strictEqual(g.perBandhColPx[0].length, 2);
  assert.strictEqual(g.perBandhColPx[1].length, 3);
  // each bandh's targets present, summing to the same total
  const t0 = Object.values(g.bandhTargets[0]).reduce((a, b) => a + b, 0);
  const t1 = Object.values(g.bandhTargets[1]).reduce((a, b) => a + b, 0);
  assert.ok(Math.abs(t0 - t1) < 1, "different-shape bandhs share one total width");
}
console.log("computeTargetGrid different-shape OK");

// ── uniformSlotPx: the single grid-slot size (Option A) ──────────────────────
{
  // One bandh, GRID 6. A span-3 text cell of natural 180 needs 60/slot.
  const bandhs = [{ GRID: 6, cells: [{ natural: 180, span: 3 }, { natural: 120, span: 3 }] }];
  assert.strictEqual(AshaarMatrix.uniformSlotPx(bandhs, { mode: "auto-fit", pagePx: 1000, headroom: 0 }), 60);
}
{
  // headroom scales the raw need: (180 × 1.2) / 3 = 72.
  const bandhs = [{ GRID: 6, cells: [{ natural: 180, span: 3 }] }];
  assert.strictEqual(AshaarMatrix.uniformSlotPx(bandhs, { mode: "auto-fit", pagePx: 1000, headroom: 0.2 }), 72);
}
{
  // auto-fit caps so the widest bandh fits the page: slot ≤ pagePx / GRID.
  const bandhs = [{ GRID: 2, cells: [{ natural: 800, span: 1 }] }];
  assert.strictEqual(AshaarMatrix.uniformSlotPx(bandhs, { mode: "auto-fit", pagePx: 1000, headroom: 0 }), 500);
}
{
  // fixed %: slot = (pct/100 × pagePx) / GRID, regardless of content.
  const bandhs = [{ GRID: 2, cells: [{ natural: 100, span: 1 }] }];
  assert.strictEqual(AshaarMatrix.uniformSlotPx(bandhs, { mode: "fixed", pct: 50, pagePx: 1000 }), 250);
}
{
  // different-shape bandhs share one slot; cap uses the largest GRID.
  const bandhs = [
    { GRID: 2, cells: [{ natural: 300, span: 1 }] },
    { GRID: 4, cells: [{ natural: 300, span: 1 }] },
  ];
  // rawSlot=300; cap = pagePx/maxGRID = 1000/4 = 250 → slot 250.
  assert.strictEqual(AshaarMatrix.uniformSlotPx(bandhs, { mode: "auto-fit", pagePx: 1000, headroom: 0 }), 250);
}
console.log("uniformSlotPx OK");

console.log("natural-width-matrix.test.js OK");
