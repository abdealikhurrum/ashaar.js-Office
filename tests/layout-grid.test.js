const assert = require("assert");
const { gridToTemplate, gridToSpec, specToGrid } = require("../src/taskpane/layout-grid");

// Matrix rows are in READING order: index 0 = first misra = visual right (sadr).
// '#' = on (bubble filled), '.' = off (gap). Each row is 12 columns.
const R = (s) => s.split("").map((c) => c === "#");
const spans = (cells) => cells.map((c) => c.span);
const aligns = (cells) => cells.map((c) => c.align);
const roles = (cells) => cells.map((c) => c.role);

// ── gridToTemplate (the Draw Table path — must be exact) ────────────────────

// Couplet: misra1 (5) | gap (2) | misra2 (5)
{
  const t = gridToTemplate([R("#####..#####")]);
  assert.equal(t.columnCount, 12, "12 columns");
  assert.equal(t.rows.length, 1);
  assert.deepEqual(spans(t.rows[0]), [5, 2, 5], "couplet spans");
  assert.deepEqual(roles(t.rows[0]), ["misra", "gap", "misra"], "couplet roles");
  assert.deepEqual(aligns(t.rows[0]), ["right", "center", "left"], "first content right, last left");
}

// Centered solo
{
  const t = gridToTemplate([R("...######...")]);
  assert.deepEqual(spans(t.rows[0]), [3, 6, 3]);
  assert.deepEqual(roles(t.rows[0]), ["gap", "misra", "gap"]);
  assert.equal(t.rows[0][1].align, "center", "lone content cell is centred");
}

// Full width
{
  const t = gridToTemplate([R("############")]);
  assert.deepEqual(spans(t.rows[0]), [12]);
  assert.deepEqual(roles(t.rows[0]), ["misra"]);
  assert.equal(t.rows[0][0].align, "center");
}

// Multi-misra (3 content runs) → right / center / left
{
  const t = gridToTemplate([R("###.###.###.")]); // 3+1+3+1+3+1 = 12
  const content = t.rows[0].filter((c) => c.role === "misra");
  assert.equal(content.length, 3, "three misra cells");
  assert.deepEqual(content.map((c) => c.align), ["right", "center", "left"]);
  assert.equal(spans(t.rows[0]).reduce((a, b) => a + b, 0), 12, "spans sum to 12");
}

// All-off row is skipped
{
  const t = gridToTemplate([R("............"), R("############")]);
  assert.equal(t.rows.length, 1, "blank row skipped");
}

// Indented / right-anchored: content then trailing gap
{
  const t = gridToTemplate([R("########....")]);
  assert.deepEqual(spans(t.rows[0]), [8, 4]);
  assert.deepEqual(roles(t.rows[0]), ["misra", "gap"]);
}

// Every row's spans sum to 12
{
  const t = gridToTemplate([R("#####..#####"), R("...######..."), R("############")]);
  t.rows.forEach((row) => assert.equal(spans(row).reduce((a, b) => a + b, 0), 12));
}

// ── gridToSpec (Numbers view — sequential numbering, reading order) ──────────

assert.equal(gridToSpec([R("#####..#####")]), "1 - 2", "couplet → pair spec");
assert.equal(gridToSpec([R("#####..#####"), R("...######...")]), "1 - 2\n<3>",
  "numbering continues across rows");
assert.equal(gridToSpec([R("############")]), "<1>", "full width → centred");
assert.equal(gridToSpec([R("###.###.###.")]), "1 | 2 | 3", "multi → pipe spec");

// ── specToGrid + round-trip ─────────────────────────────────────────────────

// Each produced row sums to 12 and has the right number of content runs.
function contentRuns(rowBools) {
  let runs = 0, prev = false;
  rowBools.forEach((b) => { if (b && !prev) runs++; prev = b; });
  return runs;
}

{
  const m = specToGrid("1 - 2");
  assert.equal(m.length, 1);
  assert.equal(m[0].length, 12, "12 columns");
  assert.equal(contentRuns(m[0]), 2, "pair → two content runs");
}
{
  const m = specToGrid("<1>");
  assert.equal(contentRuns(m[0]), 1, "centred → one content run");
}
{
  const m = specToGrid("1 | 2 | 3");
  assert.equal(contentRuns(m[0]), 3, "multi → three content runs");
}

// Round-trip: spec → grid → spec is stable for supported constructs
["1 - 2", "<1>", "1 | 2 | 3"].forEach((spec) => {
  assert.equal(gridToSpec(specToGrid(spec)), spec, "round-trips: " + spec);
});

console.log("layout-grid tests passed");
