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

// ── keysForTarget: this ⊂ bandh ⊂ poem, content/spacing never mix ───────────
// Two bandhs (tables). Table 0: solo row [c,g,c] → A1, (gap A#1), A2.
// Table 1: couplet row [c,g,c] then solo-flanked row [g,c,g] →
//   content A1,A2,B1 ; spacing A#1,B#1,B#2 (emission order).
{
  const tables = [
    [["c", "g", "c"]],
    [["c", "g", "c"], ["g", "c", "g"]],
  ];
  const table1Map = AshaarCellMap.buildBandhCellMap(tables[1]);
  const AshaarOverrides = require("../src/taskpane/cell-overrides");

  // Current cell = table 1's content "A1".
  const curContentKey = AshaarOverrides.overrideKey(1, "A1");
  assert.deepStrictEqual(
    AshaarCellMap.keysForTarget(table1Map, "content", "this", curContentKey, tables),
    [curContentKey],
    "this → exactly the current key"
  );
  const bandhContent = AshaarCellMap.keysForTarget(table1Map, "content", "bandh", curContentKey, tables);
  assert.deepStrictEqual(
    bandhContent,
    ["1:A1", "1:A2", "1:B1"],
    "bandh → all content keys in table 1 only"
  );
  const poemContent = AshaarCellMap.keysForTarget(table1Map, "content", "poem", curContentKey, tables);
  assert.deepStrictEqual(
    poemContent,
    ["0:A1", "0:A2", "1:A1", "1:A2", "1:B1"],
    "poem → all tables' content keys"
  );

  // Current cell = table 1's spacing "A#1".
  const curSpacingKey = AshaarOverrides.overrideKey(1, "A#1");
  assert.deepStrictEqual(
    AshaarCellMap.keysForTarget(table1Map, "spacing", "this", curSpacingKey, tables),
    [curSpacingKey],
    "this → exactly the current key (spacing)"
  );
  const bandhSpacing = AshaarCellMap.keysForTarget(table1Map, "spacing", "bandh", curSpacingKey, tables);
  assert.deepStrictEqual(
    bandhSpacing,
    ["1:A#1", "1:B#1", "1:B#2"],
    "bandh → all spacing keys in table 1 only"
  );
  const poemSpacing = AshaarCellMap.keysForTarget(table1Map, "spacing", "poem", curSpacingKey, tables);
  assert.deepStrictEqual(
    poemSpacing,
    ["0:A#1", "1:A#1", "1:B#1", "1:B#2"],
    "poem → all tables' spacing keys"
  );

  // Containment: this ⊂ bandh ⊂ poem, for both kinds.
  assert.ok(bandhContent.every((k) => poemContent.indexOf(k) !== -1), "bandh content ⊂ poem content");
  assert.ok(bandhSpacing.every((k) => poemSpacing.indexOf(k) !== -1), "bandh spacing ⊂ poem spacing");
  assert.ok([curContentKey].every((k) => bandhContent.indexOf(k) !== -1), "this ⊂ bandh (content)");
  assert.ok([curSpacingKey].every((k) => bandhSpacing.indexOf(k) !== -1), "this ⊂ bandh (spacing)");

  // Content and spacing enumerations never mix.
  bandhContent.concat(poemContent).forEach((k) => assert.ok(k.indexOf("#") === -1, "content key has no '#': " + k));
  bandhSpacing.concat(poemSpacing).forEach((k) => assert.ok(k.indexOf("#") !== -1, "spacing key has '#': " + k));

  // ── mode "position": same-label cell across EVERY table, not just the
  // current one — "Same cell in all bandhs". Table 0's content is A1,A2;
  // table 1's is A1,A2,B1 — position on table 1's A1 must hit both tables'
  // A1 (and only A1: never A2/B1, and never table 1's own bandh-mates).
  const positionContent = AshaarCellMap.keysForTarget(table1Map, "content", "position", curContentKey, tables);
  assert.deepStrictEqual(
    positionContent,
    ["0:A1", "1:A1"],
    "position → same label (A1) across all tables, no other labels"
  );
  const positionSpacing = AshaarCellMap.keysForTarget(table1Map, "spacing", "position", curSpacingKey, tables);
  assert.deepStrictEqual(
    positionSpacing,
    ["0:A#1", "1:A#1"],
    "position → same slot (A#1) across all tables"
  );
  // Content and spacing still never mix under "position".
  positionContent.forEach((k) => assert.ok(k.indexOf("#") === -1, "position content key has no '#': " + k));
  positionSpacing.forEach((k) => assert.ok(k.indexOf("#") !== -1, "position spacing key has '#': " + k));
  // Containment: position ⊆ poem (same kind), for both kinds.
  assert.ok(positionContent.every((k) => poemContent.indexOf(k) !== -1), "position content ⊆ poem content");
  assert.ok(positionSpacing.every((k) => poemSpacing.indexOf(k) !== -1), "position spacing ⊆ poem spacing");
  // A label that exists in only one table (B1) yields just that table's key.
  const curB1Key = AshaarOverrides.overrideKey(1, "B1");
  assert.deepStrictEqual(
    AshaarCellMap.keysForTarget(table1Map, "content", "position", curB1Key, tables),
    ["1:B1"],
    "position on a table-unique label returns only that table's key"
  );
}

// ── columnGroupKey: harmony-pooling fix. Mapped cells pool across ROWS by
// structural slot (row token-shape + ordinal), never by row letter — a
// couplet's rows stacked in one table (same shape) pool; marsiya rows of
// different content-count (3/1/2) never cross-pool despite sharing an
// ordinal number. ──────────────────────────────────────────────────────────
{
  // Three same-shape couplet rows [c,c] stacked in one table.
  const pattern = [["c", "c"], ["c", "c"], ["c", "c"]];
  const map = AshaarCellMap.buildBandhCellMap(pattern);
  const a1 = map.find((e) => e.label === "A1");
  const b1 = map.find((e) => e.label === "B1");
  const c1 = map.find((e) => e.label === "C1");
  const a2 = map.find((e) => e.label === "A2");
  assert.strictEqual(
    AshaarCellMap.columnGroupKey(pattern, a1),
    AshaarCellMap.columnGroupKey(pattern, b1),
    "same row-shape, same ordinal (misra 1) → pools across rows"
  );
  assert.strictEqual(
    AshaarCellMap.columnGroupKey(pattern, a1),
    AshaarCellMap.columnGroupKey(pattern, c1)
  );
  assert.notStrictEqual(
    AshaarCellMap.columnGroupKey(pattern, a1),
    AshaarCellMap.columnGroupKey(pattern, a2),
    "different ordinal (misra 1 vs misra 2) never pools"
  );
}
{
  // Marsiya: rows of DIFFERENT shape — 3 content cells, then a solo (1), then
  // a flanked pair (2). Ordinal "1" appears in every row but must not
  // cross-pool cells whose widths genuinely differ (different layouts).
  const pattern = [["c", "c", "c"], ["g", "c", "g"], ["c", "g", "c"]];
  const map = AshaarCellMap.buildBandhCellMap(pattern);
  const rowAFirst = map.find((e) => e.label === "A1");
  const rowBFirst = map.find((e) => e.label === "B1");
  const rowCFirst = map.find((e) => e.label === "C1");
  assert.notStrictEqual(
    AshaarCellMap.columnGroupKey(pattern, rowAFirst),
    AshaarCellMap.columnGroupKey(pattern, rowBFirst),
    "different row shape → never pools even at the same ordinal"
  );
  assert.notStrictEqual(
    AshaarCellMap.columnGroupKey(pattern, rowAFirst),
    AshaarCellMap.columnGroupKey(pattern, rowCFirst)
  );
}
{
  // Content and spacing never collide even at the same ordinal number.
  const pattern = [["c", "g"]];
  const map = AshaarCellMap.buildBandhCellMap(pattern);
  const content1 = map.find((e) => e.kind === "content");
  const spacing1 = map.find((e) => e.kind === "spacing");
  assert.notStrictEqual(
    AshaarCellMap.columnGroupKey(pattern, content1),
    AshaarCellMap.columnGroupKey(pattern, spacing1),
    "content and spacing ordinals never collide"
  );
}
{
  // Rows past 26: rowLetter falls back to "R27", "R28"… — the pooling key
  // must NOT round-trip through the label string (a leading-letters strip
  // would leak "27" into the ordinal: "R27" + "1" → ordinal "271"). A long
  // qasida with 28 baits as rows of ONE table must pool row 0 with row 27.
  const row = ["c", "g", "c"];
  const pattern = [];
  for (let r = 0; r < 28; r++) pattern.push(row.slice());
  const map = AshaarCellMap.buildBandhCellMap(pattern);
  const r0 = map.filter((e) => e.row === 0);
  const r27 = map.filter((e) => e.row === 27);
  // Sanity: row 27 labels/slots use the fallback letter scheme.
  assert.strictEqual(r27[0].label, "R281", "row 27 content label uses R-fallback");
  assert.strictEqual(r27[1].slot, "R28#1", "row 27 spacing slot uses R-fallback");
  // Content: first misra of row 0 pools with first misra of row 27…
  assert.strictEqual(
    AshaarCellMap.columnGroupKey(pattern, r0[0]),
    AshaarCellMap.columnGroupKey(pattern, r27[0]),
    "row 0 and row 27 first-content cells share a key"
  );
  // …and second with second, but never first with second.
  assert.strictEqual(
    AshaarCellMap.columnGroupKey(pattern, r0[2]),
    AshaarCellMap.columnGroupKey(pattern, r27[2]),
    "row 0 and row 27 second-content cells share a key"
  );
  assert.notStrictEqual(
    AshaarCellMap.columnGroupKey(pattern, r0[0]),
    AshaarCellMap.columnGroupKey(pattern, r27[2]),
    "first and second misra never pool, even across the letter boundary"
  );
  // Spacing: row 0's gap pools with row 27's gap.
  assert.strictEqual(
    AshaarCellMap.columnGroupKey(pattern, r0[1]),
    AshaarCellMap.columnGroupKey(pattern, r27[1]),
    "row 0 and row 27 spacing slots share a key"
  );
  // Content/spacing still never collide past the letter boundary.
  assert.notStrictEqual(
    AshaarCellMap.columnGroupKey(pattern, r27[0]),
    AshaarCellMap.columnGroupKey(pattern, r27[1]),
    "content and spacing never collide in R-fallback rows"
  );
}

console.log("bandh-cell-map.test.js OK");
