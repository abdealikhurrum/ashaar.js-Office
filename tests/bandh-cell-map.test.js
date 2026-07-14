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

console.log("bandh-cell-map.test.js OK");
