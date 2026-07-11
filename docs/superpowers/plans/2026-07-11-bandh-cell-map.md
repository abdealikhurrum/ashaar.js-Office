# Bandh Cell-Map (SP1) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Give every grid cell in an Ashaar poem an authoritative content/spacing tag and a positional label within its bandh (`A1, A2, B1, C1, …`), derived and persisted when the table is created, so justification and creation stop guessing structure from emptiness/geometry.

**Architecture:** A new pure module `bandh-cell-map.js` (`AshaarCellMap`) labels a per-row content/gap **pattern** (`[["c","g","c"], …]`) into an ordered cell-map and validates a live table's shape against a stored pattern. `word-html.js` derives that pattern from the existing deterministic poem row model (`stanzaCellPattern`/`poemCellPatterns`) and persists it in the content-control tag (`cells`, payload `v2`). `taskpane.js` passes the pattern to the tag at insert, and `justifySelection`/`applyProfileToQaseeda` read it back: label becomes the natural-width-matrix key and the tag decides content-vs-spacing (empty content cells are no longer mistaken for gaps). A read-only pane view lists the block-at-cursor's map. Maps-absent tables fall back to today's geometric behavior.

**Tech Stack:** Vanilla JS (ES5/UMD, `var`/`function`, no build step), Office.js v1, Node `assert` tests. Never edit `src/vendor/`.

**Spec:** `docs/superpowers/specs/2026-07-11-bandh-cell-map-design.md`
**Roadmap:** `docs/superpowers/specs/2026-07-11-cell-configurations-roadmap.md` (this is SP1; SP2 = per-cell overrides + editing UI, later).

## Global Constraints

- ES5/UMD only (`var`, `function`); never edit `src/vendor/`.
- Pure logic (`AshaarCellMap`, `stanzaCellPattern`/`poemCellPatterns`, tag round-trip) is **node-tested**; the Office.js glue in `taskpane.js` (justify rewire, pane view, insert wiring) is **manual-verify** (final task).
- Pattern token vocabulary is exactly `"c"` (content) and `"g"` (spacing/gap). A pattern is `Array<Array<"c"|"g">>`: outer = rows, inner = cells in OOXML emission order.
- Label scheme: rows → letters `A,B,C,…` top-to-bottom; content cells numbered `1,2,3,…` in emission order (which is misra reading order = **rightmost-first in RTL**, since the generator emits sadr first). Spacing cells get no number — a `slot` key `"<letter>#<n>"` instead.
- `BASE_CPM = 6`; for the poem path a solo/centered row always emits `[gap, content, gap]` and a K-misra row emits `[c, g, c, …, c]` (K contents, K−1 gaps, no outer pads). Verified against `BASE_CPM` and confirmed by the Task 3 cross-check test.
- Payload `v` bumps `1 → 2`; a `v1` tag (or any tag without `cells`) parses to `cells: null`, which triggers the geometric fallback — behavior for those tables is unchanged.
- `npm test` green after every task.

---

### Task 1: `AshaarCellMap` — label a pattern + validate table shape (pure)

**Files:**
- Create: `src/taskpane/bandh-cell-map.js`
- Test: `tests/bandh-cell-map.test.js`
- Modify: `package.json` (test script), `src/taskpane/taskpane.html` (script load list)

**Interfaces:**
- Produces `AshaarCellMap.buildBandhCellMap(pattern) → Array<{index, row, kind:"content"|"spacing", label:string|null, slot:string|null}>` — flat, in emission order.
- Produces `AshaarCellMap.alignPatternToTable(perRowCounts, pattern) → boolean` — true iff `pattern` has the same number of rows as `perRowCounts` and each row's cell count matches.

- [ ] **Step 1: Write the failing test**

Create `tests/bandh-cell-map.test.js`:

```js
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
```

- [ ] **Step 2: Run to verify it fails**

Run: `node tests/bandh-cell-map.test.js`
Expected: FAIL — `Cannot find module '../src/taskpane/bandh-cell-map'`.

- [ ] **Step 3: Write the module**

Create `src/taskpane/bandh-cell-map.js`:

```js
/**
 * AshaarCellMap — label a bandh's content/gap PATTERN into positional cell ids
 * and validate a live Word table's shape against a stored pattern (pure, UMD;
 * node-testable, no DOM/Office dependency).
 *
 * A pattern is Array<Array<"c"|"g">>: outer = rows top-to-bottom, inner = cells
 * in OOXML emission order. Rows map to letters A,B,C…; content cells ("c") are
 * numbered 1,2,3… in emission order (= misra reading order = rightmost-first in
 * RTL), giving labels A1,A2,B1,C1,C2. Gap cells ("g") get no number — a `slot`
 * key "<letter>#<n>" instead (the future anchor for hemistich symbols / bandh
 * numbers / annotations).
 *
 * See docs/superpowers/specs/2026-07-11-bandh-cell-map-design.md.
 */
(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) module.exports = factory();
  else root.AshaarCellMap = factory();
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  var LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  function rowLetter(r) { return LETTERS.charAt(r) || ("R" + (r + 1)); }

  // pattern → flat ordered cell-map (emission order preserved).
  function buildBandhCellMap(pattern) {
    var out = [];
    var idx = 0;
    (pattern || []).forEach(function (row, r) {
      var letter = rowLetter(r);
      var contentN = 0, spacingN = 0;
      (row || []).forEach(function (tok) {
        if (tok === "c") {
          contentN++;
          out.push({ index: idx++, row: r, kind: "content", label: letter + contentN, slot: null });
        } else {
          spacingN++;
          out.push({ index: idx++, row: r, kind: "spacing", label: null, slot: letter + "#" + spacingN });
        }
      });
    });
    return out;
  }

  // Does a stored pattern match the live table's shape? perRowCounts is the
  // number of cells in each live row, in order. A mismatch (hand-edited/adopted
  // table) tells the caller to fall back to geometric inference.
  function alignPatternToTable(perRowCounts, pattern) {
    if (!pattern || !perRowCounts) return false;
    if (pattern.length !== perRowCounts.length) return false;
    for (var i = 0; i < pattern.length; i++) {
      if (!pattern[i] || pattern[i].length !== perRowCounts[i]) return false;
    }
    return true;
  }

  return {
    buildBandhCellMap: buildBandhCellMap,
    alignPatternToTable: alignPatternToTable,
  };
}));
```

- [ ] **Step 4: Run to verify it passes**

Run: `node tests/bandh-cell-map.test.js`
Expected: PASS — `bandh-cell-map.test.js OK`.

- [ ] **Step 5: Register the test**

In `package.json`, append to the `test` script (after `natural-width-matrix.test.js`):
```
... && node tests/natural-width-matrix.test.js && node tests/bandh-cell-map.test.js"
```

- [ ] **Step 6: Register the module for the browser**

In `src/taskpane/taskpane.html`, add to the `srcs` array right after `"./natural-width-matrix.js",`:
```js
          "./natural-width-matrix.js",
          "./bandh-cell-map.js",
          "./fonts.js",
```

- [ ] **Step 7: Full suite**

Run: `npm test`
Expected: green.

- [ ] **Step 8: Commit**

```bash
git add src/taskpane/bandh-cell-map.js tests/bandh-cell-map.test.js package.json src/taskpane/taskpane.html
git commit --no-gpg-sign -m "feat(cellmap): AshaarCellMap — label pattern + validate table shape"
```

---

### Task 2: Derive the pattern from the poem row model (pure)

**Files:**
- Modify: `src/taskpane/word-html.js` (add `baytCellPatternRows`, `stanzaCellPattern`, `poemCellPatterns` + exports)
- Test: `tests/word-html.test.js`

**Interfaces:**
- Produces `AshaarWord.stanzaCellPattern(stanza, opts) → Array<Array<"c"|"g">>` — one bandh's pattern.
- Produces `AshaarWord.poemCellPatterns(text, opts, Ashaar) → Array<pattern>` — one pattern per stanza/table, in document order (the order `renderForWordOoxml` emits tables).

- [ ] **Step 1: Write the failing test**

Append to `tests/word-html.test.js`:

```js
// ── cell patterns: mirror the OOXML generator's content/gap cell order ───────
{
  const Ashaar2 = require("../src/vendor/ashaar");
  // A simple couplet (sadr \ ajuz) → one table, one row [c,g,c].
  const pats = AshaarWord.poemCellPatterns("دل ناداں \\ آخر اس درد", { layoutMode: "balanced" }, Ashaar2);
  assert.equal(pats.length, 1, "one stanza → one pattern");
  assert.deepStrictEqual(pats[0], [["c", "g", "c"]], "couplet row = content,gap,content");
}
{
  const Ashaar2 = require("../src/vendor/ashaar");
  // Solo single misra (|) → one table, one row [g,c,g].
  const pats = AshaarWord.poemCellPatterns("تنہا مصرعہ |", { layoutMode: "balanced" }, Ashaar2);
  assert.deepStrictEqual(pats[0], [["g", "c", "g"]], "solo row = gap,content,gap");
}
{
  const Ashaar2 = require("../src/vendor/ashaar");
  // Stacked couplet → two solo rows.
  const pats = AshaarWord.poemCellPatterns("دل ناداں \\ آخر اس درد", { layoutMode: "stacked" }, Ashaar2);
  assert.deepStrictEqual(pats[0], [["g", "c", "g"], ["g", "c", "g"]], "stacked = two solo rows");
}
```

- [ ] **Step 2: Run to verify it fails**

Run: `node tests/word-html.test.js`
Expected: FAIL — `AshaarWord.poemCellPatterns is not a function`.

- [ ] **Step 3: Implement**

In `src/taskpane/word-html.js`, add these functions just before `stanzaTableOoxml` (line ~1341). They mirror the branches of `baytRowsOoxml` at the KIND level only (spans/indents don't affect cell order):

```js
  // The content/gap KIND sequence a bayt contributes, per row — mirrors
  // baytRowsOoxml's branches (solo → [g,c,g]; K-misra row → [c,g,c,…,c]; stacked
  // → one solo row per misra). Spans/alignment/indents don't change cell order,
  // so only the kind matters here. This is the single derivation the persisted
  // pattern and (via AshaarCellMap) all labels come from; the Task-3-style
  // cross-check test locks it to the actual OOXML the generator emits.
  function baytCellPatternRows(bayt, opts) {
    var stacked = (opts || {}).layoutMode === "stacked";
    function solo() { return ["g", "c", "g"]; }
    function misra(K) { var r = []; for (var i = 0; i < K; i++) { r.push("c"); if (i < K - 1) r.push("g"); } return r; }
    if (bayt.type === "row") {
      var K = (bayt.misras || []).length;
      if (K === 0) return [];
      if (K === 1) return [solo()];
      if (stacked) { var rows = []; for (var i = 0; i < K; i++) rows.push(solo()); return rows; }
      return [misra(K)];
    }
    if (!bayt.ajuz) return [solo()];
    if (stacked) return [solo(), solo()];
    return [misra(2)];
  }

  function stanzaCellPattern(stanza, opts) {
    var rows = [];
    (stanza.bayts || []).forEach(function (b) {
      baytCellPatternRows(b, opts).forEach(function (r) { rows.push(r); });
    });
    return rows;
  }

  function poemCellPatterns(text, opts, Ashaar) {
    var poems = parsePoetry(String(text || ""), Ashaar);
    var pats = [];
    poems.forEach(function (poem) {
      (poem.stanzas || []).forEach(function (stanza) { pats.push(stanzaCellPattern(stanza, opts)); });
    });
    return pats;
  }
```

Add to the module exports (near `renderForWordOoxml: renderForWordOoxml,`):
```js
    stanzaCellPattern: stanzaCellPattern,
    poemCellPatterns: poemCellPatterns,
```

- [ ] **Step 4: Run to verify it passes**

Run: `node tests/word-html.test.js`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/taskpane/word-html.js tests/word-html.test.js
git commit --no-gpg-sign -m "feat(cellmap): derive bandh cell pattern from the poem row model"
```

---

### Task 3: Cross-check the pattern against the emitted OOXML (pure)

Locks `stanzaCellPattern` to what `renderForWordOoxml` actually emits, so the two can't drift: for each generated table, the number of rows and per-row cells must equal the pattern, and each cell's kind must match (a content `<w:tc>` contains a `<w:r>` run; a gap `<w:tc>` is an empty `<w:p/>`).

**Files:**
- Test: `tests/word-html.test.js`

- [ ] **Step 1: Write the cross-check test**

Append to `tests/word-html.test.js`:

```js
// ── cross-check: pattern shape/kind == the generator's actual <w:tc> cells ───
{
  const Ashaar2 = require("../src/vendor/ashaar");
  // Split a rendered OOXML body into its tables, each table into rows, each row
  // into cells; classify a cell as content when it carries a run, else gap.
  function tablesOf(xml) { return xml.match(/<w:tbl>[\s\S]*?<\/w:tbl>/g) || []; }
  function rowsOf(tblXml) { return tblXml.match(/<w:tr>[\s\S]*?<\/w:tr>/g) || []; }
  function cellsOf(trXml) { return trXml.match(/<w:tc>[\s\S]*?<\/w:tc>/g) || []; }
  function kindOf(tcXml) { return /<w:r[ >]/.test(tcXml) ? "c" : "g"; }

  const cases = [
    { src: "دل ناداں \\ آخر اس درد", opts: { layoutMode: "balanced" } },
    { src: "تنہا مصرعہ |", opts: { layoutMode: "balanced" } },
    { src: "دل ناداں \\ آخر اس درد", opts: { layoutMode: "stacked" } },
    { src: "الف \\ ب\n\nج \\ د", opts: { layoutMode: "balanced" } }, // two stanzas
  ];
  cases.forEach(function (c, ci) {
    const body = AshaarWord.renderForWordOoxml(c.src, c.opts, Ashaar2, 9360);
    const pats = AshaarWord.poemCellPatterns(c.src, c.opts, Ashaar2);
    const tbls = tablesOf(body);
    assert.equal(tbls.length, pats.length, "case " + ci + ": table count == pattern count");
    tbls.forEach(function (tbl, ti) {
      const rows = rowsOf(tbl);
      assert.equal(rows.length, pats[ti].length, "case " + ci + " tbl " + ti + ": row count");
      rows.forEach(function (tr, ri) {
        const kinds = cellsOf(tr).map(kindOf);
        assert.deepStrictEqual(kinds, pats[ti][ri], "case " + ci + " tbl " + ti + " row " + ri + ": cell kinds");
      });
    });
  });
}
```

- [ ] **Step 2: Run**

Run: `node tests/word-html.test.js`
Expected: PASS. If a case FAILS, `stanzaCellPattern` (Task 2) does not match the generator for that shape — fix the pattern derivation (not the test) until kinds align, then re-run.

- [ ] **Step 3: Full suite + commit**

Run: `npm test`; Expected: green.
```bash
git add tests/word-html.test.js
git commit --no-gpg-sign -m "test(cellmap): cross-check pattern vs emitted OOXML cells"
```

---

### Task 4: Persist the pattern in the content-control tag

**Files:**
- Modify: `src/taskpane/word-html.js` (`contentControlTag`, `parseContentControlTag`)
- Modify: `src/taskpane/taskpane.js` (the two poem-insert tag sites)
- Test: `tests/word-html.test.js`

**Interfaces:**
- `AshaarWord.contentControlTag(text, opts, cellPatterns?)` — optional 3rd arg `cellPatterns` (from `poemCellPatterns`); when present, stored as `cells`; payload `v` is `2`.
- `AshaarWord.parseContentControlTag(tag).cells` — the stored patterns, or `null` when absent (v1 tags, grid/template tags).

- [ ] **Step 1: Write the failing test**

Append to `tests/word-html.test.js`:

```js
// ── tag round-trips the cells pattern; absent → null ─────────────────────────
{
  const pat = [[["c", "g", "c"]], [["g", "c", "g"]]]; // two stanzas
  const tag = AshaarWord.contentControlTag("poem", { tableWidthPct: 50 }, pat);
  const parsed = AshaarWord.parseContentControlTag(tag);
  assert.deepStrictEqual(parsed.cells, pat, "cells round-trip");
  assert.equal(parsed.v, 2, "payload version bumped to 2");
}
{
  // No 3rd arg → no cells (grid/template paths); parses to null.
  const tag = AshaarWord.contentControlTag("grid", { tableWidthPct: 50 });
  assert.strictEqual(AshaarWord.parseContentControlTag(tag).cells, null, "absent cells → null");
}
```

- [ ] **Step 2: Run to verify it fails**

Run: `node tests/word-html.test.js`
Expected: FAIL — `parsed.cells` is `undefined` / `parsed.v` is `1`.

- [ ] **Step 3: Implement in `word-html.js`**

In `contentControlTag`, change the signature and payload. Replace the `function contentControlTag(text, opts) {` line with:
```js
  function contentControlTag(text, opts, cellPatterns) {
```
Change `v: 1,` to `v: 2,` in the payload. Then, immediately before `return "ashaar:" + …`, add:
```js
    if (cellPatterns && cellPatterns.length) payload.cells = cellPatterns;
```

In `parseContentControlTag`, ensure `cells` surfaces as `null` when absent. The function parses into `var payload` and returns it; add the normalization right before `return payload;`:
```js
      if (typeof payload.qaseeda !== "string") payload.qaseeda = "";
      payload.cells = payload.cells || null;
      return payload;
```

- [ ] **Step 4: Run to verify it passes**

Run: `node tests/word-html.test.js`
Expected: PASS.

- [ ] **Step 5: Wire the insert path to compute + pass the pattern**

In `src/taskpane/taskpane.js`, the **table** poem-insert sets its tag in `insertPoem` (line ~998, the `renderForWordOoxml` path). Replace (anchored by its unique `replaceSelection` context):
```js
      control.tag = AshaarWord.contentControlTag(source, opts);
```
with:
```js
      control.tag = AshaarWord.contentControlTag(source, opts, AshaarWord.poemCellPatterns(source, opts, Ashaar));
```
Leave the OTHER `contentControlTag(source, opts)` call (line ~1121) UNCHANGED — that is the **tab-stop/paragraph** path (`AshaarTabStop.poemToOoxml`), which produces no tables, so a cell-map doesn't apply (it correctly falls back to no-map). Also leave the grid/template tag sites (`"grid"`, `"template"`, `"grid12"`, `"template:"+…`) unchanged — they pass no 3rd arg and get the fallback.

- [ ] **Step 6: Syntax + suite**

Run: `node --check src/taskpane/taskpane.js`; Expected: no output.
Run: `npm test`; Expected: green (existing `contentControlTag(src, opts)` callers in poetry-corpus/word-html tests still pass — the 3rd arg is optional; `v:2` isn't asserted anywhere).

- [ ] **Step 7: Commit**

```bash
git add src/taskpane/word-html.js src/taskpane/taskpane.js tests/word-html.test.js
git commit --no-gpg-sign -m "feat(cellmap): persist cell pattern in the content-control tag (v2)"
```

---

### Task 5: Justification consumes the map

**Files:**
- Modify: `src/taskpane/taskpane.js` (`justifySelection`: read the cc tag; build a per-table map; key `__matKey` on label + skip spacing. `applyProfileToQaseeda`: same, per block.)

**Interfaces:**
- Consumes `AshaarWord.parseContentControlTag`, `AshaarCellMap.buildBandhCellMap`, `AshaarCellMap.alignPatternToTable`. No node test (Office.js); verified in Task 7. The label/align logic it calls is already node-tested (Tasks 1–3).

**Design:** When the block's tag carries `cells`, and a table's live shape matches its stanza pattern, each cell gets `__matKey = label` (content) / is marked spacing (skipped); otherwise fall back to today's geometric `AshaarMatrix.positionKey` + `AshaarMatrix.isContentCell`. `cells[t]` corresponds to the t-th table of the block in document order.

- [ ] **Step 1: `justifySelection` — read the tag's cells alongside the title**

In `justifySelection`, the enclosing content control is loaded at line ~1363 (`cc.load("title")`). Change it to also load the tag:
```js
      cc.load("title,tag");
      await context.sync();

      var ccCells = null;
      if (!cc.isNullObject && cc.title === "Ashaar Poem") {
        var ccPayload = AshaarWord.parseContentControlTag(cc.tag);
        ccCells = ccPayload && ccPayload.cells;
      }
```

- [ ] **Step 2: `justifySelection` — assign label/kind per table with geometric fallback**

Replace the `allCells` build loop (the `tables.items.forEach(function (tbl) { tbl.rows.items.forEach(function (row, ri) { … cell.__matKey = AshaarMatrix.positionKey(...) … }) })` block from Task 5 of the justification work) so it prefers the map. Use the table index into `ccCells`:
```js
      var allCells = [];
      tables.items.forEach(function (tbl, ti) {
        // Prefer the persisted bandh map for this table; fall back to geometry.
        var pattern = ccCells && ccCells[ti];
        var perRowCounts = tbl.rows.items.map(function (row) { return row.cells.items.length; });
        var useMap = AshaarCellMap.alignPatternToTable(perRowCounts, pattern);
        var tblMap = useMap ? AshaarCellMap.buildBandhCellMap(pattern) : null;
        var cellSeq = 0;
        tbl.rows.items.forEach(function (row, ri) {
          var cols = row.cells.items.length;
          row.cells.items.forEach(function (cell, ci) {
            allCells.push(cell);
            var mapped = tblMap ? tblMap[cellSeq] : null;
            cellSeq++;
            if (mapped) {
              cell.__kind = mapped.kind;                    // "content" | "spacing"
              cell.__matKey = mapped.label || mapped.slot;  // label for content harmony
            } else {
              cell.__kind = null;                            // unknown → infer downstream
              cell.__matKey = AshaarMatrix.positionKey({ row: ri, col: ci, span: cols });
            }
            cell.body.load("text");
            cell.body.font.load("name,size");
```
(Leave the remaining loaded fields — `cell.body.paragraphs.load("alignment")`, the word-range loads, etc. — exactly as they are. Only the loop header, `allCells.push`, and the `__kind`/`__matKey` assignment change.)

- [ ] **Step 3: `justifySelection` — honor the explicit spacing tag**

In the matrix pre-pass and the phase-1 loop, a cell is content when the map says so (an empty content cell must still be treated as content). Where the code currently decides content-vs-spacing from emptiness, prefer `__kind`:
- In the matrix pre-pass (`allCells.forEach` that pushes `matrixCells`), replace `if (!AshaarMatrix.isContentCell(base)) return;` with:
```js
        var isContent = cell.__kind === "content" || (cell.__kind == null && AshaarMatrix.isContentCell(base));
        if (!isContent) return; // tagged spacing (even with stray text) excluded from the matrix
```
- At the top of the phase-1 per-cell loop, right after `var current = (cell.body.text || "").trim();`, add an explicit spacing skip:
```js
        if (cell.__kind === "spacing") return; // structural gap — never justified
```

- [ ] **Step 4: `applyProfileToQaseeda` — per-block map**

`gatherQaseedaBlocks` already loads `items/title,items/tag`, so each block's tag is available. Thread each block's per-table pattern through the table flattening, then into the cell capture.

**(a)** Replace the `allTables` flattening (`taskpane.js` ~486–488):
```js
        var allTables = [];
        var allTablePatterns = [];
        var blockCells = blocks.map(function (cc) {
          var p = AshaarWord.parseContentControlTag(cc.tag);
          return (p && p.cells) || null;
        });
        blockTables.forEach(function (t, bi) {
          t.items.forEach(function (tbl, j) {
            allTables.push(tbl);
            allTablePatterns.push(blockCells[bi] ? blockCells[bi][j] : null);
          });
        });
        if (!allTables.length) { summary = "Qaseeda “" + name + "” has no tables to size."; return; }
```

**(b)** Replace the `tableInfos` capture (`taskpane.js` ~507, the whole `var tableInfos = allTables.map(function (tbl) { … });`) with the map-threading version:
```js
        var tableInfos = allTables.map(function (tbl, ai) {
          var pattern = allTablePatterns[ai];
          var perRowCounts = tbl.rows.items.map(function (row) { return row.cells.items.length; });
          var tblMap = AshaarCellMap.alignPatternToTable(perRowCounts, pattern)
            ? AshaarCellMap.buildBandhCellMap(pattern) : null;
          var seq = 0;
          var cells = [];
          tbl.rows.items.forEach(function (row, ri) {
            var cols = row.cells.items.length;
            row.cells.items.forEach(function (cell, ci) {
              var f = cell.body.font;
              var current = (cell.body.text || "").trim();
              var base = stripJustification(current);
              var mapped = tblMap ? tblMap[seq] : null;
              seq++;
              cells.push({
                cell: cell,
                current: current,
                base: base,
                measure: base.replace(/\s+/g, " ").trim(),
                matKey: mapped ? (mapped.label || mapped.slot) : AshaarMatrix.positionKey({ row: ri, col: ci, span: cols }),
                kind: mapped ? mapped.kind : null,
                fontName: (f && f.name) || "",
                fontSize: (f && f.size) || 0
              });
            });
          });
          return { tbl: tbl, cells: cells };
        });
```

**(c)** In the re-justify loop, skip spacing cells: replace `if (!c.base) return;` with:
```js
            if (c.kind === "spacing") return;              // structural gap — never justified
            if (!c.base && c.kind !== "content") return;   // empty & not known-content → skip
```
The `qMatrix` pre-pass and re-justify already key on `c.matKey` (now the label when a map exists) — no other change.

- [ ] **Step 5: Syntax + suite**

Run: `node --check src/taskpane/taskpane.js`; Expected: no output.
Run: `npm test`; Expected: green (glue only; matrix/label logic already covered).

- [ ] **Step 6: Commit**

```bash
git add src/taskpane/taskpane.js
git commit --no-gpg-sign -m "feat(cellmap): justify keys harmony on labels + skips tagged gaps"
```

---

### Task 6: Read-only cell-map view in the pane

**Files:**
- Modify: `src/taskpane/taskpane.html` (a container + a button in the justify area)
- Modify: `src/taskpane/taskpane.js` (a `showCellMap` handler + `bind()` wiring)

**Interfaces:**
- Consumes `AshaarWord.parseContentControlTag`, `AshaarCellMap.buildBandhCellMap`. Manual-verify.

- [ ] **Step 1: Add the container + button**

In `src/taskpane/taskpane.html`, right after the `justify-selection` button (`<button id="justify-selection" …>`), add:
```html
        <button id="show-cell-map" type="button" class="button--secondary">Show cell structure</button>
        <div id="cell-map-view" class="cell-map-view" hidden></div>
```

- [ ] **Step 2: Implement the handler**

In `src/taskpane/taskpane.js`, add a function near `getQaseedaAtSelection` (it uses the same block-at-cursor lookup):
```js
  // Read-only: show the bandh cell-map (labels + gaps) for the Ashaar Poem block
  // at the cursor. No document mutation. Labels can't be shown on the Word page
  // itself (no native per-cell text overlay), so the pane is their home.
  async function showCellMap() {
    var view = document.getElementById("cell-map-view");
    if (!view) return;
    if (typeof Word === "undefined") { setMessage("Open this task pane inside Word."); return; }
    var patterns = null;
    try {
      await Word.run(async function (context) {
        var cc = context.document.getSelection().parentContentControlOrNullObject;
        cc.load("title,tag");
        await context.sync();
        if (!cc.isNullObject && cc.title === "Ashaar Poem") {
          var payload = AshaarWord.parseContentControlTag(cc.tag);
          patterns = payload && payload.cells;
        }
      });
    } catch (e) { /* leave patterns null */ }

    if (!patterns || !patterns.length) {
      view.hidden = false;
      view.textContent = "No cell map on the block at the cursor (older or hand-drawn table).";
      return;
    }
    var html = "";
    patterns.forEach(function (pattern, bi) {
      var map = AshaarCellMap.buildBandhCellMap(pattern);
      html += "<div class=\"cell-map-bandh\"><b>Bandh " + (bi + 1) + "</b>";
      var lastRow = -1, rowHtml = "";
      function flush() { if (rowHtml) html += "<div class=\"cell-map-row\">" + rowHtml + "</div>"; rowHtml = ""; }
      map.forEach(function (e) {
        if (e.row !== lastRow) { flush(); lastRow = e.row; }
        rowHtml += e.kind === "content"
          ? "<span class=\"cell-map-cell\">" + e.label + "</span>"
          : "<span class=\"cell-map-gap\">(gap)</span>";
      });
      flush();
      html += "</div>";
    });
    view.hidden = false;
    view.innerHTML = html;
  }
```

- [ ] **Step 3: Wire the button in `bind()`**

Near the existing `document.getElementById("justify-selection").addEventListener("click", justifySelection);`, add:
```js
    var showMapBtn = document.getElementById("show-cell-map");
    if (showMapBtn) showMapBtn.addEventListener("click", showCellMap);
```

- [ ] **Step 4: Minimal styling (optional but keeps it legible)**

In `src/taskpane/taskpane.css`, append:
```css
.cell-map-view { margin-top: 8px; font-size: 12px; }
.cell-map-bandh { margin-bottom: 6px; }
.cell-map-row { display: flex; gap: 4px; margin: 2px 0; direction: rtl; }
.cell-map-cell { padding: 1px 6px; border: 1px solid #888; border-radius: 3px; }
.cell-map-gap { padding: 1px 6px; color: #999; }
```
(`direction: rtl` so `A1` shows on the right, matching the RTL reading order.)

- [ ] **Step 5: Syntax + suite**

Run: `node --check src/taskpane/taskpane.js`; Expected: no output.
Run: `npm test`; Expected: green.

- [ ] **Step 6: Commit**

```bash
git add src/taskpane/taskpane.html src/taskpane/taskpane.js src/taskpane/taskpane.css
git commit --no-gpg-sign -m "feat(cellmap): read-only cell-structure view in the pane"
```

---

### Task 7: Manual Word verification

**Files:** none. `npm start` (opens `test-documents/marsiya-test.docx`).

- [ ] **Step 1: Insert tags the map.** Insert a couplet poem (Conversion → Insert as Table). Confirm no error. (Optional: inspect the content-control tag via the debugger — `parseContentControlTag(cc.tag).cells` is a non-empty array of patterns.)

- [ ] **Step 2: Read-only view.** Put the cursor in the inserted block, click **Show cell structure**. Confirm the pane lists the bandh(s) with labels (`A1`, `A2`, …) on the right and `(gap)` markers, one row per table row. Multi-stanza poem → one "Bandh N" group per stanza.

- [ ] **Step 3: Gaps never justified.** Justify the block (Kashida). Confirm the empty gap cells receive **no** tatweels/spaces (they stay empty), in both Cell-fit and Natural-fit.

- [ ] **Step 4: Empty content cell stays content.** Insert/clear a poem so one **content** cell is empty but is a real misra slot. Justify. Confirm that cell is treated as content (not silently skipped as a gap) — i.e. it participates in its position's harmony rather than being ignored.

- [ ] **Step 5: Cross-bandh harmony by label.** In a multi-stanza (same-shape) qaseeda, Natural-fit justify. Confirm corresponding cells across bandhs (all `A1`, all `A2`) land at one width — now matched by label, not geometry.

- [ ] **Step 6: Fallback.** Adopt / hand-draw a table (no `cells` tag) and justify. Confirm it still justifies via the geometric fallback (behavior unchanged), and **Show cell structure** reports "No cell map …".

## Self-Review notes

- **Spec coverage:** §1 model (labels, RTL numbering, spacing `slot`) → Task 1; §2 derivation from the row model + 1:1 with generated cells → Tasks 2 + 3 (cross-check); §3 persistence (`cells`, `v2`, parse, absent→null) → Task 4; §4 justify consumes map (label key + explicit content/spacing + geometric fallback) → Task 5; §5 read-only pane view → Task 6; §6 testing → node Tasks 1–4 + manual Task 7; §7 non-goals untouched (spacing `slot` left as the future anchor, nothing built on it).
- **Type consistency:** pattern is `Array<Array<"c"|"g">>` everywhere (Tasks 1–5); `buildBandhCellMap` entry `{index,row,kind,label,slot}` consumed identically in Tasks 5 & 6; `parseContentControlTag(...).cells` is `Array<pattern>|null` in Tasks 4/5/6; `contentControlTag(text,opts,cellPatterns?)` optional 3rd arg (Task 4) matches the two call sites wired in Task 4 Step 5.
- **Reused unchanged:** `AshaarMatrix.*` (natural-width matrix + target math — only the *key* feeding it changes), `justifyRunsConcentrated`, the Cell-fit/Natural-fit branches, `renderForWordOoxml`, `wrapOoxml`, `getQaseedaAtSelection` pattern.
- **Risk note:** `stanzaCellPattern` mirrors the generator's cell order rather than refactoring the generator to a single source (lower regression risk in a well-tested area); Task 3's cross-check test against the real emitted OOXML is what prevents drift. If a future layout mode adds a new cell shape, Task 3 fails first and points at the pattern to update.
- **Deferred:** styling, special-gap purposes, per-cell overrides, and map *editing* are SP2 / backlog (roadmap), not here.
```
