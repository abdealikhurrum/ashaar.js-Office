# Poetry Justification Modes (Cell-fit & Natural-fit) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Give poetry justification two explicit modes that share the concentrated-tatweel engine and one organizing quantity (the per-position natural-width matrix): **Cell-fit** (fills to the true cell edge, residual filled by Word's `jc=distribute`) and **Natural-fit** (fills to a per-position matrix width extended by strength, residual filled by capped micro-spaces).

**Architecture:** A new pure module `natural-width-matrix.js` computes, per grid position, the longest tatweel-free ("natural") width across all cells at that position, plus the two modes' target math. `profiles.js` gains a `justify.fillMode` field (default `"natural-fit"`). `word-html.js` gains a `misraDistributeXml` OOXML emitter (a paragraph with `<w:jc w:val="distribute"/>`) for the Cell-fit write path. `justifySelection` and `applyProfileToQaseeda` in `taskpane.js` build the matrix and branch on `fillMode`: Cell-fit targets `colPx` (no buffer) and writes distribute-OOXML with no injected spaces; Natural-fit retargets the existing (committed) concentrate + `capMicroSpaces` behavior from `colPx` to the matrix `target`. A mode toggle in the pane stores the choice on the qaseeda profile.

**Tech Stack:** Vanilla JS (ES5/UMD, `var`/`function`, no build step), Office.js v1, Node `assert` tests. Never edit `src/vendor/`.

**Spec:** `docs/superpowers/specs/2026-07-11-poetry-justification-modes-design.md`
**Builds on (committed):** `docs/superpowers/plans/2026-07-11-poetry-kashida-concentration.md` — the shared engine (`AshaarJustify.justifyRunsConcentrated`, `AshaarWord.strengthToElongationShare`, `AshaarWord.strengthToMaxPositions`) already exists and is unchanged.

## Global Constraints

- ES5/UMD only (`var`, `function`); never edit `src/vendor/`.
- Pure logic (matrix math, profile fields, OOXML string emitters) is **node-tested**; the Office.js glue in `taskpane.js` is **manual-verify** (final task).
- `φ = AshaarWord.strengthToElongationShare(s) = (clamp(s,1,10) − 1) / 9`. Reuse it — never re-derive φ inline.
- `perPositionEm = 0.5`; `maxPositions = AshaarWord.strengthToMaxPositions(s)` (K(1..3)=1,2,3; s≥4 → 0 = unbounded). Both modes pass these to `justifyRunsConcentrated`.
- Buffer `= 0.28em = 0.28 · fontSizePt · 96/72` px. **Cell-fit uses NO buffer** (distribute fills the true edge). **Natural-fit's `reach = colPx − buffer`** (the overflow-tolerance band `[reach, colPx]` absorbs discrete micro-space overshoot).
- `CELL_MARGIN_PT = 5.76`; `colPx = max(1, (columnWidth − 2·CELL_MARGIN_PT)) · 96/72`.
- Default `fillMode = "natural-fit"`.
- Micro-space glyph: hair space `U+200A` (fallback thin space `U+2009` when the hair space measures 0) — reuse the existing `MICRO_SPACE` locals.
- `npm test` green after every task.

---

### Task 1: Natural-width matrix module (pure)

**Files:**
- Create: `src/taskpane/natural-width-matrix.js`
- Test: `tests/natural-width-matrix.test.js`
- Modify: `package.json` (add the test file to the `test` script), `src/taskpane/taskpane.html` (script load list)

**Interfaces:**
- Produces `AshaarMatrix.positionKey({row, col, span}) → string` — a stable signature grouping cells that occupy the same position across bandhs.
- Produces `AshaarMatrix.buildMatrix(cells) → { [key]: Wpos }` where `cells = [{key, natural}]`; `Wpos` = the max `natural` at each `key`.
- Produces `AshaarMatrix.isContentCell(text) → boolean` — true when `text` has non-whitespace content (a spacing/structural cell is excluded from the matrix and from justification).
- Produces `AshaarMatrix.naturalFitTarget(Wpos, reach, phi) → number` = `Wpos + phi·max(0, reach − Wpos)`.
- Produces `AshaarMatrix.cellFitBudget(natural, colPx, phi) → number` = `natural + phi·max(0, colPx − natural)`.

- [ ] **Step 1: Write the failing test**

Create `tests/natural-width-matrix.test.js`:

```js
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
```

- [ ] **Step 2: Run to verify it fails**

Run: `node tests/natural-width-matrix.test.js`
Expected: FAIL — `Cannot find module '../src/taskpane/natural-width-matrix'`.

- [ ] **Step 3: Write the module**

Create `src/taskpane/natural-width-matrix.js`:

```js
/**
 * AshaarMatrix — the per-position natural-width matrix and the two justification
 * modes' target math (pure, UMD; no DOM/Office dependency, node-testable).
 *
 * Visual harmony means corresponding cells across bandhs (stanzas) share a
 * width. A content cell's POSITION is its grid signature (row within the bandh,
 * grid-column start, span). The matrix maps each position → Wpos, the longest
 * tatweel-free ("natural") width among all content cells at that position.
 *
 *   Cell-fit budget:  natural + φ·(colPx − natural)   — φ = elongation share
 *   Natural-fit target: Wpos   + φ·(reach − Wpos)     — φ = misra width dial
 *
 * See docs/superpowers/specs/2026-07-11-poetry-justification-modes-design.md.
 */
(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) module.exports = factory();
  else root.AshaarMatrix = factory();
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  // Stable signature for a grid position. Cells with the same (row, col, span)
  // across bandhs occupy the "same position" and balance to one width.
  function positionKey(sig) {
    sig = sig || {};
    return (sig.row || 0) + ":" + (sig.col || 0) + ":" + (sig.span || 0);
  }

  // A content cell holds misra text (participates in the matrix and is
  // justified); a spacing cell is a structural gap (no text) — excluded.
  function isContentCell(text) {
    return !!(text && String(text).replace(/\s+/g, "").length);
  }

  // cells: [{key, natural}] → { key → max natural }. The baseline Wpos per
  // position is the LONGEST natural width, so shorter misras elongate up to it.
  function buildMatrix(cells) {
    var out = {};
    (cells || []).forEach(function (c) {
      if (!c || !c.key) return;
      var n = Number(c.natural) || 0;
      if (!(c.key in out) || n > out[c.key]) out[c.key] = n;
    });
    return out;
  }

  // Natural-fit total fill target: harmony baseline extended toward the
  // container by strength. φ=0 → Wpos (pure harmony); φ=1 → reach.
  function naturalFitTarget(Wpos, reach, phi) {
    Wpos = Number(Wpos) || 0;
    return Wpos + (Number(phi) || 0) * Math.max(0, (Number(reach) || 0) - Wpos);
  }

  // Cell-fit elongation budget: strength is the elongation:spacing ratio, so φ
  // is the fraction of the cell gap the tatweel engine tries to cover; the rest
  // is left for Word's distribute residual. φ=0 → natural (spacing-only).
  function cellFitBudget(natural, colPx, phi) {
    natural = Number(natural) || 0;
    return natural + (Number(phi) || 0) * Math.max(0, (Number(colPx) || 0) - natural);
  }

  return {
    positionKey: positionKey,
    isContentCell: isContentCell,
    buildMatrix: buildMatrix,
    naturalFitTarget: naturalFitTarget,
    cellFitBudget: cellFitBudget,
  };
}));
```

- [ ] **Step 4: Run to verify it passes**

Run: `node tests/natural-width-matrix.test.js`
Expected: PASS — `natural-width-matrix.test.js OK`.

- [ ] **Step 5: Register the test in the suite**

In `package.json`, append the new test file to the `test` script. Change the end of the `test` value from:
```
... && node tests/kashida-residual.test.js"
```
to:
```
... && node tests/kashida-residual.test.js && node tests/natural-width-matrix.test.js"
```

- [ ] **Step 6: Register the module for the browser**

In `src/taskpane/taskpane.html`, add the module to the `srcs` array (it must load before `taskpane.js`, and `profiles.js` is fine as a neighbor). Insert after `"./profiles.js",`:
```js
          "./profiles.js",
          "./natural-width-matrix.js",
          "./fonts.js",
```

- [ ] **Step 7: Full suite**

Run: `npm test`
Expected: green (all files, including the new one).

- [ ] **Step 8: Commit**

```bash
git add src/taskpane/natural-width-matrix.js tests/natural-width-matrix.test.js package.json src/taskpane/taskpane.html
git commit -m "feat(justify): natural-width matrix + Cell-fit/Natural-fit target math"
```

---

### Task 2: Profile `fillMode` field

**Files:**
- Modify: `src/taskpane/profiles.js`
- Test: `tests/profiles.test.js`

**Interfaces:**
- `defaultProfile(name).justify.fillMode === "natural-fit"`.
- Produces `AshaarProfiles.normalizeFillMode(m) → "cell-fit" | "natural-fit"` (default `"natural-fit"` for anything unrecognised).
- `normalizeProfile` carries a stored `justify.fillMode` through unchanged (the existing deep merge of the `justify` bucket already does this once the default exists).

- [ ] **Step 1: Write the failing test**

Append to `tests/profiles.test.js` (after the existing `strengthToTargetFill`/`normalizeStrength` assertions):

```js
// ── fillMode: default natural-fit, normalized, preserved through merge ────────
{
  const p = defaultProfile("Karbala");
  assert.equal(p.justify.fillMode, "natural-fit", "defaults to natural-fit");
}
assert.equal(AshaarProfiles.normalizeFillMode("cell-fit"), "cell-fit");
assert.equal(AshaarProfiles.normalizeFillMode("natural-fit"), "natural-fit");
assert.equal(AshaarProfiles.normalizeFillMode("bogus"), "natural-fit", "unknown → default");
assert.equal(AshaarProfiles.normalizeFillMode(undefined), "natural-fit");
{
  // A stored profile carrying cell-fit survives normalizeProfile.
  const stored = { name: "Q", justify: { mode: "kashida", strength: 7, fillMode: "cell-fit" } };
  const n = normalizeProfile(stored);
  assert.equal(n.justify.fillMode, "cell-fit", "stored fillMode preserved");
}
```

Also extend the `require` destructuring at the top of `tests/profiles.test.js` to pull in the new export:
```js
const {
  defaultProfile,
  normalizeProfile,
  mergeProfile,
  applyFontCorrection,
  deriveSharedWidths,
  columnPointsFromContentPx,
  strengthToTargetFill,
  normalizeStrength,
  normalizeFillMode,
} = require("../src/taskpane/profiles");
```

- [ ] **Step 2: Run to verify it fails**

Run: `node tests/profiles.test.js`
Expected: FAIL — `p.justify.fillMode` is `undefined` / `normalizeFillMode is not a function`.

- [ ] **Step 3: Implement**

In `src/taskpane/profiles.js`, add `fillMode` to the default `justify` bucket:
```js
      justify: { mode: "kashida", strength: 6, fillMode: "natural-fit" },   // "kashida" | "spacing" | "css" | "none"
```

Add the helper near `normalizeStrength` (before the `return {`):
```js
  // Sanitise a profile's stored fill mode. Cell-fit = fill to the true cell edge
  // (Word distribute residual); Natural-fit = fill to the per-position matrix
  // width (capped micro-space residual). Anything else defaults to natural-fit.
  function normalizeFillMode(mode) {
    return mode === "cell-fit" ? "cell-fit" : "natural-fit";
  }
```

Add it to the exports object:
```js
    strengthToTargetFill: strengthToTargetFill,
    normalizeStrength: normalizeStrength,
    normalizeFillMode: normalizeFillMode,
```

- [ ] **Step 4: Run to verify it passes**

Run: `node tests/profiles.test.js`
Expected: PASS.

- [ ] **Step 5: Full suite**

Run: `npm test`
Expected: green.

- [ ] **Step 6: Commit**

```bash
git add src/taskpane/profiles.js tests/profiles.test.js
git commit -m "feat(justify): profile justify.fillMode field (default natural-fit)"
```

---

### Task 3: `misraDistributeXml` OOXML emitter (Cell-fit write path)

**Files:**
- Modify: `src/taskpane/word-html.js`
- Test: `tests/word-html.test.js`

**Interfaces:**
- Produces `AshaarWord.misraDistributeXml(runs, sizePtFallback) → string` — one `<w:p>` with `<w:jc w:val="distribute"/>`; `runs = [{text, csName, sizePt?}]`; each run gets `<w:rtl/>`, its own `<w:rFonts w:cs=…/>` (when `csName` is set), and its size in half-points (per-run `sizePt` or the fallback). Injects NO spaces — the distribute `jc` is the residual.

- [ ] **Step 1: Write the failing test**

Append to `tests/word-html.test.js`:

```js
// ── misraDistributeXml: Cell-fit distribute paragraph ────────────────────────
{
  const xml = AshaarWord.misraDistributeXml(
    [{ text: "دل", csName: "Fatemi Maqala" }, { text: " ", csName: "Fatemi Maqala" }, { text: "ناداں", csName: "Fatemi Maqala" }],
    16
  );
  assert.match(xml, /<w:jc w:val="distribute"\/>/, "distribute jc");
  assert.match(xml, /<w:bidi\/>/, "rtl paragraph");
  assert.match(xml, /<w:rtl\/>/, "rtl runs");
  assert.match(xml, /<w:rFonts w:cs="Fatemi Maqala"\/>/, "per-run cs font");
  assert.match(xml, /<w:sz w:val="32"\/>/, "16pt → 32 half-points");
  assert.ok(xml.indexOf("دل") !== -1 && xml.indexOf("ناداں") !== -1, "carries text");
  // No injected hair/thin spaces (distribute is the residual, not micro-spaces).
  assert.ok(xml.indexOf(" ") === -1 && xml.indexOf(" ") === -1, "no micro-spaces injected");
}
{
  // Per-run size override wins over the fallback.
  const xml = AshaarWord.misraDistributeXml([{ text: "x", csName: "A", sizePt: 20 }], 16);
  assert.match(xml, /<w:sz w:val="40"\/>/, "per-run 20pt → 40 half-points");
}
```

- [ ] **Step 2: Run to verify it fails**

Run: `node tests/word-html.test.js`
Expected: FAIL — `AshaarWord.misraDistributeXml is not a function`.

- [ ] **Step 3: Implement**

In `src/taskpane/word-html.js`, add the function immediately after `runsToMisraXml` (which ends at line 1242, just before `baytRowsOoxml`):

```js
  // Cell-fit residual is Word Distributed justification: emit the (tatweel'd)
  // misra as a paragraph with <w:jc w:val="distribute"/> so Word stretches the
  // inter-word gaps to the true cell edge. Each run keeps its own cs font (+
  // size); NO micro-spaces are injected (that is the Natural-fit residual).
  // runs: [{text, csName, sizePt?}].
  function misraDistributeXml(runs, sizePtFallback) {
    var body = (runs || []).map(function (r) {
      var sz = r.sizePt || sizePtFallback;
      var szXml = sz ? '<w:sz w:val="' + Math.round(sz * 2) + '"/><w:szCs w:val="' + Math.round(sz * 2) + '"/>' : "";
      var cs = r.csName ? '<w:rFonts w:cs="' + r.csName + '"/>' : "";
      return "<w:r><w:rPr><w:rtl/>" + cs + szXml + "</w:rPr>" +
        '<w:t xml:space="preserve">' + escapeXml(r.text) + "</w:t></w:r>";
    }).join("");
    return "<w:p><w:pPr><w:bidi/><w:spacing w:after=\"80\"/><w:jc w:val=\"distribute\"/></w:pPr>" + body + "</w:p>";
  }
```

Add it to the module's export object (near `runsToMisraXml: runsToMisraXml,` at line ~1495):
```js
    runsToMisraXml: runsToMisraXml,
    misraDistributeXml: misraDistributeXml,
```

- [ ] **Step 4: Run to verify it passes**

Run: `node tests/word-html.test.js`
Expected: PASS.

- [ ] **Step 5: Full suite**

Run: `npm test`
Expected: green.

- [ ] **Step 6: Commit**

```bash
git add src/taskpane/word-html.js tests/word-html.test.js
git commit -m "feat(justify): misraDistributeXml emitter for Cell-fit distribute residual"
```

---

### Task 4: Mode-toggle UI + wiring (free-form + profile panel)

**Files:**
- Modify: `src/taskpane/taskpane.html` (two `<select>` controls)
- Modify: `src/taskpane/taskpane.js` (`options()`, `panelToProfile`, `profileToPanel`, the DOM var block, `bind()` label sync)

**Interfaces:**
- Produces `opts.fillMode` (from `options()`) — `"cell-fit" | "natural-fit"`, default `"natural-fit"`.
- Produces `profile.justify.fillMode` written by `panelToProfile` and read back by `profileToPanel`.

No node test (DOM/Office glue). Verified in Task 6.

- [ ] **Step 1: Add the free-form mode control**

In `src/taskpane/taskpane.html`, insert a new field right after the Justification `<select>` field (which closes at line 54, before the `tatweel-count` field at line 55):

```html
        <div class="field">
          <label for="justify-fill-mode">Fill mode</label>
          <select id="justify-fill-mode">
            <option value="natural-fit" selected>Natural-fit (harmony)</option>
            <option value="cell-fit">Cell-fit (precise)</option>
          </select>
        </div>
```

- [ ] **Step 2: Add the profile mode control**

In `src/taskpane/taskpane.html`, inside the qaseeda "Advanced" body, add a field right after the `qaseeda-justify-mode` field (which closes at line 285, before the `qaseeda-corr-font` field at line 286):

```html
              <div class="field">
                <label for="qaseeda-fill-mode">Fill mode</label>
                <select id="qaseeda-fill-mode">
                  <option value="natural-fit" selected>Natural-fit (harmony)</option>
                  <option value="cell-fit">Cell-fit (precise)</option>
                </select>
              </div>
```

- [ ] **Step 3: Declare the DOM references**

In `src/taskpane/taskpane.js`, next to the existing `var justifyMode = document.getElementById("justify-mode");` (line 10), add:
```js
  var justifyFillMode = document.getElementById("justify-fill-mode");
```
And next to `var qaseedaJustifyMode = document.getElementById("qaseeda-justify-mode");` (line 78), add:
```js
  var qaseedaFillMode = document.getElementById("qaseeda-fill-mode");
```

- [ ] **Step 4: Expose `fillMode` from `options()`**

In `options()` (line 239), add the field to the returned object (after the `justify:` line, line 242):
```js
      justifyMode: justifyMode.value,
      justify: justifyMode.value === "none" ? false : justifyMode.value,
      fillMode: (justifyFillMode && justifyFillMode.value) || "natural-fit",
```

- [ ] **Step 5: Read/write the profile field in panel ↔ profile**

In `panelToProfile()` (line 684), after `p.justify.mode = qaseedaJustifyMode.value;` (line 688) add:
```js
      p.justify.fillMode = (qaseedaFillMode && qaseedaFillMode.value) || "natural-fit";
```

In `profileToPanel(profile)` (line 700), after `qaseedaJustifyMode.value = p.justify.mode;` (line 704) add:
```js
      if (qaseedaFillMode) qaseedaFillMode.value = AshaarProfiles.normalizeFillMode(p.justify.fillMode);
```

- [ ] **Step 6: Syntax check + suite**

Run: `node --check src/taskpane/taskpane.js`
Expected: no output (valid).
Run: `npm test`
Expected: green (no new node coverage; regression check only).

- [ ] **Step 7: Commit**

```bash
git add src/taskpane/taskpane.html src/taskpane/taskpane.js
git commit -m "feat(justify): Cell-fit/Natural-fit mode toggle (pane + profile)"
```

---

### Task 5: `justifySelection` — build matrix + branch on `fillMode`

**Files:**
- Modify: `src/taskpane/taskpane.js` (`justifySelection`: the `allCells` build ~1410, a new matrix pre-pass before phase 1 ~1651, the Jameel/Mehr/generic branches ~1675–1811, and the phase-2 write ~1857–1896)

**Interfaces:**
- Consumes `AshaarMatrix.positionKey/buildMatrix/isContentCell/naturalFitTarget/cellFitBudget`, `AshaarWord.misraDistributeXml`, `AshaarWord.strengthToElongationShare`, `AshaarWord.strengthToMaxPositions`.
- `opts.fillMode` selects the branch. No node test (Office.js); verified in Task 7.

**Design (read before editing):**
- **Cell-fit:** `budget = cellFitBudget(natural, colPx, φ)` (no buffer). Concentrate tatweels to `budget`. Emit distribute-OOXML (`misraDistributeXml`) preserving each run's real cs font — **no** `capMicroSpaces`/`injectSpaceRuns`/`distributeMicroSpaces`. Word's `jc=distribute` is the residual.
- **Natural-fit (retarget of the committed behavior):** `reach = colPx − 0.28em`; `Wpos = matrix[key]` (fallback: the cell's own natural); `target = naturalFitTarget(Wpos, reach, φ)`. Concentrate to `target`; `capMicroSpaces` backfills `target − achieved`. Same write paths as today (Jameel → `runsToMisraXml`; Mehr → flat; generic → run-aware `outTexts`).

- [ ] **Step 1: Capture each cell's grid position**

In the `allCells` build loop (line 1411), annotate each cell with its signature. Replace:
```js
      var allCells = [];
      tables.items.forEach(function (tbl) {
        tbl.rows.items.forEach(function (row) {
          row.cells.items.forEach(function (cell) {
            allCells.push(cell);
            cell.body.load("text");
            cell.body.font.load("name,size");
```
with:
```js
      var allCells = [];
      tables.items.forEach(function (tbl) {
        tbl.rows.items.forEach(function (row, ri) {
          var cols = row.cells.items.length;
          row.cells.items.forEach(function (cell, ci) {
            allCells.push(cell);
            // Grid signature for the natural-width matrix. Cross-table (multi-
            // stanza) cells at the same (rowCount-guarded row, col) group so
            // corresponding misras balance to one width. `span` carries the
            // row's cell count so different-shaped rows never group together.
            cell.__matKey = AshaarMatrix.positionKey({ row: ri, col: ci, span: cols });
            cell.body.load("text");
            cell.body.font.load("name,size");
```
(Leave the remaining lines of that inner block — `cell.body.paragraphs.load("alignment");` etc. — unchanged.)

- [ ] **Step 2: Compute φ, fillMode, and the natural-width matrix before phase 1**

`elongShare` is already computed at the top of `justifySelection` (line 1318). Add the matrix pre-pass immediately before the `// Phase 1` comment (line 1652), i.e. after the font force-load block ends at line 1650:

```js
      // Natural-width matrix (harmony): the longest tatweel-free width per grid
      // position across every content cell in the work range. Natural-fit fills
      // each cell up to its position's Wpos (φ=1 pushes further, to the edge).
      var fillMode = opts.fillMode === "cell-fit" ? "cell-fit" : "natural-fit";
      var matrixCells = [];
      allCells.forEach(function (cell) {
        var base = stripJustification(cell.body.text || "").replace(/\s+/g, " ").trim();
        if (!AshaarMatrix.isContentCell(base)) return;
        var mf = cell.body.font;
        var mnm = (mf && mf.name) || repName, msz = (mf && mf.size) || repSize;
        var natPx = 0;
        if (canvasCtx) { canvasCtx.font = runFontStr(mnm, msz, false, false); natPx = canvasCtx.measureText(base).width; }
        cell.__natPx = natPx;
        matrixCells.push({ key: cell.__matKey, natural: natPx });
      });
      var widthMatrix = AshaarMatrix.buildMatrix(matrixCells);
```

- [ ] **Step 3: Jameel branch — branch on fillMode**

Replace the Jameel block body (lines 1686–1701, from `var jT = colPx - …` through `return; // handled …`) with:
```js
          var jNatural = wb.reduce(function (a, b) { return a + b; }, 0);
          if (fillMode === "cell-fit") {
            // Cell-fit: swap fasls up to the φ elongation budget (no buffer),
            // then let Word distribute the residual to the true edge.
            var jBudget = AshaarMatrix.cellFitBudget(jNatural, colPx, elongShare);
            var jSelC = AshaarKashidaFontswap.selectSwapRuns(fss, wb, ww, jBudget);
            var jRunsC = jSelC.runs.map(function (r) {
              return { text: r.text, csName: r.swap ? (cellDesc.kasheedaName || repName) : (cellDesc.wordName || repName), sizePt: repSize };
            });
            plans.push({ cell: cell, ooxml: AshaarWord.misraDistributeXml(jRunsC, repSize) });
            return;
          }
          // Natural-fit: fill to the position's matrix width (φ pushes toward the
          // buffered edge); capped hair-spaces backfill what the swaps miss.
          var jReach = colPx - 0.28 * repSize * 96 / 72;
          var jWpos = widthMatrix[cell.__matKey] || jNatural;
          var jTarget = AshaarMatrix.naturalFitTarget(jWpos, jReach, elongShare);
          var sel = AshaarKashidaFontswap.selectSwapRuns(fss, wb, ww, jTarget);
          var jGaps = 0;
          for (var jgi = 0; jgi < sel.runs.length; jgi++) { if (sel.runs[jgi].text === " ") jGaps++; }
          canvasCtx.font = repSize + "pt " + baseCss;
          var jSpacePx = canvasCtx.measureText(MICRO_SPACE).width || 1;
          var jn = AshaarResidual.capMicroSpaces(jTarget - sel.fill * jTarget, jGaps, jSpacePx, repSize * 96 / 72);
          var jRuns = AshaarResidual.injectSpaceRuns(sel.runs, jn, MICRO_SPACE);
          var swapXml = AshaarWord.runsToMisraXml(jRuns, cellAlign, opts, repSize);
          plans.push({ cell: cell, ooxml: swapXml });
          return; // handled — skip the tatweel/spacing paths for this cell
```
(The lines above this block — `var wideCss`, `var baseCss`, `var fss`, the `wb`/`ww` measurement loop — are unchanged.)

- [ ] **Step 4: Mehr branch — branch on fillMode**

Replace the Mehr block's target + residual (lines 1725–1738, from `var mT = colPx - …` through `if (mfinal !== current) plans.push(…)`) with:
```js
          var mNatural = mwb.reduce(function (a, b) { return a + b; }, 0);
          if (fillMode === "cell-fit") {
            var mBudget = AshaarMatrix.cellFitBudget(mNatural, colPx, elongShare);
            var mselC = AshaarKashidaFontswap.selectSwapRuns(mtoks, mwb, mww, mBudget);
            var moutC = mselC.runs.map(function (r, i) { return (r.swap && mww[i] > mwb[i]) ? melong[i] : mtoks[i]; }).join("");
            plans.push({ cell: cell, ooxml: AshaarWord.misraDistributeXml([{ text: moutC, csName: cellDesc.wordName || repName, sizePt: repSize }], repSize) });
            return;
          }
          var mReach = colPx - 0.28 * repSize * 96 / 72;
          var mWpos = widthMatrix[cell.__matKey] || mNatural;
          var mTarget = AshaarMatrix.naturalFitTarget(mWpos, mReach, elongShare);
          var msel = AshaarKashidaFontswap.selectSwapRuns(mtoks, mwb, mww, mTarget);
          var mout = msel.runs.map(function (r, i) { return (r.swap && mww[i] > mwb[i]) ? melong[i] : mtoks[i]; }).join("");
          var mGaps = mout.split(" ").length - 1;
          canvasCtx.font = mehrFont;
          var mSpacePx = canvasCtx.measureText(MICRO_SPACE).width || 1;
          var mn = AshaarResidual.capMicroSpaces(mTarget - msel.fill * mTarget, mGaps, mSpacePx, repSize * 96 / 72);
          var mfinal = AshaarWord.distributeMicroSpaces([mout], mn, MICRO_SPACE)[0];
          if (mfinal !== current) plans.push({ cell: cell, flat: mfinal, align: cellAlignOf(cell) });
          return;
```
(Note the `align:` added to the Mehr flat plan — consumed in Step 6.)

- [ ] **Step 5: Generic branch — branch on fillMode**

Replace the generic kashida block (lines 1793–1811, the whole `if (opts.justifyMode === "kashida" && !anyWhitespaceRun) { … }` body up to the matching `} else {`) with:
```js
        if (opts.justifyMode === "kashida" && !anyWhitespaceRun) {
          var gNatural = primRuns.reduce(function (a, r) { return a + r.measure(r.text); }, 0);
          var gMax = { perPositionEm: 0.5, maxPositions: AshaarWord.strengthToMaxPositions(opts.tatweelCount) };
          if (fillMode === "cell-fit") {
            // Cell-fit: concentrate tatweels to the φ budget (no buffer); Word's
            // distribute jc stretches the inter-word gaps to the true edge.
            var gBudgetC = AshaarMatrix.cellFitBudget(gNatural, colPx, elongShare);
            var concC = AshaarJustify.justifyRunsConcentrated(primRuns, gBudgetC, Object.assign({}, calibParams, gMax));
            var cfRuns = concC.runs.map(function (r, i) { return { text: r.text, csName: runs[i].name, sizePt: runs[i].size }; });
            plans.push({ cell: cell, ooxml: AshaarWord.misraDistributeXml(cfRuns, repSize) });
            return;
          }
          // Natural-fit: fill to the position's matrix width; capped micro-spaces
          // backfill whatever the concentrated tatweels didn't cover.
          var gReach = colPx - 0.28 * repSize * 96 / 72;
          var gWpos = widthMatrix[cell.__matKey] || gNatural;
          var gTarget = AshaarMatrix.naturalFitTarget(gWpos, gReach, elongShare);
          var conc = AshaarJustify.justifyRunsConcentrated(primRuns, gTarget, Object.assign({}, calibParams, gMax));
          outTexts = conc.runs.map(function (r) { return r.text; });
          var gGaps = primRuns.reduce(function (a, r) { return a + (r.text.split(" ").length - 1); }, 0);
          canvasCtx.font = runFontStr(repName, repSize, false, false);
          var gSpacePx = canvasCtx.measureText(MICRO_SPACE).width || 1;
          var gN = AshaarResidual.capMicroSpaces(gTarget - conc.achievedPx, gGaps, gSpacePx, repSize * 96 / 72);
          outTexts = AshaarWord.distributeMicroSpaces(outTexts, gN, MICRO_SPACE);
        } else {
```
(The `else { // spacing/scale … }` block below is unchanged.)

Then, so the run-aware and flat generic plans can clear a stale `distribute` jc on a mode switch, attach the cell's alignment. Change the two generic-path `plans.push` at the end of the phase-1 loop (lines 1848 and 1850):
```js
        if (!alignedOk) { plans.push({ cell: cell, flat: outTexts.join(" "), align: cellAlignOf(cell) }); return; }

        plans.push({ cell: cell, runs: runs, outTexts: outTexts, sp: sp, align: cellAlignOf(cell) });
```

- [ ] **Step 6: Phase-2 write — reset alignment for flat/run plans**

So that re-justifying a cell that was previously Cell-fit (paragraph `jc=distribute`) with a Natural-fit flat/run write clears the distribute (Office.js has no "distribute" alignment enum, so we set the cell's own alignment explicitly). In the phase-2 loop, map `align` → Office enum and apply it. Add this helper just before the `for (var pi = 0; …)` loop (line 1857):
```js
      function officeAlign(a) {
        if (a === "right") return Word.Alignment.right;
        if (a === "left") return Word.Alignment.left;
        return Word.Alignment.centered;
      }
```
In the `if (p.flat != null) { … }` block (line 1870), set alignment alongside the insert:
```js
        if (p.flat != null) {
          var flatPara = p.cell.body.paragraphs.getFirst();
          flatPara.insertText(p.flat, Word.InsertLocation.replace);
          if (p.align) flatPara.alignment = officeAlign(p.align);
          await context.sync();
          changed++;
          continue;
        }
```
In the run-aware `try { … }` block, after the `p.runs.forEach(…)` loop and before `if (cellChanged)` (line 1888), add:
```js
          if (p.align) { p.cell.body.paragraphs.getFirst().alignment = officeAlign(p.align); cellChanged = true; }
          if (cellChanged) { await context.sync(); changed++; }
```

- [ ] **Step 7: Syntax check + suite**

Run: `node --check src/taskpane/taskpane.js`
Expected: no output.
Run: `npm test`
Expected: green (regression only — this glue has no node coverage).

- [ ] **Step 8: Commit**

```bash
git add src/taskpane/taskpane.js
git commit -m "feat(justify): justifySelection builds matrix + branches Cell-fit/Natural-fit"
```

---

### Task 6: `applyProfileToQaseeda` — matrix + `fillMode` branch

**Files:**
- Modify: `src/taskpane/taskpane.js` (`applyProfileToQaseeda`: read `fillMode` ~463, build a cross-block matrix over `tableInfos` ~503, branch the per-cell justify ~617–644)

**Interfaces:**
- Consumes `AshaarProfiles.normalizeFillMode`, `AshaarMatrix.*`, `AshaarWord.misraDistributeXml`. No node test; verified in Task 7.

**Design:** The qaseeda path already equalises corresponding columns across same-shape blocks (the `sameShape` resize branch) — that IS the harmony sizing the spec's "cell sizing (self-sufficient)" describes; we add the matrix so the fill *target* (not just the column width) is per-position. `reach` is the cell's `colPx`.

- [ ] **Step 1: Read the fill mode**

In `applyProfileToQaseeda` (line 459), after `var elongShare = AshaarWord.strengthToElongationShare(strength);` (line 464) add:
```js
    var fillMode = AshaarProfiles.normalizeFillMode(profile.justify.fillMode);
```

- [ ] **Step 2: Tag each captured cell with a grid signature**

In the `tableInfos` capture (line 503), record the position signature per cell. Replace:
```js
        var tableInfos = allTables.map(function (tbl) {
          var cells = [];
          tbl.rows.items.forEach(function (row) {
            row.cells.items.forEach(function (cell) {
              var f = cell.body.font;
              var current = (cell.body.text || "").trim();
              var base = stripJustification(current);
              cells.push({
                cell: cell,
                current: current,
                base: base,
                measure: base.replace(/\s+/g, " ").trim(),
                fontName: (f && f.name) || "",
                fontSize: (f && f.size) || 0
              });
            });
          });
          return { tbl: tbl, cells: cells };
        });
```
with:
```js
        var tableInfos = allTables.map(function (tbl) {
          var cells = [];
          tbl.rows.items.forEach(function (row, ri) {
            var cols = row.cells.items.length;
            row.cells.items.forEach(function (cell, ci) {
              var f = cell.body.font;
              var current = (cell.body.text || "").trim();
              var base = stripJustification(current);
              cells.push({
                cell: cell,
                current: current,
                base: base,
                measure: base.replace(/\s+/g, " ").trim(),
                matKey: AshaarMatrix.positionKey({ row: ri, col: ci, span: cols }),
                fontName: (f && f.name) || "",
                fontSize: (f && f.size) || 0
              });
            });
          });
          return { tbl: tbl, cells: cells };
        });
```

- [ ] **Step 3: Build the cross-block matrix**

Immediately before the re-justify loop (line 614, `tableInfos.forEach(function (info) {` that runs the justify), add a pre-pass. Insert after the `MICRO_SPACE` glyph resolution (lines 611–613) and before `tableInfos.forEach`:
```js
        // Cross-block natural-width matrix: longest natural per position across
        // ALL blocks of this qaseeda (harmony). Measured in each cell's own font.
        var qMatrixCells = [];
        tableInfos.forEach(function (info) {
          info.cells.forEach(function (c) {
            if (!AshaarMatrix.isContentCell(c.measure)) return;
            var fnm = c.fontName || repName, fsz = c.fontSize || repSize;
            canvasCtx.font = fsz + "pt \"" + fnm + "\"";
            c.natPx = AshaarProfiles.applyFontCorrection(canvasCtx.measureText(c.measure).width, fnm, profile.fontCorrections);
            qMatrixCells.push({ key: c.matKey, natural: c.natPx });
          });
        });
        var qMatrix = AshaarMatrix.buildMatrix(qMatrixCells);
```

- [ ] **Step 4: Branch the per-cell justify**

Replace the `if (doKashida && colPx > 0) { … }` body inside the re-justify loop (lines 619–640, from `var fname = …` through the `justified = AshaarWord.distributeMicroSpaces(…)` line) with:
```js
            if (doKashida && colPx > 0) {
              var fname = c.fontName || repName;
              var csize = c.fontSize || repSize;
              canvasCtx.font = csize + "pt \"" + fname + "\"";
              var runOne = [{
                text: c.base, fontSize: csize, fontProfile: null,
                measure: function (s) { canvasCtx.font = csize + "pt \"" + fname + "\""; return canvasCtx.measureText(s).width; }
              }];
              var cNatural = c.natPx != null ? c.natPx : canvasCtx.measureText(c.base).width;
              var cMax = { perPositionEm: 0.5, maxPositions: AshaarWord.strengthToMaxPositions(strength) };
              if (fillMode === "cell-fit") {
                // Cell-fit: concentrate to the φ budget (no buffer); the distribute
                // OOXML paragraph fills the residual to the true edge.
                var cBudgetCf = AshaarMatrix.cellFitBudget(cNatural, colPx, elongShare);
                var cConcCf = AshaarJustify.justifyRunsConcentrated(runOne, cBudgetCf, cMax);
                c.ooxml = AshaarWord.misraDistributeXml([{ text: cConcCf.runs[0].text, csName: fname, sizePt: csize }], csize);
                justified = null; // written via OOXML below
              } else {
                // Natural-fit: fill to this position's matrix width; capped
                // micro-spaces backfill whatever the tatweels didn't cover.
                var cReach = colPx - 0.28 * csize * 96 / 72;
                var cWpos = qMatrix[c.matKey] || cNatural;
                var cTarget = AshaarMatrix.naturalFitTarget(cWpos, cReach, elongShare);
                var cConc = AshaarJustify.justifyRunsConcentrated(runOne, cTarget, cMax);
                var cElong = cConc.runs[0].text;
                var cGaps = cElong.split(" ").length - 1;
                canvasCtx.font = csize + "pt \"" + fname + "\"";
                var cSpacePx = canvasCtx.measureText(MICRO_SPACE).width || 1;
                var cN = AshaarResidual.capMicroSpaces(cTarget - cConc.achievedPx, cGaps, cSpacePx, csize * 96 / 72);
                justified = AshaarWord.distributeMicroSpaces([cElong], cN, MICRO_SPACE)[0];
              }
            }
```

- [ ] **Step 5: Write via OOXML for Cell-fit**

The existing write is `if (justified !== c.current) { c.cell.body.paragraphs.getFirst().insertText(justified, …); changed++; }` (lines 641–644). Replace it with a branch that uses OOXML when Cell-fit produced one:
```js
            if (c.ooxml) {
              c.cell.body.clear();
              c.cell.body.insertOoxml(AshaarWord.wrapOoxml(c.ooxml), Word.InsertLocation.replace);
              changed++;
            } else if (justified != null && justified !== c.current) {
              c.cell.body.paragraphs.getFirst().insertText(justified, Word.InsertLocation.replace);
              changed++;
            }
```
(Note: `justified` is initialised to `c.base` at the top of the loop; the Cell-fit branch sets it to `null` so this branch is skipped in favour of the OOXML write. The Natural-fit branch and the non-kashida case leave `justified` as text.)

- [ ] **Step 6: Syntax check + suite**

Run: `node --check src/taskpane/taskpane.js`
Expected: no output.
Run: `npm test`
Expected: green.

- [ ] **Step 7: Commit**

```bash
git add src/taskpane/taskpane.js
git commit -m "feat(justify): qaseeda apply builds cross-block matrix + Cell-fit/Natural-fit"
```

---

### Task 7: Manual Word verification

**Files:** none.

Load the add-in against a real document and confirm behavior. `npm start` (or reload the pane) — the debugging config already opens `test-documents/marsiya-test.docx`.

- [ ] **Step 1: Mode toggle present & default**

Open the pane. Confirm the **Fill mode** control shows under Justification with **Natural-fit (harmony)** selected by default; the qaseeda **Advanced** panel has its own Fill mode select, also defaulting to Natural-fit.

- [ ] **Step 2: Natural-fit harmony (multi-bandh)**

Insert / open a qaseeda with several stanzas that share a template (e.g. the marsiya bandh `A3 A2 A1 / B1 / C2 C1`). Set Fill mode = Natural-fit, strength ~5, Justify. Confirm **corresponding cells across bandhs land at one width** (all "A1" cells equal, all "A2" equal), and a different-meter position (B1) does not drag the A row. Sweep strength **1 → 10**: at 1 each cell sits at its position's natural (shortest cells grow up to the longest at that position; longest ~unchanged); at 10 cells fill toward the column edge.

- [ ] **Step 3: Cell-fit precise fill (per font)**

Set Fill mode = Cell-fit. Justify. Confirm each misra fills to the **true cell edge** via Word Distributed justification (inter-word gaps stretch), composing with the tatweels/swaps. Repeat with the font set to **Jameel** (whole-span swaps + distribute), **Mehr** (trailing tatweels + distribute), **generic/Fatemi Maqala** (concentrated tatweels + distribute). Sweep strength 1/5/10 — low = spacing-dominant (distribute does most of the fill), high = elongation-dominant.

- [ ] **Step 4: Cell-fit needs accurate widths; Natural-fit forgives**

Make a table deliberately too wide for its text. Cell-fit → the line still reaches the edge but with wide distribute gaps (operator is expected to size cells). Switch to Natural-fit → the misra fills only to its matrix width (no gaping), independent of the oversized column.

- [ ] **Step 5: Mode switch clears stale distribute**

On a cell justified with **Cell-fit** (now `jc=distribute`), switch to **Natural-fit** and re-justify. Confirm the paragraph no longer distributes (no leftover wide gaps) — the alignment reset in the write path cleared it. Do the reverse (Natural-fit → Cell-fit) and confirm distribute engages.

- [ ] **Step 6: Idempotent re-justify**

Re-justify the same cells twice at a fixed strength in each mode. Confirm **no compounding** (tatweel/space counts stable) — `stripJustification` strips U+0640 + micro-spaces before recomputing.

- [ ] **Step 7: Qaseeda apply consistency**

Save a qaseeda profile (name it, pick a fill mode + strength), assign a block, "Save & Apply to all". Confirm every tagged block justifies consistently in the chosen mode; Cell-fit blocks distribute to their edges, Natural-fit blocks share per-position widths across blocks.

## Self-Review notes

- **Spec coverage:** §1 matrix + content/spacing tagging → Task 1 (`buildMatrix`, `positionKey`, `isContentCell`); §1 harmony scope (enclosing content control / cross-block) → Task 5 pre-pass (work-range cells) + Task 6 pre-pass (all qaseeda blocks). §2 shared engine → reused unchanged (committed). §3 Cell-fit (target=edge, `jc=distribute`, no buffer, OOXML write for flat paths, tatweel-only idempotency) → Task 3 emitter + Task 5/6 branches. §4 Natural-fit (retarget `colPx`→matrix `target`, caps kept, `reach`=container=`colPx−buffer`) → Task 5/6 Natural-fit branches. §5 mode toggle (pane control, default Natural-fit, stored on profile, both paths branch) → Tasks 2 + 4. §6/§8 → Tasks 1–3 node tests + Task 7 manual.
- **Deferred (documented non-goals, per spec §7 + resolved decision 3):** **Per-cell variances / under-resolved-cell override editing** is explicitly tied to the parked "pane reflects the active Ashaar block's settings" UI-sync feature; without that editing surface a stored variance schema has no consumer (YAGNI), so it is left to the follow-on spec. **Spacing-cell styling** (colors, fill glyphs, `*` markers) is its own spec; this plan only introduces the `isContentCell` tagging it will build on. **Explicit matrix-driven free-form column resizing** is satisfied by the existing auto-fit + the qaseeda per-column equalisation (Task 6 design note) rather than a new resize pass.
- **Type consistency:** `AshaarMatrix.naturalFitTarget(Wpos, reach, phi)` and `cellFitBudget(natural, colPx, phi)` signatures identical across Tasks 1/5/6; `positionKey({row,col,span})` identical in Tasks 1/5/6; `misraDistributeXml(runs, sizePtFallback)` with `runs=[{text,csName,sizePt?}]` identical in Tasks 3/5/6; `normalizeFillMode` returns exactly `"cell-fit"|"natural-fit"` consumed by Tasks 4/6; `opts.fillMode` produced in Task 4, consumed in Task 5.
- **Reused unchanged:** `justifyRunsConcentrated`, `strengthToElongationShare`, `strengthToMaxPositions`, `capMicroSpaces`, `injectSpaceRuns`, `distributeMicroSpaces`, `selectSwapRuns`, `runsToMisraXml`, `wrapOoxml`, `stripJustification`, `contentPx`, `cellAlignOf`.
