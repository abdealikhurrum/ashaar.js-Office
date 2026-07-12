# Per-Cell Overrides + Active-Block Sync (SP2) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** When the cursor enters an Ashaar Poem block, the pane reflects that block's settings; a sub-panel lets the cell at the cursor (by its SP1 label, e.g. `A1`) deviate on strength / target width / cap-lift, stored per-cell on the block tag and consumed at justify time.

**Architecture:** A new pure module `cell-overrides.js` (`AshaarOverrides`) formats the per-cell key and merges an override onto the block's base justify settings. `word-html.js` gains `setTagOverride` (write/remove one override in the content-control tag payload, mirroring `setTagQaseeda`). `justifySelection`/`applyProfileToQaseeda` read `tag.overrides` and apply the resolved per-cell strength/target/cap. `taskpane.js` registers a debounced `DocumentSelectionChanged` handler that (a) resyncs the pane's block controls when the active block changes and (b) shows a per-cell override editor for the content cell at the cursor; editing writes the tag and re-justifies the block via the existing `justifySelection()`.

**Tech Stack:** Vanilla JS (ES5/UMD, `var`/`function`, no build step), Office.js v1, Node `assert` tests. Never edit `src/vendor/`.

**Spec:** `docs/superpowers/specs/2026-07-11-per-cell-overrides-design.md`
**Builds on (implemented):** SP1 bandh cell-map (`AshaarCellMap`, tag `cells`, per-cell `__matKey` label in justify) and the Cell-fit/Natural-fit justification work.

## Global Constraints

- ES5/UMD only (`var`, `function`); never edit `src/vendor/`.
- Pure logic (`AshaarOverrides`, `setTagOverride`) is **node-tested**; all Office.js glue (selection-changed handler, cell/tableIndex detection, block sync, editor, justify consumption) is **manual-verify** (Task 6). The §6a `TableCell → tableIndex` detection cannot be exercised headless — it is validated in Task 6 and has a documented label-only fallback.
- Override payload: `{ strength?: 1–10, widthPt?: number, capEm?: number }`; absent field = inherit. Key = `"<tableIndex>:<label>"`.
- `capEm` applies only in Natural-fit; `widthPt` applies only in Natural-fit; `strength` applies in both. `fillMode` is NOT overridable.
- Edit feedback = write tag, then call the existing `justifySelection()` (whole-block re-justify) — no bespoke single-cell path.
- Block controls resync only when the **active block changes** (track last-reflected block), never on cursor moves within the same block.
- `npm test` green after every task.

---

### Task 1: Pure override helpers (`AshaarOverrides` + `setTagOverride`)

**Files:**
- Create: `src/taskpane/cell-overrides.js`
- Modify: `src/taskpane/word-html.js` (`setTagOverride` + export)
- Test: `tests/cell-overrides.test.js`
- Modify: `package.json` (test script), `src/taskpane/taskpane.html` (script load list)

**Interfaces:**
- `AshaarOverrides.overrideKey(tableIndex, label) → "<ti>:<label>"`.
- `AshaarOverrides.resolveCellOverride(base, override) → { strength, fillMode, widthPt, capEm }` — `base` is the block justify `{ strength, fillMode }`; `override` (or null) supplies deviations. Returns effective values: `strength` from override else base; `fillMode` always from base (not overridable); `widthPt` = override's or `null`; `capEm` = override's or `null` (null → caller uses the 0.28 default).
- `AshaarWord.setTagOverride(tag, key, override) → tag` — returns a copy of the `ashaar:` tag with `payload.overrides[key]` set (non-empty override) or deleted (null/empty override). Non-ashaar tags returned unchanged.

- [ ] **Step 1: Write the failing test**

Create `tests/cell-overrides.test.js`:

```js
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
  const tag0 = AshaarWord.contentControlTag("poem", { qaseeda: "Q", tableWidthPct: 60 }, [[["c","g","c"]]]);
  const tag1 = AshaarWord.setTagOverride(tag0, "0:A1", { strength: 9 });
  const p1 = AshaarWord.parseContentControlTag(tag1);
  assert.deepStrictEqual(p1.overrides, { "0:A1": { strength: 9 } }, "override added");
  assert.equal(p1.qaseeda, "Q", "other payload fields intact");
  assert.deepStrictEqual(p1.cells, [[["c","g","c"]]], "cells intact");

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
```

- [ ] **Step 2: Run to verify it fails**

Run: `node tests/cell-overrides.test.js`
Expected: FAIL — `Cannot find module '../src/taskpane/cell-overrides'`.

- [ ] **Step 3: Create `cell-overrides.js`**

```js
/**
 * AshaarOverrides — per-cell justify override key + merge (pure, UMD;
 * node-testable, no DOM/Office). An override deviates one cell from its block's
 * justify defaults on strength / target width / cap-lift. fillMode is NOT
 * overridable (block-wide choice). See
 * docs/superpowers/specs/2026-07-11-per-cell-overrides-design.md.
 */
(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) module.exports = factory();
  else root.AshaarOverrides = factory();
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  // Per-cell key: the SP1 label prefixed by the cell's table (bandh) index in
  // the block, so an override targets exactly one cell.
  function overrideKey(tableIndex, label) {
    return (tableIndex || 0) + ":" + label;
  }

  // Merge a cell override onto the block's base justify settings. Absent fields
  // inherit. fillMode always comes from base. widthPt/capEm are null when unset
  // (the caller then uses the matrix target / the 0.28em default).
  function resolveCellOverride(base, override) {
    base = base || {};
    override = override || {};
    var s = (override.strength != null) ? override.strength : base.strength;
    return {
      strength: s,
      fillMode: base.fillMode,
      widthPt: (override.widthPt != null) ? override.widthPt : null,
      capEm: (override.capEm != null) ? override.capEm : null
    };
  }

  return { overrideKey: overrideKey, resolveCellOverride: resolveCellOverride };
}));
```

- [ ] **Step 4: Add `setTagOverride` to `word-html.js`**

Immediately after `setTagQaseeda` (ends ~line 939), add:
```js
  // Return a copy of an "ashaar:" tag with one per-cell override set or removed.
  // A null/empty override deletes the key. Non-ashaar tags returned unchanged.
  function setTagOverride(tag, key, override) {
    var payload = parseContentControlTag(tag);
    if (!payload) return tag;
    var ov = payload.overrides && typeof payload.overrides === "object" ? payload.overrides : {};
    var has = override && typeof override === "object" &&
      (override.strength != null || override.widthPt != null || override.capEm != null);
    if (has) {
      var clean = {};
      if (override.strength != null) clean.strength = override.strength;
      if (override.widthPt != null) clean.widthPt = override.widthPt;
      if (override.capEm != null) clean.capEm = override.capEm;
      ov[key] = clean;
    } else {
      delete ov[key];
    }
    payload.overrides = ov;
    return "ashaar:" + encodeURIComponent(JSON.stringify(payload));
  }
```
And in `parseContentControlTag`, normalize `overrides` alongside `cells` — change:
```js
      payload.cells = payload.cells || null;
      return payload;
```
to:
```js
      payload.cells = payload.cells || null;
      payload.overrides = (payload.overrides && typeof payload.overrides === "object") ? payload.overrides : {};
      return payload;
```
Add to the exports (near `setTagQaseeda: setTagQaseeda,`):
```js
    setTagOverride: setTagOverride,
```

- [ ] **Step 5: Run to verify it passes**

Run: `node tests/cell-overrides.test.js`
Expected: PASS — `cell-overrides.test.js OK`.

- [ ] **Step 6: Register test + module**

`package.json` test script — append `&& node tests/cell-overrides.test.js`.
`src/taskpane/taskpane.html` — add `"./cell-overrides.js",` right after `"./bandh-cell-map.js",`.

- [ ] **Step 7: Full suite + commit**

Run: `npm test`; Expected: green.
```bash
git add src/taskpane/cell-overrides.js src/taskpane/word-html.js tests/cell-overrides.test.js package.json src/taskpane/taskpane.html
git commit --no-gpg-sign -m "feat(overrides): AshaarOverrides merge + setTagOverride tag writer"
```

---

### Task 2: Justify consumes per-cell overrides

**Files:**
- Modify: `src/taskpane/taskpane.js` (`justifySelection` + `applyProfileToQaseeda`)

**Interfaces:** Consumes `AshaarOverrides.overrideKey/resolveCellOverride`, `parseContentControlTag(...).overrides`. Manual-verify (Office glue); the merge math is node-tested (Task 1).

**Design:** Each content cell already carries its `(tableIndex,label)` via SP1. Look up `overrides[overrideKey(ti,label)]`, merge with `resolveCellOverride`, and let it drive that cell's φ / target / cap.

- [ ] **Step 1: `justifySelection` — read overrides + per-cell table index**

`justifySelection` already parses `ccPayload` and sets `ccCells`. Right after `ccCells = ccPayload && ccPayload.cells;`, add:
```js
        var ccOverrides = (ccPayload && ccPayload.overrides) || {};
```
In the `allCells` build loop, the per-table index `ti` is already in scope (`tables.items.forEach(function (tbl, ti)`). Where each cell gets `cell.__matKey`/`cell.__kind` (the `if (mapped) { … }` block), also stamp the override key and resolved override for content cells:
```js
            if (mapped) {
              cell.__kind = mapped.kind;
              cell.__matKey = mapped.label || mapped.slot;
              cell.__ovKey = mapped.kind === "content" && mapped.label
                ? AshaarOverrides.overrideKey(ti, mapped.label) : null;
            } else {
              cell.__kind = null;
              cell.__matKey = AshaarMatrix.positionKey({ row: ri, col: ci, span: cols });
              cell.__ovKey = null;
            }
```
- [ ] **Step 2: `justifySelection` — apply the override in the generic Natural-fit branch**

The block's base justify = `{ strength: opts.tatweelCount, fillMode: fillMode }`. In the phase-1 loop, resolve per cell near the top (after `var colPx = contentPx(cell);`), then use it. Add:
```js
        var cellOv = cell.__ovKey ? ccOverrides[cell.__ovKey] : null;
        var resolved = AshaarOverrides.resolveCellOverride({ strength: opts.tatweelCount, fillMode: fillMode }, cellOv);
        var cellPhi = AshaarWord.strengthToElongationShare(resolved.strength);
        var cellMaxPos = AshaarWord.strengthToMaxPositions(resolved.strength);
```
Then, in the generic Natural-fit branch, replace the uses of `elongShare`/`gMax`/target with the resolved ones:
- `var gMax = { perPositionEm: 0.5, maxPositions: cellMaxPos };`
- Cell-fit budget: `AshaarMatrix.cellFitBudget(gNatural, colPx, cellPhi)`.
- Natural-fit target: if `resolved.widthPt != null` use `resolved.widthPt * 96 / 72` else `AshaarMatrix.naturalFitTarget(gWpos, gReach, cellPhi)`.
- Natural-fit residual cap: `AshaarResidual.capMicroSpaces(gTarget - conc.achievedPx, gGaps, gSpacePx, repSize * 96 / 72, resolved.capEm != null ? resolved.capEm : undefined)`.
Apply the same `cellPhi`/`cellMaxPos`/`resolved.widthPt`/`resolved.capEm` substitutions in the **Jameel** and **Mehr** branches (they use `elongShare` and a `maxPositions` and a `capMicroSpaces` call each).

- [ ] **Step 3: `applyProfileToQaseeda` — same, per block**

Where `blockCells` is built, also capture each block's overrides. After `var blockCells = blocks.map(...)`, add:
```js
        var blockOverrides = blocks.map(function (cc) {
          var p = AshaarWord.parseContentControlTag(cc.tag);
          return (p && p.overrides) || {};
        });
        var allTableOverrides = [];
        blockTables.forEach(function (t, bi) {
          t.items.forEach(function () { allTableOverrides.push(blockOverrides[bi] || {}); });
        });
```
In the `tableInfos` capture, stamp `ovKey` on each cell (the table index within its block is `mapped.row`? no — it's the bandh/table index). Track it: the `allTablePatterns[ai]` is per flattened table; use the flattened `ai` for the pattern but the **block-local** table index for the key. Simplest: derive the block-local index while flattening — extend the earlier flatten to also push it:
```js
        var allTableBlockIdx = [];
        blockTables.forEach(function (t) {
          t.items.forEach(function (tbl, j) { allTableBlockIdx.push(j); });
        });
```
Then in the cell capture, set:
```js
                ovKey: (mapped && mapped.kind === "content" && mapped.label)
                  ? AshaarOverrides.overrideKey(allTableBlockIdx[ai], mapped.label) : null,
```
and in the re-justify loop, resolve + apply (mirror Step 2) using `allTableOverrides[ai][c.ovKey]` — thread `ai` into the loop (it maps `tableInfos`↔`allTables` 1:1, so iterate with index). Use `resolveCellOverride({ strength: strength, fillMode: fillMode }, ov)` and drive `cMax.maxPositions`, the Cell-fit/Natural-fit target, and the `capMicroSpaces` cap exactly as in Step 2.

- [ ] **Step 4: Syntax + suite**

Run: `node --check src/taskpane/taskpane.js`; Expected: no output.
Run: `npm test`; Expected: green (glue; merge math already covered).

- [ ] **Step 5: Commit**

```bash
git add src/taskpane/taskpane.js
git commit --no-gpg-sign -m "feat(overrides): justify applies per-cell strength/width/cap overrides"
```

---

### Task 3: Active-block sync (selection-changed handler)

**Files:**
- Modify: `src/taskpane/taskpane.js` (`bind()` + a `reflectActiveContext` handler + a `syncBlockControls` helper)

**Interfaces:** Consumes `parseContentControlTag`, `profileToPanel`, `AshaarProfiles.normalizeFillMode`. Manual-verify.

- [ ] **Step 1: Add the handler + helpers**

Add near `getQaseedaAtSelection` (block-at-cursor domain):
```js
  var _lastBlockTag = null;      // last-reflected block tag (resync only on change)
  var _reflectPending = false;   // debounce guard
  var _reflectBusy = false;      // suppress while our own justify runs

  // Populate the pane's block-level controls from a parsed tag payload.
  function syncBlockControls(payload) {
    if (!payload) return;
    if (payload.fontMode) fontMode.value = payload.fontMode;
    if (payload.justifyMode) justifyMode.value = payload.justifyMode;
    if (justifyFillMode && payload.fillMode) justifyFillMode.value = AshaarProfiles.normalizeFillMode(payload.fillMode);
    if (payload.tatweelCount != null) { tatweelCount.value = payload.tatweelCount; if (tatweelValue) tatweelValue.textContent = String(payload.tatweelCount); }
    if (payload.tableWidthPct != null && tableWidth) { tableWidth.value = payload.tableWidthPct; if (tableWidthValue) tableWidthValue.textContent = String(payload.tableWidthPct); }
    if (payload.qaseeda) { var st = loadProfileStore()[payload.qaseeda]; if (st) profileToPanel(st); }
  }

  // Reflect the Ashaar block (and, in Task 5, the cell) at the cursor.
  async function reflectActiveContext() {
    if (typeof Word === "undefined" || _reflectBusy) return;
    try {
      await Word.run(async function (context) {
        var sel = context.document.getSelection();
        var cc = sel.parentContentControlOrNullObject;
        cc.load("title,tag");
        await context.sync();
        var isBlock = !cc.isNullObject && cc.title === "Ashaar Poem";
        var payload = isBlock ? AshaarWord.parseContentControlTag(cc.tag) : null;
        // Resync block controls only when the active block changes.
        if (isBlock && cc.tag !== _lastBlockTag) { _lastBlockTag = cc.tag; syncBlockControls(payload); }
        if (!isBlock) _lastBlockTag = null;
        // (Task 5 extends this with the per-cell editor.)
      });
    } catch (e) { /* selection transient — ignore */ }
  }

  // Debounced entry point for the selection-changed event.
  function onSelectionChanged() {
    if (_reflectPending) return;
    _reflectPending = true;
    window.setTimeout(function () { _reflectPending = false; reflectActiveContext(); }, 150);
  }
```

- [ ] **Step 2: Register the handler in `bind()`**

In `bind()`, guard on Word and register:
```js
    if (typeof Office !== "undefined" && Office.context && Office.context.document &&
        Office.context.document.addHandlerAsync && typeof Word !== "undefined") {
      Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChanged);
    }
```

- [ ] **Step 3: Suppress reflection during our own justify**

`justifySelection` mutates the selection (which fires the event). At the very top of `justifySelection` set `_reflectBusy = true;` and in a `finally` (or before each early return and at the end) reset `_reflectBusy = false;`. Simplest: wrap the existing body so `_reflectBusy` is true for the duration. Add `_reflectBusy = true;` as the first line and `_reflectBusy = false;` immediately before every `return` path is heavy — instead set it true at entry and reset it in the outer `withWord(...)`/`try` completion. Concretely, set `_reflectBusy = true;` first, and add `_reflectBusy = false;` as the last line of the function (after the `await withWord(...)` resolves) and in the `catch`/early-return branches that exist. (Manual-verify the flag is cleared on all paths in Task 6.)

- [ ] **Step 4: Syntax + suite + commit**

Run: `node --check src/taskpane/taskpane.js`; `npm test`; Expected: green.
```bash
git add src/taskpane/taskpane.js
git commit --no-gpg-sign -m "feat(overrides): active-block sync via DocumentSelectionChanged"
```

---

### Task 4: Per-cell override editor + cell detection

**Files:**
- Modify: `src/taskpane/taskpane.html` (editor sub-panel), `src/taskpane/taskpane.css` (styles), `src/taskpane/taskpane.js` (cell detection in `reflectActiveContext`, editor populate/write handlers, `bind()` wiring)

**Interfaces:** Consumes `AshaarCellMap.buildBandhCellMap`, `AshaarOverrides.overrideKey`, `AshaarWord.setTagOverride`. Manual-verify. **Contains the §6a `TableCell → tableIndex` detection — the plan's one real risk; validate first in Task 6.**

- [ ] **Step 1: Add the editor markup**

In `taskpane.html`, after the `cell-map-view` div (from SP1), add:
```html
        <div id="cell-override" class="cell-override" hidden>
          <div class="cell-override-head">Cell <span id="cell-override-label"></span></div>
          <label>Strength <input id="cell-ov-strength" type="number" min="1" max="10" placeholder="inherit"></label>
          <label>Target width (pt) <input id="cell-ov-width" type="number" min="1" placeholder="inherit"></label>
          <label>Cap lift (em) <input id="cell-ov-cap" type="number" min="0" step="0.01" placeholder="0.28"></label>
          <button id="cell-ov-clear" type="button" class="button--secondary">Clear cell override</button>
        </div>
```

- [ ] **Step 2: Detect the cell + table index in `reflectActiveContext`**

Inside the `Word.run` of `reflectActiveContext`, after the block block, add (only when `isBlock`):
```js
        var editor = document.getElementById("cell-override");
        var activeOvKey = null, activeLabel = null;
        if (isBlock && payload && payload.cells) {
          var selRange = sel.getRange();
          var tbls = cc.getRange().tables;
          tbls.load("items");
          var tcell = sel.parentTableCellOrNullObject;
          tcell.load("rowIndex,cellIndex,isNullObject");
          await context.sync();
          if (!tcell.isNullObject) {
            // §6a: which block table contains the selection?
            var tIdx = -1, inters = tbls.items.map(function (tbl) {
              var r = tbl.getRange().intersectWithOrNullObject(selRange); r.load("isNullObject"); return r;
            });
            await context.sync();
            for (var k = 0; k < inters.length; k++) { if (!inters[k].isNullObject) { tIdx = k; break; } }
            if (tIdx >= 0 && payload.cells[tIdx]) {
              var map = AshaarCellMap.buildBandhCellMap(payload.cells[tIdx]);
              var inRow = map.filter(function (e) { return e.row === tcell.rowIndex; });
              var entry = inRow[tcell.cellIndex];
              if (entry && entry.kind === "content") {
                activeLabel = entry.label;
                activeOvKey = AshaarOverrides.overrideKey(tIdx, entry.label);
                _activeBlockOverrides = payload.overrides || {};
                _activeOvKey = activeOvKey;
              }
            }
          }
        }
        // Show/populate or hide the editor (DOM work is safe inside the queue).
        if (editor) {
          if (activeOvKey) { populateCellEditor(activeLabel, (payload.overrides || {})[activeOvKey]); editor.hidden = false; }
          else { editor.hidden = true; _activeOvKey = null; }
        }
```
Add module-scope vars near `_lastBlockTag`: `var _activeOvKey = null; var _activeBlockOverrides = {};`.

- [ ] **Step 3: Editor populate + write handlers**

Add:
```js
  function populateCellEditor(label, ov) {
    ov = ov || {};
    document.getElementById("cell-override-label").textContent = label || "";
    document.getElementById("cell-ov-strength").value = (ov.strength != null) ? ov.strength : "";
    document.getElementById("cell-ov-width").value = (ov.widthPt != null) ? ov.widthPt : "";
    document.getElementById("cell-ov-cap").value = (ov.capEm != null) ? ov.capEm : "";
  }

  function readCellEditor() {
    function num(id) { var v = document.getElementById(id).value; return v === "" ? null : Number(v); }
    var ov = {};
    var s = num("cell-ov-strength"); if (s != null) ov.strength = Math.max(1, Math.min(10, s));
    var w = num("cell-ov-width"); if (w != null) ov.widthPt = w;
    var c = num("cell-ov-cap"); if (c != null) ov.capEm = c;
    return ov;
  }

  // Write the current editor state to the active cell's override on the block
  // tag, then re-justify the whole block for instant feedback.
  async function applyCellOverride(clear) {
    if (!_activeOvKey || typeof Word === "undefined") return;
    var ov = clear ? null : readCellEditor();
    _reflectBusy = true;
    try {
      await Word.run(async function (context) {
        var cc = context.document.getSelection().parentContentControlOrNullObject;
        cc.load("title,tag");
        await context.sync();
        if (cc.isNullObject || cc.title !== "Ashaar Poem") return;
        cc.tag = AshaarWord.setTagOverride(cc.tag, _activeOvKey, ov);
        await context.sync();
        _lastBlockTag = cc.tag;
      });
    } catch (e) { /* ignore */ } finally { _reflectBusy = false; }
    if (clear) populateCellEditor(document.getElementById("cell-override-label").textContent, null);
    await justifySelection(); // instant feedback via the existing path
  }
```

- [ ] **Step 4: Wire the editor in `bind()`**

```js
    ["cell-ov-strength", "cell-ov-width", "cell-ov-cap"].forEach(function (id) {
      var el = document.getElementById(id);
      if (el) el.addEventListener("change", function () { applyCellOverride(false); });
    });
    var ovClear = document.getElementById("cell-ov-clear");
    if (ovClear) ovClear.addEventListener("click", function () { applyCellOverride(true); });
```

- [ ] **Step 5: Styles**

Append to `taskpane.css`:
```css
.cell-override { margin-top: 8px; padding: 6px; border: 1px solid #ccc; border-radius: 4px; font-size: 12px; }
.cell-override label { display: block; margin: 4px 0; }
.cell-override input { width: 90px; }
.cell-override-head { font-weight: 600; margin-bottom: 4px; }
```

- [ ] **Step 6: Syntax + suite + commit**

Run: `node --check src/taskpane/taskpane.js`; `npm test`; Expected: green.
```bash
git add src/taskpane/taskpane.html src/taskpane/taskpane.css src/taskpane/taskpane.js
git commit --no-gpg-sign -m "feat(overrides): per-cell override editor + cell/tableIndex detection"
```

---

### Task 5: Manual Word verification (incl. §6a spike)

**Files:** none. `npm start`.

- [ ] **Step 1 — §6a first (the risk).** Insert a multi-bandh poem. Click into several different cells across different bandhs. Confirm the editor header shows the correct label (`A1`, `A2`, `B1`, …) and that clicking the *same* label in *different* bandhs yields **different** keys (verify the applied override lands only on the clicked cell in Task step 4). If the table-index is wrong/flaky across your Word host → STOP and report; the fallback is label-only keying (needs sign-off).
- [ ] **Step 2 — block sync.** Click into a block → pane's font/justify/fill/strength/width reflect that block. Click into a *different* block (different settings) → pane resyncs. Move the cursor within one block while mid-editing a control → controls are NOT reset.
- [ ] **Step 3 — per-cell override.** In a Natural-fit poem, select a cell, set **Strength 10** → only that cell elongates more on the auto re-justify; siblings unchanged. Set **Target width** → that cell fills to it. Set **Cap lift** (e.g. 0.6) on an under-resolved cell → it reaches its harmony width. **Clear** → reverts to the block default.
- [ ] **Step 4 — scope.** Confirm an override on bandh 2's `A1` does NOT change bandh 1's `A1`.
- [ ] **Step 5 — gaps / no-map.** A spacing/gap cell shows no editor. An adopted (no `cells`) block shows block sync but no per-cell editor.
- [ ] **Step 6 — no compounding / clean flag.** Repeatedly move the cursor and edit; confirm no runaway re-justify loops (the `_reflectBusy` guard holds) and re-justify is idempotent.

## Self-Review notes

- **Spec coverage:** §1 data model → Task 1 (`resolveCellOverride`, `overrideKey`, `setTagOverride`, tag `overrides`); §2 cursor detection → Task 3 (block) + Task 4 (cell/tableIndex, §6a); §3 block sync → Task 3 (`syncBlockControls`, resync-on-change); §4 editor → Task 4; §5 justify consumes → Task 2; §6 feasibility → Task 4 code + Task 5 Step 1 validation; §7 testing → node Task 1 + manual Task 5; §8 non-goals untouched.
- **Type consistency:** override `{strength,widthPt,capEm}` and `resolveCellOverride` return `{strength,fillMode,widthPt,capEm}` identical across Tasks 1/2/4; key `"<ti>:<label>"` from `overrideKey` used in Tasks 2/4; `setTagOverride(tag,key,override)` in Tasks 1/4.
- **Risk:** Task 4 §6a (TableCell→tableIndex) is the one unproven piece — isolated to the detection block, validated Task 5 Step 1, with a documented label-only fallback. `capEm`/`widthPt` are Natural-fit-only (Cell-fit ignores them) — consistent with the spec.
- **Reused unchanged:** `justifySelection` as the re-justify path, `AshaarMatrix`/`AshaarCellMap`, `capMicroSpaces` (its existing `capEm` 5th arg), `profileToPanel`, `setTagQaseeda` (template for `setTagOverride`).
- **Deferred:** per-cell style overrides, special-gap purposes, manual load-from-cursor fallback — backlog.
```
