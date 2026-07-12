# Unified Formatting Model + Span-Safe Width Engine — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make Save & Apply size and harmonize span-based poetry tables correctly (equal per-position widths, no wrap) in both auto-fit and fixed-% modes, and drive all formatting from one block-first profile model shared by insert, adopt, justify, and apply.

**Architecture:** A pure module computes the target column grid (per-position widths, equalized across same-shape bandhs, sized per auto-fit/fixed-%) from canvas measurements; the apply engine sets those widths on the table (in-place `setWidth` if a Word spike confirms it works on span tables, else a `stanzaTableOoxml` rebuild) and reuses the *same computed widths* as the kashida fill target so text cannot overflow. The two duplicated control sets collapse into one profile panel; every code path reads the profile instead of the top controls.

**Tech Stack:** Vanilla ES5/UMD JS, Office.js v1 (WordApi 1.3 / WordApiDesktop 1.3–1.4), Node `assert` tests (no framework), no build step. Dev server: `npm run dev-server` (port 3000, `Cache-Control: no-store`); add-in launch: `npm start`.

## Global Constraints

- Never edit `src/vendor/` (synced from the submodule).
- All library modules use the UMD wrapper (node + browser). Pure modules must have **no DOM/Office dependency** so they are node-testable.
- Widths are computed and set — **never read back** `Table.columns` or `TableCell.columnWidth` (both are "uniform tables only" and throw / mislead on span tables).
- `shadingColor` clear value is `"#FFFFFF"` (build rejects `""` and `"No color"`). See `~/.claude/.../memory/office-js-word-constraints.md`.
- Office.js assignment errors surface at `context.sync()`, not the assignment — never rely on a `try/catch` around a property set.
- Commit messages end with the repo's Co-Authored-By + Claude-Session trailers; commit with `--no-gpg-sign` (signing agent locked this session).
- Poetry tables from `stanzaTableOoxml` are fixed-layout: `<w:tblLayout w:type="fixed"/>`, `<w:tblW dxa>`, uniform `<w:gridCol w:w="cwt"/>`. `cwt = round(textWidthTwips / GRID)`; a cell of span S has width `S·cwt`. Twips↔px: `px = twips · 96/1440`; points↔px: `px = pt · 96/72`.
- Spec: `docs/superpowers/specs/2026-07-11-unified-formatting-width-engine-design.md`.

---

## File Structure

- `src/taskpane/natural-width-matrix.js` (`AshaarMatrix`) — **extend**: add `computeTargetGrid()` (pure grid math). Where the harmony/auto-fit/fixed-% width vector is computed.
- `tests/natural-width-matrix.test.js` — **extend**: `computeTargetGrid` cases.
- `src/taskpane/profiles.js` (`AshaarProfiles`) — **extend**: `DEFAULT_PROFILE_NAME`, `resolveProfileName(tag)` (fallback to Default). Profile schema already carries `width`/`justify`.
- `tests/profiles.test.js` — **extend**: default-name resolution.
- `src/taskpane/word-html.js` (`AshaarWord`) — **extend**: `stanzaGridTwips(source, opts, Ashaar, pageTwips)` helper exposing the computed `{GRID, cwt, colTwips}` per stanza for the rebuild path (reuses existing `stanzaGridInfo`).
- `src/taskpane/taskpane.js` — **modify**: rewrite width step of `applyProfileToQaseeda`; add `setTableGridWidths()` (spike-selected mechanism); route `insertPoem`/`adoptTable`/`justifySelection` through the active profile; profile-dropdown wiring; remove top-control reads from `options()` formatting fields.
- `src/taskpane/taskpane.html` — **modify**: remove top width/justify/fill/strength controls; add Profile dropdown + "＋ New profile"; keep structural controls (layout spec, misra/bandh counts, font-for-drawing).
- Spike: throwaway handler in `taskpane.js` behind a temporary button (removed before final commit) — Task 0.

---

## Phase 0 — Spike (gates the width-set mechanism)

### Task 0: Word spike — does `setWidth` work on a span table?

**Files:**
- Modify (temporary): `src/taskpane/taskpane.html` (one temp button), `src/taskpane/taskpane.js` (one temp handler). **Reverted at end of task.**

**Interfaces:**
- Produces: a decision recorded in the plan — `MECHANISM = "setWidth"` or `MECHANISM = "rebuild"` — consumed by Task 3.

- [ ] **Step 1: Add a temporary spike button + handler**

In `taskpane.html`, add near the justify buttons: `<button id="spike-setwidth" type="button">SPIKE setWidth</button>`.

In `taskpane.js` `bind()`, wire it:

```js
document.getElementById("spike-setwidth").addEventListener("click", async function () {
  try {
    await Word.run(async function (context) {
      var sel = context.document.getSelection();
      var t = sel.parentTableOrNullObject;
      t.load("isNullObject");
      await context.sync();
      if (t.isNullObject) { setMessage("Put cursor in a poem table first."); return; }
      // Try the tolerant, method-based width set (no .columns load).
      t.columns.setWidth(60, "SameWidth"); // 60 pt per column
      await context.sync();
      setMessage("SPIKE OK: setWidth accepted on this table.");
    });
  } catch (e) { setMessage("SPIKE FAIL: " + describeError(e)); }
});
```

- [ ] **Step 2: Run and observe**

`npm start`, adopt a marsiya bandh so it is a span-based poem table, click into it, click **SPIKE setWidth**.
Record which happens:
- Pane shows **"SPIKE OK"** and the columns visibly change → set `MECHANISM = "setWidth"`.
- Pane shows **"SPIKE FAIL: … mixed cell widths"** (or any GeneralException/InvalidArgument) → set `MECHANISM = "rebuild"`.

- [ ] **Step 3: Record the decision in this plan**

Edit the line below to the observed value, then remove the temp button + handler from `taskpane.html`/`taskpane.js`.

> **SPIKE RESULT (2026-07-11):** `MECHANISM = setWidth` — `TableColumnCollection.setWidth(pt, "SameWidth")` was accepted on a span-based marsiya table and changed the columns. Phase 2 uses the in-place `setWidth` path; **Task 4b (rebuild) is skipped**, and `stanzaGridTwips` is not needed.

- [ ] **Step 4: Commit the removal (no spike code ships)**

```bash
git add src/taskpane/taskpane.html src/taskpane/taskpane.js docs/superpowers/plans/2026-07-11-unified-formatting-width-engine.md
git commit -m "chore(spike): determine span-table setWidth viability (no ship)"
```

---

## Phase 1 — Pure width-grid math (`AshaarMatrix.computeTargetGrid`)

### Task 1: `computeTargetGrid` — auto-fit + fixed-% + harmony

**Files:**
- Modify: `src/taskpane/natural-width-matrix.js`
- Test: `tests/natural-width-matrix.test.js`

**Interfaces:**
- Consumes: existing `buildMatrix(cells)`.
- Produces: `computeTargetGrid(bandhs, opts)` where
  - `bandhs`: `[{ GRID, cells: [{ key, natural, col, span }] }]` — one per block, all with the same `GRID` (grid column count). `key` is `positionKey`, `natural` is the tatweel-free width in **px**, `col`/`span` are grid coordinates.
  - `opts`: `{ mode: "auto-fit"|"fixed", pct, pagePx, headroom }` — `headroom` is the auto-fit kashida margin fraction (e.g. `0.18`), `pagePx` the usable page text width in px.
  - Returns: `{ sameShape, colPx: [Number×GRID], bandhTargets: [{ [key]: targetPx }] }`
    - `colPx`: the shared per-grid-column width vector (px). Same for every same-shape bandh → harmony.
    - `bandhTargets`: per block, each content position's fill target px (= sum of the `colPx` its cell spans, i.e. the box the text must fill).

- [ ] **Step 1: Write the failing test**

Append to `tests/natural-width-matrix.test.js`:

```js
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `node tests/natural-width-matrix.test.js`
Expected: FAIL — `AshaarMatrix.computeTargetGrid is not a function`.

- [ ] **Step 3: Implement `computeTargetGrid`**

Insert before the `return {` block in `src/taskpane/natural-width-matrix.js`:

```js
  // Compute the shared per-grid-column width vector (px) for a set of same-shape
  // bandhs plus each bandh's per-position fill target.
  //   auto-fit: each position width = longest natural × (1 + headroom); total capped at pagePx.
  //   fixed:    scale so total = (pct/100) × pagePx, proportions from the matrix.
  // colPx[j] is the width of grid column j; a position's target = sum of its spanned colPx.
  function computeTargetGrid(bandhs, opts) {
    bandhs = bandhs || [];
    opts = opts || {};
    var pagePx = Number(opts.pagePx) || 0;
    var headroom = Number(opts.headroom) || 0;

    // sameShape: every bandh shares one GRID.
    var GRID = bandhs.length ? Number(bandhs[0].GRID) || 0 : 0;
    var sameShape = GRID > 0 && bandhs.every(function (b) { return Number(b.GRID) === GRID; });

    // Cross-bandh matrix: longest natural per position (px).
    var flat = [];
    bandhs.forEach(function (b) { (b.cells || []).forEach(function (c) { flat.push(c); }); });
    var matrix = buildMatrix(flat);

    // Per-position width need = longest natural × (1 + headroom).
    // Distribute a position's need evenly across the grid columns it spans; a
    // column's width is the max need imposed by any position covering it.
    var need = {};
    Object.keys(matrix).forEach(function (k) { need[k] = matrix[k] * (1 + headroom); });

    var colPx = [];
    for (var j = 0; j < GRID; j++) colPx.push(0);
    // Use the first bandh's positions to map key → (col, span) (same-shape share layout).
    var layout = {};
    (bandhs[0] && bandhs[0].cells || []).forEach(function (c) { layout[c.key] = { col: c.col, span: c.span }; });
    Object.keys(need).forEach(function (k) {
      var L = layout[k]; if (!L || !L.span) return;
      var per = need[k] / L.span;
      for (var j = L.col; j < L.col + L.span && j < GRID; j++) colPx[j] = Math.max(colPx[j], per);
    });

    var total = colPx.reduce(function (a, b) { return a + b; }, 0);
    if (opts.mode === "fixed") {
      var want = (Number(opts.pct) || 100) / 100 * pagePx;
      if (total > 0 && want > 0) { var kf = want / total; colPx = colPx.map(function (w) { return w * kf; }); }
    } else if (pagePx > 0 && total > pagePx) {
      var ka = pagePx / total; colPx = colPx.map(function (w) { return w * ka; });
    }

    // Each bandh's per-position target = sum of spanned colPx (the box to fill).
    var bandhTargets = bandhs.map(function (b) {
      var t = {};
      (b.cells || []).forEach(function (c) {
        var sum = 0;
        for (var j = c.col; j < c.col + c.span && j < GRID; j++) sum += colPx[j];
        t[c.key] = sum;
      });
      return t;
    });

    return { sameShape: sameShape, colPx: colPx, bandhTargets: bandhTargets };
  }
```

Add `computeTargetGrid: computeTargetGrid,` to the returned object.

- [ ] **Step 4: Run test to verify it passes**

Run: `node tests/natural-width-matrix.test.js`
Expected: `natural-width-matrix.test.js OK` and `computeTargetGrid OK`.

- [ ] **Step 5: Commit**

```bash
git add src/taskpane/natural-width-matrix.js tests/natural-width-matrix.test.js
git commit -m "feat(matrix): computeTargetGrid — harmony/auto-fit/fixed-% column vector"
```

### Task 2: Different-shape fallback (equalize on total)

**Files:**
- Modify: `src/taskpane/natural-width-matrix.js`
- Test: `tests/natural-width-matrix.test.js`

**Interfaces:**
- Consumes/Produces: same `computeTargetGrid` signature; when `sameShape === false`, `colPx` is `null` and `bandhTargets` still returns per-bandh targets computed from each bandh's own grid, but every bandh is scaled to one shared **total** width.

- [ ] **Step 1: Write the failing test**

Append:

```js
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
  // each bandh's targets present, summing to the same total
  const t0 = Object.values(g.bandhTargets[0]).reduce((a, b) => a + b, 0);
  const t1 = Object.values(g.bandhTargets[1]).reduce((a, b) => a + b, 0);
  assert.ok(Math.abs(t0 - t1) < 1, "different-shape bandhs share one total width");
}
console.log("computeTargetGrid different-shape OK");
```

- [ ] **Step 2: Run test to verify it fails**

Run: `node tests/natural-width-matrix.test.js`
Expected: FAIL — `sameShape` currently derived but `colPx` not null / totals differ.

- [ ] **Step 3: Implement the branch**

Replace the single-vector body from Task 1 with a `sameShape` guard. When `!sameShape`, compute each bandh's own `colPx` vector (same per-position logic, over that bandh's own cells), then scale every bandh so its total equals `min(pagePx, max total across bandhs)` (auto-fit) or `pct×pagePx` (fixed); return `colPx: null` and per-bandh targets. Keep the same-shape path returning the shared vector. Full code:

```js
  function computeTargetGrid(bandhs, opts) {
    bandhs = bandhs || [];
    opts = opts || {};
    var pagePx = Number(opts.pagePx) || 0;
    var headroom = Number(opts.headroom) || 0;
    var mode = opts.mode === "fixed" ? "fixed" : "auto-fit";
    var pctWant = (Number(opts.pct) || 100) / 100 * pagePx;

    var GRID = bandhs.length ? Number(bandhs[0].GRID) || 0 : 0;
    var sameShape = GRID > 0 && bandhs.every(function (b) { return Number(b.GRID) === GRID; });

    // Build a colPx vector for one bandh from its cells (need spread over spans).
    function vectorFor(cells, grid, matrix) {
      var v = []; for (var j = 0; j < grid; j++) v.push(0);
      cells.forEach(function (c) {
        var w = ((matrix ? matrix[c.key] : c.natural) || 0) * (1 + headroom);
        var per = w / c.span;
        for (var j = c.col; j < c.col + c.span && j < grid; j++) v[j] = Math.max(v[j], per);
      });
      return v;
    }
    function scaleTo(v, targetTotal) {
      var tot = v.reduce(function (a, b) { return a + b; }, 0);
      if (tot <= 0 || targetTotal <= 0) return v;
      var k = targetTotal / tot; return v.map(function (w) { return w * k; });
    }
    function targetsFrom(cells, v, grid) {
      var t = {};
      cells.forEach(function (c) {
        var sum = 0; for (var j = c.col; j < c.col + c.span && j < grid; j++) sum += v[j];
        t[c.key] = sum;
      });
      return t;
    }

    if (sameShape) {
      var flat = []; bandhs.forEach(function (b) { (b.cells || []).forEach(function (c) { flat.push(c); }); });
      var matrix = buildMatrix(flat);
      var colPx = vectorFor(bandhs[0].cells || [], GRID, matrix);
      // Positions absent from bandh[0] but present elsewhere still influence cols via matrix:
      Object.keys(matrix).forEach(function (k) {
        var any = null;
        bandhs.some(function (b) { return (b.cells || []).some(function (c) { if (c.key === k) { any = c; return true; } }); });
        if (!any) return;
        var per = matrix[k] * (1 + headroom) / any.span;
        for (var j = any.col; j < any.col + any.span && j < GRID; j++) colPx[j] = Math.max(colPx[j], per);
      });
      var total = colPx.reduce(function (a, b) { return a + b; }, 0);
      if (mode === "fixed") colPx = scaleTo(colPx, pctWant);
      else if (pagePx > 0 && total > pagePx) colPx = scaleTo(colPx, pagePx);
      var bandhTargets = bandhs.map(function (b) { return targetsFrom(b.cells || [], colPx, GRID); });
      return { sameShape: true, colPx: colPx, bandhTargets: bandhTargets };
    }

    // Different shapes: per-bandh vectors, all scaled to one shared total.
    var vecs = bandhs.map(function (b) { return vectorFor(b.cells || [], Number(b.GRID) || 0, null); });
    var totals = vecs.map(function (v) { return v.reduce(function (a, b) { return a + b; }, 0); });
    var shared = mode === "fixed" ? pctWant : Math.min(pagePx || Infinity, Math.max.apply(null, totals.concat([0])));
    var scaled = vecs.map(function (v) { return scaleTo(v, shared); });
    var bt = bandhs.map(function (b, i) { return targetsFrom(b.cells || [], scaled[i], Number(b.GRID) || 0); });
    return { sameShape: false, colPx: null, bandhTargets: bt };
  }
```

- [ ] **Step 4: Run tests to verify all pass**

Run: `node tests/natural-width-matrix.test.js`
Expected: all `OK` lines print, no assertion errors.

- [ ] **Step 5: Commit**

```bash
git add src/taskpane/natural-width-matrix.js tests/natural-width-matrix.test.js
git commit -m "feat(matrix): different-shape fallback equalizes on total width"
```

---

## Phase 2 — Apply engine rewrite

### Task 3: `setTableGridWidths` helper (spike-selected mechanism)

**Files:**
- Modify: `src/taskpane/taskpane.js` (add helper near `applyProfileToQaseeda`)
- Modify: `src/taskpane/word-html.js` (add `stanzaGridTwips` if `MECHANISM === "rebuild"`)

**Interfaces:**
- Produces: `async function setTableGridWidths(context, table, colTwips)` — sets grid column widths (twips) on `table`. Uses `MECHANISM` from Task 0.
- Consumes (rebuild path only): `AshaarWord.stanzaGridTwips`.

- [ ] **Step 1: Implement the width-set helper**

If `MECHANISM === "setWidth"`, add to `taskpane.js`:

```js
// Set a fixed-layout table's per-grid-column widths (twips) WITHOUT reading
// Table.columns width (which throws on span tables). setWidth is method-based
// and tolerant. colTwips length must equal the table's grid column count.
async function setTableGridWidths(context, table, colTwips) {
  var cols = table.columns;
  cols.load("items");            // items count only — NOT width
  await context.sync();
  var n = Math.min(cols.items.length, colTwips.length);
  for (var j = 0; j < n; j++) {
    cols.items[j].setWidth(colTwips[j] / 20, "None"); // twips → points
  }
  await context.sync();
}
```

If `MECHANISM === "rebuild"`, instead the width is applied by regenerating the table (Task 4 handles rebuild inline); `setTableGridWidths` is a no-op stub that returns `false` to signal "caller must rebuild":

```js
async function setTableGridWidths() { return false; } // rebuild path: see applyProfileToQaseeda
```

- [ ] **Step 2: (rebuild path only) add `stanzaGridTwips` to word-html.js**

Only if `MECHANISM === "rebuild"`. Add near `stanzaTableOoxml`:

```js
  // Expose the computed grid for a stanza (GRID column count + per-column twips)
  // so callers can rebuild a table at a target width. Mirrors stanzaTableOoxml.
  function stanzaGridTwips(stanza, opts, textWidthTwips) {
    var si = stanzaGridInfo(stanza, opts, textWidthTwips);
    return { GRID: si.GRID, cwt: si.cwt, totalW: si.GRID * si.cwt };
  }
```

Export it in the returned object: `stanzaGridTwips: stanzaGridTwips,`.

- [ ] **Step 3: Sanity-parse**

Run: `node --check src/taskpane/taskpane.js && node --check src/taskpane/word-html.js`
Expected: both print nothing (exit 0). (No node unit test — Word behavior is covered by Task 11 manual verification.)

- [ ] **Step 4: Commit**

```bash
git add src/taskpane/taskpane.js src/taskpane/word-html.js
git commit -m "feat(apply): setTableGridWidths — span-safe width set (spike-selected)"
```

### Task 4: Rewrite the width step of `applyProfileToQaseeda`

**Files:**
- Modify: `src/taskpane/taskpane.js` — the `canResize` block inside `applyProfileToQaseeda` (currently ~lines 855–912) and the per-cell `colPx` source.

**Interfaces:**
- Consumes: `AshaarMatrix.computeTargetGrid`, `setTableGridWidths`, existing `tableInfos` capture, `section.pageLayout`.
- Produces: after this task, `tableInfos[i]` carries `c.targetPx` (the computed fill box per content cell) and `info.colTwipsSet` (the grid applied), used by Task 5's justify.

- [ ] **Step 1: Replace the `canResize`/columns block**

Delete the entire `var canResize = …` block through its trailing `await context.sync();` (the one that builds `colColls` from `info.tbl.columns` and sets `col.width`). Replace with grid computation + span-safe set:

```js
        // Build the per-bandh cell descriptors for the matrix (content cells only).
        var pagePx = pagePt * 96 / 72;
        var HEADROOM = doKashida ? 0.18 : 0.06; // kashida needs more room to stretch
        var bandhsForGrid = tableInfos.map(function (info) {
          var cells = [];
          info.cells.forEach(function (c) {
            if (!AshaarMatrix.isContentCell(c.measure)) return;
            var fnm = c.fontName || repName, fsz = c.fontSize || repSize;
            canvasCtx.font = fsz + "pt \"" + fnm + "\"";
            c.natPx = AshaarProfiles.applyFontCorrection(canvasCtx.measureText(c.measure).width, fnm, profile.fontCorrections);
            cells.push({ key: c.matKey, natural: c.natPx, col: c.gridCol, span: c.gridSpan });
          });
          return { GRID: info.grid, cells: cells };
        });
        var grid = AshaarMatrix.computeTargetGrid(bandhsForGrid, {
          mode: profile.width.mode === "fixed" ? "fixed" : "auto-fit",
          pct: profile.width.pct, pagePx: pagePx, headroom: HEADROOM
        });

        // Stamp each content cell's fill-box target (px) from the grid result.
        tableInfos.forEach(function (info, bi) {
          var tg = grid.bandhTargets[bi] || {};
          info.cells.forEach(function (c) { if (c.natPx != null) c.targetPx = tg[c.matKey] || c.natPx; });
        });

        // Apply widths to each table (in-place setWidth, or rebuild fallback).
        for (var ti = 0; ti < tableInfos.length; ti++) {
          var info = tableInfos[ti];
          var vec = grid.sameShape ? grid.colPx : gridVectorForBandh(grid, bandhsForGrid[ti]); // see note
          var colTwips = vec.map(function (px) { return Math.round(px * 1440 / 96); });
          var ok = await setTableGridWidths(context, info.tbl, colTwips);
          if (ok === false && typeof MECHANISM_REBUILD !== "undefined") {
            // rebuild path — Task 4b
          }
          info.colTwipsSet = colTwips;
        }
```

Note: for `sameShape === false`, add a tiny local helper `gridVectorForBandh(grid, bandh)` that recomputes that bandh's own vector — or simpler, have `computeTargetGrid` also return `perBandhColPx` in the different-shape case. **Decision for the implementer:** extend `computeTargetGrid`'s different-shape return with `perBandhColPx: [ [twips…] per bandh ]` (add to Task 2 return + a test) so the caller needs no helper.

- [ ] **Step 2: Capture grid coords + grid count during `tableInfos` build**

Where `tableInfos` is built (the `cells.push({ … })`), add `gridCol`, `gridSpan`, and set `info.grid`. The bandh cell-map already knows spans; derive `gridCol`/`gridSpan` from the row's cumulative spans. Add to the cell push:

```js
                gridCol: runningCol,      // sum of spans before this cell in its row
                gridSpan: cols === 1 ? 1 : (spanOf ? spanOf : 1),
```

and track `runningCol` per row (reset to 0 each row, `runningCol += gridSpan`). Set `info.grid` = max over rows of total span. (Reuse `AshaarCellMap`/`perRowCounts` already computed; if spans aren't directly available, compute from `columnWidth` ratios captured pre-resize, or from the cell-map pattern `c`/`g` counts.)

- [ ] **Step 3: Sanity-parse + node suite**

Run: `node --check src/taskpane/taskpane.js && npm test`
Expected: parse OK; all 16 suite lines pass (pure modules unaffected).

- [ ] **Step 4: Commit**

```bash
git add src/taskpane/taskpane.js src/taskpane/natural-width-matrix.js tests/natural-width-matrix.test.js
git commit -m "feat(apply): compute + set span-safe grid widths, stamp per-cell targets"
```

### Task 4b: Rebuild fallback (only if `MECHANISM === "rebuild"`)

**Files:**
- Modify: `src/taskpane/taskpane.js`

**Interfaces:**
- Consumes: `AshaarWord.stanzaGridTwips`, `AshaarTableAdopt.adoptTableToSource`, `AshaarWord.renderForWordOoxml`, `AshaarWord.wrapOoxml`.

- [ ] **Step 1: Implement rebuild in the width loop**

If the spike chose rebuild: in the width loop of Task 4, instead of `setTableGridWidths`, reconstruct each block's source, render at `totalW = sum(colTwips)`, and replace the block **in place using the proven content-control replace** from `insertPoem` (Task 6 exposes `replacePoemInPlace(context, cc, ooxml, tag)`). Then re-capture cell proxies (cells were recreated — reuse the Task-earlier re-map pattern). Full code depends on Task 6's helper; sequence Task 6 before 4b if rebuild is chosen.

- [ ] **Step 2: Sanity-parse**

Run: `node --check src/taskpane/taskpane.js`
Expected: exit 0.

- [ ] **Step 3: Commit**

```bash
git add src/taskpane/taskpane.js
git commit -m "feat(apply): rebuild-at-width fallback for span tables"
```

### Task 5: Justify to the computed target (both fill + both justify modes)

**Files:**
- Modify: `src/taskpane/taskpane.js` — the per-cell justify block in `applyProfileToQaseeda` (the `if (doKashida && colPx > 0)` region).

**Interfaces:**
- Consumes: `c.targetPx` (Task 4), `AshaarMatrix.naturalFitTarget`/`cellFitBudget`, existing `AshaarJustify.justifyRunsConcentrated`, `AshaarResidual.capMicroSpaces`, `AshaarWord.distributeMicroSpaces`, `AshaarWord.misraDistributeXml`.

- [ ] **Step 1: Use `c.targetPx` as the fill box; ungate elongation from kashida-only**

Replace the `var colPx = …` derivation and the `if (doKashida && colPx > 0)` gate so that:
- `colPx` (the true box) = `c.targetPx` when set, else the previous margin-adjusted fallback.
- The elongation runs when `justify mode` is kashida **or** spacing (both fill to the box); only `"none"`/`"css"` skip. i.e. `var doFill = profile.justify.mode === "kashida" || profile.justify.mode === "spacing";` and gate on `doFill && colPx > 0`.
- Kashida mode uses tatweels (existing `justifyRunsConcentrated`); Spacing mode uses micro-spaces only (set `cPhi`/elongation share to 0 so the fill is space-driven — reuse the existing `distributeMicroSpaces` path with `cN` covering `target − natural`).
- Natural-fit: `cReach = colPx`; `cTarget = naturalFitTarget(c.targetPx, cReach, cPhi)` (with `c.targetPx` already the harmony box, `reach == target` so it fills exactly the box).
- Cell-fit: `cTarget = cellFitBudget(cNatural, colPx, cPhi)` unchanged, with `colPx = c.targetPx`.
- **Invariant:** never let the computed `cTarget` exceed `colPx` (the box) — clamp `cTarget = Math.min(cTarget, colPx)` — this is the no-wrap guarantee.

- [ ] **Step 2: Add a node test for the no-wrap invariant on the pure math**

Append to `tests/natural-width-matrix.test.js`:

```js
// no-wrap invariant: naturalFitTarget(box, box, phi) never exceeds box
[0, 0.5, 1].forEach(function (phi) {
  assert.ok(AshaarMatrix.naturalFitTarget(300, 300, phi) <= 300 + 1e-9, "target ≤ box at phi=" + phi);
});
console.log("no-wrap invariant OK");
```

- [ ] **Step 3: Run tests + parse**

Run: `node tests/natural-width-matrix.test.js && node --check src/taskpane/taskpane.js`
Expected: `no-wrap invariant OK`; parse exit 0.

- [ ] **Step 4: Commit**

```bash
git add src/taskpane/taskpane.js tests/natural-width-matrix.test.js
git commit -m "feat(apply): fill to computed box; enable fill under kashida+spacing; clamp no-wrap"
```

---

## Phase 3 — Route insert/adopt/justify through the profile

### Task 6: Insert/adopt/justify read the active profile; size via the width engine

**Files:**
- Modify: `src/taskpane/taskpane.js` — `insertPoem`, `adoptTable`, `justifySelection`/`reRender` option sourcing.

**Interfaces:**
- Consumes: `getProfile(name)`, active-block profile from tag (`AshaarWord.parseContentControlTag`), `AshaarProfiles`.
- Produces: `activeProfile()` — returns the profile for the block at the cursor, else the currently-selected dropdown profile, else Default. Used everywhere formatting is needed.

- [ ] **Step 1: Add `activeProfile()`**

```js
// The profile driving the current action: the block at the cursor's profile if
// we're in one, else the dropdown selection, else Default.
function activeProfile() {
  var name = (qaseedaSelect && qaseedaSelect.value) || AshaarProfiles.DEFAULT_PROFILE_NAME;
  return getProfile(name);
}
```

- [ ] **Step 2: Source formatting from the profile, not the top controls**

In `insertPoem`/`adoptTable`, replace reads of `opts.justifyMode`, `opts.fillMode`, `opts.tableWidthPct`, `opts.autoFitWidth`, `opts.gapWidth`, `opts.tatweelCount` with the corresponding `activeProfile()` fields (`justify.mode`, `justify.fillMode`, `width.pct`, `width.mode`, `gap`, `justify.strength`). Keep structural fields (`layoutSpec`, `misraCount`, `bandhCount`, `fontMode`) from `options()`.

- [ ] **Step 3: Size a freshly-inserted block via the same width engine**

After `insertPoem` creates the block, call `applyProfileToQaseeda(activeProfile().name)` (or factor the width+justify core into `formatBlocks(context, blocks, profile)` and call it directly) so a new poem is sized/justified identically to an applied one. DRY: extract the Phase-2 core into `formatBlocks` and have both `applyProfileToQaseeda` and insert call it.

- [ ] **Step 4: Parse + node suite**

Run: `node --check src/taskpane/taskpane.js && npm test`
Expected: parse OK; suite green.

- [ ] **Step 5: Commit**

```bash
git add src/taskpane/taskpane.js
git commit -m "feat(format): route insert/adopt/justify through the active profile"
```

---

## Phase 4 — UI unification

### Task 7: Profile dropdown + "＋ New profile"

**Files:**
- Modify: `src/taskpane/taskpane.html` (qaseeda panel), `src/taskpane/taskpane.js` (wiring).

- [ ] **Step 1: Replace the qaseeda name input with a `<select>` + add button**

In `taskpane.html`, swap the `qaseeda-name` input/datalist for:

```html
<label for="qaseeda-select">Profile</label>
<select id="qaseeda-select"></select>
<button id="qaseeda-new" type="button" class="button--secondary">＋ New profile</button>
```

- [ ] **Step 2: Populate + wire**

`populateQaseedaNames()` fills `#qaseeda-select` with `Object.keys(loadProfileStore())` plus Default. `#qaseeda-new` prompts for a name, `putProfile(defaultProfile(name))`, repopulates, selects it. Changing `#qaseeda-select` loads that profile into the panel via `profileToPanel(getProfile(value))`.

- [ ] **Step 3: Manual smoke (Word)**

`npm start`; confirm dropdown lists profiles, "＋ New profile" adds one, selecting one loads its settings. (UI wiring — no node test.)

- [ ] **Step 4: Commit**

```bash
git add src/taskpane/taskpane.html src/taskpane/taskpane.js
git commit -m "feat(ui): profile dropdown with inline New profile"
```

### Task 8: Block-first active-context sync

**Files:**
- Modify: `src/taskpane/taskpane.js` — `reflectActiveContext`.

- [ ] **Step 1: On selection change, select the block's profile in the dropdown**

Extend `reflectActiveContext` (already reads the active block's tag) so it sets `#qaseeda-select.value` to the block's `qaseeda` (or Default) and calls `profileToPanel`. When not in a block, leave the dropdown as-is (insert defaults).

- [ ] **Step 2: Manual smoke (Word)**

Click between two blocks with different profiles; confirm the dropdown + panel follow the cursor.

- [ ] **Step 3: Commit**

```bash
git add src/taskpane/taskpane.js
git commit -m "feat(ui): block-first active-context selects the block's profile"
```

### Task 9: Remove the duplicated top formatting controls

**Files:**
- Modify: `src/taskpane/taskpane.html` (remove `#justify-mode`, `#justify-fill-mode`, `#table-width`, `#width-mode`, `#auto-fit*`, `#tatweel*`, `#gap-width` from the *formatting* area), `src/taskpane/taskpane.js` (drop those reads from `options()`; keep structural).

- [ ] **Step 1: Remove the controls from HTML**

Delete the formatting inputs listed above from the top panel. Keep layout-spec, misra/bandh counts, font mode (structural).

- [ ] **Step 2: Trim `options()`**

Remove `justifyMode`, `justify`, `fillMode`, `widthMode`, `tableWidthPct`, `autoFitWidth`, `gapWidth`, `tatweelCount` from `options()`. Grep for each removed field and repoint remaining consumers to `activeProfile()` (Task 6 handled the main ones; fix any stragglers the grep finds).

- [ ] **Step 3: Grep for orphans + parse + suite**

Run: `grep -n "justifyFillMode\|tableWidthPct\|autoFitWidth\|\.tatweelCount\|gapWidth" src/taskpane/taskpane.js` → resolve each. Then `node --check src/taskpane/taskpane.js && npm test`.
Expected: no orphan references to removed controls; parse OK; suite green.

- [ ] **Step 4: Commit**

```bash
git add src/taskpane/taskpane.html src/taskpane/taskpane.js
git commit -m "refactor(ui): remove duplicated top formatting controls; profile is sole source"
```

---

## Phase 5 — Migration & fallback

### Task 10: Default profile + older-Word fallback message

**Files:**
- Modify: `src/taskpane/profiles.js` (+ test), `src/taskpane/taskpane.js`.

**Interfaces:**
- Produces: `AshaarProfiles.DEFAULT_PROFILE_NAME = "Default"`; `AshaarProfiles.resolveProfileName(tagQaseeda)` → the name or Default.

- [ ] **Step 1: Write the failing test**

Append to `tests/profiles.test.js`:

```js
assert.strictEqual(AshaarProfiles.DEFAULT_PROFILE_NAME, "Default");
assert.strictEqual(AshaarProfiles.resolveProfileName(""), "Default");
assert.strictEqual(AshaarProfiles.resolveProfileName("Karbala"), "Karbala");
console.log("profile default-name OK");
```

- [ ] **Step 2: Run → fail**

Run: `node tests/profiles.test.js`
Expected: FAIL — `DEFAULT_PROFILE_NAME` undefined.

- [ ] **Step 3: Implement**

In `profiles.js`: `var DEFAULT_PROFILE_NAME = "Default";` and `function resolveProfileName(q){ return (q && String(q).trim()) || DEFAULT_PROFILE_NAME; }`; export both. In `taskpane.js`, `gatherQaseedaBlocks` and `activeProfile` use `resolveProfileName`.

- [ ] **Step 4: Fallback message when width can't be set**

In the width loop (Task 4), if `setTableGridWidths` couldn't run (no WordApiDesktop 1.3 and rebuild off), set a flag; in the final `summary`, append `" (widths unchanged — this Word build can't resize; text fills current cells)"`. Because the kashida target is `c.targetPx` clamped to the box, it still won't wrap.

- [ ] **Step 5: Run tests + parse**

Run: `node tests/profiles.test.js && node --check src/taskpane/taskpane.js`
Expected: `profile default-name OK`; parse exit 0.

- [ ] **Step 6: Commit**

```bash
git add src/taskpane/profiles.js tests/profiles.test.js src/taskpane/taskpane.js
git commit -m "feat(profiles): Default profile + resolveProfileName; apply width fallback message"
```

---

## Phase 6 — Manual Word verification (Task 7 successor)

### Task 11: Manual verification checklist

**Files:** none (verification only). `npm start` opens `test-documents/marsiya-test.docx`.

- [ ] **Step 1: Setup** — Restore/adopt the two marsiya bandhs as two managed blocks; assign both to one profile via the dropdown.

- [ ] **Step 2: Auto-fit harmony** — Profile Width = Auto-fit, Justify = Kashida, Fill = Natural-fit, strength 5, **Apply**. Confirm: matching-position cells across both bandhs are **equal width**, **no word-wrap**, red tatweels present.

- [ ] **Step 3: Fixed %** — Width = Fixed %, set 60 then 85, Apply each. Confirm the block visibly resizes to the % and both bandhs stay equal; no wrap.

- [ ] **Step 4: Strength sweep** — 1 → 10. At 1 cells sit near natural widths; at 10 filled to the box edge. No wrap at any step.

- [ ] **Step 5: Fill mode under both justify modes** — Toggle Fill = Cell-fit vs Natural-fit under Kashida, then under Spacing. Confirm Fill mode **visibly changes** the result in **both** justify modes (no silent no-op).

- [ ] **Step 6: Per-block override** — Set a per-cell override on one cell; Apply; confirm it layers on top of the profile.

- [ ] **Step 7: Idempotent re-apply** — Apply twice at fixed settings; confirm no compounding (tatweel/space counts stable).

- [ ] **Step 8: Fallback (if reproducible)** — On a Word without WordApiDesktop 1.3, confirm the fallback message shows and text still doesn't wrap.

- [ ] **Step 9: Record results** — Note pass/fail per step in a short handoff doc; file any residual issues as follow-ups (re-render/multi-table wrap, adopt-review checkbox remain known-open).

---

## Self-Review notes

- **Spec coverage:** §1 unified model → Tasks 6–9; §2 width engine → Tasks 1–5 (+0 spike); §3 folded fixes (fill under both modes → Task 5; fixed-% → Tasks 1/4; older-Word fallback → Task 10); §4 data flow/migration → Tasks 6, 10; §5 testing → node tests in Tasks 1,2,5,10 + spike Task 0 + manual Task 11.
- **Mechanism branch:** Tasks 3/4b are gated on the Task 0 `MECHANISM` result — exactly one path ships. If rebuild, sequence Task 6 (for `replacePoemInPlace`) before Task 4b.
- **Open decision surfaced to implementer:** Task 4 Step 1 note — add `perBandhColPx` to `computeTargetGrid`'s different-shape return (with a test) rather than a caller-side helper, to keep the width math in one place.
- **Known-open (out of scope):** re-render / multi-table Adopt content-control re-wrap; Adopt "Review before replacing" no-op.
