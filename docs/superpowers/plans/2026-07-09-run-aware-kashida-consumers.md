# Run-Aware Kashida Consumers Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Wire the Word add-in and the web preview to the multi-run kashida engine so a misra with mixed weight/size justifies correctly, preserving each run's styling on write-back.

**Architecture:** Both consumers discover an ordered array of styled runs from their own rendered content, call `AshaarJustify.justifyRuns` (kashida) / `AshaarJustify.computeRunSpacing` (spacing/scale), and map the same-length result back run-by-run. Word discovers runs by splitting a cell into word-ranges and coalescing by style, then replaces each run's range in place (fonts preserved); the web walks the misra span's child nodes and writes each child's text plus a single `word-spacing`/per-child scale.

**Tech Stack:** Vanilla JS (ES5/UMD, no build step), Office.js v1, Node `assert` tests, Ashaar.js submodule.

## Global Constraints

- No build step, no transpilation — ES5-compatible UMD only.
- Never edit `src/vendor/` directly. The web change is authored in `vendor/ashaar-js/ashaar.js`, then `npm run sync:ashaar` copies it in; tests run against the synced `src/vendor/` copy.
- Tests are pure Node (`node tests/<file>.test.js`), no jest/mocha/jsdom. Office.js and DOM code is NOT node-testable — it is verified manually.
- Justification must stay idempotent/reducible: strip tatweels + micro-spaces (`stripJustification`) before every pass so results never compound.
- Full suite: `npm test`.

**What is NOT node-verifiable (manual only):** every `Word.run` call in `taskpane.js` (needs Word) and every DOM read/write in `ashaar.js` (`getComputedStyle`, `getBoundingClientRect`, child-node text) (needs a browser). Plan isolates pure logic into helpers so those ARE tested; the glue is reviewed and manually verified.

---

# Part A — Word add-in consumer (this repo)

Adds pure helpers to `word-html.js` (tested) and rewrites the per-cell apply loop of `justifySelection` in `taskpane.js` (manual verify).

### Task A1: `coalesceRuns` — merge adjacent same-style words into runs

**Files:**
- Modify: `src/taskpane/word-html.js` (add function + export)
- Test: `tests/word-html.test.js` (append cases)

**Interfaces:**
- Produces: `AshaarWord.coalesceRuns(words)` where `words = [{ text, name, size, bold, italic }]` (per Office.js word-range read). Returns `[{ text, name, size, bold, italic }]`; adjacent words with an identical `(name, size, bold, italic)` tuple merge, their `text` joined by a single space. Empty input → `[]`.

- [ ] **Step 1: Write the failing test** — append to `tests/word-html.test.js`:

```js
// ── coalesceRuns ────────────────────────────────────────────────────────────
{
  // Two words, same style → one run.
  const r = AshaarWord.coalesceRuns([
    { text: "درد", name: "Amiri", size: 16, bold: false, italic: false },
    { text: "دل",  name: "Amiri", size: 16, bold: false, italic: false },
  ]);
  assert.deepEqual(r, [{ text: "درد دل", name: "Amiri", size: 16, bold: false, italic: false }]);
}
{
  // Style change (bold) splits into two runs, order preserved.
  const r = AshaarWord.coalesceRuns([
    { text: "درد", name: "Amiri", size: 16, bold: false, italic: false },
    { text: "دل",  name: "Amiri", size: 16, bold: true,  italic: false },
  ]);
  assert.equal(r.length, 2);
  assert.equal(r[0].text, "درد");
  assert.equal(r[1].text, "دل");
  assert.equal(r[1].bold, true);
}
{
  // Size change splits; empty input → [].
  const r = AshaarWord.coalesceRuns([
    { text: "الف", name: "Amiri", size: 24, bold: false, italic: false },
    { text: "ب",   name: "Amiri", size: 16, bold: false, italic: false },
    { text: "ج",   name: "Amiri", size: 16, bold: false, italic: false },
  ]);
  assert.deepEqual(r.map(x => x.text), ["الف", "ب ج"]);
  assert.deepEqual(AshaarWord.coalesceRuns([]), []);
}
```

- [ ] **Step 2: Run to verify it fails**

Run: `node tests/word-html.test.js`
Expected: FAIL — `AshaarWord.coalesceRuns is not a function`.

- [ ] **Step 3: Implement** — add near the other justify helpers in `src/taskpane/word-html.js` (before the exports block), and add `coalesceRuns: coalesceRuns,` to the returned object:

```js
// Merge adjacent words that share an identical style tuple into runs. A run's
// text is its words joined by a single space. Word-aligned (exact per the
// foundation spec); mid-word style changes are not represented here — the
// caller reads one style per whole word.
function coalesceRuns(words) {
  var runs = [];
  (words || []).forEach(function (w) {
    var prev = runs[runs.length - 1];
    if (prev && prev.name === w.name && prev.size === w.size &&
        prev.bold === w.bold && prev.italic === w.italic) {
      prev.text += " " + w.text;
    } else {
      runs.push({ text: w.text, name: w.name, size: w.size, bold: w.bold, italic: w.italic });
    }
  });
  return runs;
}
```

- [ ] **Step 4: Run to verify it passes**

Run: `node tests/word-html.test.js`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/taskpane/word-html.js tests/word-html.test.js
git commit -m "feat(justify): coalesceRuns — group same-style words into runs"
```

### Task A2: `distributeMicroSpaces` — realize wordSpacing as micro-space glyphs

**Files:**
- Modify: `src/taskpane/word-html.js` (add function + export)
- Test: `tests/word-html.test.js`

**Interfaces:**
- Produces: `AshaarWord.distributeMicroSpaces(runTexts, n, spaceChar)`. `runTexts` = `string[]` (each run's words joined by `" "`). Distributes `n` copies of `spaceChar` across all **intra-run** word gaps (the spaces already inside each run's text), round-robin so no gap gets two before every gap has one, keeping each inserted glyph inside its own run. Returns a new same-length `string[]`. `n <= 0` or no gaps → input unchanged. Inter-run delimiter spaces are outside any run's text and are intentionally not stretched (documented limitation).

- [ ] **Step 1: Write the failing test**:

```js
// ── distributeMicroSpaces ───────────────────────────────────────────────────
{
  const HAIR = " ";
  // Two runs; gaps: run0 has 1 ("a b"), run1 has 1 ("c d") → 2 gaps. n=2 → 1 each.
  const out = AshaarWord.distributeMicroSpaces(["a b", "c d"], 2, HAIR);
  assert.deepEqual(out, ["a " + HAIR + "b", "c " + HAIR + "d"]);
}
{
  const HAIR = " ";
  // n=3 over 2 gaps → first gap gets 2, second gets 1 (round-robin).
  const out = AshaarWord.distributeMicroSpaces(["a b", "c d"], 3, HAIR);
  assert.deepEqual(out, ["a " + HAIR + HAIR + "b", "c " + HAIR + "d"]);
}
{
  const HAIR = " ";
  // No gaps (single words) or n<=0 → unchanged.
  assert.deepEqual(AshaarWord.distributeMicroSpaces(["a", "b"], 5, HAIR), ["a", "b"]);
  assert.deepEqual(AshaarWord.distributeMicroSpaces(["a b"], 0, HAIR), ["a b"]);
}
```

- [ ] **Step 2: Run to verify it fails** — Run: `node tests/word-html.test.js`; Expected: FAIL — not a function.

- [ ] **Step 3: Implement** (add function + `distributeMicroSpaces: distributeMicroSpaces,` export):

```js
// Insert n micro-space glyphs across the intra-run word gaps, round-robin so
// distribution stays even. Each glyph is placed at a gap that belongs to one
// run, so it stays within that run's range on write-back. Inter-run gaps (the
// split delimiters) are outside every run's text and are left unstretched.
function distributeMicroSpaces(runTexts, n, spaceChar) {
  var texts = (runTexts || []).map(String);
  if (n <= 0) return texts;
  // Enumerate global gap slots: { ri, gapIndex } for each intra-run gap.
  var slots = [];
  texts.forEach(function (t, ri) {
    var gaps = t.split(" ").length - 1;
    for (var g = 0; g < gaps; g++) slots.push({ ri: ri, gap: g });
  });
  if (!slots.length) return texts;
  var counts = {}; // "ri:gap" -> extra glyphs
  for (var i = 0; i < n; i++) {
    var s = slots[i % slots.length];
    var key = s.ri + ":" + s.gap;
    counts[key] = (counts[key] || 0) + 1;
  }
  return texts.map(function (t, ri) {
    var parts = t.split(" ");
    var out = parts[0] || "";
    for (var g = 0; g < parts.length - 1; g++) {
      var extra = counts[ri + ":" + g] || 0;
      out += " " + new Array(extra + 1).join(spaceChar) + parts[g + 1];
    }
    return out;
  });
}
```

- [ ] **Step 4: Run to verify it passes** — Run: `node tests/word-html.test.js`; Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/taskpane/word-html.js tests/word-html.test.js
git commit -m "feat(justify): distributeMicroSpaces — even micro-space fill across run gaps"
```

### Task A3: run-aware apply loop in `justifySelection` (Office.js glue — manual verify)

**Files:**
- Modify: `src/taskpane/taskpane.js` — the read of cell fonts (`~1120-1130`) and the apply loop (`~1244-1284`).

**Interfaces:**
- Consumes: `AshaarWord.coalesceRuns`, `AshaarWord.distributeMicroSpaces`, `AshaarJustify.justifyRuns`, `AshaarJustify.computeRunSpacing`.

This task has no node test (Office.js). Verified manually (Task A4). Steps:

- [ ] **Step 1: Read per-word styling.** After the existing `row.cells.load("items/columnWidth")` sync, for each cell get word ranges and load their fonts:

```js
// Per cell: split into word ranges so we can read a font per word.
allCells.forEach(function (cell) {
  cell.__wordRanges = cell.body.getRange().getTextRanges([" "], true);
  cell.__wordRanges.load("items");
});
await context.sync();
allCells.forEach(function (cell) {
  cell.__wordRanges.items.forEach(function (wr) {
    wr.load("text");
    wr.font.load("name,size,bold,italic");
  });
});
await context.sync();
```

- [ ] **Step 2: Build runs per cell** (replace the single `cell.body.font` read). For each cell, map its `__wordRanges.items` to `{ text, name, size, bold, italic }` (stripping justification from each word's text), drop empty words, then `AshaarWord.coalesceRuns(words)`. Keep the parallel array of the underlying ranges so run *i* maps to the union range `words[start].expandTo(words[end])`. Skip cells with no visible text.

- [ ] **Step 3: Kashida apply.** Build primitive runs — for each coalesced run set `measure` from a canvas ctx whose `font` is `run.size + "pt \"" + run.name + (run.bold ? "\" bold" : "\"")` (weight via the ctx font shorthand), and `fontProfile` = the poem-level profile (unchanged), `fontSize = run.size`. Call:

```js
var out = AshaarJustify.justifyRuns(primRuns, colPx, calibParams); // same length/order
```

Then write each run's justified text into its union range: `unionRange.insertText(out[i].text, Word.InsertLocation.replace)`. Never touch the paragraph, so `jc`/spacing survive.

- [ ] **Step 4: Spacing apply.** Instead of `justifyRuns`:

```js
var sp = AshaarJustify.computeRunSpacing(primRuns, colPx, calibParams); // {wordSpacing, fontScale}
```

Apply `fontScale`: for each run union range, `range.font.size = run.size * sp.fontScale` (skip when `sp.fontScale === 1`). Realize `wordSpacing`: total extra px = `sp.wordSpacing * totalIntraRunGaps`; `n = Math.max(0, Math.round(extra / spaceGlyphPx))` where `spaceGlyphPx = canvasCtx.measureText(" ").width` (fall back to `" "` if 0); `var spacedTexts = AshaarWord.distributeMicroSpaces(runTexts, n, spaceChar)`; write each `spacedTexts[i]` back into its union range.

- [ ] **Step 5: Fallback flag.** Wrap Steps 3–4 write-back in a try; on an Office.js range error, log to `diags` and fall back to the existing whole-cell `insertText` on the flattened result for that cell (single-font B2-lite), so a range-replace failure never aborts the batch.

- [ ] **Step 6: Preserve debug diagnostics.** Keep the existing `debug`/`diags` block, computing per-cell `nat`/`fin`/`fill` from the summed run measures rather than a single `measureText`.

- [ ] **Step 7: Commit**

```bash
git add src/taskpane/taskpane.js
git commit -m "feat(justify): run-aware cell justification in justifySelection"
```

### Task A4: manual verification in Word

**Files:** none (uses `test-documents/`).

- [ ] **Step 1:** Add a mixed-style stanza to the marsiya generator (`scripts/make-test-doc.mjs`) or hand-edit a test doc: one misra with a **bold refrain word** and one with a **larger first word**.
- [ ] **Step 2:** `npm start`, click inside the Ashaar Poem content control, choose Kashida, click Justify. Confirm: bold/large words keep their styling, tatweels appear in both runs, `jc`/spacing intact.
- [ ] **Step 3:** Switch to Spacing mode, Justify. Confirm the larger word stays proportionally larger (uniform `fontScale`), line fills toward the target, styling preserved.
- [ ] **Step 4:** Re-justify twice. Confirm no compounding (idempotent).

---

# Part B — Web preview consumer (upstream submodule)

Authored in `vendor/ashaar-js/ashaar.js`, then synced. Extracts a pure run-spec helper (tested) and rewrites `justifyMisra` to be run-aware (manual browser verify).

### Task B1: pure `misraRunSpecs` helper in the submodule

**Files:**
- Modify: `vendor/ashaar-js/ashaar.js` (add function + export)
- Sync: `npm run sync:ashaar` (copies to `src/vendor/ashaar.js`)
- Test: `tests/ashaar-misra-runs.test.js` (new; add to `npm test` script in `package.json`)

**Interfaces:**
- Produces: `Ashaar.misraRunSpecs(childStyles)` where `childStyles = [{ text, fontKey, fontSize }]` (one entry per misra child node; `fontKey` is a caller-built string identity of the child's computed font — family+weight+style). Returns `[{ text, fontKey, fontSize }]` unchanged in order (identity pass-through in v1 — no coalescing, since DOM child nodes are already the run boundaries), but **filters out** entries whose `text` is empty/whitespace-only. This isolates the "which nodes are runs" rule so it is node-testable; the DOM walk that produces `childStyles` is the untested glue.

- [ ] **Step 1: Write the failing test** — new file `tests/ashaar-misra-runs.test.js`:

```js
const assert = require("assert");
const Ashaar = require("../src/vendor/ashaar");

// Empty/whitespace child nodes are dropped; order and styling preserved.
{
  const specs = Ashaar.misraRunSpecs([
    { text: "درد ",  fontKey: "Amiri/400/normal", fontSize: 16 },
    { text: "دل",    fontKey: "Amiri/700/normal", fontSize: 16 },
    { text: "   ",   fontKey: "Amiri/400/normal", fontSize: 16 },
  ]);
  assert.equal(specs.length, 2);
  assert.equal(specs[0].fontKey, "Amiri/400/normal");
  assert.equal(specs[1].text, "دل");
}
assert.deepEqual(Ashaar.misraRunSpecs([]), []);
```

- [ ] **Step 2: Run to verify it fails**

Run: `node tests/ashaar-misra-runs.test.js`
Expected: FAIL — `Ashaar.misraRunSpecs is not a function` (function absent from the synced copy).

- [ ] **Step 3: Implement in the submodule.** In `vendor/ashaar-js/ashaar.js` add the function and add `misraRunSpecs: misraRunSpecs` to the returned module object:

```js
// Given one entry per misra child node, return the ordered run specs to justify,
// dropping whitespace-only nodes. DOM-free so it is unit-testable; the caller
// supplies computed-font identity (fontKey) and size per child.
function misraRunSpecs(childStyles) {
  return (childStyles || []).filter(function (c) {
    return c && typeof c.text === "string" && c.text.trim();
  }).map(function (c) {
    return { text: c.text, fontKey: c.fontKey, fontSize: c.fontSize };
  });
}
```

- [ ] **Step 4: Sync + run to verify it passes**

Run: `npm run sync:ashaar && node tests/ashaar-misra-runs.test.js`
Expected: PASS.

- [ ] **Step 5: Wire into `npm test`.** In `package.json`, append ` && node tests/ashaar-misra-runs.test.js` to the `test` script.

- [ ] **Step 6: Commit**

```bash
git add vendor/ashaar-js src/vendor/ashaar.js tests/ashaar-misra-runs.test.js package.json
git commit -m "feat(web): misraRunSpecs pure helper (submodule + sync)"
```

### Task B2: run-aware `justifyMisra` (DOM glue — manual verify)

**Files:**
- Modify: `vendor/ashaar-js/ashaar.js` — `justifyMisra` (`~319-377`), `blockTargets` (`~407-433`) natural-width; then `npm run sync:ashaar`.

**Interfaces:**
- Consumes: `Ashaar.misraRunSpecs`, `AshaarJustify.justifyRuns`, `AshaarJustify.computeRunSpacing`.

No node test (DOM). Verified manually (Task B3). Steps:

- [ ] **Step 1: Cache styled markup for reducibility.** Replace the `dataset.ashaarOriginal` (plain-text) cache with `dataset.ashaarOriginalHtml`: on first pass store `spanEl.innerHTML`; on every pass restore it before discovery. This keeps child styling across re-justifies and stops compounding.

- [ ] **Step 2: Discover runs.** Walk `spanEl.childNodes`. For each node build `{ node, text, fontKey, fontSize, measure }`: for an element node use `getComputedStyle(node)`; for a text node use `getComputedStyle(spanEl)`. `fontKey = family + "/" + weight + "/" + style`; `fontSize = parseFloat(cs.fontSize)`; `measure(s)` restyles the single shared probe (family/weight/style/size/features) and returns `probeWidth(probe, s)`. Pass the `{text, fontKey, fontSize}` list through `Ashaar.misraRunSpecs` to get the runs to justify (keeping the parallel `node` array for write-back).

- [ ] **Step 3: Kashida.** `var out = Justify.justifyRuns(primRuns, effectiveTarget, opts)` then write each `out[i].text` into its node (`node.nodeType===TEXT ? node.nodeValue = t : node.textContent = t`). Element children keep their styling.

- [ ] **Step 4: Spacing/scale.** `var sp = Justify.computeRunSpacing(primRuns, available, opts)`. Set `spanEl.style.wordSpacing = sp.wordSpacing + "px"` (one value, whole misra). Apply `sp.fontScale` per child: for each run node with an element, `el.style.fontSize = (fontSize * sp.fontScale) + "px"`; for the bare-text run, set `spanEl.style.fontSize`. This scales every run robustly regardless of px/em source units.

- [ ] **Step 5: `blockTargets` natural width.** Where it currently does one `probeWidth(probe, text)` per misra, sum per-run widths (discover runs as in Step 2 and add their `measure(text)`), so balancing uses the correct mixed-style natural width.

- [ ] **Step 6: Sync + commit**

```bash
npm run sync:ashaar
git add vendor/ashaar-js src/vendor/ashaar.js
git commit -m "feat(web): run-aware justifyMisra over styled child nodes"
```

### Task B3: manual verification in the browser preview

**Files:** none.

- [ ] **Step 1:** In the taskpane preview, enter poetry and mark a word bold and another word larger (or use markup that renders child spans).
- [ ] **Step 2:** Trigger justify (kashida). Confirm bold/larger words keep styling, tatweels land across runs, misras balance to the longest.
- [ ] **Step 3:** Spacing mode: confirm `word-spacing` fills the line and the larger word stays proportionally larger.
- [ ] **Step 4:** Re-render/re-justify: confirm no compounding (styled markup restored from `ashaarOriginalHtml`).

---

## Final

- [ ] Run full suite: `npm test` — all green.
- [ ] Push branch `feat/run-aware-kashida-consumers`.

## Self-review notes

- **Spec coverage:** A1/A2/A3 = Word run discovery + write-back (spec §Sub-project 1); B1/B2 = Web discovery + apply (spec §Sub-project 2); A4/B3 = the spec's "manual verification" testing rows. Micro-space inter-run-gap limitation documented in A2 matches the spec's boundary decision.
- **Not covered by node tests (by design, per Global Constraints):** all Office.js/DOM glue — isolated so the pure rules (coalescing, gap distribution, run filtering) are tested.
- **Type consistency:** `coalesceRuns` returns `{text,name,size,bold,italic}`; A3 consumes exactly those. `distributeMicroSpaces(runTexts,n,spaceChar)` consumed in A3 Step 4. `misraRunSpecs(childStyles)→{text,fontKey,fontSize}` consumed in B2 Step 2.
