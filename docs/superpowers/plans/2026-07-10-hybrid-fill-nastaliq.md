# Hybrid Fill (Nastaliq) — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make the Nastaliq calligraphic mechanisms actually fill the column: after Mehr (tatweel) / Jameel (font-swap) elongation undershoots, close the residual with **capped** inter-word micro-spacing, accepting a slightly short line if the cap is hit.

**Architecture:** One new pure module `src/taskpane/kashida-residual.js` (`AshaarResidual`) with two node-tested functions — `capMicroSpaces` (residual px → capped glyph count) and `injectSpaceRuns` (spread hair-spaces across a font-swap run-list's space runs). The two existing elongation branches in `justifySelection` (`taskpane.js`) chain onto their output: measure achieved width, compute the capped residual count, realize it (Mehr via the existing `distributeMicroSpaces` on its flat text; Jameel via `injectSpaceRuns` on its run-list before `runsToMisraXml`). No new user mode; no engine/submodule change.

**Tech Stack:** Vanilla JS (ES5/UMD, no build step), Office.js v1, Node `assert` tests (no framework), `adm-zip` only for doc generators.

**Spec:** `docs/superpowers/specs/2026-07-10-hybrid-fill-design.md`

## Global Constraints

- **ES5/UMD only** — `var`, `function`; no arrow/`const`/template literals in shipped `src/` files. Match surrounding style.
- **Never edit `src/vendor/`** — not touched by this plan.
- **Pure logic is node-tested** (`node tests/<file>.test.js`); **Office.js glue in `taskpane.js` is NOT node-testable** — verified manually in Word (Task 5).
- **Reuse, do not duplicate:** `AshaarWord.distributeMicroSpaces`, `AshaarWord.runsToMisraXml`, `AshaarKashidaFontswap.selectSwapRuns`, the in-scope `MICRO_SPACE`/`canvasCtx`/`repSize` locals in `justifySelection`.
- **Cap:** total added spacing ≤ `0.28 · sizePx · gaps` (0.28em/gap), matching the spacing-mode ceiling.
- **Policy:** fill to the column edge (`colPx`); **accept-short** when the cap binds. No glyph-shrink, no column-width change.
- **Order (fixed):** elongate → measure achieved → residual → capped micro-spacing → accept short. Never spacing-first.
- **Out of scope (tracked, do NOT build):** strength-as-kashida-share, hybrid on the generic path, per-run mixed mechanisms → [ashaar-js#7]; prose justification → [ashaar-js#8]. The generic path already fills via `justifyRuns`; whitespace fonts (Gulzar/Noto) are unchanged.
- `npm test` green after every task.

---

### Task 1: `capMicroSpaces` — new pure module `AshaarResidual`

**Files:**
- Create: `src/taskpane/kashida-residual.js`
- Create: `tests/kashida-residual.test.js`
- Modify: `src/taskpane/taskpane.html` (add to the `srcs` array)
- Modify: `package.json` (append to the `test` script)

**Interfaces:**
- Produces: `AshaarResidual.capMicroSpaces(residualPx, gaps, spaceGlyphPx, sizePx, capEm) -> integer`.
  Returns `0` when `residualPx <= 0`, `gaps <= 0`, or `spaceGlyphPx <= 0`. `capEm` defaults to `0.28`. Total added width is capped at `capEm * sizePx * gaps`; the returned count is `round(min(residualPx, cap) / spaceGlyphPx)`.

- [ ] **Step 1: Write the failing test** — create `tests/kashida-residual.test.js`:

```js
"use strict";
var assert = require("assert");
var AshaarResidual = require("../src/taskpane/kashida-residual");

// zero / invalid inputs → 0
assert.strictEqual(AshaarResidual.capMicroSpaces(0, 5, 2, 16), 0);      // no residual
assert.strictEqual(AshaarResidual.capMicroSpaces(100, 0, 2, 16), 0);    // no gaps
assert.strictEqual(AshaarResidual.capMicroSpaces(100, 5, 0, 16), 0);    // no glyph width

// under the cap: fill the residual exactly (rounded to glyph count)
// cap = 0.28*16*4 = 17.92px; residual 10 < cap → round(10/2) = 5
assert.strictEqual(AshaarResidual.capMicroSpaces(10, 4, 2, 16), 5);

// cap binds: residual 100 > cap 17.92 → round(17.92/2) = round(8.96) = 9
assert.strictEqual(AshaarResidual.capMicroSpaces(100, 4, 2, 16), 9);

// explicit capEm honored: cap = 0.5*16*4 = 32px → round(min(100,32)/2)=16
assert.strictEqual(AshaarResidual.capMicroSpaces(100, 4, 2, 16, 0.5), 16);

console.log("kashida-residual tests passed");
```

- [ ] **Step 2: Run to verify it fails** — Run: `node tests/kashida-residual.test.js`; Expected: FAIL — `Cannot find module '../src/taskpane/kashida-residual'`.

- [ ] **Step 3: Implement** — create `src/taskpane/kashida-residual.js`:

```js
/**
 * AshaarResidual — hybrid-fill residual spacing. After calligraphic elongation
 * (Mehr tatweel / Jameel font-swap) undershoots the column, close the gap with
 * a CAPPED number of inter-word micro-spaces. Pure (no DOM); node-testable.
 */
(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) module.exports = factory();
  else root.AshaarResidual = factory();
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  // Hair space (U+200A) — same glyph the spacing path uses to realize word gaps.
  var HAIR_SPACE = " ";

  // How many micro-space glyphs to add to close `residualPx`, never exceeding
  // capEm*sizePx per gap (default 0.28em/gap). Returns 0 for non-positive
  // residual / no gaps / unmeasurable glyph.
  function capMicroSpaces(residualPx, gaps, spaceGlyphPx, sizePx, capEm) {
    if (capEm == null) capEm = 0.28;
    if (!(residualPx > 0) || !(gaps > 0) || !(spaceGlyphPx > 0)) return 0;
    var capPx = capEm * sizePx * gaps;
    var wantPx = Math.min(residualPx, capPx);
    return Math.max(0, Math.round(wantPx / spaceGlyphPx));
  }

  return { HAIR_SPACE: HAIR_SPACE, capMicroSpaces: capMicroSpaces };
}));
```

- [ ] **Step 4: Run to verify it passes** — Run: `node tests/kashida-residual.test.js`; Expected: `kashida-residual tests passed`.

- [ ] **Step 5: Register the module for the browser** — in `src/taskpane/taskpane.html`, add `"./kashida-residual.js",` to the `srcs` array immediately after `"./kashida-fontswap.js",` (order is load-sequence; no dependency, must precede `./taskpane.js`):

```js
          "./kashida-fontswap.js",
          "./kashida-residual.js",
          "./word-html.js",
```

- [ ] **Step 6: Add to the test suite** — in `package.json`, append ` && node tests/kashida-residual.test.js` to the end of the `"test"` script string.

- [ ] **Step 7: Full suite** — Run: `npm test`; Expected: all green including `kashida-residual tests passed`.

- [ ] **Step 8: Commit**

```bash
git add src/taskpane/kashida-residual.js tests/kashida-residual.test.js src/taskpane/taskpane.html package.json
git commit -m "feat(hybrid-fill): capMicroSpaces — capped residual micro-space count"
```

---

### Task 2: `injectSpaceRuns` — spread hair-spaces across a run-list's space runs

**Files:**
- Modify: `src/taskpane/kashida-residual.js` (add function + export)
- Modify: `tests/kashida-residual.test.js` (append)

**Interfaces:**
- Consumes: nothing new.
- Produces: `AshaarResidual.injectSpaceRuns(runs, n, spaceChar) -> runs'`.
  `runs`: `[{text, swap, ...}]` (as returned by `AshaarKashidaFontswap.selectSwapRuns`). Distributes `n` `spaceChar` glyphs (default `HAIR_SPACE`) as evenly as possible across the runs whose `text` is exactly `" "`, appending them to that run's text. Returns a **new** array of shallow-copied run objects (originals untouched); `swap` flags and order preserved. `n <= 0` or no space runs → structural copy, text unchanged.

- [ ] **Step 1: Write the failing test** — append to `tests/kashida-residual.test.js` (before the final `console.log`):

```js
// injectSpaceRuns — hair-spaces appended to the " " runs only
(function () {
  var runs = [{ text: "كہہ", swap: true }, { text: " ", swap: false },
              { text: "رہے", swap: false }, { text: " ", swap: false },
              { text: "تھے", swap: false }];
  var out = AshaarResidual.injectSpaceRuns(runs, 3);      // 2 space runs, n=3
  assert.strictEqual(out.length, 5);
  assert.strictEqual(out[0].text, "كہہ");                 // words untouched
  assert.strictEqual(out[0].swap, true);                  // swap flag preserved
  assert.strictEqual(out[1].text, " " + "  ");  // first gap: 2 (remainder)
  assert.strictEqual(out[3].text, " " + " ");        // second gap: 1
  assert.strictEqual(runs[1].text, " ");                  // original NOT mutated

  // n <= 0 → text unchanged (structural copy)
  var out0 = AshaarResidual.injectSpaceRuns(runs, 0);
  assert.strictEqual(out0[1].text, " ");
  assert.notStrictEqual(out0, runs);

  // no space runs → unchanged
  var solid = [{ text: "ابجد", swap: false }];
  assert.strictEqual(AshaarResidual.injectSpaceRuns(solid, 5)[0].text, "ابجد");

  // custom spaceChar honored
  var out2 = AshaarResidual.injectSpaceRuns(runs, 2, " ");
  assert.strictEqual(out2[1].text, " " + " ");
  assert.strictEqual(out2[3].text, " " + " ");
})();
```

- [ ] **Step 2: Run to verify it fails** — Run: `node tests/kashida-residual.test.js`; Expected: FAIL — `injectSpaceRuns is not a function`.

- [ ] **Step 3: Implement** — in `src/taskpane/kashida-residual.js`, add the function and export it:

```js
  // Spread `n` spaceChar glyphs across the inter-word (" ") runs of a font-swap
  // run-list, as evenly as possible (earlier gaps take the remainder). Returns a
  // new array of copied run objects; input is never mutated.
  function injectSpaceRuns(runs, n, spaceChar) {
    if (spaceChar == null) spaceChar = HAIR_SPACE;
    var out = (runs || []).map(function (r) {
      var c = {}; for (var k in r) { if (r.hasOwnProperty(k)) c[k] = r[k]; } return c;
    });
    if (!(n > 0)) return out;
    var gapIdx = [];
    for (var i = 0; i < out.length; i++) { if (out[i].text === " ") gapIdx.push(i); }
    if (!gapIdx.length) return out;
    var base = Math.floor(n / gapIdx.length), rem = n % gapIdx.length;
    for (var g = 0; g < gapIdx.length; g++) {
      var add = base + (g < rem ? 1 : 0);
      if (add > 0) out[gapIdx[g]].text += new Array(add + 1).join(spaceChar);
    }
    return out;
  }
```

Update the return to: `return { HAIR_SPACE: HAIR_SPACE, capMicroSpaces: capMicroSpaces, injectSpaceRuns: injectSpaceRuns };`

- [ ] **Step 4: Run to verify it passes** — Run: `node tests/kashida-residual.test.js`; Expected: `kashida-residual tests passed`.

- [ ] **Step 5: Full suite** — Run: `npm test`; Expected: all green.

- [ ] **Step 6: Commit**

```bash
git add src/taskpane/kashida-residual.js tests/kashida-residual.test.js
git commit -m "feat(hybrid-fill): injectSpaceRuns — even hair-space spread across gap runs"
```

---

### Task 3: Chain residual spacing onto the Jameel (font-swap) branch

**Files:**
- Modify: `src/taskpane/taskpane.js` — the `if (mechanism === "font-swap")` branch in `justifySelection`'s phase-1 loop (currently ~`1471–1486`).

**Interfaces:**
- Consumes: `AshaarResidual.capMicroSpaces`, `AshaarResidual.injectSpaceRuns`; in-scope locals `canvasCtx`, `repSize`, `colPx`, `MICRO_SPACE`, `cellAlign`, `opts`; `sel.fill` (achieved / colPx) from `selectSwapRuns`.

No node test (Office.js) — verified in Task 5. `sel.fill = achievedWidth / colPx`, so `achieved = sel.fill * colPx` and `residual = colPx - achieved`.

- [ ] **Step 1: Insert residual spacing between `selectSwapRuns` and `runsToMisraXml`.** Replace these two lines:

```js
          var sel = AshaarKashidaFontswap.selectSwapRuns(fss, wb, ww, colPx);
          var swapXml = AshaarWord.runsToMisraXml(sel.runs, cellAlign, opts, repSize);
```

with:

```js
          var sel = AshaarKashidaFontswap.selectSwapRuns(fss, wb, ww, colPx);
          // Hybrid fill: font-swap elongation undershoots (only fasls with a
          // Kasheeda variant widen) — close the residual with capped hair-spaces
          // in the inter-word gap runs. Accept-short if the cap binds.
          var jGaps = 0;
          for (var jgi = 0; jgi < sel.runs.length; jgi++) { if (sel.runs[jgi].text === " ") jGaps++; }
          canvasCtx.font = repSize + "pt " + baseCss;
          var jSpacePx = canvasCtx.measureText(MICRO_SPACE).width || 1;
          var jn = AshaarResidual.capMicroSpaces(colPx - sel.fill * colPx, jGaps, jSpacePx, repSize * 96 / 72);
          var jRuns = AshaarResidual.injectSpaceRuns(sel.runs, jn, MICRO_SPACE);
          var swapXml = AshaarWord.runsToMisraXml(jRuns, cellAlign, opts, repSize);
```

- [ ] **Step 2: Syntax check** — Run: `node --check src/taskpane/taskpane.js`; Expected: no output (exit 0).

- [ ] **Step 3: Full suite** — Run: `npm test`; Expected: all green (no regression; branch is Office.js, exercised manually in Task 5).

- [ ] **Step 4: Commit**

```bash
git add src/taskpane/taskpane.js
git commit -m "feat(hybrid-fill): Jameel font-swap residual micro-spacing (accept-short)"
```

---

### Task 4: Chain residual spacing onto the Mehr (tatweel) branch

**Files:**
- Modify: `src/taskpane/taskpane.js` — the `if (mechanism === "tatweel" && opts.justifyMode === "kashida")` branch (currently ~`1494–1513`).

**Interfaces:**
- Consumes: `AshaarResidual.capMicroSpaces`, `AshaarWord.distributeMicroSpaces`; in-scope locals `canvasCtx`, `repSize`, `colPx`, `MICRO_SPACE`, `mehrFont`, `current`; `msel.fill` from `selectSwapRuns`.

No node test (Office.js) — verified in Task 5. Mehr's output `mout` is flat text with single spaces between words; `distributeMicroSpaces([mout], n, MICRO_SPACE)[0]` inserts the hair-spaces at those gaps.

- [ ] **Step 1: Add residual spacing after `mout` is built.** Replace:

```js
          var mout = msel.runs.map(function (r, i) { return (r.swap && mww[i] > mwb[i]) ? melong[i] : mtoks[i]; }).join("");
          if (mout !== current) plans.push({ cell: cell, flat: mout });
          return;
```

with:

```js
          var mout = msel.runs.map(function (r, i) { return (r.swap && mww[i] > mwb[i]) ? melong[i] : mtoks[i]; }).join("");
          // Hybrid fill: Mehr elongates only at whitelisted word-endings, so it
          // undershoots — close the residual with capped hair-spaces at the word
          // gaps (reusing distributeMicroSpaces). Accept-short if the cap binds.
          var mGaps = mout.split(" ").length - 1;
          canvasCtx.font = mehrFont;
          var mSpacePx = canvasCtx.measureText(MICRO_SPACE).width || 1;
          var mn = AshaarResidual.capMicroSpaces(colPx - msel.fill * colPx, mGaps, mSpacePx, repSize * 96 / 72);
          var mfinal = AshaarWord.distributeMicroSpaces([mout], mn, MICRO_SPACE)[0];
          if (mfinal !== current) plans.push({ cell: cell, flat: mfinal });
          return;
```

- [ ] **Step 2: Syntax check** — Run: `node --check src/taskpane/taskpane.js`; Expected: no output (exit 0).

- [ ] **Step 3: Full suite** — Run: `npm test`; Expected: all green.

- [ ] **Step 4: Commit**

```bash
git add src/taskpane/taskpane.js
git commit -m "feat(hybrid-fill): Mehr tatweel residual micro-spacing (accept-short)"
```

---

### Task 5: Manual end-to-end verification in Word

**Files:** none.

- [ ] **Step 1:** `npm start` (or reload the task pane if the add-in is already running). Insert an Arabic/Urdu poem via Conversion → Insert as Table.
- [ ] **Step 2 — Jameel:** set the poem font to **Jameel Noori Kasheeda** (pane Font = Jameel), Justification = Kashida, then **Justify Selected Text**. Confirm each misra **fills to the column edge**: fasls with a Kasheeda variant widen AND the remaining gap is closed by small inter-word spacing (no giant word gaps; shaping intact — spaces sit at word boundaries).
- [ ] **Step 3 — Mehr:** set the font to **Mehr Nastaliq** and Justify. Confirm the whitelisted word-endings elongate AND the residual is closed by capped spacing; lines that can't reach the edge stay **slightly short** (no over-spacing, no shrink).
- [ ] **Step 4 — cap/accept-short:** pick a misra with few/no eligible elongation points (e.g. a short Mehr line). Confirm it does NOT over-space — it fills up to the 0.28em/gap cap then stops short.
- [ ] **Step 5 — no regression:** confirm a **generic** font (Fatemi Maqala, Document default) still kashida-fills via the generic path (unchanged), and a **whitespace** font (Gulzar) still fills by spacing only.
- [ ] **Step 6:** Re-justify twice; confirm idempotent (no compounding — `stripJustification` strips prior tatweels/micro-spaces before each pass).

## Self-Review notes
- **Spec coverage:** `capMicroSpaces` §1 → Task 1; `injectSpaceRuns` §2 → Task 2; Mehr chaining §3 → Task 4; Jameel chaining §4 → Task 3; accept-short/0.28em policy §5 → Global Constraints + Tasks 3/4 (cap) + Task 5 Step 4 (verify). Testing § - `capMicroSpaces`/`injectSpaceRuns` node-tested (Tasks 1–2); chaining manual (Task 5).
- **Types:** `capMicroSpaces(residualPx, gaps, spaceGlyphPx, sizePx, capEm)`; `injectSpaceRuns(runs, n, spaceChar)`; both consistent across tasks. `sel.fill`/`msel.fill` are the achieved/colPx ratio (confirmed in `kashida-fontswap.js`).
- **Deferred correctly:** strength-share, generic-path hybrid, mixed mechanisms → ashaar-js#7; prose → ashaar-js#8. Not referenced by any task here.
