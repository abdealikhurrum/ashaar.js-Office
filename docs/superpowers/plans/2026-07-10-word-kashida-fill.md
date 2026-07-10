# §3 — Native Word Kashida Fill ("Let Word fill it") Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make the "Let Word fill it" mode actually justify in the Word document, using Word's **native kashida** (`w:jc="lowKashida|mediumKashida|highKashida"`) with an auto-inserted trailing soft break so single-line misras stretch — no user Shift-Enter, no font-loading (Word uses its own font).

**Architecture:** The mode's value is `justifyMode:"css"` (label reworked later in §5). Two emission points: the **insert path** (`word-html.js` `misraParaXml`, pure/OOXML-testable) and the **justify path** (`taskpane.js` `justifySelection`, Office.js, rebuilds each cell paragraph via `insertOoxml` since the Office.js alignment enum has no kashida values). Strength (`tatweelCount`, 0–24) picks the level in thirds; full strength allows ~15% qaseeda-proportional column expansion. Non-Arabic falls back to `w:jc="distribute"` (spacing, fills the last line without a break).

**Tech Stack:** Vanilla JS (ES5/UMD, no build), Office.js v1, Node `assert` tests, `adm-zip` for the verification docx.

## Global Constraints

- No build step; ES5-compatible UMD only.
- Never edit `src/vendor/` directly (not touched here).
- Tests are pure Node (`node tests/<file>.test.js`); Office.js and real Word rendering are NOT node-testable — verified manually.
- Shrunk trailing break run = `<w:r><w:rPr><w:sz w:val="4"/><w:szCs w:val="4"/></w:rPr><w:br/></w:r>` (sz 4 = 2pt, minimizes the empty line's height).
- Kashida `w:jc` values elongate **Arabic/complex script only**; Latin/mixed → `distribute`.
- Full suite: `npm test`.

---

### Task 1: Verification spike — prove the trailing-break kashida trick in Word (GATE)

**Files:**
- Create: `scripts/make-kashida-spike-doc.mjs` (throwaway generator)

**Interfaces:** none (manual verification task). **This gates the whole plan** — if Word does not kashida the single line, or the empty line is too tall to hide, stop and revise §3 before writing any product code.

- [ ] **Step 1: Write the spike generator.** A minimal `.docx` with **three one-row, two-cell RTL tables**, each cell a single Arabic misra, differing only in `w:jc`: (a) `mediumKashida` **with** the shrunk trailing `<w:br/>`, (b) `mediumKashida` **without** a break (control — expect no stretch), (c) `distribute` (control — expect spacing fill). Reuse the packaging from `scripts/make-mixed-style-doc.mjs` (adm-zip, same `[Content_Types]`/rels). Cell paragraph:

```js
// jc: "mediumKashida" | "distribute"; withBreak: boolean
function cellPara(text, jc, withBreak) {
  var brk = withBreak
    ? '<w:r><w:rPr><w:sz w:val="4"/><w:szCs w:val="4"/></w:rPr><w:br/></w:r>'
    : '';
  return '<w:p><w:pPr><w:bidi/><w:jc w:val="' + jc + '"/></w:pPr>' +
    '<w:r><w:rPr><w:rtl/><w:rFonts w:cs="Scheherazade New"/>' +
    '<w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>' +
    '<w:t xml:space="preserve">' + text + '</w:t></w:r>' + brk + '</w:p>';
}
```

Use a short misra (e.g. `العلم نور والجهل ظلام`) in a ~2.5" cell so there is clear room to stretch. Write to `test-documents/kashida-spike.docx`.

- [ ] **Step 2: Generate.** Run: `node scripts/make-kashida-spike-doc.mjs`
Expected: "Created …/kashida-spike.docx". Confirm valid zip: `node -e "new (require('adm-zip'))('test-documents/kashida-spike.docx').readAsText('word/document.xml').length"` prints a positive number.

- [ ] **Step 3: Open in Word and OBSERVE (manual).** Open `test-documents/kashida-spike.docx` in desktop Word.
Expected / record findings:
  - (a) `mediumKashida` **+ break** → the misra's letters **elongate to fill the cell**; the trailing empty line is barely visible (2pt).
  - (b) `mediumKashida` **no break** → line does **not** stretch (confirms the last-line rule).
  - (c) `distribute` → line fills via **spacing**, not kashida.

- [ ] **Step 4: GATE decision.** If (a) stretches acceptably and the empty line is tolerable → proceed. If not, **stop** and record what Word did (no stretch? too much empty-line height? needs `highKashida`? needs a different shrink like exact line spacing?) — revise the spec's §3 approach before continuing.

- [ ] **Step 5: Commit the spike** (kept as a reproducible check; the `.docx` is a generated artifact — don't commit it).

```bash
git add scripts/make-kashida-spike-doc.mjs
git commit -m "spike: Word kashida + trailing-break verification doc for §3"
```

### Task 2: `strengthToKashidaLevel` — map 0–24 to a kashida jc value (thirds)

**Files:**
- Modify: `src/taskpane/word-html.js` (add function + export)
- Test: `tests/word-html.test.js`

**Interfaces:**
- Produces: `AshaarWord.strengthToKashidaLevel(strength)` → `"lowKashida" | "mediumKashida" | "highKashida"`. Thirds of 0–24: `[0,8) → low`, `[8,16) → medium`, `[16,24] → high`. Non-numeric/undefined → `"mediumKashida"`.

- [ ] **Step 1: Write the failing test** — append to `tests/word-html.test.js`:

```js
// ── strengthToKashidaLevel ──────────────────────────────────────────────────
assert.equal(AshaarWord.strengthToKashidaLevel(0),  "lowKashida");
assert.equal(AshaarWord.strengthToKashidaLevel(7),  "lowKashida");
assert.equal(AshaarWord.strengthToKashidaLevel(8),  "mediumKashida");
assert.equal(AshaarWord.strengthToKashidaLevel(15), "mediumKashida");
assert.equal(AshaarWord.strengthToKashidaLevel(16), "highKashida");
assert.equal(AshaarWord.strengthToKashidaLevel(24), "highKashida");
assert.equal(AshaarWord.strengthToKashidaLevel(undefined), "mediumKashida");
```

- [ ] **Step 2: Run to verify it fails** — Run: `node tests/word-html.test.js`; Expected: FAIL — not a function.

- [ ] **Step 3: Implement** (add near the run-aware helpers; add `strengthToKashidaLevel: strengthToKashidaLevel,` to exports):

```js
// Map the 0–24 stretch slider to Word's three native kashida jc levels, in
// thirds. Used by "Let Word fill it" (§3).
function strengthToKashidaLevel(strength) {
  var s = Number(strength);
  if (!isFinite(s)) return "mediumKashida";
  if (s < 8) return "lowKashida";
  if (s < 16) return "mediumKashida";
  return "highKashida";
}
```

- [ ] **Step 4: Run to verify it passes** — Run: `node tests/word-html.test.js`; Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/taskpane/word-html.js tests/word-html.test.js
git commit -m "feat(word-fill): strengthToKashidaLevel — 0-24 to kashida level in thirds"
```

### Task 3: `containsArabic` + `wordFillJc` — kashida for Arabic, distribute otherwise

**Files:**
- Modify: `src/taskpane/word-html.js`
- Test: `tests/word-html.test.js`

**Interfaces:**
- Produces:
  - `AshaarWord.containsArabic(text)` → boolean (any char in the Arabic ranges).
  - `AshaarWord.wordFillJc(text, strength)` → the jc value: a kashida level (Arabic) or `"distribute"` (non-Arabic).

- [ ] **Step 1: Write the failing test**:

```js
// ── containsArabic / wordFillJc ─────────────────────────────────────────────
assert.equal(AshaarWord.containsArabic("العلم"), true);
assert.equal(AshaarWord.containsArabic("hello"), false);
assert.equal(AshaarWord.containsArabic("۱۲۳"), true);   // Urdu digits are in-range
assert.equal(AshaarWord.wordFillJc("العلم نور", 4),  "lowKashida");
assert.equal(AshaarWord.wordFillJc("العلم نور", 20), "highKashida");
assert.equal(AshaarWord.wordFillJc("hello world", 20), "distribute");
```

- [ ] **Step 2: Run to verify it fails** — Run: `node tests/word-html.test.js`; Expected: FAIL.

- [ ] **Step 3: Implement** (add functions + exports):

```js
// Any Arabic-script character (Arabic, Arabic Supplement, presentation forms).
function containsArabic(text) {
  return /[؀-ۿݐ-ݿﭐ-﷿ﹰ-﻿]/.test(String(text || ""));
}

// jc for "Let Word fill it": native kashida (level from strength) for Arabic,
// else distribute (spacing — fills the last line without a trailing break).
function wordFillJc(text, strength) {
  return containsArabic(text) ? strengthToKashidaLevel(strength) : "distribute";
}
```

- [ ] **Step 4: Run to verify it passes** — Run: `node tests/word-html.test.js`; Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/taskpane/word-html.js tests/word-html.test.js
git commit -m "feat(word-fill): containsArabic + wordFillJc (kashida vs distribute)"
```

### Task 4: Emit native kashida jc + trailing break on insert (`misraParaXml`)

**Files:**
- Modify: `src/taskpane/word-html.js` (`misraParaXml`, ~1089)
- Test: `tests/word-html.test.js`

**Interfaces:**
- Consumes: `wordFillJc`, `strengthToKashidaLevel`, `containsArabic`.
- Changes `misraParaXml` behavior when `opts.justifyMode === "css"`: the paragraph's `w:jc` becomes the word-fill value, and (for Arabic) a shrunk trailing `<w:br/>` run is appended. Signature unchanged.

- [ ] **Step 1: Write the failing test** (append; `misraParaXml` is not exported — test via the public `renderForWordOoxml` OR export `misraParaXml`; export it for a focused test):

Add `misraParaXml: misraParaXml,` to exports, then:

```js
// ── misraParaXml: word-fill mode ────────────────────────────────────────────
{
  const opts = { justifyMode: "css", tatweelCount: 20 };
  const xml = AshaarWord.misraParaXml("العلم نور", "center", false, opts, 0);
  assert.ok(xml.indexOf('w:jc w:val="highKashida"') !== -1, "Arabic → highKashida jc");
  assert.ok(xml.indexOf("<w:br/>") !== -1, "Arabic word-fill appends a trailing break");
  assert.ok(xml.indexOf('w:sz w:val="4"') !== -1, "trailing break run is shrunk");
}
{
  const opts = { justifyMode: "css", tatweelCount: 20 };
  const xml = AshaarWord.misraParaXml("hello world", "center", false, opts, 0);
  assert.ok(xml.indexOf('w:jc w:val="distribute"') !== -1, "non-Arabic → distribute");
  assert.ok(xml.indexOf("<w:br/>") === -1, "distribute needs no trailing break");
}
{
  // Non-word-fill modes unchanged: jc still by position, no break.
  const xml = AshaarWord.misraParaXml("العلم", "right", false, { justifyMode: "kashida" }, 0);
  assert.ok(xml.indexOf('w:jc w:val="right"') !== -1, "kashida mode keeps positional jc");
  assert.ok(xml.indexOf("<w:br/>") === -1, "no break outside word-fill");
}
```

- [ ] **Step 2: Run to verify it fails** — Run: `node tests/word-html.test.js`; Expected: FAIL.

- [ ] **Step 3: Implement** — replace the body of `misraParaXml`:

```js
function misraParaXml(text, align, isRefrain, opts, indTwips) {
  opts = opts || {};
  var jc = align === "right" ? "right" : align === "left" ? "left" : "center";
  var trailingBreak = "";
  if (opts.justifyMode === "css") {
    // "Let Word fill it": native Word justification. Arabic → kashida level +
    // a shrunk trailing break so the single (last) line actually stretches;
    // non-Arabic → distribute (fills the last line without a break).
    jc = wordFillJc(text, Number(opts.tatweelCount || 0));
    if (containsArabic(text)) {
      trailingBreak = '<w:r><w:rPr><w:sz w:val="4"/><w:szCs w:val="4"/></w:rPr><w:br/></w:r>';
    }
  }
  var rpr = "<w:rPr><w:rtl/>";
  if (isRefrain) rpr += '<w:color w:val="A7352A"/>';
  if ((opts || {}).fontMode === "nastaliq") rpr += '<w:rFonts w:cs="Noto Nastaliq Urdu"/>';
  else if ((opts || {}).fontMode === "arabic-serif") rpr += '<w:rFonts w:cs="Scheherazade New"/>';
  rpr += "</w:rPr>";
  var ind = indTwips ? '<w:ind w:left="' + indTwips + '"/>' : "";
  return "<w:p>" +
    "<w:pPr><w:bidi/><w:spacing w:after=\"80\"/><w:jc w:val=\"" + jc + "\"/>" + ind + "</w:pPr>" +
    "<w:r>" + rpr + '<w:t xml:space="preserve">' + escapeXml(text) + "</w:t></w:r>" +
    trailingBreak +
    "</w:p>";
}
```

- [ ] **Step 4: Run to verify it passes** — Run: `node tests/word-html.test.js`; Expected: PASS.

- [ ] **Step 5: Full suite** — Run: `npm test`; Expected: all green (no regression in existing `renderForWordOoxml` tests).

- [ ] **Step 6: Commit**

```bash
git add src/taskpane/word-html.js tests/word-html.test.js
git commit -m "feat(word-fill): misraParaXml emits native kashida jc + trailing break"
```

### Task 5: `kashidaExpansionFraction` — 0 → ~0.15 column expansion by strength

**Files:**
- Modify: `src/taskpane/word-html.js`
- Test: `tests/word-html.test.js`

**Interfaces:**
- Produces: `AshaarWord.kashidaExpansionFraction(strength)` → number in `[0, 0.15]`, linear in strength/24 (0 at 0, ~0.15 at 24). Consumed by the justify path (Task 6) to widen columns qaseeda-proportionally.

- [ ] **Step 1: Write the failing test**:

```js
// ── kashidaExpansionFraction ────────────────────────────────────────────────
assert.equal(AshaarWord.kashidaExpansionFraction(0), 0);
assert.equal(AshaarWord.kashidaExpansionFraction(24), 0.15);
assert.equal(AshaarWord.kashidaExpansionFraction(12), 0.075);
assert.equal(AshaarWord.kashidaExpansionFraction(999), 0.15); // clamped
```

- [ ] **Step 2: Run to verify it fails** — Run: `node tests/word-html.test.js`; Expected: FAIL.

- [ ] **Step 3: Implement** (add + export):

```js
// Column expansion allowance for "Let Word fill it": 0 at strength 0, ~15% at
// full strength (24). Applied qaseeda-proportionally by the justify path.
function kashidaExpansionFraction(strength) {
  var s = Math.max(0, Math.min(24, Number(strength) || 0));
  return Math.round((0.15 * s / 24) * 1000) / 1000;
}
```

- [ ] **Step 4: Run to verify it passes** — Run: `node tests/word-html.test.js`; Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/taskpane/word-html.js tests/word-html.test.js
git commit -m "feat(word-fill): kashidaExpansionFraction — up to 15% at full strength"
```

### Task 6: Apply word-fill in `justifySelection` (Office.js glue — manual verify)

**Files:**
- Modify: `src/taskpane/taskpane.js` (`justifySelection`, ~1056)

**Interfaces:**
- Consumes: `AshaarWord.wordFillJc`, `AshaarWord.containsArabic`, `AshaarWord.kashidaExpansionFraction`.

No node test (Office.js). Verified in Task 7. The Office.js alignment enum has **no** kashida values, so the paragraph must be rebuilt via `insertOoxml`.

- [ ] **Step 1: Branch the mode.** Near the top of `justifySelection`, treat `opts.justifyMode === "css"` as the word-fill path (today it isn't handled). Skip the probe/calibrate/tatweel machinery for it.

- [ ] **Step 2: Widen columns first (qaseeda-proportional).** Compute `frac = AshaarWord.kashidaExpansionFraction(opts.tatweelCount)`. If `frac > 0` and the desktop TableColumn API is available (the existing `canResize` check, `taskpane.js:1162`), scale every column of each table by `1 + frac`, capped so the table doesn't exceed page width (reuse the page-width computation already in the auto-fit block). Sync.

- [ ] **Step 3: Rebuild each cell paragraph via OOXML — reuse the tested emitter.** The Office.js alignment enum has no kashida values, so replace the cell's paragraph with OOXML built by the **same `misraParaXml` from Task 4** (DRY — it already emits the word-fill jc + shrunk break, with escaping), wrapped by the existing `AshaarWord.wrapOoxml` (the wrapper used by `insertPoem`/`insertTabStopPoem`):

```js
var base = stripJustification(cell.body.text || "").trim();
if (!base) return;
var paraXml = AshaarWord.misraParaXml(base, "center", false, opts, 0); // opts.justifyMode==="css" → kashida jc + break
cell.body.insertOoxml(AshaarWord.wrapOoxml(paraXml), Word.InsertLocation.replace);
```

Do not set `rFonts` here — omit it so Word keeps the cell's own font. Sync per cell inside a try/catch so one failure doesn't abort the batch (mirror the run-aware fallback pattern already in `justifySelection`). **Verify in Task 7 that a `wrapOoxml`-wrapped single paragraph inserts cleanly into a cell body** — if Word rejects the wrapped package at cell scope, fall back to inserting the bare `paraXml` fragment, or set `paragraph.alignment = Word.Alignment.justified` (jc=both, no kashida) as a degraded path and note it.

- [ ] **Step 4: Report via the message line** (the §1 result panel is a later plan): `setMessage("Filled N cell(s) with Word kashida.")`.

- [ ] **Step 5: Commit**

```bash
git add src/taskpane/taskpane.js
git commit -m "feat(word-fill): justifySelection applies native kashida jc + break + column expansion"
```

### Task 7: Manual end-to-end verification in Word

**Files:** none.

- [ ] **Step 1:** `npm start` (opens Word with the add-in). Insert or open a poem with Arabic misras.
- [ ] **Step 2:** Choose "Let Word fill it" (justify mode `css`), set strength low / mid / high, and **Justify**. Confirm each cell's letters **kashida-stretch** at increasing intensity (low/med/high), the trailing empty line is negligible, and columns widen up to ~15% at full strength.
- [ ] **Step 3:** Insert a fresh poem with the same mode selected; confirm the **insert path** (Task 4) produces the same kashida fill without a separate Justify.
- [ ] **Step 4:** Try a Latin line; confirm it falls back to `distribute` (spacing) with no trailing break.
- [ ] **Step 5:** Re-justify / switch modes; confirm no compounding and that switching away behaves sanely (full cleanup is §6 Reset, a later plan — just confirm nothing breaks).

## Self-review notes

- **Spec coverage:** Task 1 = the §3 verify-in-Word risk (gate); Tasks 2–5 = the pure mappings (kashida level thirds, Arabic detection, jc selection, 15% expansion); Task 4 = insert-path emission; Task 6 = justify-path emission + expansion; Task 7 = §3 manual verification.
- **Deferred (noted):** the **tab-stop path** (`word-tabstop.js:150`, hardcoded `jc="center"`) is out of scope for this plan (tables are the primary layout); add a follow-up if tab-stop word-fill is needed. The §1 result-panel reporting and §6 reset of the trailing break / jc are their own plans.
- **Type consistency:** `strengthToKashidaLevel(strength)→string`, `containsArabic(text)→bool`, `wordFillJc(text,strength)→string`, `kashidaExpansionFraction(strength)→number` — all consumed with these signatures in Tasks 4 & 6.
- **Placeholder scan:** the OOXML-escape/wrap helper in Task 6 Step 3 must be the file's existing one — confirm its name when implementing (do not invent).


---

## REVISION (2026-07-10): re-render architecture (supersedes Task 6)

**Finding:** per-cell `cell.body.insertOoxml(package, replace)` drops the kashida `jc` (verified via saved runtime OOXML). Authoring at **selection scope** preserves it. `insertPoem(true)` already does exactly that: `renderForWordOoxml(source, opts)` (→ `misraParaXml`, which emits word-fill jc) → `wrapOoxml` → `selection.insertOoxml(..., replace)` → re-wrap CC. And `AshaarTableAdopt.adoptTableToSource(rows, {direction})` already reconstructs source from cell text. So word-fill justify = reconstruct + `insertPoem(true)` with `justifyMode:"css"`.

**Architecture (user-approved):** TWO PATHS. Engine-mode kashida (tatweel) keeps the in-place run-aware `justifySelection` (preserves per-run fonts). Word-fill/spacing/geometry-adjust use the re-render primitive.

**Task R1 — misraParaXml: add shrunk paragraph mark (word-html.js).** In the word-fill branch, also emit `<w:pPr><w:rPr><w:sz w:val="4"/><w:szCs w:val="4"/></w:rPr>…` so the trailing empty line (from the break) is ~2pt. Spike confirmed the empty-line height comes from the paragraph mark, not the break run. TDD: assert the pPr rPr sz=4 present in word-fill output; absent in non-word-fill. Also drop the now-unused `wordFillFont` param if trivial (font comes from the render, not per-cell) — optional.

**Task R2 — justifySelectionWordFill → reconstruct + insertPoem (taskpane.js, Word-verify).** Replace the per-cell insertOoxml body (commits bdc065b/6619340/ec97f86 superseded) with: find enclosing "Ashaar Poem" CC or selection tables → read each table's rows→cells text → `AshaarTableAdopt.adoptTableToSource(rows, {direction:"rtl"})` per table, join stanzas with "\n\n" → set `input.value = source` (like adoptTable) → select the CC/table range → `await insertPoem(true)` (opts already carry justifyMode:"css" + strength/width/gap/fontMode from the pane). Column expansion (§4 ~15%) is realized by the width the render uses (scaledTextWidth), not by poking columns. Verify in Word on marsiya-test.docx: cells kashida-stretch, tiny empty line, no error.

**Deferred / coordinate with other worktree:** §8 Nastaliq cursor visibility + wider Urdu Nastaliq justification live in a separate worktree; do not touch here (keep mergeable).
