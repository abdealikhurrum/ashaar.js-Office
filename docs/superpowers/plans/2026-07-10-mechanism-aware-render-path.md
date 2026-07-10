# Mechanism-Aware Render Path Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: superpowers:subagent-driven-development or superpowers:executing-plans. Steps use checkbox (`- [ ]`) syntax.

**Goal:** Make the render/insert path apply the **same per-font kashida mechanism** as the Justify button, so **Insert (Conversion), §3 word-fill re-render, and Justify** all agree: Jameel → font-swap, Mehr → discrete trailing tatweel, whitespace fonts (Noto/Gulzar/Arabic-serif/document) → spacing, non-registry fonts → generic tatweel engine.

**Why:** `justifyText`/`renderForWordOoxml` (the insert path) is mechanism-unaware — it always runs the generic tatweel engine for kashida mode. So Conversion-insert (and word-fill re-render, which calls `insertPoem`→`renderForWordOoxml`) emits generic tatweels for Jameel/Mehr instead of their real mechanisms. The dispatch exists only inline in `justifySelection` (`taskpane.js`).

**Architecture:** Extract a single per-misra dispatcher used by both paths. It returns a **full cell paragraph OOXML** (because font-swap emits multi-run `w:cs`, not just modified text). Signature (pure; caller supplies a `measure(text, faceKey)` fn so it's node-testable):

```
renderMisraCellXml(text, colPx, align, isRefrain, opts, measure) -> paragraph OOXML string
```

Dispatch by `AshaarFonts.mechanismOf(opts.fontMode)`:
- **font-swap (Jameel):** `AshaarKashidaFontswap.splitSpans` → measure each fasl base vs Kasheeda → `selectSwapRuns` → `runsToMisraXml(runs, align, opts)`.
- **tatweel (Mehr):** discrete trailing-tatweel subset selection (extract the logic currently inline at `taskpane.js:~1464`) → text with trailing tatweels → `misraParaXml(text, ...)`.
- **whitespace (Noto/Gulzar/arabic-serif/document) in kashida mode:** downgrade to spacing (`justifyWordSpacing`) → `misraParaXml(text, ...)`.
- **generic (non-registry font, kashida mode):** `AshaarJustify.justifyLine` → `misraParaXml(text, ...)` (today's behavior).
- **spacing mode / none:** existing `justifyText` behavior → `misraParaXml`.

## Global Constraints
- ES5/UMD, no build step. Never edit `src/vendor/`.
- Pure logic node-tested (fake `measure`); Office.js/Word rendering verified manually.
- Reuse existing helpers — do NOT duplicate: `AshaarFonts.mechanismOf/wordNameOf/kasheedaNameOf/tatweelRulesOf`, `AshaarKashidaFontswap.splitSpans/selectSwapRuns`, `AshaarWord.runsToMisraXml/misraParaXml/justifyWordSpacing`, `AshaarJustify.justifyLine`.
- `npm test` green.

---

### Task 1: Extract the Mehr discrete-tatweel selection into a pure helper
**Files:** `src/taskpane/word-html.js` (or a shared module) + `tests/word-html.test.js`.
Extract the inline Mehr logic from `taskpane.js:~1464-1478` to `AshaarWord.discreteTatweelFill(text, colPx, opts, measure)` → returns the text with trailing tatweels on selected words (uses `tatweelRulesOf(...).finalInto` + `AshaarKashidaFontswap.selectSwapRuns`). TDD with a fake `measure` where the elongated form is wider only for whitelisted-final words. Then refactor `justifySelection`'s Mehr block to call it (no behavior change — verify tests + a Word re-check of Mehr).

### Task 2: The dispatcher `renderMisraCellXml`
**Files:** `src/taskpane/word-html.js` + `tests/word-html.test.js`.
Implement the dispatcher above. TDD each branch with a fake `measure`:
- font-swap → output contains per-run `w:cs="…Kasheeda"` + base runs (assert via `runsToMisraXml`).
- tatweel → output is a `misraParaXml` with trailing tatweels on the selected words.
- whitespace kashida → `misraParaXml` with micro-spaces.
- generic kashida → `misraParaXml` with tatweels from `justifyLine`.
Keep `align`/`isRefrain`/`indTwips` handling identical to `misraParaXml`.

### Task 3: Wire `renderForWordOoxml`'s misra emitters to the dispatcher
**Files:** `src/taskpane/word-html.js` + tests.
Replace the `misraParaXml(justifyText(text, opts, px), …)` calls in `misraRow` and `soloRow` (and any other misra emitters) with `renderMisraCellXml(text, px, align, isRefrain, opts, measure)`, where `measure` wraps `opts._justifyCtx.measureText` at the right face (base/Kasheeda). Preserve the soloRow misra-width + pad-cell layout (a font-swap solo still occupies the BASE_CPM cell). Update/extend the poetry-corpus/word-html OOXML assertions.

### Task 4: `insertPoem` — preload the mechanism's faces in the measurement canvas
**Files:** `src/taskpane/taskpane.js` (Office.js — manual verify).
Mirror `justifySelection`'s preload: before rendering, if the selected font's mechanism is font-swap or tatweel, `await document.fonts.load` the base (and Kasheeda) faces so `renderForWordOoxml`'s `measure` is accurate. Otherwise the dispatcher mis-measures (same class of bug as the Mehr canvas fix).

### Task 5: (Optional) Unify `justifySelection` onto the dispatcher
Refactor the Justify-button path to call `renderMisraCellXml`/`discreteTatweelFill` so both paths share one implementation. Reduces the duplication that caused this divergence. Guard with the existing tests + Word re-check.

### Task 6: Word verification
Insert (Conversion → Insert as Table) Arabic poems with **Jameel** (expect font-swap Kasheeda, no tatweel chars), **Mehr** (expect discrete trailing tatweels), **Gulzar/Noto** (expect spacing). Confirm §3 word-fill and engine-Justify still agree. Confirm the Jameel distinct-family font renders.

## Notes
- Coordinate with the `worktree-nastaliq-kashida-fonts` branch — this deeply integrates its mechanisms; keep the dispatcher's behavior consistent with that engine's gates (G/G2).
- The Jameel Kasheeda distinct-family font (`~/Library/Fonts/JameelNooriNastaleeqKasheeda-DistinctFamily.ttf`) is a local install prerequisite (private font; not committed).
