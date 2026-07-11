# In-Place Adjust + Per-Cell Mechanism Auto-Detect — Design

**Date:** 2026-07-10
**Status:** IMPLEMENTED + Word-verified (per-cell auto-detect, size-preserving re-render, in-place table-width). Commits `d959239`, `84f2032`, `f23684c`.
**Branch:** `feat/guided-justification-ux`

## Known limitation (documented 2026-07-11)

**Middle gap has no in-place / dedicated re-render trigger for engine-kashida poems.** Because the gap is structural in the fixed grid (see "Design finding" below), a gap change must **re-render** the poem. That re-render is now **size-preserving**, but the only paths that trigger it are **Insert as Table** (Conversion) and **"Word justify"** (css word-fill) mode — both of which reconstruct + replace the selection. For a poem being justified via the **engine kashida** path (Jameel/Mehr/generic, in place), there is **no button that rebuilds the existing managed table with a new gap**. Mode/strength and table-width % apply in place on Justify; **gap does not**. Workaround today: change the gap and re-insert / use Word-justify mode (size is preserved). A dedicated "rebuild this managed table with the current gap" action for engine-mode poems is a tracked follow-up, not yet built.
**Scope:** Two related changes to the justify/adjust workflow: **(1)** resolve the kashida mechanism **per cell from its real font** (so any dropdown, incl. Document default, justifies each font correctly and in place); **(2)** apply geometry + mode changes — **middle gap, table width %, kashida sub-mode** — to the **existing** managed table **in place** (column resize + in-place re-justify), instead of the current re-render-and-replace-selection, which rebuilds the table and reverts font size to Word's 12pt default.

## Problem (from Word testing, 2026-07-10)

- Jameel/Mehr/generic kashida justify works great and is idempotent (verified). But two gaps remain:
- **B — Document default breaks Jameel/Mehr.** Mechanism dispatch keys off the pane font dropdown (`justifySelection`, `mechanism = AshaarFonts.mechanismOf(fontId)`); under "Document default" (mechanism `whitespace`) the Jameel/Mehr branches don't fire, so their cells fall to the generic path, which injects medial tatweels that shatter Jameel/Mehr shaping.
- **A — Adjusting properties requires copy-and-replace.** Middle gap (`gapWidth`), table width (`tableWidthPct`), and the font-based kashida sub-mode only reach the document through `renderForWordOoxml` → replace the selection. That rebuild reverts font size (`misraParaXml` emits `w:rFonts`/`w:jc` but **no `w:sz`**; only the in-place `runsToMisraXml` emits size) and churns the table/content-control. There is no way to push a gap/width/mode change onto the existing table.

## Key insight

Both are the same move: **make the mechanism follow the real font per cell, and apply geometry in place** — so the everyday adjust/justify never rebuilds the table. The native-Word-kashida ("Word justify"/`justifyMode:"css"`) mode is the one exception that *must* re-render (Office.js paragraph-alignment API has no kashida value; per-cell `insertOoxml` drops the kashida `jc` — verified in the §3 spike); it stays as-is and out of scope here.

## Decisions locked (from brainstorm)

- **Per-cell** mechanism detection (not per-run mixed — that stays ashaar-js#7).
- Dropdown becomes the **fallback** for cells with no resolvable font, not the mechanism source.
- In-place adjust is scoped to **managed tables** (Ashaar Poem content controls) first; raw hand-made tables are out of scope for gap adjust.
- Native-Word-kashida (`css`) mode keeps its re-render path (API limit); keep-vs-retire is deferred (see Open questions).

---

## Component 1 — Per-cell mechanism auto-detect

### 1a. Registry: `descriptorForFontName`
**File:** `src/taskpane/fonts.js` (+ `tests/fonts.test.js`).

```
AshaarFonts.descriptorForFontName(name) -> descriptor
```
Returns the registry descriptor (`{id, wordName, kasheedaName, mechanism, tatweelRules, …}`) whose `wordName` or `kasheedaName` matches `name` (trimmed, exact); for an unrecognised/empty name, returns a synthetic **generic** descriptor `{id:"generic", mechanism:"generic", wordName:null, kasheedaName:null, tatweelRules:null}`. `mechanismForFontName` (already shipped) becomes a thin wrapper returning `descriptorForFontName(name).mechanism`. Pure; node-tested (each registry font → its descriptor; unknown/"" → generic; Kasheeda name → the Jameel descriptor).

### 1b. Per-cell dispatch in `justifySelection`
**File:** `src/taskpane/taskpane.js` (Office.js — manual verify).

Move the mechanism decision **inside** the phase-1 per-cell loop. For each cell, resolve its descriptor from the cell's **dominant real font** (the representative run font; fall back to the pane dropdown font only when the cell reports none). Dispatch the cell to the existing verified branch by `descriptor.mechanism`, using the **descriptor's** `wordName`/`kasheedaName`/`tatweelRules` — not the global `fontId`:
- `font-swap` → Jameel branch (measure base vs `descriptor.kasheedaName`, `selectSwapRuns`, `runsToMisraXml` + Task-3 residual spacing).
- `tatweel` → Mehr branch (`descriptor.tatweelRules`, discrete trailing tatweel + Task-4 residual spacing).
- `generic` → generic run-aware `justifyRuns` path.
- `whitespace` → spacing.

The existing branches are reused unchanged except that their font identifiers come from the per-cell descriptor. The current top-level `mechanism`/`fontId` (dropdown) is retained only as the per-cell fallback.

### 1c. Force-load each cell's faces
Extend the existing pre-measure `document.fonts.load` block so that, for every cell resolved to `font-swap`, its `kasheedaName` face is loaded (and each cell's base face) — not only the single dropdown font's faces. Without this the canvas mis-measures a Jameel cell under a non-Jameel dropdown. (The distinct-run-font force-load added earlier already covers base faces; this adds the per-cell Kasheeda faces.)

**Result:** under any dropdown, each cell justifies by its true font — in place, size preserved, no breakage.

---

## Component 2 — In-place Apply for geometry + mode

### 2a. The operation
A single **"Apply"** action reads the current pane settings (gap, width %, justify mode/strength) and pushes them onto the managed table under the cursor **in place** — no `renderForWordOoxml`, no selection replace. Steps:

1. **Locate** the enclosing Ashaar Poem content control; read its tag payload (`parseContentControlTag`) for the current `layoutMode`/`misraPattern`/`misraCount`/`gapWidth`/`tableWidthPct` and the current cell texts.
2. **Recompute the grid** deterministically from those layout params + current cell texts using the existing model (`BASE_CPM`, `allocateSpans`, `misraSpans`, `gapCols`), yielding the intended per-column twip widths for the new `tableWidthPct` and `gapWidth` — including **which grid columns are gaps**.
3. **Resize columns in place** via the Word Table column API (`tbl.columns.items[j].width`, the same API `applyProfileToQaseeda` uses): scale content columns to the new table width %, set the gap column(s) to the new gap width.
4. **Re-justify each cell in place** via Component 1 (per-cell mechanism), so a mode/strength/mechanism change is realized without rebuild.
5. **Update the content-control tag** with the new geometry so the block stays self-describing.

Font size, manual cell edits, and the table/content-control identity are all **preserved** because the table is never rebuilt.

### 2b. Pure helper (node-testable)
Extract the grid→column-width computation as a pure function, e.g.:
```
AshaarWord.gridColumnWidths(layoutParams, cellTexts, pageTwips) -> { widths:[twips…], gapIndices:[…] }
```
so the twip math + gap-column identification is unit-tested independently of Office.js. The Office.js resize/justify glue is manual-verified.

### 2c. Capability + fallback
In-place column resize needs **WordApiDesktop 1.3** (`Office.context.requirements.isSetSupported("WordApiDesktop","1.3")` — the existing `canResize` check). When available (desktop Word): full in-place Apply. When **not** available: fall back to the current re-render path and surface a message that geometry was applied by rebuild (documented degradation), OR apply mode-only in place and skip resize — the plan picks one and states it.

### 2d. UI
Reuse/keep the existing gap / table-width / mode controls; the change is that an **Apply** action now targets the existing table in place. (Whether Justify and Apply are one button or two is an implementation detail for the plan; the natural choice is that **Justify already re-applies mode in place**, and **gap/width changes gain an in-place Apply** rather than requiring re-insert.)

---

## Testing

- **Node (pure):** `descriptorForFontName` (Component 1a); `gridColumnWidths` (Component 2b — width scaling, gap-column indices, edge cases).
- **Manual in Word (Office.js):** per-cell dispatch under Document default (Jameel/Mehr/Fatemi/Noto cells in one selection each justify by their own font, no shatter); in-place Apply of a gap change, a width change, and a mode change on a managed table — confirm the table is **not** rebuilt, font size is **retained**, and the content control persists; idempotent re-Apply.

## Out of scope (tracked)

- **Per-run mixed mechanisms** (a Jameel word beside a Fatemi word in one misra) → ashaar-js#7.
- **Prose justification** → ashaar-js#8.
- **Raw (non-managed) table** gap adjust — needs layout inference; managed tables first.
- **Native-Word-kashida (`css`) mode** stays re-render (Office.js API limit); no change here.
- Font-size preservation in the `css`/insert *render* path is a separate, smaller fix (carry `w:sz` through `misraParaXml`); note it but don't bundle it — the in-place Apply makes it unnecessary for the adjust workflow.

## Open questions

- **Keep vs retire the "Word justify" (`css`) mode** now that in-place engine kashida covers all fonts. Deferred; revisit after Component 1+2 land.
- **One button or two** (Justify vs Apply) — resolve in the plan against the current pane layout.
