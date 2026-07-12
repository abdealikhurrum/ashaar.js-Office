# Session Handoff — 2026-07-12 (width-engine rebuild + justify polish)

**Read this first.** Branch: `feat/guided-justification-ux`. All work **local, committed, unpushed, unsigned** (`--no-gpg-sign`). Node suite green (17 files). Dev server runs via `npm start` (port 3000, `Cache-Control: no-store`); Word reloads JS only when you **reload the task pane**.

## Headline: the width engine now works in Word (verified)

Last session shipped "Option A" (resize span tables via `columns.setWidth(slotPt,"SameWidth")`). **That approach is dead** — this session proved it garbles span (merged-cell) tables. The apply engine was re-architected to **rebuild the OOXML** and it now works end-to-end in Word.

### What was wrong & the fix (root-caused live)
- `columns.setWidth(pt,"SameWidth")` does NOT set per-grid-column widths on a span table. `RulerStyle.sameWidth` sets *all cells* to one value → grid columns collapse, equal-span cells render at wildly different widths (one misra went 1-char-per-line **vertical**). Per-column `setWidth` throws; `cell.columnWidth` is "uniform tables only". **There is no Office.js API to set per-grid-column widths on a span table.** (Memory `office-js-word-constraints.md` item 2b corrected; new memory `width-engine-rebuild-not-setwidth.md`.)
- **Fix = rebuild.** Only re-authoring the table OOXML (`renderForWordOoxml` writes `<w:gridCol>`/`<w:tcW>`) resizes correctly — same thing the **Re-render** button does.

### Commits this session (all Word-verified except where noted)
- `dad48a3` `stanzaCellGeometry`/`poemCellGeometry` — source-derived grid spans (Word can't report them for span tables). Node cross-check locks spans/kinds/cols to the generator.
- `bd99715` first Option-A apply rewrite (setWidth) — **superseded**, keep for history.
- `36c230f` **rebuild-based apply**: two passes — SIZE (rebuild every block at one shared `targetTwips` via `uniformSlotPx`, margin-aware → no wrap; same width for all bandhs → same-GRID bandhs get identical `cwt` = harmony) then JUSTIFY (fill each cell to `box = span×(target/GRID) − margins`, kashida + spacing, clamped no-wrap). Factored `captureQaseedaTables` (load/measure/geometry) shared by both passes.
- `c005489` **SDT wrap**: `AshaarWord.wrapOoxmlControl` emits the content control *inside* the OOXML (block-level `w:sdt`) so it spans the whole table (not just row 1). Also closes the old multi-table re-wrap gap.
- `f53f84e` removed temp DBG readout.
- `1c1e1c1` **rebuild-gating**: only rebuild when the width signature (`targetTwips` + block source hashes, cached in-session in `_appliedSizeSig`) changes. Strength / fill mode / per-cell override are justify-only → skip the destructive rebuild → fast, non-destructive, and fixed the Word **crash** from repeated content-control surgery.
- `e7dde61` **Cell-fit uses real micro-spaces** (not Word `distribute`, which no-ops on a single line). Both fill modes now share the micro-space(+tatweel) path; they differ only in target — Cell-fit → cell edge, Natural-fit → harmony width. Fill mode is now visibly distinct in both justify modes.
- `9f165ef` space debug tint via `font.highlightColor` (spaces have no ink for `font.color`).
- `7ee982f` **Nastaliq fix**: the justify pass injected generic tatweels for every font → shattered Nastaliq (Noto = whitespace-shaping, Jameel = font-swap). Added the per-cell guard `justifySelection` uses (`AshaarFonts.mechanismForFontName`): only `generic`/`tatweel`(Mehr) fonts elongate; `whitespace`/`font-swap` fall to spacing fill even under a kashida profile. **Needs a Word re-check on a Nastaliq qaseeda.**

## Open observations / follow-ups
- **Cell-fit "doesn't do much"; Natural-fit preferred (user, 2026-07-12).** Expected under auto-fit: cells are sized snug to the text, so "fill to cell edge" (Cell-fit) ≈ "fill to harmony width" (Natural-fit). They diverge under fixed-% (wide cells). Consider making **Natural-fit the default** (and/or de-emphasizing Cell-fit) when Phase 4 UI lands.
- **Font-aware kashida in apply is partial.** `applyProfileToQaseeda` only does generic tatweels (+ spacing fallback). It does NOT yet do Jameel **font-swap** or Mehr's own kashida the way `justifySelection` does — those get safe spacing fill in the multi-block apply. Porting the full per-run mechanism dispatch into the apply/`formatBlocks` core is a real follow-up (do it during Phase 3).

### Verified in Word this session (user-confirmed)
Auto-fit + fixed-% sizing; harmony across bandhs; no vertical text / no wrap; kashida tatweels + spacing both fill; full-table content-control wrap; strength sweep; spacing mode; idempotent re-apply; per-cell override; rebuild-gating (override/strength now light, no crash). Space-highlight + final Cell-fit-vs-Natural-fit distinction were the **last things handed for a look** — assume good unless the next session hears otherwise.

## ⏭️ NEXT — remaining plan phases (plan: `docs/superpowers/plans/2026-07-11-unified-formatting-width-engine.md`)
The width engine (Phases 1–2) is DONE. Left:
- **Phase 3 (Task 6)** — `activeProfile()`; route `insertPoem`/`adoptTable`/`justifySelection` through the profile (not the top controls) so a fresh insert is sized/justified like an applied one. Factor the SIZE+JUSTIFY core into `formatBlocks` and call from both.
- **Phase 4 (Tasks 7–9)** — profile dropdown + "＋ New"; block-first active-context sync; **remove the duplicated top formatting controls** (`#justify-mode`, `#table-width`, `#width-mode`, `#auto-fit*`, `#tatweel*`, `#gap-width`) so the profile is the sole source.
- **Phase 5 (Task 10)** — `AshaarProfiles.DEFAULT_PROFILE_NAME`/`resolveProfileName`; older-Word fallback message (node-testable, independent — good quick start).
- **Phase 6 (Task 11)** — manual Word verification checklist.

## Key facts so the next session doesn't rediscover them
- **Never edit `src/vendor/`.** Pure modules (`natural-width-matrix.js`, `word-html.js` helpers) stay DOM/Office-free + node-tested.
- **Resize span tables ONLY by OOXML rebuild.** No `columns.setWidth`, no `cell.columnWidth`. See memory `width-engine-rebuild-not-setwidth.md`.
- **Grid geometry comes from the SOURCE** (`poemCellGeometry`), never from Word. Geometry uses each block's OWN stored opts (tag payload gapWidth/layoutMode) so it matches the render.
- **Content control = block-level `w:sdt` in the OOXML** (`wrapOoxmlControl`), not `insertContentControl()` on an inserted range (that caught only row 1 on Mac Word).
- **Rebuild is gated** by `_appliedSizeSig[name]` (in-session). Justify-only changes skip it. First apply per session always rebuilds.
- `applyProfileToQaseeda` = two passes: pass 1 SIZE (rebuild if width changed) → pass 2 JUSTIFY (fresh gather, fill to box). `captureQaseedaTables` is the shared load/measure/geometry helper.
- `stripJustification` makes re-apply idempotent (strips old tatweels/spaces before re-filling).

## Continuity memory (auto-loaded)
`~/.claude/projects/-Users-abdealikhurrum-ashaar-js-Office/memory/`: `width-engine-rebuild-not-setwidth.md` (new), `office-js-word-constraints.md` (2b corrected), `justification-modes-state.md` (updated), `cell-configurations-state.md`, `MEMORY.md` index.
