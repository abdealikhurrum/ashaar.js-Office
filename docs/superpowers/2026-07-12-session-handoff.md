# Session Handoff — 2026-07-12

**Read this first.** Branch: `feat/guided-justification-ux`. All work is **local, committed, unsigned** (`--no-gpg-sign`, 1Password agent locked) — **not pushed, no PR**. Node suite green (16 files). Dev server was running via `npm start` (port 3000, `Cache-Control: no-store`).

## What this session did

Started as Task 7 (manual Word verification of the justification modes) and turned into: **fix 3 real Word bugs**, then discover a **design gap** in the width/harmony engine, brainstorm + spec + plan it, and begin implementing.

### Bugs found & fixed (committed, Word-verified)
- `cc03e55` **Re-render in-place replace** — `insertOoxml("Replace")` on a whole content-control range throws `GeneralException`; now detect the enclosing "Ashaar Poem" control via range `intersectWithOrNullObject` and replace in place (insert after + wrap + delete old). Also added `describeError()` (surfaces `debugInfo.errorLocation` + `surroundingStatements` to the pane — this is how every Office.js error got diagnosed).
- `e760669` **Qaseeda apply** — (a) `TableCell.shadingColor` rejects `""` and `"No color"`; clear with `"#FFFFFF"`. (b) auto-fit column resize invalidated captured cell proxies → re-map fresh proxies after resize.
- Justify Selected Text (kashida) confirmed working.

### Design gap found (the reason for the new spec/plan)
Save & Apply **cannot size/harmonize span-based poetry tables**: `width.pct` is dead code in apply; auto-fit crashes because `Table.columns`/`TableCell.columnWidth` are "uniform tables only"; and formatting is set in two disconnected places (top controls vs qaseeda profile). Full analysis in the spec.

### Brainstormed → spec → plan (committed)
- Spec: `docs/superpowers/specs/2026-07-11-unified-formatting-width-engine-design.md`
- Plan: `docs/superpowers/plans/2026-07-11-unified-formatting-width-engine.md`
- Design in one line: **one block-first profile model** (drop the duplicated top controls; profile dropdown you can add to) + a **span-safe width engine** that computes widths from structure and sets them, reusing the same box as the kashida target (no wrap).

### Execution progress (inline, via executing-plans)
- ✅ **Spikes done.** Per-column `setWidth` is **impossible** on span tables (`columns.load` throws "mixed cell widths"); only collection-level uniform `columns.setWidth(pt, "SameWidth")` works. → **Decision: Option A — scale equal slots.** (Non-uniform cell widths survive because they come from integer spans of equal slots.) Recorded as a plan AMENDMENT block above Phase 2.
- ✅ **Phase 1 pure math done + node-tested + committed:** `AshaarMatrix.computeTargetGrid` (`ed2c4c2`, `286b2ee` — the Option-B/per-column reference, kept but NOT on the shipped path) and **`AshaarMatrix.uniformSlotPx`** (`f197af5` — the Option-A slot sizer). Existing `buildMatrix`/`naturalFitTarget`/`cellFitBudget` reused.

## ⏭️ NEXT STEP — the apply-engine rewrite (Option A)

This is exactly where to resume. The math is ready; the Word orchestration is not. Verifiable only in Word (no headless test).

1. **`word-html.js` geometry helper** — expose per-cell grid geometry from the poem SOURCE (GRID + each cell's span, row-major), reusing the private `stanzaGridInfo`/`stanzaColSpans`. Word cannot report this for span tables, so it must come from structure.
2. **Rewrite the width step of `applyProfileToQaseeda`** (`taskpane.js`, the old `canResize`/`info.tbl.columns` block — now deleted-worthy):
   - Reconstruct each block's source (as re-render does) → geometry helper → per content cell: `natural` (canvas px), `span`, table `GRID`.
   - `slot = AshaarMatrix.uniformSlotPx(bandhs, {mode, pct, pagePx, headroom})`; set it with `table.columns.setWidth(slot_pt, "SameWidth")` (the ONLY width op that works).
   - Stamp `c.box = span × slot`, `c.wpos = buildMatrix()[key]`.
3. **Rewrite per-cell justify** to fill to the box under **both** kashida and spacing (ungate from `doKashida`): natural-fit `target = naturalFitTarget(wpos, box, phi)`; cell-fit `cellFitBudget(natural, box, phi)`; **clamp `target ≤ box`** (no-wrap). Reuse `justifyRunsConcentrated`/`distributeMicroSpaces`.
4. Then Phase 3 (route insert/adopt/justify through the active profile), Phase 4 (UI: profile dropdown + block-first + remove top formatting controls), Phase 5 (Default profile + older-Word fallback), Phase 6 (manual Word verification checklist).

## Key facts so the next session doesn't rediscover them

- **Never edit `src/vendor/`.** Pure modules stay DOM/Office-free (node-testable).
- **Office.js width reality (memory `office-js-word-constraints.md`):** span tables expose NO column geometry; only uniform `columns.setWidth(pt,"SameWidth")` works; `shadingColor` clears with `"#FFFFFF"`; resize invalidates cell proxies; assignment errors surface at `context.sync()` (a `try/catch` around a property set never fires).
- **`marsiya-test.docx` is PLAIN tables** (no managed blocks). To manage one: **Adopt Existing Table** (structure-aware). **Load Selection is the WRONG tool for tables** — it grabs raw `selection.text` (tabs/returns, no misra `\`) and garbles the rebuild; that was the source of earlier corruption confusion.
- **Poetry tables are fixed-layout** (`stanzaTableOoxml`: `tblLayout fixed`, `tblW dxa`, uniform `gridCol=cwt`; cell of span S = S·cwt). Twips↔px `×96/1440`; pt↔px `×96/72`.
- Two justify entry points share the per-cell model: `justifySelection` (one block) and `applyProfileToQaseeda` (all blocks of a qaseeda).

## Known-open follow-ups (out of scope for the current plan)
- Re-render / multi-table Adopt **content-control re-wrap** (partial wrap of multi-table ranges).
- Adopt **"Review before replacing"** checkbox has no effect.
- Fill mode being a no-op in Spacing mode — **folded into** the current plan (Phase 2 Task 5), not separate.

## Continuity memory (auto-loaded)
`~/.claude/projects/-Users-abdealikhurrum-ashaar-js-Office/memory/`: `office-js-word-constraints.md` (new), `justification-modes-state.md`, `cell-configurations-state.md`, `MEMORY.md` index.
