# Design — Unified formatting model + span-safe width engine

**Date:** 2026-07-11
**Status:** Approved (brainstorm), pending implementation plan
**Branch:** `feat/guided-justification-ux`

## Problem

Task 7 (manual Word verification of the justification modes) surfaced that **Save & Apply cannot size or harmonize span-based poetry tables**, and that formatting parameters are set in **two disconnected places**.

Concrete defects found in Word:

1. **Fixed % is dead code in apply.** `applyProfileToQaseeda` never reads `profile.width.pct`, so changing the % does nothing — the table is never sized to the page.
2. **Auto-fit crashes on span tables.** The resize path calls `table.columns.load("width")`, but `Table.columns` / `TableCell.columnWidth` are documented "applicable to uniform tables." The marsiya tables use `gridSpan` cells of varying widths, so Word throws `GeneralException: Cannot access individual columns … the table has mixed cell widths`.
3. **Kashida wraps.** Because widths are never equalized and `colPx` is read from the unreliable `TableCell.columnWidth`, the natural-fit target can exceed the real (fixed-layout) cell width; the cell can't grow, so text word-wraps.
4. **Fill mode is a silent no-op in Spacing mode.** All elongation is gated behind `doKashida = profile.justify.mode === "kashida"`, so Fill mode (natural-fit / cell-fit) does nothing when the justify mode is Spacing.
5. **Duplicated controls.** The top-of-pane insert controls (`options()`) and the Qaseeda profile both set width, justify mode, fill mode, strength, and gap — at different scopes/lifetimes, implemented on **separate code paths** (the apply path being the broken one). This is historical layering never reconciled.

Root insight: the poetry tables built by `stanzaTableOoxml` are **fixed-layout** (`tblLayout fixed`, `tblW dxa`, uniform `gridCol = cwt`). Widths are therefore deterministic and known from structure — the bug is purely that apply (a) never resizes them correctly and (b) reads width back through an API that rejects span tables, instead of computing it.

## Goals

- Save & Apply produces **equal per-position widths across all same-shape bandhs** of a profile (harmony), sized correctly, with **no word-wrap**, in **both** auto-fit and fixed-% modes.
- **One** formatting model (width / justify / fill / strength / gap / font) is the single source of truth, shared by insert, adopt, justify, and apply — so the two paths can never disagree.
- Width and fill are **implemented once**, structure-driven (computed, not read back).

## Non-goals (separate follow-ups)

- The re-render / multi-table Adopt **content-control re-wrap** bug (partial wrap of multi-table ranges).
- Adopt's **"Review before replacing"** checkbox having no effect.
- Any **task-pane layout overhaul** (tabs/panels/ordering) beyond folding the duplicated formatting controls into one panel.

## Section 1 — Unified formatting model (data + UX)

**One concept: the profile (a.k.a. qaseeda).** Holds all *formatting*: `width {mode: "auto-fit"|"fixed", pct}`, `justify {mode: "kashida"|"spacing", strength, fillMode: "natural-fit"|"cell-fit"}`, `gap`, `font`, plus decoration/correction extras. Already exists in `profiles.js` and in each block's content-control tag — consolidated onto, not invented.

**Formatting vs. structure.** Only *formatting* unifies. *Structural* choices (misra count, spans, bandh layout, layout spec, font-for-drawing) stay in the Table Input / insert flow and remain in `options()`.

**One formatting panel, block-first / active-context:**
- A **Profile dropdown** lists every profile in the document plus **"＋ New profile…"** to add one inline.
- Cursor enters a block → the dropdown selects that block's profile and the panel fills with its settings (reuses the existing active-block sync, `reflectActiveContext`).
- Not in a block → the panel shows what a new Insert/Adopt will use.
- Edits stage against the selected profile.
- **Apply** re-formats every block bound to that profile in one batch (explicit; no auto-apply). Per-block overrides (SP2) still layer on top.

**The top-of-pane width / justify / fill / strength controls are removed.** Those parameters live only in this panel. Insert, Adopt, Justify Selected Text, and Save & Apply all read the same profile.

## Section 2 — Span-safe width engine

Runs on **Apply** and at **Insert/Adopt** (to size a fresh block). Principle: **compute all widths from the known table structure, set them, and reuse the same numbers for kashida** — never read back `Table.columns` / `TableCell.columnWidth`.

Steps:

1. **Gather** every block bound to the profile; read each cell's text + font (as today).
2. **Measure & matrix** — canvas-measure each cell's natural width; build the cross-block **position matrix**: for each position label (A1, A2, …), the longest natural line across all bandhs.
3. **Compute the target grid** (pure, testable math):
   - *Auto-fit:* each position's width = its longest natural line + kashida headroom; total capped at the page (text width = page − margins).
   - *Fixed %:* total = `pct × pageWidth`; positions sized proportionally to the **matrix** widths (cross-bandh longest per position), not each cell's own natural — so the vector is shared.
   - *Harmony:* because both modes size from the shared position matrix, same-shape bandhs all receive the **identical** width vector → matching cells line up by construction. Different-shape bandhs equalize on total width (existing proportional fallback).
4. **Set the widths** — spike-gated in-place resize (`TableColumnCollection.setWidth`, WordApiDesktop 1.3). If the spike shows the API rejects span tables, fall back to **rebuilding** that table via the proven fixed-layout `stanzaTableOoxml` at the computed grid.
5. **Justify** — each cell fills to *its computed target* (natural-fit → position-matrix width; cell-fit → the cell edge). Because the target is the box just set, text cannot overflow → **no wrap**.

**Fallback when no width API and rebuild is off:** justify at current widths with a clear message; the kashida target still uses the *known structural* width, so it still won't wrap.

**Where it lives:** the pure grid/matrix math extends `natural-width-matrix.js` (node-tested); the Word orchestration stays in the apply function, now much simpler (no `.columns`).

## Section 3 — Folded-in fixes & explicit scope

Folded in (part of "implement width/fill once, correctly"):
- **Fill mode works under both justify modes.** Fill mode drives the width-fill under both Kashida (tatweels) and Spacing (micro-spaces) — no silent no-op.
- **Fixed % actually sizes the table** (currently dead code in apply).
- **Older-Word fallback is explicit:** resize API absent + rebuild off → justify at current widths and say so plainly.

Out of scope (logged as separate follow-ups): re-render / multi-table Adopt re-wrap bug; Adopt "Review" checkbox no-op; task-pane layout overhaul.

## Section 4 — Data flow & migration

| Action | Reads | Width behavior |
|--------|-------|----------------|
| Insert as Table / Adopt | current profile (dropdown / active block) | sizes the new block via the width engine |
| Justify Selected Text | the active block's profile | fills cells to computed targets |
| Save & Apply | the selected profile | width engine across all blocks bound to it (harmony) |

- **Profile ↔ block binding:** a block stores its profile name in its content-control tag (existing `qaseeda` field). The dropdown lists profiles from the document's saved store. Re-selecting a profile for the block at the cursor re-binds it (takes effect on next Apply).
- **Migration (lazy, no document rewrite on load):** existing v2-tagged blocks keep working — their stored recipe maps onto a profile; blocks with no `qaseeda` resolve to a **"Default"** profile so the dropdown always has a selection; a block adopts its profile when first applied/edited.
- **Removing the top controls:** insert/justify paths switch from `options()` to the active profile; `options()` keeps only structural fields (layout spec, misra/bandh counts, font-for-drawing) that Table Input needs.

## Section 5 — Testing

**Pure math (node — primary confidence):**
- Position-matrix: longest-natural per position across multiple bandhs, mixed fonts.
- Target-grid: auto-fit (longest + headroom, page cap), fixed-% (pct × page, proportional), same-shape harmony (identical width vector), different-shape equalize-on-total.
- Kashida target = computed box (natural-fit → matrix width; cell-fit → edge); invariant **target ≤ box** (no-wrap guarantee).
- Migration/defaults: no-`qaseeda` block → "Default"; v2 tag → profile mapping.

**Word spike (gates the mechanism — first implementation step):** insert a real span-based poem table, try in-place `setWidth`, confirm Word accepts it and the grid changes. Decides in-place resize vs. rebuild fallback. Throwaway, not shipped.

**Manual Word verification (Task 7 successor):**
- Two same-shape bandhs, one profile → matching cells equal width, no wrap, in both auto-fit and fixed %.
- Strength sweep 1→10 behaves (natural widths → filled to box).
- Fill mode visibly differs under both Kashida and Spacing.
- Older-Word path (if reproducible) shows fallback message, still no wrap.
- Per-block override still layers on top.

**Not automated:** Office.js resize/proxy behavior (covered by the spike + manual pass).

## Key technical decisions

- **Approach #1 (in-place `setWidth`), spike-gated, with rebuild (`stanzaTableOoxml`) as the fallback.** Rejected: bake-width-at-insert-only (breaks auto-fit, which must recompute across siblings at apply time).
- **Width is computed, never read back** from `Table.columns` / `TableCell.columnWidth` (both "uniform tables" only).
- **Harmony = one shared width vector** across same-shape bandhs.

## Reference

Office.js constraints that forced this design are recorded in memory `office-js-word-constraints.md`: `insertOoxml`-replace on content-control ranges, `TableCell.shadingColor` clear value, resize invalidates cell proxies, and `Table.columns`/`TableCell.columnWidth` being uniform-tables-only.
