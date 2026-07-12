# Per-Cell Overrides + Active-Block Sync (SP2) — Design Spec

**Date:** 2026-07-11
**Status:** Draft — user-approved (brainstorming); ready for implementation planning.
**Sub-project:** SP2 of `2026-07-11-cell-configurations-roadmap.md` (SP1 = bandh cell-map, implemented).
**Builds on:** SP1 `2026-07-11-bandh-cell-map-design.md` (per-cell labels `A1/A2/…` + the persisted `cells` map are the per-cell identity this hangs off) and the justification-modes work (Cell-fit/Natural-fit, `AshaarMatrix`, `strengthToElongationShare`/`strengthToMaxPositions`, `capMicroSpaces`).

## Overview

Two connected capabilities, both "reflect what's under the cursor in the pane":

1. **Active-block sync.** When the cursor enters an Ashaar Poem block, the pane reflects that block's settings (font / justify mode / fill mode / strength / width, and its qaseeda profile if named). Editing those and re-applying is the **"change all cells"** lever.
2. **Per-cell justify override.** A sub-panel shows the cell at the cursor (by its SP1 label, e.g. `A1`) and lets it deviate on **strength / target width / cap-lift**. The override is stored on the block, scoped to that **one cell**, and consumed at justify time. This is the **"change one cell"** lever — the §4 under-resolved-cell fix.

The two levers are deliberately distinct: broad changes go through the profile/block settings; a single deviating cell gets a per-cell override.

## §1 — Data model

- **Override payload** (all fields optional; absent = inherit the block/profile default):
  ```
  { strength?: 1–10, widthPt?: number, capEm?: number }
  ```
  - `strength` → the cell's own `φ = strengthToElongationShare(strength)` and `maxPositions = strengthToMaxPositions(strength)`.
  - `widthPt` → the Natural-fit fill **target** for the cell (`px = widthPt·96/72`), bypassing the matrix `Wpos`. (Ignored in Cell-fit, which always fills to the edge.)
  - `capEm` → the `capMicroSpaces` per-gap cap for the cell (default `0.28`), raised to close a residual that otherwise binds short.
- **Scope = one cell.** Keyed by `"<tableIndex>:<label>"` (e.g. `"2:A1"`) — the label from SP1, prefixed by the cell's table (bandh) index within the block. Other bandhs' `A1` are unaffected.
- **Storage.** On the block's content-control tag: `payload.overrides = { "2:A1": {…}, … }`. Payload stays `v2` (`overrides` is optional, written by the editor). `parseContentControlTag` already returns the whole payload; `overrides` surfaces as `{}`/absent when unset.
- **Pure merge** `resolveCellOverride(base, override) → resolved` — given the block's base justify settings (`{strength, fillMode, …}`) and a cell's override, returns the effective per-cell `{strength, fillMode, widthPt|null, capEm}`. Node-testable, no Office/DOM.

## §2 — Cursor → context detection

- Register `Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, handler)` on Office ready; **debounce ~150 ms**.
- On fire, in one `Word.run`:
  - **Block:** `selection.parentContentControlOrNullObject`; load `title,tag`. Not an "Ashaar Poem" → clear the per-cell editor, leave block controls as-is.
  - **Cell:** `selection.parentTableCellOrNullObject`; load `rowIndex,cellIndex`. Determine the cell's **tableIndex within the block** by enumerating the block's tables and finding the one whose range contains the selection (see §6a). Resolve `(tableIndex, label)` via the block's SP1 `cells` map + `AshaarCellMap.buildBandhCellMap` (`rowIndex,cellIndex` → the map entry's label; a spacing cell → no override editor).
- **Thrash control:** sync the block-level controls only when the **active block changes** (track the last-reflected block), so moving the cursor within a block never clobbers controls the user is mid-edit. The per-cell editor updates on every selection change.

## §3 — Active-block sync

When the active block changes, populate the existing pane controls from the block's tag recipe (`fontMode`, `justifyMode`, `fillMode`, `tatweelCount`→strength, `tableWidthPct`) and, if the tag carries a `qaseeda` name, load that profile into the qaseeda panel (`profileToPanel`). Applying block-wide changes uses the existing **Justify Selected Text** / **Save & Apply to all** buttons — no new apply path. When no Ashaar block is at the cursor, controls are left untouched (so a fresh-insert setup isn't disturbed).

## §4 — Per-cell override editor

- A sub-panel beneath the justify actions, hidden until the cursor is in a **content** cell of an Ashaar block. Shows the cell's label (`A1`) and three inputs: **Strength** (1–10 or blank), **Target width** (pt or blank), **Cap lift** (em, blank = default 0.28). Each blank = inherit.
- Populated from `overrides["<ti>:<label>"]` for the active cell.
- **On edit:** write the merged override back to the block's `cc.tag` (re-find the block, `setTagOverride`, sync), then call `justifySelection()` to re-justify the whole block — instant feedback, reusing the existing justify path (no bespoke single-cell path). Editing an empty field removes that key; a **Clear cell override** button removes the whole entry.
- A new pure helper `setTagOverride(tag, key, override) → tag` (mirrors `setTagQaseeda`) writes/removes one override in the payload. Node-testable.

## §5 — Justify consumes overrides

`justifySelection` and `applyProfileToQaseeda` already tag each cell with its `(tableIndex,label)` key (SP1). Additionally read `tag.overrides` and, per content cell, compute `resolved = resolveCellOverride(blockJustify, overrides[key])`:
- **strength set** → use its `φ` and `maxPositions` for that cell instead of the block's.
- **widthPt set** (Natural-fit) → use `widthPt·96/72` as the fill `target` instead of `naturalFitTarget(Wpos, reach, φ)`.
- **capEm set** → pass as the `capEm` argument to `AshaarResidual.capMicroSpaces` for that cell. Applies **only in Natural-fit** (Cell-fit's residual is Word `distribute` — no micro-spaces — so `capEm` is inert there).
- No override / absent field → today's behavior (block/profile default). Cells with no map key (fallback tables) never carry overrides.

Rest of the self-review: `strength` works in both modes (it drives φ/positions); `widthPt` applies in Natural-fit only (§1); `fillMode` is not overridable, so `resolveCellOverride` always carries the block's `fillMode` through unchanged.

## §6 — Feasibility (validate first)

- **(a) TableCell → tableIndex within the block.** Office exposes no stable table id. Resolve by enumerating the block's tables and selecting the one whose range contains the selection (range comparison / `intersectWithOrNullObject`). **Spike this first.** If unreliable across hosts, the fallback is label-only keying (`"A1"`, position-wide **within the block**) — a documented scope change from "one cell" that would need user sign-off.
- **(b) Selection-changed lifecycle.** Register once on Office ready; guard against re-entrancy while a justify is running; debounce; tolerate hosts that fire sparsely (the §3 controls just won't auto-refresh — a manual path is out of scope for MVP but the block-sync must not error when the event is quiet).
- **(c) Tag write-back.** Re-find the block by walking `document.contentControls` (the selection's `parentContentControl`), set `cc.tag`, sync. Must not disturb the block's content.

## §7 — Testing

- **Node (pure):** `resolveCellOverride` (each field overrides / inherits; empty override → base unchanged); `setTagOverride` (adds, replaces, removes a key; round-trips via `parseContentControlTag`; leaves other payload fields intact); override-key formatting (`"<ti>:<label>"`).
- **Manual (Word):** cursor into a block auto-fills the pane; cursor into `A1` shows its override editor; set strength/width/cap → that one cell changes on the instant re-justify, siblings unchanged; **Clear** reverts it; moving within the same block doesn't reset block controls; a spacing/gap cell shows no editor; a block with no `cells` map (adopted) shows block sync but no per-cell editor.

## §8 — Non-goals

- **Per-cell style overrides** (color/shade) and **special-gap purposes** (hemistich symbols, bandh numbers, annotations) — backlog, own specs.
- **Per-cell fill-mode** override — fill mode stays a block-wide choice (not among the chosen levers).
- **Manual "load from cursor" fallback** — auto-only for MVP.

## Resolved decisions

1. MVP = per-cell justify override **+ full active-block sync**.
2. Override scope = **one cell**, keyed `"<tableIndex>:<label>"`, stored on the block tag. Broad changes go through profile/block settings, not a widened override.
3. Cursor tracking = **auto** via `DocumentSelectionChanged` (debounced); block controls resync only on **active-block change**.
4. Per-cell levers = **strength, target width, cap lift** (no fill mode).
5. Edit feedback = write tag then **re-justify the whole block** via the existing `justifySelection()` (instant, no new single-cell path).
6. Cell→tableIndex detection is **spiked first**; label-only fallback needs sign-off.
