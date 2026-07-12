# Spacing-Cell Decorations — Symbols + Fill Color (SP3) — Design Spec

**Date:** 2026-07-11
**Status:** Draft — user-approved (brainstorming); ready for implementation planning.
**Sub-project:** SP3 of `2026-07-11-cell-configurations-roadmap.md` (backlog "special gaps" + fill color).
**Builds on:** SP1 bandh cell-map (spacing cells carry a `slot` key `A#1` — the anchor for this feature; the persisted `cells` tag makes a decorated gap authoritatively still "spacing"), SP2 per-cell overrides + active-block sync (the editor + tag-writer + resolve pattern this reuses), and the qaseeda profile model.

## Overview

Put content and color into the structural gap cells that today render empty: a **hemistich / decorative symbol** and a **fill (background) color**, with an optional symbol **text color**. Configured two ways (as approved): a **profile default by slot-position** (uniform across bandhs) plus a **per-slot override** (one specific gap). Auto-numbering (bandh/verse) is a deferred fast-follow. Decorated gaps stay tagged `spacing`, so justification keeps skipping them — decoration is orthogonal to fill.

## §1 — Decoration payload

Per spacing slot:
```
{ symbol?: string, fill?: "RRGGBB", color?: "RRGGBB" }
```
- `symbol` — free-text glyph(s) placed in the gap (e.g. `؎`, `*`, an ornament). Free input (paste any mark); a preset picker is a later nicety.
- `fill` — cell background shade (hex, no `#`).
- `color` — the symbol run's text color (hex). Ignored when there's no symbol.
- Absent field = none. An **empty string in an override = explicit none** (suppress a profile default); absent in an override = inherit the profile default.

## §2 — Config model (both)

- **Profile default by position:** `profile.spacingDecor = { "<slot>": { symbol, fill, color }, … }`, keyed by the SP1 slot-position (e.g. `"A#1"`). Applied to every bandh's slot at that position. This supersedes the currently-unused flat `profile.misraSymbol` / `profile.symbolColor` fields (folded into this structured, per-position form).
- **Per-slot override:** new content-control-tag field `slotDecor = { "<tableIndex>:<slot>": { … }, … }` (e.g. `"0:A#1"`), overriding/suppressing the profile default for one specific gap. Key built with the existing `AshaarOverrides.overrideKey(tableIndex, slot)` (the "label" arg carries the slot string).
- **Pure merge** `resolveSlotDecor(profileDecor, override) → { symbol, fill, color }` — per field, `override` has the key → use it (empty string = none); else inherit `profileDecor`. Mirrors `resolveCellOverride`. Node-testable.

## §3 — Rendering (OOXML)

A decorated spacing cell:
- Its `<w:tcPr>` gains `<w:shd w:val="clear" w:color="auto" w:fill="RRGGBB"/>` when `fill` is set.
- Its paragraph carries a centered RTL run with the `symbol` (colored via `<w:color w:val="RRGGBB"/>` when `color` set), instead of the empty `<w:p/>`.
- No `fill` and no `symbol` → unchanged empty `<w:p/>` (plain gap).

Pure emitters: `spacingDecorParaXml(decor, sizePt) → "<w:p>…</w:p>"` (empty-para string when no symbol) and a `shdXml(fill) → "<w:shd …/>"` fragment; `tcXml` extended to accept an optional shd fragment in its `<w:tcPr>`.

## §4 — Apply paths

- **Insert (bake-in):** the gap/pad cell builders (`gapTc`/`padTc` in `baytRowsOoxml`) consume the resolved profile decoration for each slot-position and emit the shd + symbol directly. Requires threading the resolved decor-by-slot into the generator (via `opts`).
- **Existing blocks (decorate pass):** part of **Save & Apply** and run after a per-slot edit — walk the block's spacing cells (located via the SP1 `cells` map), and for each write the symbol (`cell.body` insert) + set `cell.shadingColor` (Office.js `TableCell.shadingColor`, or OOXML rebuild if unavailable). Reuses the same block-walking `applyProfileToQaseeda` already does.
- Decorated gaps remain `kind:"spacing"` in the map → `justifySelection`/`applyProfileToQaseeda` continue to skip them for justification.

## §5 — UI

- **Per-slot editor** — extend SP2's active-cell reflection: when the cursor is in a **spacing** cell (map `kind==="spacing"`), show a **decoration editor** (symbol text / fill color / text color) instead of the justify-override editor; editing writes `slotDecor` on the block tag (new `AshaarWord.setTagSlotDecor`, mirroring `setTagOverride`) and runs the decorate pass on that block. A **Clear** button removes the per-slot entry (reverting to the profile default).
- **Profile defaults** — a small "spacing decorations" section in the qaseeda panel to set `spacingDecor` by slot-position (the read-only SP1 cell-map view already lists the gaps/labels, so the operator can see which slot is which).
- The cell-map view (SP1) renders decorations inline (show the symbol / a color swatch next to `(gap)`).

## §6 — Testing

- **Node (pure):** `resolveSlotDecor` (each field overrides / inherits / empty-string-suppresses); `spacingDecorParaXml` + `shdXml` (shd present with fill; colored symbol run; empty decor → plain `<w:p/>`); `tcXml` with shd; `setTagSlotDecor` add/replace/remove round-trip via `parseContentControlTag` (leaves `cells`/`overrides` intact); profile `spacingDecor` survives `normalizeProfile`.
- **Manual (Word):** a profile symbol/color shows in every bandh's slot; a per-slot override changes one gap and Clear reverts it; fill color renders; editing via the pane re-decorates instantly; justify still leaves decorated gaps untouched (no tatweels/spaces added to them); an adopted block (no `cells`) shows no decor editor.

## §7 — Non-goals

- **Auto-numbering** (bandh/verse counters in a slot) — the fast-follow spec (own numbering scheme: which slot, per-bandh vs global, format).
- **Content-cell shading / styling** — MVP scopes fill color to spacing cells only.
- **Multi-glyph rules / decorative borders / repeated-fill leaders** — later.
- **Preset symbol picker** — free-text input for MVP.

## Resolved decisions

1. Payload = `{ symbol?, fill?, color? }`; fill color scoped to **spacing cells** for MVP; symbol is **free-text**.
2. Config = **both**: profile default keyed by slot-position (`"A#1"`) + per-slot override on the block tag keyed `"<tableIndex>:<slot>"`. Empty-string override field = explicit none.
3. Decorated gaps stay `kind:"spacing"` — justification skips them (SP1's persisted tag makes this authoritative even though the cell is no longer empty).
4. Applied at insert (bake-in) and via a decorate pass on Save & Apply / per-slot edit.
5. Per-slot editing reuses the SP2 active-cell editor (spacing cell → decoration editor); profile defaults in the qaseeda panel.
6. Supersedes the unused flat `profile.misraSymbol`/`symbolColor` with structured `profile.spacingDecor`.
