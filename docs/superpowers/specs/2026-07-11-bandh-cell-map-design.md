# Bandh Cell-Map — Content/Spacing Tag + Positional Labels — Design Spec

**Date:** 2026-07-11
**Status:** Draft — user-approved (brainstorming); ready for implementation planning.
**Sub-project:** SP1 of `2026-07-11-cell-configurations-roadmap.md` (SP2 = per-cell overrides + editing UI, a future session).
**Builds on / revises:** `2026-07-11-poetry-justification-modes-design.md` §1 — that spec introduced content/spacing tagging and a *geometric* position signature (`AshaarMatrix.positionKey({row,col,span})`, content-vs-spacing inferred from emptiness). This spec makes both **first-class and persisted**, and replaces the geometric harmony key with a stable **positional label**.

## Overview

Give every grid cell in an Ashaar poem two authoritative properties, derived and persisted when the table is created:

1. a **content/spacing tag** — a cell either holds a misra (content) or is a structural gap (spacing);
2. a **positional label within the bandh** — `A1, A2, B1, C1, C2, …` (letters = rows top-to-bottom, numbers = content cells in misra reading order).

Everything else (hemistich symbols, bandh numbers, annotations, block-level roles like sadr/ajuz/refrain, per-cell overrides, styling) is deferred; this is the minimum data model that makes **table creation and justification predictable and user-friendly**. The label is the identity all later features hang off.

The current justification path *guesses* these — content-vs-spacing from text emptiness (so a not-yet-filled content cell looks like a gap) and cross-bandh position from raw table geometry. This spec replaces the guesses with a persisted **bandh cell-map**.

## §1 — The cell-map model (pure)

- **Bandh = one Word table** (a stanza / repeating unit). A poem block (Ashaar Poem content control) holds one or more bandhs.
- **Rows → letters.** Each table row, top-to-bottom, gets a letter: `A`, `B`, `C`, …
- **Content cells → numbers in misra reading order.** Within a row, content cells are numbered `1, 2, 3, …` in reading order — **sadr before ajuz**, i.e. **rightmost-first in RTL**. Label = letter + number: `A1`, `A2`, `B1`, `C1`, `C2`.
- **Spacing cells** (`kind:"spacing"`): structural gaps (the middle gap, padding columns). **Excluded from content numbering** — they get no letter+number label. They DO get a stable **slot key** (`slot`, a within-row spacing index, e.g. `A#1`) so enumeration and future "special gaps" (hemistich symbols, bandh numbers, annotations) can target them. This is their sole reason to exist in the map today.
- **Cell-map** = an ordered list per bandh, in the same order the OOXML generator emits cells:

```
[
  { index: 0, kind: "content", label: "A1", gridCol: 0, span: 5 },
  { index: 1, kind: "spacing", slot: "A#1", gridCol: 5, span: 2 },
  { index: 2, kind: "content", label: "A2", gridCol: 7, span: 5 },
  { index: 3, kind: "content", label: "B1", gridCol: 0, span: 12 },
  ...
]
```

- **Pure function** `buildBandhCellMap(rows) → cellMap`, where `rows` is the existing internal row model (`structInfo`/`layoutTablesForTemplate` output — rows of type pair/refrain/solo/multi with their content vs gap cells and grid spans). No DOM/Office dependency → **node-testable**.
- **Extension point (future, NOT built here):** a spacing entry is where a future `purpose` (`"hemistich-symbol" | "bandh-number" | "annotation"`) and its payload will attach. The model leaves `kind:"spacing"` open to gain those subtypes later without a data migration.

## §2 — Derivation

The map is derived from the **same deterministic row model** the OOXML generator already walks (`baytRowsOoxml` / `layoutTablesForTemplate`, driven by `misraPattern`, `misraCount`, `gapWidth`, source text). Because it is built from that model, the map's entry order is **1:1 with the Word cells** the generator emits — so at read-back time, walking the live table's cells in order re-associates each with its map entry positionally.

## §3 — Persistence

- Extend the content-control tag payload (`AshaarWord.contentControlTag`) with a compact **`cells`** field: the per-bandh **pattern** (each row as a sequence of `"c"`/`"g"` tokens), which fully re-derives labels/slots via `buildBandhCellMap`. Bump the payload `v` (currently `1` → `2`).
  - Normal case: all bandhs share one template → store **one** shared pattern.
  - Bandhs of differing shape → store a per-table pattern list.
- `AshaarWord.parseContentControlTag` reads `cells` back (absent on `v1` tags → treated as "no map", triggering the §4 fallback).
- Storing the small pattern (not every label) keeps the tag compact; labels/slots are a pure function of the pattern.

## §4 — Justification consumes the map

`justifySelection` and `applyProfileToQaseeda` (in `taskpane.js`) currently key harmony on `AshaarMatrix.positionKey({row,col,span})` and skip cells via `AshaarMatrix.isContentCell(emptiness)`. Change:

- **content/spacing** comes from the map, not emptiness → a still-empty **content** cell is justified as content (not skipped as a gap); a **spacing** cell is never elongated/spaced.
- **harmony key** = the **label** (`A1` matches `A1` across every bandh), replacing the raw geometric signature. Same-label cells across bandhs form one natural-width-matrix position.
- **Fallback:** when a block has no `cells` map (v1 tag, adopted/hand-drawn table, external table), fall back to today's geometric `positionKey` + emptiness inference — behavior unchanged for those.

The `AshaarMatrix` module (natural-width matrix + target math) is unchanged; only the **key** feeding it changes (label when a map exists, geometric signature otherwise).

## §5 — Read-only pane surface

When the cursor is inside an Ashaar Poem block, the pane shows that bandh's cell-map — content labels (`A1`, `A2`, …) and gap markers (`(gap)`) laid out by row — **read-only**. Reuses the existing block-at-cursor lookup (`getQaseedaAtSelection` pattern). No document mutation. This is the seed for SP2's per-cell editing UI.

- Labels cannot be shown legibly on the Word page itself (Word has no per-cell text overlay; cell shading is a color not a label and belongs to the deferred styling feature; per-cell content controls / comments / hidden text all pollute the document). The pane is therefore the home for labels.

## §6 — Testing

- **Node (pure):**
  - `buildBandhCellMap` — RTL reading-order numbering (sadr = `A1`, ajuz = `A2`); gaps excluded from numbering but given `slot` keys; multi-row letters (`A/B/C`); solo full-width row → single `A1`; refrain rows numbered like any content row.
  - Tag round-trip — `contentControlTag` writes `cells`; `parseContentControlTag` recovers the pattern; `v1` tag (no `cells`) parses as "no map".
  - Label-match — two same-shaped bandhs yield identical labels; different-shaped bandhs do not collide.
- **Manual (Word):** insert → tag carries the map; justify keys harmony on labels (corresponding cells across bandhs land at one width); a deliberately empty content cell is still treated as content; a gap is never justified; the pane shows the block-at-cursor's map; an adopted/hand-drawn table justifies via the geometric fallback.

## §7 — Non-goals / future

- **Special gaps** — hemistich symbols, bandh numbers, annotations, in-gap labels. The spacing-cell `slot` is their future anchor; not built here.
- **Block-level roles** — sadr/ajuz/solo/refrain semantic classification and role-based styling (colors, shades, fill glyphs). Later, block/qaseeda level.
- **Per-cell overrides + select-cell→pane editing** — SP2 (justification variance + style variance keyed by label).
- **Editing the map** — SP1's pane view is read-only; changing which cells are gaps is SP2.

## Resolved decisions

1. Cell id = **human positional label** `A1/A2/B1…` (letters = rows, numbers = content cells in misra reading order, **rightmost = 1** in RTL). Gap cells excluded from numbering; keyed by `slot`.
2. Tag = **binary** content/spacing; `kind:"spacing"` left open to future subtypes.
3. Source of truth = **derive + persist at creation** on the content-control tag (`cells` pattern, payload `v2`); **geometric inference is the fallback** for maps-absent tables.
4. Observable deliverable = justification correctness **plus** a **read-only** cell-map view in the pane; labels are **not** rendered into the Word document.
