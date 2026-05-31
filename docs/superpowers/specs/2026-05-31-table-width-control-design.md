# Table Width Control — Design

**Date:** 2026-05-31
**Status:** Approved

## Problem

Inserted/drawn tables span the full page text column (100%). That's usually too wide for
poetry, and narrowing a table by dragging its borders/cells in Word is painful. Provide a
**Table width** control so tables are inserted at a chosen fraction of the page (default 50%),
centred — no manual border resizing.

## Design

- **Control:** a **Table width** slider in the **shared controls** (applies to both Table
  Input and Convert). Range **25–100%**, step 5%, **default 50%**. Value = table width as a
  percentage of the page's text column.
- **Applies to** every table-producing action: *Insert as Table* (`renderForWordOoxml`),
  *Draw Table* incl. Grid mode and the span-based path (`templateToOoxml`), *Apply template*
  (`templateToOoxml`), *Drop Grid* (`generateBareGrid12Ooxml`), and *Adopt* (inserts via the
  *Insert as Table* path, so it inherits). *Insert as Paragraphs* (tab-stops, not a table) is
  unaffected.
- **Mechanism:** `taskpane.js` already computes `textWidthTwips` from the page layout before
  calling each generator. Multiply it by `pct/100` first. Because every table emits
  `tblW` (dxa) + `jc="center"`, a smaller width is simply a **smaller, centred table** — no
  cell/border changes. Kashida targets (`_textWidthPx`) derive from the same `textWidthTwips`,
  so they scale automatically and justification stays balanced.
- **Plumbing:** add `tableWidthPct` (default 50, clamped 25–100) to `options()` and to the
  content-control tag, so the value is remembered. 100% reproduces today's behaviour.

## Files

- `taskpane.html` — the slider in the shared controls.
- `taskpane.js` — read the slider into `options()`; scale `textWidthTwips` by `pct/100` in the
  insert paths (insertPoem, insertStructure grid + spanBased, applyTemplate, drop grid);
  include `tableWidthPct` in the content-control tag.
- `taskpane.css` — minor (slider label).
- `word-html.js` — **unchanged**; it already honours whatever width it is given.

## Testing

- **Unit (word-html.test.js):** assert the OOXML width scales — e.g. `renderForWordOoxml` at
  `textWidthTwips` 4680 produces `gridCol`/`tblW` twips half those at 9360. (This proves the
  generators honour a scaled width; the scaling itself is a scalar in taskpane.js.)
- **Browser smoke test:** the slider exists, reads into `options().tableWidthPct`, default 50.
- **In-Word (manual):** insert at 50% → a centred half-width table; 100% → full width as before.

## Edge cases

- Clamp to 25–100%. Never produce a zero/negative width (floor at 25%).
- Existing documents / re-justify: justification reads the live page width and the poem's own
  cell widths, so it is unaffected by this control after insertion.

## Out of scope

- Per-table width overrides after insertion (use Word's table tools).
- Tab-stop paragraph width.
