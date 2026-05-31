# Visual Layout Grid ("Grid mode") — Design

**Date:** 2026-05-31
**Status:** Approved (pending spec review)

## Problem

The Table Input **Layout spec** is a text DSL (`2 - 1`, `<4>`, `3 | 2 | 1`, indentation).
It's powerful but opaque — users found "adding my own tables" uncertain. Provide a
**visual, scantron-style grid** that translates the desired ashaar shape into cells and
spacing directly, as an alternative view of the same layout, with the text spec still
available for those who prefer typing.

## Chosen approach

A **Grid / Numbers toggle** on the Layout spec area. *Numbers* is today's textarea
(unchanged). *Grid* is a new bubble widget: rows of 12 column-bubbles, tapped on/off.

The grid serializes to the **existing 12-column span template** —
`{ columnCount: 12, rows: [ [ {span, align} … ] ] }` — the same structure that
**Capture from Word** produces and `templateToOoxml` / Draw Table already consume. This
reuses two tested pipelines and adds only a visual editor plus pure converters.

Rejected alternatives: (2) inventing a richer text-spec syntax for lossless sync — adds an
unreadable format and undercuts the friendly numbers mode; (3) limiting the grid to the
coarse vocabulary — removes the precision that makes the grid worthwhile.

## Interaction model

- **Tap each bubble** (true scantron) — chosen over drag/two-ends for simplicity.
- A **run of contiguous "on" bubbles = one misra cell** (span = run length).
- **"Off" runs = spacing/gap cells.**
- Columns read **right-to-left** (poetry): the rightmost cell is misra 1 (sadr).
- Cells **auto-number by reading order** (right-to-left, top-to-bottom). Numbers are just
  cell labels — Draw Table inserts a blank table the user types into.

## Scope decisions

- The grid defines **one bandh's rows**; the existing **Bandhs** count repeats it. In Grid
  mode, *Misras per bandh* is derived from the drawn rows and is ignored.
- **Each mode draws from its own representation when active**: Grid → span template →
  `templateToOoxml`; Numbers → existing `templateGrid` text path. The toggle converts
  between views **best-effort** (a convenience bridge, not a lossy bottleneck).
- Tap-only (no drag) for this version.

## Components

### New pure module: `src/taskpane/layout-grid.js` (UMD, exported, tested)

- **`gridToTemplate(matrix)`** — `matrix` = array of rows, each row = 12 booleans. Walks each
  row grouping contiguous `true`s into content cells and `false`s into gap cells; assigns
  align (first content cell = right, last = left, lone = center, middle = center); numbers
  sequentially. Returns `{ columnCount: 12, rows: [ [ {span, align, role} ] ] }`.
  All-off rows are ignored. This is what Draw Table uses in Grid mode.
- **`gridToSpec(matrix)`** — serializes the matrix to the nearest text-spec line for the
  Numbers view: two content cells → `R - L`; one centred → `<n>`; three+ → `a | b | c`;
  a single right/left-anchored cell → `n>` / `<n`; leading/trailing gaps → indentation.
  Best-effort and may simplify shapes the text spec can't express exactly.
- **`specToGrid(text)`** — parses the coarse spec into a 12-col boolean matrix so
  Numbers→Grid shows an existing layout: pair → cols 1–5 + gap + 8–12; center → middle span;
  multi → even split across the row; indent → shifted run. Best-effort.

### `taskpane.js`

- Render the grid widget (rows of bubbles), tap-to-toggle, **+ Add row** / per-row remove,
  and a live mini-preview of the resulting shape.
- The **Grid / Numbers toggle**: on switch to Numbers, write `gridToSpec(matrix)` into the
  textarea; on switch to Grid, build the matrix via `specToGrid(textarea)`.
- **Draw Table** branch: if Grid mode is active, build `gridToTemplate(matrix)` and render
  via the existing template-OOXML path; otherwise use the existing text path.

### `taskpane.html` / `taskpane.css`

- The toggle control and a grid container in the Table Input panel; load `layout-grid.js`
  (with cache-buster); bubble + preview styling.

## Data flow

```
Grid mode:   bubbles → matrix → gridToTemplate() → {columnCount:12, rows:[{span}]}
               → templateToOoxml() → Draw Table

Toggle:      Grid → Numbers : gridToSpec(matrix)  → textarea
             Numbers → Grid : specToGrid(text)    → bubbles  (best-effort)

Numbers mode: textarea → parseLayoutSpec/templateGrid → Draw Table  (unchanged)
```

## Error handling & edge cases

- **All-off row** → ignored (no empty cell emitted).
- **Minimum one row**; removing the last row leaves one empty row.
- **Leading/trailing gaps** → indentation / centering, preserved as gap cells.
- **Grid→Numbers loss** → when a drawn shape can't be expressed exactly in the text spec,
  show a small note that the Numbers view is approximate (Grid remains the precise source
  while it's the active mode).
- **Unparseable spec** on Numbers→Grid → start from a sensible default (e.g. one paired row)
  rather than erroring.

## Testing

- **Unit (TDD), `tests/layout-grid.test.js`:** `gridToTemplate` for couplet, centred solo,
  full-width, indented, multi-misra, interleaved gaps (assert spans, align, numbering,
  all-off ignored); `gridToSpec` for each shape; `specToGrid` for pair/center/multi/indent;
  round-trip stability (`gridToSpec(specToGrid(x))`) for supported constructs.
- **Browser smoke test:** toggle works, tapping toggles bubbles, add/remove rows, preview
  updates, no console errors.
- **In-Word (manual):** Grid mode → Draw Table produces the expected blank table shape.

## Module boundaries

`layout-grid.js` holds the pure matrix↔template↔spec logic (Word-free, fully testable).
`taskpane.js` owns the widget UI and wiring. `word-html.js` (templateToOoxml, parseLayoutSpec,
templateGrid) is reused unchanged.

## Out of scope

- Drag / two-ends bubble interaction.
- Encoding refrain colour or content in the grid (the template is blank; colour happens at
  justify/content time).
- A new richer text-spec syntax.
