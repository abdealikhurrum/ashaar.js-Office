# Qaseeda Profiles — Design

**Date:** 2026-05-31
**Status:** In review

## Problem

An Ashaar poem is often split across several "Ashaar Poem" content controls (separate
inserts, stanzas interleaved with prose). Their column widths and justification drift apart,
and there's no way to apply one consistent look — or to re-fit an existing block — across the
document. We also can't reliably justify some fonts because the task-pane canvas can't always
measure what Word renders.

A **qaseeda** is a named link across blocks that share one **profile** (width, justification,
gap, symbol, colors). Applying/refreshing a profile keeps every block consistent and lets an
existing block be re-fit in place.

## Decisions (from brainstorming)

- **Identity:** named qaseeda. A block's content-control tag carries the qaseeda **name**.
- **Storage:** the profiles live in a **document-embedded store keyed by name** (Word
  document settings or a custom XML part — pick the cleaner during build). One authoritative
  copy; travels with the .docx.
- **Sync:** **hybrid auto-refresh.** Justifying a block applies the stored profile locally;
  if the block has a misra wider than the profile knows, the shared widths refresh and
  re-apply to all the qaseeda's blocks.
- **Width:** **per-profile choice** — `auto-fit` (widths derived to fit the widest misra per
  column + kashida headroom across all blocks) or `fixed` (%). Default auto-fit.
- **Apply method:** **in-place resize + re-justify** (preserve the table and manual edits);
  rebuild only as a fallback when a table's columns can't be resized cleanly.

## The profile (data model)

```
Profile {
  name
  width:   { mode: "auto-fit" | "fixed", pct }
  justify: { mode: "kashida" | "spacing" | "css" | "none", strength /* → targetFill */ }
  gap                                  // middle gap
  misraSymbol, symbolColor             // gap symbol + its colour
  debugColors: { tatweel, space }      // colour inserted artifacts so they're visible
  font?                                // optional explicit font
  fontCorrections: { <fontName>: factor }   // per-font fill nudge (see Measurement)
  derived: { colWidthVector, calibrationRecipe }   // computed; refreshed on the hybrid trigger
}
```

Authoritative fields are user-set; `derived` is computed from the qaseeda's content.

## Measurement accuracy (font resolution + correction)

Kashida fill is measured in the task-pane canvas, which only matches Word when the WebView
can resolve the font. System/bundled fonts (Arial, Noto Nastaliq) resolve and measure
correctly; fonts the WebView can't see — not installed (Arabic Typesetting on macOS) or
user-installed but not exposed to the sandboxed WKWebView (FatemiMaqala) — fall back and the
fill % is meaningless.

- **Availability check:** **canvas width-comparison** — measure a test string in the font
  layered over generic families; if the width changes vs the generic baseline the font
  resolved, otherwise it fell back. (`document.fonts.check` is unreliable here: it returns
  true for unknown system-font names.) Surfaced in the Debug panel's `res` column. When a font
  is **not** resolved, warn that justify metrics are approximate for it.
- **Bundled fonts (preferred fix for owned/redistributable fonts):** vendor the font as a git
  submodule (`vendor/<font>/`, mirroring `vendor/ashaar-js`), sync its web file into
  `assets/fonts/`, and declare an `@font-face` whose `font-family` matches the document's font
  name exactly. The WebView then resolves the font natively (`document.fonts.check` → true) and
  the canvas measures the same outlines Word renders — no correction needed. Justify must
  `await document.fonts.load(font)` before measuring. Requires the bundled web file to be the
  same outlines/version as the installed desktop font.
- **Per-font correction (`fontCorrections`):** the fallback for fonts we can't bundle
  (third-party / non-redistributable). A remembered nudge per font name to bridge the residual
  measure-vs-render gap; tuned by eye once, stored in the profile.

## Apply / refresh engine

Applying a profile to a qaseeda:
1. Find all content controls whose tag names that qaseeda.
2. If `width.mode === "auto-fit"`, derive the shared column-width vector (widest misra per
   column across all blocks, with kashida headroom; using `fontCorrections` for the measured
   font); else use the fixed %.
3. **Resize each block's columns in place** to the shared widths and **re-justify** with the
   profile's params (and symbol/colors). Rebuild a block only if it can't be resized.
4. Cache the derived vector in the profile.

**Hybrid trigger:** Justify applies the stored profile to the current block; if its content
exceeds the cached widths, refresh the vector and re-apply to all blocks.

## In-place column resize (load-bearing capability)

Set existing table column widths via Office.js (`TableCell.columnWidth` /
`Table.fixedColumnWidths()`), non-destructively. **Verify this API path early** — it's the
assumption the whole "apply in place" rests on; fall back to rebuild for a block if it can't
be resized cleanly.

## UI (consolidation)

A profile panel: name/pick a qaseeda; edit its params (width auto/fixed, gap, misra symbol +
symbol colour, justification type/strength, debug tatweel/space colours, per-font correction);
**Apply to all** (and it auto-applies on justify per the hybrid rule). A `res`/warning shows
when the current font isn't resolvable. Inserts let you assign a qaseeda name (or "none").

## Build phasing

- **P1 — In-place column resize + auto-fit an existing block on Justify.** The load-bearing
  capability; useful on its own (re-fit any existing table, incl. grid-made).
- **P2 — Profile data model + document store + tagging blocks with a name.**
- **P3 — Apply/refresh engine across instances (hybrid auto-refresh).**
- **P4 — Profile UI + debug/symbol colours + per-font correction + font-availability warning.**

## Modules

- **New `profiles.js`** (pure, UMD, tested): profile shape, merge/defaults, derive-shared-widths
  math, fill-correction application.
- **`taskpane.js`:** Office.js orchestration — tagging, the document store, gather/apply,
  in-place resize, font-availability check.
- **`word-html.js`:** generator gains misra-symbol colouring and debug colouring of inserted
  tatweels/spaces.

## Testing

- Pure `profiles.js` logic (derive widths, corrections, merge) via Node tests.
- In-place resize, store, gather/apply, font-check — verified in Word (Office.js).

## Out of scope (future overhaul)

- **A general font-bundling system (option c)** — a managed way to bundle/serve *arbitrary*
  (incl. third-party) fonts as `@font-face`, with discovery/UI. Bundling the developer's *own*
  font(s) via submodule (above) is in scope now; the general system comes later.
- Cross-document profile sharing (export/import) beyond the in-document store.
