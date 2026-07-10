# Run-Aware Kashida — Consumers Design

**Date:** 2026-07-09
**Status:** Approved, ready for implementation plans
**Depends on:** `2026-07-09-run-aware-kashida-design.md` (foundation primitive — done)
**Scope:** The two follow-up consumers named by the foundation spec. Wire the app layers to the multi-run engine (`justifyRuns`, `computeRunSpacing`, `applySlotsMulti`, `measureRunsNatural`) so a single misra with mixed styling justifies correctly. Two independent sub-projects, each with its own implementation plan.

## Problem

The foundation primitive justifies a misra whose runs each carry their own font, but nothing calls it. Both existing consumers are single-font and flatten styling on write-back:

- **Word** (`taskpane.js:1244-1284`) reads one `cell.body.font`, calls `justifyLine` / `justifyWordSpacing` per cell, and writes back with `cell.body.paragraphs.getFirst().insertText(justified, replace)` — destroys any runs in the cell.
- **Web** (`ashaar.js:319-377`) reads one computed style off the misra span, sets `wordSpacing` + `fontSize` + tatweel on the whole `textContent`, then `spanEl.textContent = best` — flattens child spans.

## Motivating content

Same font **family** throughout a misra; **weight** and **size** vary run-to-run:

- **Bold / emphasis / refrain** — a word or refrain set heavier. Measured width differs run-to-run even with no family change.
- **Mixed size** — a larger opening word or a smaller annotation. `fontSize` varies; best served by the spacing/scale path's uniform `fontScale`.

The honorific-different-family case is **not** a v1 driver, so cross-family shaping is moot — but run discovery keys on the full style tuple, so a family change is handled for free if it ever appears.

## Goals / non-goals

- **Goal:** Make **both** justify modes run-aware in **both** consumers — kashida (`justifyRuns`) and spacing/scale (`computeRunSpacing`) — preserving each run's styling on write-back.
- **Non-goal (v1):**
  - Exact mid-word cross-style shaping. Word-aligned runs are exact (foundation spec boundary decision); mid-word style changes collapse to the run's dominant style (Word) / split at the node boundary (Web). Documented, not a bug.
  - Per-run font profiles. Family is constant, so one poem-level `fontProfile` applies to every run; the correctness gain is entirely per-run **measurement**.
  - Any new Ashaar source markup. Styling comes from rendered content (Word ranges / DOM nodes), never from the poetry source.

## Decomposition & ordering

| | **Word add-in consumer** | **Web preview consumer** |
|---|---|---|
| Code | `taskpane.js` + `word-html.js` (this repo) | `ashaar.js` (upstream submodule → sync) |
| Read runs | word-range split + coalesce by style | misra span child nodes + computed style |
| Write-back | in-place range replace (OOXML fallback) | per-child text + parent `word-spacing`/scale |
| Tests | pure-node (fake range/ctx) | pure helper node-testable; DOM glue manual |

**Order: Word first, then Web.** Word is the shipping app, lives entirely in this repo (no submodule round-trip), and is where the flattening `insertText` loses formatting. Web is lower-risk (DOM properties) and follows. Each gets its own implementation plan.

---

## Sub-project 1 — Word add-in consumer

Changes `justifySelection` (`taskpane.js`) and adds run-aware helpers to `word-html.js`.

### Run discovery (A1: word-range split + coalesce)

Office.js exposes no run collection. Per cell paragraph:

1. `paragraph.getRange().getTextRanges([" "], true)` → one range per word.
2. Load `font.name/size/bold/italic` on each word range.
3. Coalesce adjacent words whose full style tuple `(name, size, bold, italic)` is identical into a **run**; the run's range is the union `first.expandTo(last)`.

Word-aligned, which the foundation spec calls exact. A mid-word style change collapses to the run's dominant style — documented limitation.

Each run becomes `{ text, measure, fontProfile, fontSize, range }`, where `measure(s)` uses a canvas `ctx` whose `font` is set to the run's `size`/family/`bold` (the real correctness gain), and `range` is retained for write-back.

### Write-back (B1: in-place range replace; B2 OOXML fallback)

`justifyRuns` / `computeRunSpacing` return same-length, same-order results, so results map index-aligned to the discovered run ranges.

- **Kashida:** for each run, `run.range.insertText(justifiedText, Word.InsertLocation.replace)`. The range keeps its own font; the paragraph is never touched, so `jc` / spacing / indents survive automatically.
- **Spacing/scale:** apply `run.range.font.size = run.fontSize * fontScale` per run (uniform scale, preserves relative sizes); realize `wordSpacing` as micro-spaces (U+200A hair, U+2009 thin fallback) totalling `wordSpacing × gaps` px across all gaps, reusing `justifyWordSpacing`'s insertion driven by the run-aware total width.

**Fallback (B2):** if sequential `insertText(replace)` across disjoint tracked ranges proves unstable in one sync batch, rebuild the cell paragraph as OOXML — `<w:r><w:rPr>…</w:rPr><w:t>…</w:t></w:r>` per run — and `cell.body.insertOoxml(xml, replace)`, carrying the paragraph's `jc`/spacing forward by hand. Heavier and reconstructs paragraph props, hence fallback only.

### Font profile

Unchanged. Family is constant across runs, so the existing poem-level probe/calibrate (`AshaarTune.probeFont` + `calibrate`) yields one `fontProfile` passed to every run. No per-run probing in v1.

### Edge cases

- Cell with one style → one run → identical to today's single-font result (parity).
- Empty / whitespace cell → skipped (as today).
- `fontScale === 1` → sizes untouched.
- Re-justify: strip existing kashida/micro-spaces (`stripJustification`) before discovery, so results never compound (idempotent/reducible — matches the primitive).

### Testing (`tests/` pure node, deterministic fakes)

- Run coalescing: adjacent same-style words merge; a style change splits; single-style cell → one run.
- Kashida mapping: `justifyRuns` output maps index-aligned to the correct run ranges (against a fake range model — Office.js is not node-loadable).
- Spacing: `fontScale` multiplies each run's size; `wordSpacing × gaps` → correct micro-space total.
- Idempotence: strip → justify → strip returns bare text; re-justify doesn't compound.

`Word.run` glue (range split, sequential replace) is not node-testable — verified manually in Word against a mixed-style stanza in `test-documents/` (bold refrain + larger first word).

### Risk

Sequential range-replace instability (B1) → OOXML fallback (B2) per cell.

---

## Sub-project 2 — Web preview consumer

Changes `justifyMisra` / `justifyEl` in the upstream submodule `vendor/ashaar-js/ashaar.js`, then `npm run sync:ashaar`, then tests. **Never edit `src/vendor/` directly.**

### Run discovery

A misra span (`.ashaar-misra--sadr` / `--ajuz`) may hold child nodes — `<b>`, size-styled `<span>`, honorific spans. Walk its child nodes (text + element); for each build `{ text, measure, fontProfile, fontSize }` where `measure` uses a probe styled from that child's `getComputedStyle` (family/weight/style/size/features), mirroring the existing `createProbe`. A plain text node → a run styled by the parent misra's computed style. Reuse a **single** probe element restyled per run (not N probes) to keep cost down.

### Reducibility

Cache `spanEl.dataset.ashaarOriginalHtml = innerHTML` once; restore it at the start of every pass, then re-discover runs. Re-justify re-derives from the bare styled markup and never compounds — same contract as the pure primitive. (Replaces today's flattening `dataset.ashaarOriginal` + `textContent = best`.)

### Apply

- **Kashida:** `justifyRuns(runs, target, params)` → write each returned `text` back into its owning child node's text; element children keep their styling; index-aligned mapping.
- **Spacing/scale:** `computeRunSpacing(runs, target, params)` → set `spanEl.style.wordSpacing` once (one value for the whole misra, exactly CSS's model) and apply `fontScale` by scaling **each run/child's** `font-size` — robust to px or em and preserving relative sizes, unlike setting the parent's `%` (which silently fails when children carry absolute sizes, i.e. the larger-first-word case).

### Block balancing

`blockTargets` (balance all misras to the longest) stays; per-misra "natural width" becomes the run-aware sum (`measureRunsNatural`-style) instead of a single-probe measure.

### Testing

Extract run-discovery + target math into a **pure helper** (misra element → run specs; balancing math) node-testable with a fake measure. DOM glue (`getComputedStyle`, `getBoundingClientRect`, writing child text) stays thin and is verified manually in the live preview. Matches the repo's no-jsdom, pure-node convention.

### Risks

- `measure` cost per pass → single reused probe.
- Upstream drift → sync before tests (plan step, not a design risk).

---

## Data flow (both consumers)

```
consumer discovers ordered runs of a misra, binds a per-run measure()
  → kashida:  justifyRuns(runs, targetWidth, params)      → [{text}]  → write each run's text back (styling preserved)
  → spacing:  computeRunSpacing(runs, targetWidth, params) → {wordSpacing, fontScale}
                Word: micro-spaces (wordSpacing×gaps) + per-run size×fontScale
                Web:  spanEl.style.wordSpacing + per-child font-size×fontScale
```

## Out of scope (future)

- Per-run font profiles (probe bold/size variants separately).
- Exact mid-word cross-style shaping.
- Width-balanced kashida across runs (primitive ranks by quality/count, not equal per-run fill).
