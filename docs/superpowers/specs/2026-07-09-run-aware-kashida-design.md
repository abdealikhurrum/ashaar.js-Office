# Run-Aware Kashida — Foundation Primitive Design

**Date:** 2026-07-09
**Status:** Approved, ready for implementation plan
**Scope:** The shared justification **core only**, in `ashaar-justify.js`. Multi-font/multi-style justification of a single misra, computed once in the vendored node/browser library so both the web preview and the Word add-in can consume it. Consumer wiring (DOM property application, Word per-run read + OOXML write-back) is **out of scope** and tracked as follow-up specs.

## Problem

The kashida engine assumes **one font per misra**:

- `justifyLine(text, targetWidth, ctx, params, fontProfile)` measures every candidate with a **single `ctx`** (`ashaar-justify.js:247, 263`).
- `buildSlots` addresses insertion points by word index within the whole line (`ashaar-justify.js:192, 207`) — a style run can begin or end mid-word, so a word index alone can't say which run a slot belongs to.
- One `fontProfile` scores the whole line (`ashaar-justify.js:203`) — each font has its own kashida-quality scores.
- Output is one string — a multi-run result must stay a **list of runs** so a consumer can write each back with its own styling.

Non-kashida (spacing) justification has the same single-font assumption and today lives in two consumer-specific places, not the shared library:

- **Word** (`word-html.js:40` `justifyWordSpacing`) inserts thin/hair-space characters at word gaps (text-mutating).
- **Browser** (`ashaar.js:316` `justifyMisra`) sets the CSS `word-spacing` property, then a `font-size` scale-down (`maxScaleDown`), only falling back to tatweel last (property-based).

The result: a misra with mixed fonts/sizes/bold cannot be justified correctly anywhere, and there is no shared primitive to fix it in one place.

## Goals / non-goals

- **Goal:** Add run-aware justification functions to `ashaar-justify.js` for **both** modes — kashida and spacing/scale — that measure each run in its own font. Pure and node-testable. Authored in the upstream submodule so the web library benefits, then synced.
- **Non-goal (v1):**
  - Exact mid-word cross-font shaping (two different fonts cannot shape-join across a seam; measured approximately, documented).
  - Consumer wiring: DOM property application in `ashaar.js`, OOXML spacing/scale mapping and per-run read/write-back in the Word add-in. Separate follow-up specs.
  - Width-*balanced* kashida distribution (v1 ranks by quality/count, not equal per-run fill).
  - Any new Ashaar **source markup** — styling comes from the rendered content (DOM spans / Word ranges), never from the poetry source. The parser stays plain-text.

## Key insight

Styling originates anywhere (Word formatting, DOM spans), so the primitive should not care about the source. It receives an **ordered array of already-styled runs**, each carrying its own measurement function, and justifies the misra preserving run order and count. Both consumers discover runs at render/justify time from their own content and map results back by index.

## Architecture

Authored in the submodule `vendor/ashaar-js/ashaar-justify.js` (upstream source of truth, so the web library gains multi-style kashida), then `npm run sync:ashaar` copies it to `src/vendor/ashaar-justify.js`. Tests run against the synced `src/vendor/` copy per repo convention. **Never edit `src/vendor/` directly.**

### Data model

```
run   = { text: string, measure: (s: string) => number, fontProfile?, fontSize? }
misra = run[]                  // ordered runs of one hemistich
```

`measure(text)` returns pixel width; the caller binds font + size (a canvas `ctx.measureText().width` on web/Word; a deterministic fake in tests). `fontProfile` is the per-run kashida-quality profile (kashida mode). `fontSize` is used only by spacing mode for clamp-bound reference. A run missing `measure` is a contract violation → throw `TypeError`.

### Functions

**1. Kashida — text-mutating**

```
justifyRuns(runs, targetWidth, params) → [{ text }, …]   // same length + order as input
```

One algorithm, shared with the existing single-font path:

1. `stripTatweels` each run's text (idempotent / reducible, unchanged behavior).
2. `naturalTotal = Σ run.measure(text_i)`; `target = targetWidth * (params.targetFill || 1)`.
3. If `naturalTotal >= target` or no visible text → return stripped texts unchanged.
4. Build global slots: for each run `i`, `buildSlots(text_i, params, run.fontProfile)`, tag each slot with `ri = i`; merge; sort by `score` desc with a stable tiebreak (`ri, wi, pos`). Quality ranking holds **across** runs.
5. Binary-search total tatweel count `N ∈ [1, totalLetters·8]` (unchanged bounds): for `N`, `cand = applySlotsMulti(texts, slots, N)`; `width = Σ measure_i(cand_i)`; if `width <= target` keep and raise `lo`, else lower `hi`. Adding tatweels only grows width, so width is monotonic in `N` and the search is valid.
6. Return `best.map(t => ({ text: t }))`.

**2. `applySlotsMulti(runTexts, slots, n) → string[]`** — new, generalizes `applySlots`. Slots are keyed `{ ri, wi, pos }`; round-robins the top-`n` and splices tatweels into the correct run's text (each run split on `' '` exactly as it was when its slots were built).

**3. `justifyLine(text, targetWidth, ctx, params, fontProfile)`** — refactored to delegate: wraps the args as a single run `{ text, measure: s => ctx.measureText(s).width, fontProfile }` and returns `justifyRuns([run], targetWidth, params)[0].text`. **Signature and observable behavior unchanged** — existing callers and tests are unaffected.

**4. Spacing/scale — value-returning**

```
computeRunSpacing(runs, targetWidth, params) → { wordSpacing, fontScale }
  params = { targetFill = 1, maxWordSpacing, minWordSpacing, maxScaleDown = 0.06, refFontSize }
```

Generalizes `justifyMisra`'s spacing/scale logic (`ashaar.js:330-351`):

1. `natural = Σ measure_i(text_i)`; `available = targetWidth * (params.targetFill || 1)`.
2. `gaps` = total word gaps across all runs (sum of inter-word spaces across the runs' texts).
3. `desired = gaps ? (available − natural) / gaps : 0`; `wordSpacing = clamp(desired, minWordSpacing, maxWordSpacing)`. `wordSpacing` may be negative (tighten).
4. Clamp defaults derive from a **reference font size** = max run `fontSize` (fallback `params.refFontSize`, then 16): `maxWordSpacing = ref·0.28`, `minWordSpacing = −ref·0.08` — because one `wordSpacing` value applies to the whole misra.
5. `wordSpacing` is a property, not text — so the width after spacing is computed **analytically**, not re-measured: `naturalAfterSpacing = natural + wordSpacing · gaps`. If `naturalAfterSpacing > available` (only possible when `natural` overflows and `wordSpacing` is clamped negative at `minWordSpacing`), `fontScale = max(1 − maxScaleDown, available / naturalAfterSpacing)`, applied **uniformly to all runs** so relative sizes are preserved. `fontScale ≤ 1` (shrink-only, matching today); otherwise `fontScale = 1`.
6. Returns `{ wordSpacing, fontScale }`. **Consumers apply them** as properties (web: `word-spacing` + `font-size:%`; Word: the OOXML equivalent) — that mapping is a follow-up spec.

**5. Shared internal helper — `measureRunsNatural(runs)`** — strips tatweels and returns `Σ measure_i(text_i)`; both public functions build on it.

Unchanged and reused: `buildSlots`, `applySlots`, `spreadTatweels`, `tatweelSlots`, `justifyLines`.

## Data flow

```
consumer (web / Word)
  → discovers ordered runs of a misra, binds a per-run measure()
  → kashida:  justifyRuns(runs, targetWidth, params)      → [{text}]  → write each run's text back
  → spacing:  computeRunSpacing(runs, targetWidth, params) → {wordSpacing, fontScale} → set properties
```

## Boundary decision (v1)

Slots and measurement are per-run. **Word-aligned runs are exact** — a space between runs makes independent per-run measurement correct, which is the common case (whole words are bolded / re-fonted). **Runs that split mid-word are approximate:** the summed width of independently measured substrings differs from the shaped whole, and no tatweel is inserted exactly at the seam (two fonts cannot shape-join there anyway). Documented as a known limitation, not a bug.

## Error handling / edge cases

- `runs = []` → `[]` (kashida) / `{ wordSpacing: 0, fontScale: 1 }` (spacing).
- Whitespace/empty run → preserved; contributes 0 slots and 0 natural width.
- No legal slots anywhere (Latin text, or the `لله` special case) → kashida returns stripped text unchanged.
- A run missing `measure` → throw `TypeError` (explicit contract).
- Pre-tatweeled input → identical result to the bare line (strip at entry).
- `gaps === 0` (single word) → `wordSpacing === 0`.

## Testing — `tests/ashaar-justify-runs.test.js` (pure node, deterministic fake `measure`)

Fake `measure = s => s.length` (each tatweel / space char has width 1) unless a test needs per-run weighting.

**Kashida (`justifyRuns` / `justifyLine`):**
- **Parity:** `justifyLine` output matches a golden — delegation is transparent, existing behavior preserved.
- **Two runs:** total width reaches target; tatweels land in both runs; result array length and order match the input.
- **Quality-first across runs:** run A given a boosting `fontProfile` fills before run B.
- **Edge:** `[]`; whitespace run; no-slot Latin run returned unchanged; already `≥ target` returns stripped; reducibility (pre-tatweeled input == bare input).

**Spacing (`computeRunSpacing`):**
- Two runs of differing `fontSize` → a single `wordSpacing` that fills to target.
- Clamp honored at `maxWordSpacing`.
- `fontScale < 1` only when even max `wordSpacing` still overflows; `fontScale === 1` when spacing suffices.
- `gaps === 0` (single word) → `wordSpacing === 0`.

Add the file to the `npm test` script.

## Follow-up specs (not this project)

1. **Web preview consumer** — make `ashaar.js` `justifyMisra` / `justifyEl` run-aware over styled child spans, calling `justifyRuns` / `computeRunSpacing` and applying results to the DOM.
2. **Word add-in consumer** — read a cell's runs with per-run font/size, call the primitive, and write back preserving runs (OOXML), replacing the current flattening `insertText` (`taskpane.js:1281`); map `wordSpacing`/`fontScale` to the OOXML equivalents.
