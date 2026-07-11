# Poetry Kashida — Concentration + Budget/Backfill Model — Design Spec

**Date:** 2026-07-11
**Status:** Approved (brainstorming) — ready for implementation planning
**Supersedes (in part):** `2026-07-11-poetry-strength-scale-design.md` — keeps its φ helper, 1–10 sliders, and domain rescale; **revises** how φ is applied (Tasks 4/5 of that work).
**Upstream issue:** ashaar-js#7 (strength-share / generic-path hybrid).

## Problem

At low Kashida strength (1–3) the justified poetry "doesn't feel like it prioritizes spacing over kashidas." Two root causes in the current model:

1. **Generic path spreads thin.** The vendor engine's generic tatweel distribution (`applySlotsMulti`) round-robins *one* tatweel across the top-N ranked slots, so a small elongation budget becomes *many short* kashidas rather than *a few full* ones. Jameel/Mehr don't have this (a whole-span swap / trailing tatweel is already a "full" kashida).
2. **Generic & profile paths are accept-short.** They stop at the elongation target with **no spacing backfill**, so a low-strength line shows a few kashidas and a *short line* — reading as "no spacing," the opposite of the intent.

Additionally, the old strength→`targetFill` fractions (0.90–1.0) were a workaround for not being able to size tables/cells from the task pane. Cell widths are now defined reliably via the grid-maker, so the fill target should be the true cell edge (minus a small buffer), not a strength-scaled fraction.

## Goal

Make Kashida **strength** a lever on the **elongation:spacing distribution**, not on how full the line gets:

- Every line fills to a fixed target `T = cell edge − buffer` regardless of strength.
- **Low strength →** a few *full* kashidas at the best-ranked positions, with **spacing always backfilling** the rest to `T`.
- **High strength →** more positions elongate (and, generic path, longer kashidas), elongation-dominant.
- Elongation concentrates on a *subset* of the best positions, growing that subset **as-needed within a line** and **with strength**.

## Non-goals / out of scope

- **Pane reflects the active Ashaar block's settings** — a separate UI↔block-state sync feature; its own spec afterward.
- **Broader table-design rework.** Cell widths are usually accurate (grid-maker defines individual misra/cell widths); the justify flow treats each cell's `columnWidth` as the desired width. Any wider table rework is a separate concern.
- **Prose semantics** (ashaar-js#8) and non-Jameel font-swap alternates remain out of scope.

## Model

### Fill target (fixed, not strength-scaled)

For each cell:

```
T = (columnWidth − 2·cellMargin) · 96/72  −  buffer
buffer = 0.28em   (0.28em expressed in px for the cell's font size)
```

The band `[T, colPx]` (where `colPx = (columnWidth − 2·cellMargin)·96/72` is the true edge, ≈`0.28em` wide) is **overflow tolerance for the total line**: because spacing is discrete (micro-spaces), the last micro-space may push the *total* width into that band rather than land exactly on `T` — that is accepted rather than clipped short, and rare overflow slightly past `colPx` is tolerated (the buffer absorbs it). Elongation itself never enters this band: it aims for `B ≤ T` (see below).

### Strength → elongation budget

```
φ = (clamp(s,1,10) − 1) / 9            // reused from the strength-scale work
natural = tatweel-free width of the line (measured per run in its own font)
B = natural + φ · (T − natural)        // the elongation budget target
```

Elongation aims for `B`; **spacing always covers the remainder** `T − achievedElongation`, on **every** path. `φ` is the elongation *share* of the gap; spacing is `(1 − φ)` of it.

### Position growth & concentration

The engine spends the budget on the **best-ranked positions first** (`buildSlots` order: calligraphic tier + font-quality bonus), and engages more positions two ways:

- **As-needed within a line:** keep adding the next-best position until `B` is reached; the rest is spacing.
- **By strength:** higher `s` → larger `B` → more positions (and longer kashidas) before spacing takes over.

**Concentration (generic path only):** instead of one tatweel each across the top-N slots, pile tatweels onto a position up to a **per-position em cap** (`perPositionEm = 0.5`, i.e. `0.5 × fontSize` per join), *then* move to the next-best position. This yields *a few long* kashidas at low `B` and *many long* kashidas at high `B`. Jameel/Mehr are already "one full elongation per position," so this changes only the generic tatweel path.

### Safety valve C (low-strength position cap)

At low strength, additionally cap the *count* of engaged positions per line at `K(s)`:

```
K(1)=1, K(2)=2, K(3)=3, K(s≥4)=∞ (unbounded)
```

so 1–3 stay deliberately sparse even if `B` could engage more; spacing absorbs the difference. `K` is passed as `maxPositions` to the engine.

## Architecture

### Where the logic lives (upstream-first)

The generalizable concentration algorithm goes **into the ashaar.js engine** (submodule), synced into `src/vendor/`. Add-in-specific glue stays app-side.

| Concern | Home |
|---|---|
| Concentration algorithm + em-cap + maxPositions + achieved-width | **Upstream** `vendor/ashaar-js/ashaar-justify.js` (closes ashaar-js#7) |
| cell-width → `T`/buffer, `φ` strength policy, per-cell mechanism dispatch, Word OOXML | App-side `taskpane.js` |
| Spacing backfill (`AshaarResidual`), Jameel/Mehr swap paths (`kashida-fontswap.js`) | App-side (already there) |

### Upstream engine API

New export in `ashaar-justify.js` (existing `justifyRuns` untouched, backward-compatible):

```
justifyRunsConcentrated(runs, targetWidth, params) → { runs: [{text}], achievedPx, positionsUsed }
```

- `runs`: same shape as `justifyRuns` (each carries `.text`, `.measure`, `.fontSize`, `.fontProfile`).
- `targetWidth`: the elongation budget `B`.
- `params.perPositionEm` (default `0.5`): max kashida length per position, in em of that run's `fontSize` → `perPositionPx = perPositionEm · fontSize · 96/72`.
- `params.maxPositions` (optional): the C valve `K(s)`.
- `params.priorityTable` / `fontProfile`: passed through to `buildSlots` unchanged.
- Behavior: walk `buildSlots` best-first; pile tatweels onto a position until its em cap, then advance; stop when total width ≥ `targetWidth`, or `maxPositions` positions have been engaged, or slots are exhausted.
- Returns `achievedPx` (measured total) so the caller does spacing backfill, and `positionsUsed`.

Reuses `buildSlots`, `measureRunsNatural`, and per-run `measure`. Pure; node-testable with mock `measure` functions.

### Per-path behavior (add-in)

Target `T = edge − 0.28em`, budget `B = natural + φ·(T − natural)`, spacing backfills `achieved → T` everywhere.

| Path | Elongation mechanism | Change from current (strength-scale) work |
|---|---|---|
| **Jameel** (`selectSwapRuns`, app-side) | whole-span swaps, greedy by gain — already few/full | retarget swap target `→ B` (swaps kept ≤ `B`, as today); residual backfills achieved → `T`; buffer added |
| **Mehr** (`selectSwapRuns`, app-side) | trailing tatweel per eligible word, discrete | same as Jameel |
| **Generic** (`justifyRuns` round-robin) | **→ `justifyRunsConcentrated`** to budget `B` (perPositionEm=0.5, maxPositions=`K(s)`), then `AshaarResidual` spacing backfill to `T` | replaces round-robin; **adds** spacing backfill (was accept-short) |
| **Profile** (`applyProfileToQaseeda`, uses `justifyLine`) | route through the **same** `justifyRunsConcentrated` + backfill | drops `justifyLine`; adds backfill (was accept-short) |

### Buffer / overflow / idempotency

- **Overflow:** elongation always aims for `B ≤ T` — continuous generic tatweels hit ~`B` (em-capped); discrete Jameel/Mehr *underfill* `B` (a step is taken only while total ≤ `B`, as `selectSwapRuns` does today). Spacing then backfills `achieved → T`; because micro-spaces are discrete, the last one may land in the `[T, colPx]` tolerance band, and rare overflow slightly past `colPx` is accepted (`capMicroSpaces` headroom). At `s=1` (`φ=0`, `B=natural`) no elongation is engaged and the whole gap is spacing — the "spacing-only" case.
- **Idempotency:** `stripJustification` already removes tatweels (U+0640) and micro-spaces (U+200A/U+2009) before re-justifying; concentration uses the same U+0640, so re-justify strips to base and recomputes with no compounding.

## Plan shape (upstream-first, 4 phases)

- **Phase A — Upstream engine.** Add `justifyRunsConcentrated` to `vendor/ashaar-js/ashaar-justify.js`; add tests to `vendor/ashaar-js/test/justify.test.js` (mock-measure, em-cap, maxPositions, achieved-width, monotonicity, empty/degenerate). Submodule `npm test` green. Commit in submodule (closes ashaar-js#7).
- **Phase B — Sync.** `npm run sync:ashaar` → copies engine into `src/vendor/ashaar-justify.js`, stamps `ASHAAR_UPSTREAM_VERSION`; bump submodule pointer in the parent repo. Parent `npm test` green.
- **Phase C — Add-in integration.** `T = edge − 0.28em`; `B = natural + φ·(T − natural)`; route generic + profile paths through `justifyRunsConcentrated` + `AshaarResidual` backfill; retarget Jameel/Mehr to `B` with backfill to `T`; wire `K(s)` maxPositions; overflow tolerance. Node-test the pure helpers (target/buffer math, `K(s)`); the Office.js glue is manual-verify.
- **Phase D — Manual Word verification.** Ratio sweep (1/5/10) per font (Jameel swaps, Mehr tatweels, generic Fatemi concentration, Noto/Gulzar spacing-only); confirm low strength = few full kashidas + spacing to edge; overflow within tolerance; idempotent re-justify; profile apply.

## Testing

- **Upstream (submodule):** `justifyRunsConcentrated` — em-cap per position honored; `maxPositions` honored; `achievedPx` matches measured width; monotonic in budget; concentrates (fewer positions than round-robin for same budget); empty/no-slots degenerate cases.
- **Add-in (node):** `T`/buffer math; `B` from φ; `K(s)` mapping; idempotent strip→re-justify at the pure level.
- **Manual (Word):** Phase D above.

## Open decisions (locked)

- `buffer = 0.28em`, overflow tolerance up to `0.5em`/edge. (Locked.)
- `perPositionEm = 0.5`. (Locked; tunable param.)
- `K(1..3) = 1,2,3`; unbounded for `s ≥ 4`. (Locked; tunable.)
- `φ = (s−1)/9` linear (reused). If 1–3 still too busy after concentration + backfill, retune `φ` or tighten `K` — not expected.
