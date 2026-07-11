# Poetry Kashida Strength — 1–10 elongation:spacing ratio — Design

**Date:** 2026-07-11
**Status:** Approved (design), ready for implementation plan
**Branch:** `feat/guided-justification-ux`
**Scope:** Redefine the **poetry** "Kashida strength" control from a 0–24 fill-amount knob into a **1–10 ratio** of **elongation → spacing**: how much of each line's fill is achieved by the font's elongation mechanism vs. inter-word spacing. Total fill always targets the column edge; strength moves the *balance*, not the *amount*. Extends the shipped hybrid-fill + per-cell-mechanism work.

## The model

For each cell: `fillTarget = colPx` (fill to the column edge), `natural` = measured bare width, `gap = max(0, fillTarget − natural)`. Strength `s ∈ [1,10]` sets the **elongation share** `φ = (s − 1) / 9 ∈ [0,1]`:

- Elongate (via the cell's mechanism) toward **`natural + φ·gap`** — only φ of the gap.
- Residual spacing (existing `capMicroSpaces`, 0.28em/gap cap, accept-short) fills the rest toward `fillTarget`.

So:
- **s = 1** → φ = 0 → **all spacing**, no elongation (spacing fills up to its cap, accept-short beyond).
- **s = 10** → φ = 1 → **all elongation** (the mechanism fills the gap; spacing only for the minor residual it can't reach).
- Between: the mechanism fills φ of the gap, spacing the remainder.

**"Elongation" = the cell's font mechanism** (per-cell auto-detect, already shipped): Jameel → Kasheeda font-swap (its "ligature alternates"); Mehr → tatweels; generic → tatweels; **Noto/Gulzar → none**, so a whitespace-mechanism cell is spacing-only at *every* strength (φ has nothing to drive).

**Ligature alternates beyond Jameel** (OpenType `calt`/`jalt`/stylistic-set widening on arbitrary fonts) is **filed on ashaar-js#7, not built here** — for now "alternates" = Jameel font-swap; other fonts elongate with tatweels only.

**Prose** uses the SAME slider with different semantics (tatweel *cap*: 1→2 tatweels, 3→alternates+3, 7→full cap+alternates, 10→cap×1.5) — **recorded on ashaar-js#8**, built when prose justification exists. Out of scope here.

## Components

### 1. `strengthToElongationShare` — new pure helper (`word-html.js`, `AshaarWord`)
```
strengthToElongationShare(strength) → φ   // (clamp(strength,1,10) − 1) / 9, in [0,1]
```
Node-unit-testable (1→0, 10→1, 5→~0.44, out-of-range clamps).

### 2. Slider rescale 0–24 → 1–10
- `taskpane.html`: `#tatweel-count` and `#qaseeda-strength` → `min="1" max="10" value="7"` (7 ≈ "mostly elongation"). Labels stay "Kashida strength".
- `options()` reads the value as before (now 1–10).

### 3. φ-target in the elongation branches (`justifySelection`, Office.js)
Each per-cell elongation branch currently targets the full column, then adds residual spacing. Change the **elongation target** to `natural + φ·gap`:
- **Jameel font-swap:** pass the φ-scaled target (not `colPx`) to `selectSwapRuns` so it swaps only enough fasls to reach it; residual spacing unchanged.
- **Mehr tatweel:** pass the φ-scaled target to the discrete `selectSwapRuns`; residual spacing unchanged.
- **generic:** pass the φ-scaled target to `AshaarJustify.justifyRuns`; residual spacing unchanged.
- `natural` per cell = sum of run natural widths (already measured for the residual computation); `gap = colPx·targetFill − natural`.

### 4. Decouple targetFill from strength (engine poetry path)
Today strength → `targetFill` (fill amount). Now strength → φ (ratio), and **`targetFill` becomes a fixed fill-to-edge** (`≈ 1.0`) for the engine poetry path. `strengthToTargetFill` is **no longer** the strength consumer for engine kashida; it is repurposed/retired for that path (the profile path, §6, follows the same rule).

### 5. Migration of the other 0–24 consumers
Rescale each to the 1–10 domain so nothing reads a stale range:
- **`strengthToKashidaLevel`** (`word-html.js`, css "Word justify" → Word kashida level): thirds of **1–10** (`≤3 low`, `≤6 medium`, else `high`) instead of thirds of 0–24. (css mode keeps level semantics — it is not the engine ratio.)
- **`kashidaExpansionFraction`** (css column expansion): domain 1–10 (0 at 1, ~0.15 at 10).
- **`word-html.js:88`** insert-path `strengthToFill` (`0.90 + count/24·0.10`): rescale to 1–10, or fold into the fixed fill-to-edge (decide in plan).
- **`strengthToTargetFill`** (`profiles.js`): see §6.

### 6. Profiles
- **Stored profiles** carry a 0–24 `justify.strength`. On load, **remap** old values to 1–10 (`round(1 + old/24·9)`, clamped) so existing profiles don't break. New profiles store 1–10.
- **`applyProfileToQaseeda`**: profile strength now means φ (poetry ratio) too, consistent with the free-form path — the profile path adopts the same φ-target model (§3) rather than `strengthToTargetFill`.

### 7. Result-panel / guide wording (deferred hooks)
The "capped stretch strength" recourse and the 0.28em accept-short guidance (hybrid §5a) now read against a 1–10 ratio — reflect that when the guided-justification §1 Result panel lands. No build here beyond the number range.

## Testing
- **Node (pure):** `strengthToElongationShare` (mapping + clamp); rescaled `strengthToKashidaLevel` (1–10 thirds); rescaled `kashidaExpansionFraction`; profile strength remap (0–24 → 1–10).
- **Manual in Word (Office.js):** at s=1 a line fills by spacing only (accept-short if wide); at s=10 by elongation with minimal spacing; mid-values shift the balance; verified for Jameel (swaps), Mehr (tatweels), generic (tatweels); Noto/Gulzar spacing-only at all s. Existing profiles still apply (remapped). css "Word justify" level still low/med/high across 1–10.

## Out of scope (tracked)
- Prose strength (tatweel cap) → ashaar-js#8.
- Ligature alternates on non-Jameel fonts (OpenType) → ashaar-js#7.
- Per-run mixed mechanisms → ashaar-js#7.

## Open questions
- Slider default: **7** proposed (mostly elongation) — confirm in plan/UX.
- `word-html.js:88` insert-path fill: rescale vs. fold into fixed fill-to-edge — plan decides.
