# Hybrid Fill (elongation + residual spacing) — Design

**Date:** 2026-07-10
**Status:** Approved (design), build-after-verification
**Branch:** `worktree-nastaliq-kashida-fonts`
**Scope:** After calligraphic elongation (Mehr tatweel / Jameel font-swap) undershoots the column, close the residual with **capped inter-word micro-spacing**; if the cap is hit, leave the line slightly short (no glyph shrink, no column change). Internal to the elongation paths — **no new user mode**. Whitespace fonts (Gulzar/Noto) are unchanged (already spacing-only).

## Motivation (validated by the demo)

Neither elongation engine can fill an arbitrary width alone:
- **Mehr** only elongates at limited, whitelisted points (and — see the demo finding below — its browser rendering of tatweel is itself unreliable).
- **Jameel** only widens fasls that *have* a Kasheeda variant.

The `demo/kashida-fitting.html` harness (real engine + real fonts) confirmed this directly: many misras reach only 52–75% from elongation and the rest is inter-word space. For Mehr lines with no eligible word-endings, elongation contributes **0** and spacing does all the work. So residual spacing isn't a nicety — it's what makes the "fill the column" promise real.

## Demo findings that shape this design

- **Jameel font-swap renders clean, real elongation** in-browser (Kasheeda `خواب` ≈ 2× base width; full lines fill and stay readable). Font-swap is the solid path.
- **Mehr tatweel does not render cleanly in the browser.** Medial U+0640 is inert (zero advance); a *trailing* U+0640 after a whitelisted word-final letter adds advance width (so canvas "measures" a fill) but the shaped result breaks/stacks rather than forming a connected kashida. **Mehr's elongation is therefore unverified for real use and is pending an on-device Word check** (browser HarfBuzz ≠ Word's shaper). Until then, treat Mehr's fill as *mostly spacing* — which this hybrid provides.
- Because our fitting depends on **canvas measurement**, any elongation the canvas can't measure cleanly can't be driven regardless of how Word renders it. This is why residual spacing (which is measurable and reliable) is the dependable half.

## Order (fixed)

elongate → measure achieved width → residual = `colPx − achieved` → **capped** micro-spacing → accept short. Never spacing-first.

## Components

### 1. `capMicroSpaces` — pure, new (`src/taskpane/kashida-residual.js`, `AshaarResidual`)
```
capMicroSpaces(residualPx, gaps, spaceGlyphPx, sizePx, capEm=0.28) → integer count
```
Total added spacing never exceeds `capEm × sizePx × gaps` (0.28em/gap — the same ceiling `computeRunSpacing` uses, so hybrid and spacing-only stay visually consistent). Returns 0 when `residualPx ≤ 0`, `gaps ≤ 0`, or `spaceGlyphPx ≤ 0`. Node-unit-testable. (Validated in the demo as an inline copy.)

### 2. `injectSpaceRuns` — pure, new (same module)
```
injectSpaceRuns(runs, n) → runs'   // runs: [{text, swap|…}]
```
Distributes `n` hair-spaces (U+200A) as evenly as possible across the inter-word `" "` runs of a font-swap/tatweel run list, so `runsToMisraXml` emits the widened gaps. Node-unit-testable (even distribution, remainder handling, zero-n passthrough).

### 3. Mehr (tatweel) chaining — in the existing `kashida` branch
After the tatweel fill produces `outTexts`, measure achieved width, `n = capMicroSpaces(...)`, then `outTexts = distributeMicroSpaces(outTexts, n, MICRO_SPACE)` (reuses the existing MICRO_SPACE + `distributeMicroSpaces` + its U+200A→regular-space fallback at `taskpane.js:1328`). **Gated on the Mehr Word verification** — if Mehr is later reclassified to whitespace, this chaining is moot (the whitespace path already spaces).

### 4. Jameel (font-swap) chaining — after `selectSwapRuns`
achieved = Σ(`run.swap ? ww[i] : wb[i]`), `n = capMicroSpaces(...)`, `sel.runs = injectSpaceRuns(sel.runs, n)`, then `runsToMisraXml` as today. Micro-spaces sit at PAW boundaries (between words) so they never disturb shaping.

### 5. Policy
Accept-short: no font-shrink (`maxScaleDown` unused). Shared 0.28em/gap cap. Column-widening (guided-justification §4, deferred) remains the future lever for lines the cap can't fill.

## Testing
- `capMicroSpaces`, `injectSpaceRuns` → Node unit tests (residual→count, cap enforcement, zero-gap/zero-residual edges; even distribution across space-runs).
- Chaining is browser-side → the `demo/kashida-fitting.html` harness is the visual verification aid; on-device confirmation rides the same Word pass as the elongation.

## Risks
- **U+200A consistency** across Word builds — already relied on by the spacing mode, with the regular-space fallback.
- **Mehr elongation unverified** (demo showed browser breakage) — the hybrid's Mehr chaining is only meaningful if Mehr's elongation is confirmed usable in Word; otherwise Mehr rides the whitespace path (which this design's spacing half already embodies).
- Re-measuring Mehr's tatweel'd text needs the font loaded — already ensured (base + Kasheeda preloads in place for font-swap; the tatweel path preloads `repName`).

## Build-after-verification
This layers on top of the elongation. Build once:
1. Jameel `w:cs="…Kasheeda"` is confirmed to resolve the wider face in Word (Task 8 on-device check), and
2. Mehr's fate is decided (its Word check).
Jameel's residual spacing (§4) can be built as soon as (1) passes, independent of Mehr.

## Out of scope
- Column-widening / expressive fill (guided-justification §4).
- Word-native `distribute` justification (rejected: can't honor the 0.28em cap; over-spaces).
- Any change to the elongation selectors themselves (this only chains onto their output).
