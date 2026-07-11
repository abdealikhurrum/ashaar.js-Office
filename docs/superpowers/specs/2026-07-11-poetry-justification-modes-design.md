# Poetry Justification Modes — Cell-fit & Natural-fit — Design Spec

**Date:** 2026-07-11
**Status:** Draft — under user review (brainstorming)
**Builds on:** `2026-07-11-poetry-kashida-concentration-design.md` (the shared concentrated-tatweel engine, already implemented: `justifyRunsConcentrated`, `strengthToMaxPositions`).
**Supersedes:** the single-mode fill of that spec — the residual fill is now **mode-dependent** (see §3/§4). The concentration *engine* is unchanged.

## Overview

Poetry justification gets two explicit modes that share one engine and one organizing quantity (the per-position **natural-width matrix**). They differ only in **where each cell's target width comes from** ("the width operator") and **how the residual is filled**:

| | **Cell-fit** (precise) | **Natural-fit** (harmony / forgiving) |
|---|---|---|
| Target width | the **cell edge** (operator keeps cell width = the position's natural width) | the computed **natural-width matrix**, + strength reach toward page/column |
| Strength means | elongation : spacing **ratio** | misra **width** (natural baseline → page/column) |
| Elongation | concentrated tatweels (em cap) | concentrated tatweels (em cap) — same engine |
| Residual fill | Word **`jc=distribute`** (no caps, no buffer) | **capped micro-spaces** (existing `capMicroSpaces`, 0.28em caps; per-instance override) |
| Table widths | must be accurate (operator-maintained) | irrelevant — grid-independent; tables auto-fit to the text |
| Best for | seasoned operators; cleanest text | good-enough regardless of table width |

Both modes fill each misra to "the natural width" for its position; Cell-fit reads that width off an accurate cell, Natural-fit computes it. The concentrated-tatweel engine, the natural-width matrix, and the content/spacing-cell model are shared.

## §1 — Shared: the natural-width matrix (harmony)

Visual harmony means corresponding cells across bandhs (stanzas) share a width. The organizing quantity is a **per-position natural-width matrix**, scoped to a **qaseeda profile**.

- **Content cells vs spacing cells.** A grid cell is either a **content cell** (holds misra text; participates in the matrix and is justified) or a **spacing cell** (a structural gap — e.g. marsiya's empty columns; no text, excluded from the matrix). This tagging is introduced here and reused by the follow-on spacing-cell-styling spec.
- **Position identity = grid signature.** A content cell's position within the repeating bandh template is `(row index within the bandh, grid-column start, grid-span)` on the shared (12-column) grid. Cells with the same signature across bandhs occupy the "same position."
- **Baseline per position** `W_pos` = the **longest natural (tatweel-free) text width** among all content cells at that position across every bandh of the profile, measured on canvas in each cell's own font (the same measurement the engine already uses).
- **Harmony scope = the qaseeda profile** (all its bandhs across blocks). For a **free-form justify** with no stored profile, the scope is the enclosing Ashaar Poem content control — `justifySelection` already gathers all its cells; the matrix is computed over those (trivial when the control holds a single bandh). All bandhs sharing a profile share one template, so signature-matching is safe. A bandh that doesn't fit the template gets its **own profile**. (Related-/sub-profiles for semantic organization are a future nicety, explicitly not in scope.)

The matrix is a pure function of `{cell text, position signature, font}` → `{position → W_pos}`, so it is **node-testable** independent of Word.

Example (marsiya bandh `A3 A2 A1 / B1 / C2 C1`):

```
A1 → 300pt   A2 → 300pt   A3 → 300pt
B1 → 325pt   C1 → …       C2 → 250pt
```

Each position balances within itself, so B1's different meter (325) does not drag A1 (300).

## §2 — Shared engine (already implemented)

`AshaarJustify.justifyRunsConcentrated(runs, targetWidth, params) → {runs, achievedPx, positionsUsed}` (concentrated tatweels, per-position em cap, `maxPositions = K(s)`). Both modes call it with a mode-specific `targetWidth`. `AshaarWord.strengthToElongationShare(s)→φ` and `strengthToMaxPositions(s)→K` are shared.

## §3 — Cell-fit mode

Target = the true cell edge; strength is the elongation:spacing ratio; Word's Distributed justification fills the residual (validated: the spike + the CSS `word-spacing` harness confirmed distribute stretches word gaps to the edge and composes with tatweels).

- `colPx = cell content width`. **No buffer** (Word distributes to the true edge).
- Tatweel budget `B = natural + φ·(colPx − natural)`, `φ=(s−1)/9`. Concentrated tatweels to `B` (em cap, `maxPositions=K(s)`).
- **Residual fill = `<w:jc w:val="distribute"/>`** on the cell paragraph. **No hair-space injection, no `capMicroSpaces`, no buffer** for this mode.
- **Write-path:** emit the cell paragraph as OOXML with `jc=distribute`.
  - Jameel already emits OOXML (`runsToMisraXml`) → set `jc=distribute` there.
  - Mehr / generic / profile currently `insertText` (which preserves the cell's existing jc, and Office.js `paragraph.alignment` has no "distribute") → switch these to emit an OOXML paragraph with `jc=distribute` + the tatweel'd text (preserving font/size), like the Jameel path.
- **Idempotency:** only tatweels (U+0640) to strip; no injected spaces.
- **Harmony:** comes from the operator keeping cell widths equal to the positions' natural widths. The natural-width matrix (§1) MAY be offered as an assist to set those widths, but Cell-fit itself just fills to the cell edge. (Auto-setting cell widths from the matrix is optional, not required.)

## §4 — Natural-fit mode

Target = the position's natural width (matrix), extended by strength toward the page; the old em-capped tatweel + capped micro-space mechanism fills it; tables auto-fit so cells follow the text. This IS the committed Tasks 4–6 behavior, **retargeted from the cell edge to the matrix, with the caps kept**.

- Per content cell at position `pos`: `target = W_pos + φ·(reach − W_pos)`, `φ=(s−1)/9`.
  - **`reach` = the cell's container width** — the page / column / cell it sits in (for a table cell, its logical column allocation = its grid-span's share of the page text-width; a free line → the page text-width). So s1 = `W_pos` (pure harmony — shorter misras elongate up to the longest at that position; the longest gets ~none), s10 = fills the container. (Exact grid-span → width derivation resolved in planning.)
- Concentrated tatweels (em cap, `maxPositions=K(s)`) to `target`; **residual = capped micro-spaces** via existing `AshaarResidual.capMicroSpaces` (0.28em/gap cap) + `distributeMicroSpaces` — unchanged from the committed paths.
- **Per-cell overrides (variances).** The Ashaar block stores per-cell setting *variances* keyed by cell position/id — a deviation from the profile defaults (e.g. a lifted cap, or a per-cell strength/width). The operator **selects that one cell in Word and adjusts it in the pane**; the variance is stored on the block/profile. The primary use is an **under-resolved** cell (caps bind short of its harmony width → override to match), but the mechanism is general per-cell customization. NOTE: the select-cell → pane-shows-its-settings editing UI overlaps the parked *pane-reflects-active-block* feature; they should be designed together.
- **Cell sizing (self-sufficient).** Natural-fit does not trust the operator's cell widths: it sizes each content cell's column to its position `target` from the matrix (deterministic per-position uniformity), then fills the text to it. (Equivalently, an auto-fit table lets the cell follow the elongated text — same visual result; the add-in setting widths from the matrix is the deterministic default.) This is what makes it "forgiving of table width."
- `jc` = normal alignment (right/center); no distribute.
- **Idempotency:** strip tatweels + micro-spaces (existing `stripJustification`), as today.

## §5 — Mode toggle (UI)

- A new control in the task pane selects **Cell-fit** vs **Natural-fit**, stored on the qaseeda profile (so a poem justifies consistently).
- **Default: Natural-fit** (works out-of-box without accurate cell widths).
- `justifySelection` and `applyProfileToQaseeda` branch on the mode to choose the target (cell edge vs matrix) and the residual (distribute vs capped micro-spaces).

## §6 — What changes vs the committed concentration work

- **Engine (Tasks 1–3):** unchanged and reused by both modes.
- **Cell-fit:** NEW `jc=distribute` residual + OOXML write-path for the flat-text paths; drop hair-space backfill/buffer for this mode.
- **Natural-fit:** the committed Tasks 4–6 (tatweels + `capMicroSpaces`) **retargeted** from `colPx` to the matrix `target`; caps kept; add per-instance override + auto-fit.
- **New:** natural-width matrix computation (§1), content/spacing-cell tagging, mode toggle (§5).
- The earlier "cap-lift" idea is **dropped** — the caps are Natural-fit's defining behavior.

## §7 — Non-goals (separate specs / future)

- **Spacing-cell styling** — colors, repeated fill glyphs, hemistich markers (`*`) in spacing cells. Depends on §1's content/spacing tagging; **own spec after this one.**
- **Pane reflects the active Ashaar block's settings** — separate UI-sync spec (parked earlier).
- **Related-/sub-profiles**, prose semantics.

## §8 — Testing

- **Node (pure):** natural-width matrix (grouping by signature + longest-natural per position, given cell texts/positions/measure); the φ target math for both modes; `K(s)`. Engine already node-tested.
- **Web harness:** Cell-fit distribute fill already validated (CSS `word-spacing` ≡ Word distribute); re-usable for regression.
- **Manual Word:** mode toggle; Cell-fit distribute per font; Natural-fit harmony (corresponding cells across bandhs land at one width; strength sweeps natural→page); per-instance override; idempotent re-justify; auto-fit behavior.

## Resolved decisions

1. Natural-fit `reach` at s10 = the cell's **container width** (page / column / cell; a table cell → its grid-span's share of the page). Exact grid-span → width derivation in planning.
2. Mode toggle **default = Natural-fit**.
3. Under-resolved cells handled by **per-cell variances** stored on the Ashaar block, edited by selecting the cell in the pane (see §4; the editing UI ties to the parked pane-reflects-active-block feature).
