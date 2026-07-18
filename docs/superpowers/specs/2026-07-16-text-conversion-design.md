# Text Conversion (Double-Press ⇄ Modern) — Design Spec

**Date:** 2026-07-16
**Status:** Draft — user-approved (brainstorming); ready for implementation planning.
**Scope:** An on-demand find-replace subsystem that converts Arabic-script text between the **legacy LD "double-press" encoding** (used by old AL-KANZ-era display fonts that lacked Unicode codepoints for Indic-extension letters and special LD symbols) and a **modern encoding** (proper Unicode where it exists, Fatemi Maqala `!keyword!` strings otherwise). Exposed as a new "Convert" tab in the task pane, alongside Table Input / Conversion / Booklet / Styles.

This was explicitly scoped out of the Styles spec (see the "Explicitly out of scope" section of `docs/superpowers/specs/2026-07-16-ashaar-styles-design.md`) and is its own subsystem.

## Overview

Old Lisan-ud-Dawat (LD) documents were typed for display fonts (chiefly **AL-KANZ**, "Al-Kanz for Windows.ttf", 2013) that had no Unicode codepoints for the Indic-extension letters (پ چ گ ٹ ڈ ڑ ں ہ ے), for certain repurposed diacritic marks, or for the special LD honorific symbols. Authors produced those glyphs by **double-pressing a base Arabic letter** (e.g. typing `حح` to display چ), by **repurposing keyboard-layout diacritics** (e.g. the sukun key standing in for khari zabar), and by typing **placeholder characters** the font shaped into honorific symbols. Modern fonts (Fatemi Maqala, Kanz-al-Marjaan) instead use real Unicode letters/marks where they exist and exclamation-wrapped `!keyword!` ligature strings where they do not.

This feature is a **preset-driven find-replace** that migrates such documents **in either direction**:
- **To Modern** (Legacy → modern) — the common case: open an old AL-KANZ document, produce clean Unicode/Fatemi text.
- **To Legacy** (modern → double-press) — still needed, because some fonts in use continue to require the old mechanism; this direction must be as easy to reach as the other.

The whole system reduces to **an ordered list of literal find→replace operations per direction**, applied to the Word document via the Office.js search/replace API so run formatting is preserved. The interesting logic (ordering, the kashida-escape rule, whole-word matching) lives in a pure, node-testable data module; a thin Office.js layer orchestrates it.

## §1 — Architecture

Mirrors the Styles feature's discipline (pure data module + thin Office.js orchestration + its own task-pane mode). `src/vendor/` is untouched.

- **`src/taskpane/word-conversion.js`** — pure, node-testable UMD module (same pattern as `word-styles.js`/`word-html.js`). Owns:
  - the **mapping table** (all rows, with category, forms, match scope, lossy flag);
  - `buildOperations(direction, enabledIds)` → an **ordered** array of literal ops `[{find, replaceWith, wholeWord}]`. The ordering encodes the kashida-escape precedence (§4); there is no runtime context-sensitivity.
  - `convert(text, direction, enabledIds)` → a pure string transform that applies that op-list to a plain string. Used **only as the test oracle** (proving the op-list is correct); it is *not* on the live document path.
- **`src/taskpane/conversion-pane.js`** — thin Office.js orchestration (same pattern as `styles-pane.js`). Wires the Convert panel DOM, runs the op-list against the document via `search()` + range-replace inside `Word.run()`, manages named presets, and reports per-category counts.
- **`scripts/generate-conversion-table.mjs`** — a build/data-generation script (parallel to `scripts/sync-ashaar-vendor.mjs`) that extracts the `mark` and `symbol` rows from the fonts (§3) and emits them into `word-conversion.js`. Reproducible and re-runnable so the table survives font updates. The `letter` tier is authored by hand (it is complete and font-independent).
- **New "Convert" tab** — a fifth mode in `taskpane.html`; `setMode()` in `taskpane.js` is extended to toggle its panel and to hide the poetry-only chrome (settings panel, justify actions, fonts strip), exactly as Styles/Booklet modes already do.
- **`tests/word-conversion.test.js`** — node `assert` (no jest/mocha), same as the existing suite.

## §2 — Data model

Each mapping is one row:

```js
{
  id: "seen-baariye",          // stable id, used by selection + presets
  category: "letter" | "mark" | "symbol",
  legacy: "سس",                // the legacy (double-press / placeholder) form
  modern: "ے",                 // the modern form (see resolution rule below)
  label: "ے — baari ye",       // human-readable label for the checklist
  wholeWord: false,            // match scope; true only where the modern form is common text
  lossy: false                 // true = not safely round-trippable; flagged in the UI
}
```

**Modern-form resolution rule (marks + symbols):** the `modern` value is a **real Unicode codepoint / sequence where one exists**, and a Fatemi **`!keyword!`** string only where none does. Decided per-row during extraction (§3).

### Category `letter` — doubled consonants (complete, ships as-authored)

The finalized letter table (bidirectional unless noted):

| Unicode (modern) | Double-press (legacy) | Notes |
|---|---|---|
| ے | سس | |
| ہ | ظظ | |
| ں | طط | |
| گ | كك | |
| چ | حح | |
| ٹ | ضض | |
| ڑ | رٌ | rā + dammatan (a *mark*, not a doubling) |
| ڈ | دٌ | dāl + dammatan |
| پ | ثث | |
| چھے | ؛ | `wholeWord: true` for the Modern→Legacy direction — only the standalone word `چھے` collapses to `؛`, never a substring inside another word |

Not differentiated (no conversion — treated as identical): ک/ك (kaf), ی/ي (yeh). No double-press form (excluded): ژ (zhe), ھ (do-chashmi he). Arabic he ه / ہ / final "jhatka" he are visually close enough that they are left alone.

Each doubled-consonant row (سس ظظ طط كك حح ضض ثث) also implies a **kashida-escape op** (§4); this is derived automatically when the row is enabled and is not a separately selectable row.

### Category `mark` — legacy Arabic-101 keyboard repurposings

Single legacy characters (emitted by specific keystrokes on the Arabic 101 layout) that old fonts displayed as different marks:

| Keystroke | Legacy char | Modern | Notes |
|---|---|---|---|
| Shift+X | sukun (U+0652) | khari zabar / dagger alef (U+0670) | **lossy** — a genuine sukun would be reinterpreted |
| Shift+C | (extract) | high jeem — codepoint if one exists, else `!keyword!` | exact legacy char + modern form confirmed in the extraction spike |
| Shift+V | (extract) | high noon — codepoint if one exists, else `!keyword!` | |

Exact legacy characters and modern targets for high jeem / high noon are finalized in the extraction spike (§3, §5) rather than pinned here.

### Category `symbol` — honorifics / special LD symbols

Legacy AL-KANZ placeholder characters ⇄ modern honorific forms (Unicode codepoint where it exists — e.g. U+08D5, U+08D6, U+FD48 and similar — else Fatemi `!keyword!`). The honorifics **are** the special LD symbols (a single tier, not a separate "expansion" table). This tier is **generated** by the extraction script (§3), not hand-authored.

## §3 — Extraction of the `mark` and `symbol` rows

Authoritative sources, all present on disk:

- **Legacy side:** `~/Downloads/Al_Kanz_Fonts_For_Windows/Al-Kanz for Windows.ttf` (2013) — the old placeholder/double-press font. Its `cmap`/GSUB reveal which placeholder characters map to which special glyphs.
- **Modern `!keyword!` side:** `assets/fonts/FatemiMaqala-Regular.ttf` (already vendored) — its GSUB ligature rules define the `!keyword!` input strings.
- **Authoritative cross-reference:** the `~/Kanz-al-Marjaan` project — `sources/KanzAlMarjaan-Regular.ufo/features.fea` (the substitution rules) and `scripts/build_honorifics.py` / `check_honorifics.py`, which define the honorific glyph set and document parity with FatemiMaqala. Its venv provides `fonttools`.

**Implementation task 1 is an extraction spike:** using `fonttools`, enumerate the honorific/symbol glyphs and their input encodings in the three sources, decide each row's canonical `modern` form via the resolution rule (§2), confirm the `mark` rows, and emit the `symbol`/`mark` rows into `word-conversion.js` via `scripts/generate-conversion-table.mjs`. The relevant fonts are copied into the add-in repo (or read from their source paths by the generator) so the build is self-contained.

## §4 — Conversion engine

`buildOperations(direction, enabledIds)` returns literal ops in **precedence order** so the kashida-escape rule needs no runtime context-sensitivity — ordering alone keeps genuine double letters distinct from double-press pairs.

The kashida-escape rule: in double-press text, a word that genuinely needs two identical base letters adjacent (which would otherwise merge into the special glyph) is written with a **tatweel (ـ, U+0640) between them**. So:

- **Legacy → Modern:** first run the contiguous double-press substitutions (`سس → ے`, …); *then* the separator-drop (`سـس → سس`). Because `سـس` contains the tatweel, the first rule never matches it, so ordering keeps them distinct.
- **Modern → Legacy:** first run "protect genuine doubles" (`سس → سـس`, inserting the separator); *then* the letter substitutions (`ے → سس`, …). The `سس` newly created from `ے` is immune because the protect rule already ran.

Within each batch, ops are emitted **longest-`find`-first** to avoid partial shadowing. Symbol/mark ops and the whole-word `چھے → ؛` op are appended in their safe positions.

`conversion-pane.js`, inside a single `Word.run()`:
1. Resolve `scope` = `context.document.body` (Whole document) or `context.document.getSelection()` (Current selection).
2. For each op **in order**: `scope.search(op.find, { matchCase: true, matchWholeWord: op.wholeWord })` → `context.load(results, "items")` → `sync` → for each hit `item.insertText(op.replaceWith, Word.InsertLocation.replace)` → `sync`. A **fresh search per op** is required (edits from a prior op would otherwise leave stale ranges).
3. Accumulate per-category replacement counts and surface them in the run report.

There are no content-control / poem-block considerations — these documents predate ashaar.js (§6).

## §5 — UI: the Convert panel

- **Direction — two explicit buttons** (never a silent toggle, so the wrong direction can't be run by accident):
  - **⟵ To Modern** (Legacy → modern)
  - **To Legacy ⟶** (modern → double-press)
- **Scope — a radio:** Whole document (default) / Current selection.
- **Preset dropdown + Save-as / manage.** Presets are **named subsets** of enabled rows, stored at the **add-in / roaming level** (they follow the user across the many legacy documents they convert), *not* per-document. Default preset "All".
- **Grouped selectable checklist:** rows grouped under "Letters", "Marks", and "Symbols / honorifics" headers; each row a checkbox showing both forms (e.g. `حح ⇄ چ`), with select-all / none per group. **All rows checked by default.** Rows with `lossy: true` (e.g. sukun → khari zabar) carry a small warning marker so the user knows to uncheck them when a document contains genuine instances of the source character.
- **Run report:** a count line after each run, e.g. "Converted 47 letters, 3 marks, 12 symbols."

## §6 — Safety / edge cases

- **No content-control / poem-block awareness.** Target documents predate ashaar.js; there are no tagged poem blocks to protect, and a lone tatweel between two identical base letters is unambiguously a genuine-double-letter separator.
- **Formatting preservation.** Range-replace via `insertText(..., "Replace")` keeps the replaced range's run formatting; intra-paragraph runs (bold, color, font) survive. This is the reason the engine uses search/replace rather than whole-text rewriting.
- **Font correction is not part of this feature.** Conversion changes characters only; switching an AL-KANZ document to Fatemi/Kanz-al-Marjaan is left to the Styles feature or to the user.
- **Lossy rows.** The sukun → khari-zabar `mark` row (and any other row flagged `lossy`) is not safely round-trippable; it is checked by default but flagged, and the user unchecks it for documents with genuine sukuns.
- **Harakat.** `رٌ` / `دٌ` intentionally use dammatan; a genuine `ر`/`د` + dammatan (vanishingly rare in these documents) would be converted — a documented known limitation.
- **Office.js feasibility spike (do first, alongside extraction):** confirm `search()` matches literal tatweel/diacritics and honors Arabic word boundaries for `matchWholeWord`, and confirm sequential per-op search/replace behaves within one `Word.run()`. If `matchWholeWord` proves unreliable on Arabic script, the `چھے → ؛` op falls back to a manual boundary check (search the bare string, then verify neighbors on each hit before replacing).

## §7 — Testing

- **Node (pure), against `convert()` as oracle:**
  - op-list ordering for both directions (kashida-escape precedence);
  - round-trip idempotency for the reversible set (`Legacy → Modern → Legacy` and back) — excluding rows flagged `lossy`;
  - kashida-escape correctness: `سـس` ⇄ genuine `سس`, distinct from `سس` ⇄ `ے`;
  - whole-word behavior: `چھے → ؛` only as a standalone word;
  - longest-`find`-first (no partial shadowing);
  - per-category counts.
- **Manual (Word):** a real AL-KANZ legacy document converted both directions; formatting preserved; Whole-document vs Current-selection; preset save / reload; lossy-row warning; reported counts correct.

## §8 — Non-goals

- **Live autocorrect as-you-type.** On-demand batch pass only. (Possible fast-follow, but technically awkward in Office.js — no reliable per-keystroke hook.)
- **Font swapping** (AL-KANZ → modern face) — delegated to the Styles feature / left to the user.
- **Content-control / poem-block awareness** — target documents predate ashaar.js.
- **`ژ` (zhe) and `ھ` (do-chashmi he) double-press forms** — never existed.
- **Mixing conversion with poetry justification** — this feature does not touch the kashida justification engine; the shared U+0640 tatweel is only ever interpreted as a double-letter separator in the narrow ordered contexts of §4.

## Resolved decisions

1. **One axis, three sub-tables.** Conversion is Legacy ⇄ Modern; the mapping table has `letter`, `mark`, and `symbol` categories that all move in the chosen direction together.
2. **Modern form = Unicode codepoint where it exists, else Fatemi `!keyword!` string** (per-row, decided in extraction).
3. **Engine = ordered list of literal find→replace ops per direction.** The kashida-escape rule is expressed purely through op ordering; no runtime context-sensitivity.
4. **Kashida (U+0640) between two identical base letters is the genuine-double-letter escape** — dropped on Legacy→Modern, inserted on Modern→Legacy.
5. **Apply via Office.js `search()` + range-replace** (Approach A) to preserve run formatting; the pure `convert()` transform is the test oracle only.
6. **On-demand batch pass, both directions first-class**, over Whole document or Current selection.
7. **Granular per-row selection**, all-on by default, with **named presets stored at the add-in/roaming level**.
8. **`symbol` and `mark` rows are generated** from on-disk fonts (`Al-Kanz for Windows.ttf`, `FatemiMaqala-Regular.ttf`, cross-checked against the Kanz-al-Marjaan `features.fea` / honorifics scripts) via `scripts/generate-conversion-table.mjs`; the `letter` tier is hand-authored and complete.
9. **Lossy rows are flagged, checked by default, individually uncheckable** (e.g. sukun → khari zabar).
10. **No content-control awareness, no font-swap, no live autocorrect** in v1.
