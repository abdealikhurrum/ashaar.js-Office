# Handoff → Fable: mixed-font qaseeda justify is over-stretching & non-idempotent

**Date:** 2026-07-12 · **Branch:** `feat/guided-justification-ux` · **Repo:** `ashaar.js-Office` (Word task-pane add-in).
**Model context:** Opus 4.8 did the work below; handing to Fable to crack the remaining bug.

---

## 0. TL;DR — the open bug

Applying a qaseeda profile ("Save & Apply") to a **mixed-font** block **over-stretches** the text and **gets worse on every re-apply** (not idempotent). Letters shatter into isolated forms and injected micro-spaces (shown teal via debug highlight) multiply each pass.

Evidence (screenshots in `docs/superpowers/handoff-assets/`):
- `mixed-font-original.png` — clean input.
- `mixed-font-apply1.png` — after 1st apply: already over-spread, teal spaces appear.
- `mixed-font-apply2.png` — after 2nd apply: markedly worse — more spread, more shattering, more teal.

**User confirmed:** the block is **mixed-font** (deliberately stress-testing), and the **latest code was loaded** (task pane reloaded) when these were captured. So the bug reproduces on the current working tree (uncommitted changes included).

Two distinct symptoms, likely two causes:
1. **Progressive growth (non-idempotency).** Each apply widens further ⇒ some per-pass measurement is inflating (target grows ⇒ more elongation ⇒ even wider next time).
2. **Shattering (isolated letterforms).** A Nastaliq run is getting generic tatweels (U+0640) injected between joined forms, OR spacing is breaking ligatures.

Your job: root-cause both, make apply **idempotent** and **mechanism-correct** for arbitrary mixed-font misras. Use `superpowers:systematic-debugging` — **do not guess-fix**; the previous session already burned several plausible-but-incomplete fixes.

---

## 0b. CLEANEST REPRO — uniform Jameel, apply twice (START HERE)

**`jameel-uniform-apply2-breaks.png`**: an **all-Jameel** block (no mixing at all). First apply looks fine; **second apply breaks** — over-spread, shattered, teal spaces. This isolates the bug from mixed fonts entirely: **the font-swap round-trip is not idempotent.** Crack this first; it almost certainly subsumes the mixed-font case.

### Root cause (traced, high confidence — verify in Word)

Jameel fills by swapping individual **fasls** (sub-word segments) from the base face to the wider **Kasheeda** face. After apply 1, a single *word* can contain **both** faces (some fasls base, some Kasheeda). Now the round-trip:

1. **`misraRunsXml` (word-html.js:1322) sets `w:ascii = w:hAnsi = w:cs = <the fasl's face>`.** So within one word, some runs have `ascii=Jameel Noori Nastaleeq`, others `ascii=Jameel Noori Nastaleeq Kasheeda`.
2. **Office.js `Font.name` returns the ASCII font.** For a word/range whose runs have *different* ascii fonts, `Font.name` returns **`""` (mixed)**. (This is also why cs-only originally broke — `.name` is ascii, not cs.)
3. On apply 2, capture reads those fonts back:
   - `captureQaseedaTables` → `cell.body.font.name` for the whole (mixed-face) cell = `""` → `baseFaceOf("")` = `""` → `measureText` with an **empty font string** substitutes a default (wrong, and different from apply 1's clean base measure) → **`natPx` / `qMatrix` / `cTarget` drift**. Bigger target ⇒ more swaps ⇒ wider ⇒ even bigger next pass = **runaway growth**.
   - `captureQaseedaCellRuns` → a **partially-swapped word** reads `font.name = ""` → `descriptorForFontName("")` = **`generic`** (wordName null). So that word becomes a *generic* segment, gets **continuous U+0640 tatweels** injected (= **shattering** of the Nastaliq), and a `flattenSegs` inter-segment space is inserted around it (= extra growth). Fully-unswapped words read `base`, fully-swapped read `Kasheeda`; only the *partial* ones go `""` → they're the poison.

So the base-face normalization added this session **doesn't fire** because the font name it's handed is `""`, not a Kasheeda name. `baseFaceOf("")` can't recover "Jameel".

> **PROBE RESULT (2026-07-12, live Word, debug dump):** the ascii assumption is **false**.
> With `ascii=hAnsi=base` on every run (fix A implemented + verified stored via `getOoxml`),
> `Font.name` still tracked the **cs** face per word: unswapped word → base name, fully-swapped
> word → Kasheeda name, partially-swapped word → `""`. So Office.js `Font.name` reads the
> COMPLEX-SCRIPT font for Arabic runs, and `""` when cs is mixed within the range. Fix (A) is
> dead — mixed cs within a word IS the swap mechanism. Also observed: the CELL-level read
> resolves through the paragraph-mark theme run to the theme default (**"Aptos"**), which
> inflated `natPx` → target 6642→8674tw → destructive rebuild pinned to Aptos (the shattering).
> **Fix (B) implemented instead:** per-word fonts persisted in the tag (`payload.runFonts`,
> `packRunWords`/`reconcileRunWords`/`setTagRunFonts` in word-html.js), healed on capture, and
> `natPx`/`qMatrix`/`repFont` recomputed from the reconciled runs in both passes.

### Candidate fixes (for evaluation — pick after confirming in Word)

- **(A) Make `font.name` read back a single family.** In the font-swap emit, set `w:ascii`/`w:hAnsi` to the **base** face for *every* fasl (both base and Kasheeda runs), and only `w:cs` to the actual (base/Kasheeda) face. Arabic renders via `cs` (the run is `w:rtl`), so Kasheeda still shows; but `Font.name` (=ascii) reads back **base uniformly** → cell + every word report "Jameel Noori Nastaleeq" → measurements normalize → idempotent. Requires `misraRunsXml` to accept a separate `asciiName` per run. **Depends on the assumption `Font.name == ascii for rtl runs` — verify.**
- **(B) Persist per-word fonts in the content-control tag** (`AshaarWord.setTag`/`parseContentControlTag`) at first apply and reuse on re-apply instead of re-reading from the rendered cell. Makes per-word font **authoritative in the tag**, sidestepping `Font.name` entirely. Most robust; larger change. (This is handoff §5.4.)
- **(C) Never classify a `""`/unresolved font as `generic`** when the block's known font family is a single registry font; treat `""` as "inherit cell/qaseeda font". Cheap guard, but doesn't fix the `natPx`/`qMatrix` drift by itself — combine with (A) or (B).
- **(D) Measure `natPx`/`cNatural` robustly when `font.name==""`** — fall back to `repName`/`profile.font` (the qaseeda's font) rather than an empty font string.

Recommended: prove the `Font.name` behavior with a one-off probe (write a run with `ascii=base, cs=Kasheeda`, read `.name`), then do **(A)+(D)**, or go straight to **(B)** if the tag is the cleaner source of truth.

---

## 1. What this project is

Word add-in for typesetting Urdu/Arabic/Persian poetry with kashida justification. No build step, vanilla ES5/UMD, Office.js. Tests are plain `node tests/*.js` (`npm test`). See `CLAUDE.md` for the full map. Dev server: `npm start` (port 3000, `Cache-Control: no-store`); **Word only reloads JS when the task pane is reloaded.**

A **qaseeda** = a named profile linked across one or more "Ashaar Poem" content-control blocks (bandhs). "Save & Apply" runs `applyProfileToQaseeda(name)` over **all** blocks tagged with that qaseeda, so touching one bandh re-applies to every bandh.

### Font mechanisms (`src/taskpane/fonts.js`)
`AshaarFonts.descriptorForFontName(realFontName)` → `{mechanism, wordName, kasheedaName, tatweelRules}`. Mechanisms:
- `whitespace` — Noto Nastaliq, Gulzar, Scheherazade. **Shatter under injected tatweels** → fill by spacing only.
- `tatweel` — **Mehr Nastaliq Web**. Discrete *trailing* tatweel on whitelisted isolated/final letters (`mehrElongate`). Medial U+0640 is zero-width in Mehr.
- `font-swap` — **Jameel Noori Nastaleeq**. Fills by swapping whole fasls (connected segments) from the base face to the wider **Kasheeda** face (`kashida-fontswap.js`). No tatweel char is added; width comes from the face.
- `generic` — anything unrecognized (e.g. Fatemi Maqala, Amiri). Continuous tatweel engine (`AshaarJustify.justifyRunsConcentrated`).

Registry match is by exact `wordName`/`kasheedaName`. **Both Jameel faces resolve to `wordName = "Jameel Noori Nastaleeq"`** (this matters for grouping).

---

## 2. The apply pipeline (where the bug lives)

`applyProfileToQaseeda(name)` in `src/taskpane/taskpane.js:953`. Two passes inside separate `Word.run`s:

### Pass 1 — SIZE (rebuild for width)
- `gatherQaseedaBlocks` (`:711`) → the tagged content controls.
- `captureQaseedaTables(context, blocks, profile)` (`:737`) → per-cell geometry, fonts, alignment, indent, `natPx` (natural width), and the cross-block harmony matrix `qMatrix`.
- **`captureQaseedaCellRuns(context, cap)` (`:907`)** — NEW. Snapshots each content cell's **per-word runs** (`{text,name,size,bold,italic,color}` coalesced via `AshaarWord.coalesceRuns`) + align + indent, keyed `"block:table:cell"`, into `origContent`. **Captured BEFORE the rebuild**, because the rebuild destroys per-word fonts.
- Compute one shared `targetTwips` (via `AshaarMatrix.uniformSlotPx`), a width `sizeSig`, and `needRebuild = _appliedSizeSig[name] !== sizeSig` (`:1014`). If a real width change, **rebuild each block's OOXML** from font-less source text (`renderForWordOoxml`) — this is the only way to resize span tables (see memory `width-engine-rebuild-not-setwidth`). The rebuild pins ONE representative font ⇒ it flattens fonts; pass 2 restores them from `origContent`.

### Pass 2 — JUSTIFY (fill each cell)
- Re-gather + re-capture the (now bare) tables.
- Force-load every `origContent` run font **+ Kasheeda face** so `measureText` uses real metrics.
- For each **content** cell → `buildContentCellOoxml(c, info, key, colPx)` (`:1096`) returns an OOXML `<w:p>`; collected in `cellPlans`; written via `cell.body.clear()` + `insertOoxml(wrapOoxml(...), replace)`, **one sync per cell**.
- Spacing cells are decorated inline (unchanged).
- Optional debug tint colours injected tatweels (font colour) and micro-spaces (highlight → the **teal** in the screenshots).

### `buildContentCellOoxml` — the heart (rewritten this session, `:1096`)
1. Rebuild `origRuns` from `origContent[key]` (fallback to a single run of `c.base` if the entry is missing or its joined text ≠ `c.base` after stripping).
2. **Regroup runs into family segments** (`baseFaceOf(name) = descriptorForFontName(name).wordName || name`): consecutive same-family+size+color runs merge. Jameel base+Kasheeda → one segment; Amiri and Fatemi stay separate.
3. `flattenSegs(segOut)` (`:1140`) re-inserts a regular space run **between** segments (fix for the inter-run-gap the old code dropped, which made cross-font words touch).
4. No-fill (`justify none/css`) → emit segments unchanged (fonts+spaces preserved).
5. Compute `cNatural` (sum of `measIn(seg.text, baseFaceOf, size)` + inter-segment spaces) and `cTarget` (cell-fit → `cellFitBudget`; natural-fit → `naturalFitTarget(qMatrix[matKey], reach, φ)`), clamped ≤ colPx.
6. Elongate **each segment by its own mechanism** toward a proportional share of `cTarget`:
   - `font-swap` → `splitSpans` + `selectSwapRuns` → per-fasl base/Kasheeda runs.
   - `tatweel` → `mehrElongate` + `selectSwapRuns` → single-font text.
   - `whitespace` or `!doKashida` → pass-through (no elongation).
   - `generic` → deferred; then all generic segments jointly absorb the remaining gap via `justifyRunsConcentrated`.
7. `flattenSegs`, then close residual with capped hair-spaces (`AshaarResidual.capMicroSpaces` + `distributeMicroSpaces`), emit via `AshaarWord.misraRunsXml`.

`misraRunsXml(runs, align, sizePtFallback, opts)` (`word-html.js:1311`) emits per-run `<w:rFonts w:ascii=.. w:hAnsi=.. w:cs=..>` (all three — see §3), optional per-run `<w:color>`, real `<w:jc>` (right/center/left, **not** distribute), and optional `<w:ind>`.

---

## 3. What was already tried this session (and why it wasn't enough)

Committed as **`fed5611`** ("run-aware, font-preserving, idempotent qaseeda justify"), plus **uncommitted** working-tree changes (the family-segment + base-face work). In order:

1. **Ported Jameel/Mehr mechanisms into apply** (were generic-or-spacing only).
2. **Per-word font preservation across the rebuild** — capture `origContent` before rebuild, re-emit run-aware OOXML.
3. **`misraRunsXml` sets `w:ascii`+`w:hAnsi`+`w:cs`** (not cs alone). Reason: Office.js `Font.name` reads back the **ascii** font; cs-only meant a re-apply read the document-default font and collapsed mixed cells. This is a real round-trip requirement — keep it.
4. **Detect whole-cell Jameel/Mehr by mechanism+family**, not exact face (a re-applied Jameel cell holds mixed base+Kasheeda faces).
5. **Inter-segment space restored** (`flattenSegs`) — cross-font words no longer touch.
6. **Natural width measured in BASE face** in both `buildContentCellOoxml` and the `qMatrix` build in `captureQaseedaTables` (`:881`). Intent: a re-applied Jameel cell reads back in the wider Kasheeda face; measuring in the base face was supposed to stop the target from inflating each pass.

**Despite #6, the screenshots still show progressive growth on a mixed block with the latest code.** So either base-face normalization isn't reaching the measurement that actually drives the target, or the growth has a different source. **This is the crux to crack.**

---

## 4. Leading hypotheses (verify, don't assume)

**H1 — `qMatrix` / target still inflates each pass.** `qMatrix[matKey]` (natural-fit harmony width) is built from `c.natPx`, measured in `captureQaseedaTables`. Check: is `c.fontName` on a re-applied cell the Kasheeda face (wider)? Does `baseFaceOf` actually normalize it there? Is `c.measure` truly clean (stripJustification) on re-capture? Add a per-cell diagnostic dumping `{fontName, measuredNatPx, qMatrix[matKey], cNatural, cTarget, colPx}` on pass 1 and pass 2 of two consecutive applies and compare. If any grows, that's the leak.

**H2 — Mixed-font `cNatural` under/over-count.** `cNatural` sums per-segment base-face widths + inter-segment spaces, but `justifyRunsConcentrated` for generic segments is told `genTarget = cTarget - nonGenAchieved`. If `nonGenAchieved` double-counts or omits the inter-segment spaces vs. what's actually emitted, the residual hair-space count (`cTarget - achievedTot`) can be positive every pass and pile up — but note hair-spaces (U+200A) ARE stripped on re-capture (`stripJustification`, `:1808`), so they shouldn't accumulate unless a **non-stripped** space (regular U+0020) is being injected mid-line.

**H3 — Regular space injected where it isn't stripped.** `stripJustification` only removes U+0640/U+200A/U+2009 — **not** regular spaces. If `flattenSegs` ever inserts an inter-segment regular space **mid-word** (i.e., a segment boundary that isn't a real word gap), that space survives re-capture, splits the word into two on the next `getTextRanges([" "],true)`, creates two segments, and adds another space next pass → growth + shattering. **When can a segment boundary fall mid-word?** Segments split on family/size/**color**. If `coalesceRuns` (now splits on `color`) or the per-word colour read produces a colour change **within** a word (e.g. a partly-red refrain, or debug tint colour from a prior pass being read back as the run colour!), a word can split mid-word. ⚠️ **Debug tint interaction:** the debug pass sets `font.color`/`highlightColor` on injected artifacts; on the next capture those colours could come back as run colours and fragment runs. Strongly suspect this — reproduce with debug colours OFF.

**H4 — Shattering = generic tatweels on a Nastaliq segment.** In a mixed cell, a Nastaliq run whose font is NOT in the registry (or whose `font.name` reads back empty/substituted) resolves to `generic` → `justifyRunsConcentrated` injects U+0640 between joined forms → isolated letters. Confirm by logging each segment's `{name, family, mechanism}` in a mixed cell. If a Mehr/Jameel word reads back with an ambiguous `font.name` (Office.js exposes **no** complex-script font name — only ascii), it will misclassify. This is the known residual edge noted in memory.

**H5 — Kasheeda self-inflation is measured in `misraRunsXml` output, not input.** `achievedTot` in `buildContentCellOoxml` measures the *emitted* runs in their emitted faces (Kasheeda = wider). That's correct for residual, but make sure the *target* (`cTarget`) never derives from an emitted/Kasheeda width.

---

## 5. Suggested attack plan

1. **Reproduce deterministically with a diagnostic build.** The `debugMode` checkbox already dumps a table in `justifySelection`; add an equivalent per-cell readout to `applyProfileToQaseeda` pass 2: for each content cell log `key, segCount, per-seg {family, mech, text}, cNatural, cTarget, colPx, achievedTot, nSp (hair-space count)`. Run apply twice, diff the two dumps. **Whatever grows is the bug.** (Office.js glue isn't node-testable — this on-screen dump is the evidence mechanism, per `systematic-debugging`.)
2. **Toggle debug colours OFF** and re-test — if growth/shatter largely stops, H3 (colour-fragmentation feedback) is confirmed; fix by never letting debug-tint colours re-enter run grouping (e.g. ignore colours that equal the debug tint, or strip artifact formatting before capture).
3. **Confirm the mechanism per segment** in a mixed cell (H4). If a Nastaliq is landing in `generic`, fix classification (the read-back `font.name` problem) — perhaps persist per-run font in the content-control tag instead of re-reading it from the document each pass (that would also make idempotency structural rather than round-trip-dependent).
4. **Consider persisting `origContent` in the block's content-control tag** (`AshaarWord.setTag`/`parseContentControlTag`) at first apply, and re-using it on re-apply instead of re-reading fonts from justified cells. This sidesteps the entire cs-vs-ascii / Kasheeda-face / colour-read-back round-trip class of bugs — the source-of-truth for per-word fonts becomes the tag, not the rendered cell. **This may be the clean structural fix** the round-trip patches have been approximating.
5. Only after root cause is pinned: make the change, run `npm test`, and **verify in Word** with the exact repro (apply twice to a mixed Mehr+Amiri and a Jameel+naskh block; confirm pass 2 == pass 1, fonts kept, spaces between fonts present, no shatter).

---

## 6. Files & anchors

| File | What |
|---|---|
| `src/taskpane/taskpane.js:953` | `applyProfileToQaseeda` (pass 1 SIZE `:963`, rebuild gate `:1014`, pass 2 JUSTIFY, `_appliedSizeSig` `:727`/`:1324`) |
| `…:737` | `captureQaseedaTables` (natPx/qMatrix build `:881`, base-face normalize just added) |
| `…:907` | `captureQaseedaCellRuns` (per-word run snapshot → `origContent`) |
| `…:1096` | `buildContentCellOoxml` (segment model; `baseFaceOf` `:1123`, `flattenSegs` `:1140`) |
| `…:1808` | `stripJustification` (removes U+0640/U+200A/U+2009 only — NOT regular space) |
| `…:2024` | `justifySelectionInner` — the **in-place** justify path (works well for two naskhs; reference implementation for per-run behavior; NOT the buggy path but the model to match) |
| `src/taskpane/word-html.js:1311` | `misraRunsXml` (ascii+hAnsi+cs, color, jc, indent) |
| `…:994` | `coalesceRuns` (now splits on `color` too) |
| `…:1100` | `mehrElongate` · `…:1017` `distributeMicroSpaces` · `…:1647` `renderForWordOoxml` (rebuild) |
| `src/taskpane/kashida-fontswap.js` | `splitSpans`, `selectSwapRuns` (Jameel) |
| `src/taskpane/kashida-residual.js` | `capMicroSpaces`, `injectSpaceRuns` |
| `src/taskpane/fonts.js:59` | `descriptorForFontName` (mechanism/family resolution) |
| `src/taskpane/natural-width-matrix.js` | `uniformSlotPx`, `cellFitBudget`, `naturalFitTarget`, `buildMatrix` |
| `src/taskpane/profiles.js` | profile model, `normalizeFillMode`, strength math |

## 7. Constraints & gotchas (learned the hard way)
- **Never edit `src/vendor/`.** Pure modules stay DOM/Office-free + node-tested.
- **Resize span tables ONLY by OOXML rebuild** — no `columns.setWidth`, no `cell.columnWidth` (memory `width-engine-rebuild-not-setwidth`).
- **Office.js exposes no complex-script (cs) font name** — `Font.name` is the ascii font. This is the root of the round-trip fragility; re-reading fonts from justified cells is inherently lossy. (Hence §5.4 tag-persistence idea.)
- **`adoptTableToSource` strips tatweels/micro-spaces**, so the rebuild width-signature is stable (not a growth source).
- **Office.js glue is not node-testable.** Pure helpers (`misraRunsXml`, `coalesceRuns`, `splitSpans`, `selectSwapRuns`, `mehrElongate`, `capMicroSpaces`) ARE tested in `tests/` — keep them so. Verify glue in Word.
- Git: commit with `--no-gpg-sign`. HEAD = `fed5611`; the family-segment/base-face work is **uncommitted** in `taskpane.js`. `npm test` is green.

## 8. Continuity memory (auto-loaded)
`~/.claude/projects/-Users-abdealikhurrum-ashaar-js-Office/memory/`: `justification-modes-state.md` (this whole saga, most detail), `width-engine-rebuild-not-setwidth.md`, `run-aware-generic-kashida-regression.md`, `font-measurement-model.md`, `office-js-word-constraints.md`.

---

**Start here:** `superpowers:systematic-debugging` → build the per-cell diagnostic dump (§5.1), apply twice, diff. Find what grows before touching code. Prime suspects: debug-colour fragmentation feeding regular spaces mid-word (H3), and re-read font misclassification (H4) — both argue for making per-word fonts authoritative in the content-control tag rather than re-derived from the rendered cell each pass.
