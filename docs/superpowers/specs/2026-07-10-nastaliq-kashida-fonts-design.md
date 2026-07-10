# Nastaliq Kashida Fonts — Design

**Date:** 2026-07-10
**Status:** Approved, ready for implementation plans
**Branch:** `worktree-nastaliq-kashida-fonts` (worktree, based on `feat/guided-justification-ux`)
**Scope:** Add first-class support for the Nastaliq fonts that can actually fill a line — **Mehr Nastaliq** (tatweel), **Jameel Noori Kasheeda** (italic-run elongation), and **Gulzar** (whitespace-only) — and make the justification engine choose the *right mechanism per font* instead of assuming every "Nastaliq" font kashidas. Introduces a font registry that drives preview, Word run naming, and — critically — which justification modes are valid for a given font.

> **Update (Gate G2, 2026-07-10):** Jameel is **font-swap**, not italic-run — see the Gate G/G2 results below.

## Problem

The task pane has a single `nastaliq` font mode that maps to **Noto Nastaliq Urdu** with `Jameel Noori Nastaleeq` as a CSS fallback string (`word-html.js:142`, `taskpane.js:256`), and emits `<w:rFonts w:cs="Noto Nastaliq Urdu"/>` into Word (`word-html.js:1093,1230,1271`, `word-tabstop.js:68`). But **Noto Nastaliq Urdu has no kashida** — it does not elongate on U+0640 tatweel. So the canvas tatweel engine (`AshaarJustify.justifyLine`) and the probe scorer (`AshaarTune.probeFont`) have nothing to bite on: the "Nastaliq" option can never truly fill a line, and the guided-justification recourse "raise stretch strength" is meaningless for it.

The three fonts that *can* fill a Nastaliq line each do it by a **different mechanism**, and one of them (Gulzar) actively breaks if tatweels are injected. A single font mode with a single justification path cannot serve them.

## Font-realm findings (verified against sources, see research trail)

| Font | Kashida mechanism | Fills a line by | License / shipping |
|---|---|---|---|
| **Mehr Nastaliq Web** | `tatweel` (U+0640), **limited** coverage, **whitelist slot rules** (Beta 2.0) | Inserting tatweels only on joins into whitelisted letters; canvas-measured | CC-BY-SA (npm pkg MIT) → **bundle** w/ attribution |
| **Jameel Noori Kasheeda** | `italic-run` — elongated forms live in the **italic style slot**; whole-run toggle, **discrete** | Italicizing chosen spans (each jumps to its wider kasheeda form) | murky/pirated → **bundle privately** (owner's decision), do not ship on public Pages without resolving |
| **Gulzar** | `whitespace` — **no elongation**; *"dramatic failures when set to justify"* with tatweels | Inter-word space only (Word `distribute`) | OFL → **bundle** |
| *Noto Nastaliq Urdu (existing)* | `whitespace` — no kashida | Inter-word space only | already referenced; **reclassify** off the tatweel path |

**Consequence:** justification is not one algorithm gated by a boolean; it is three strategies gated by a per-font tag.

**Mehr's tatweel whitelist (Beta 2.0, verified — [mehrtype how-to](https://mehrtype.blogspot.com/2022/02/how-to-use-mehr-nastaliq-web-font.html)).** Elongation is designed only on the join *leading into* these letters; every other join renders no kashida:
- **Medial/initial:** ب پ ت ٹ ث س ش ف ک گ
- **Word-final:** ب پ ت ٹ ث ف ک گ (س ش are not elongated word-final)

The blog states this is a **beta** limitation that later versions may lift. So the whitelist is authoritative *for this version* but must be versioned and probe-expandable — see §5 dispatch.

**Gate G result: FAIL (2026-07-10)** — Word does not honor Jameel *italic*→kasheeda; the `italic-run` trigger is dead. Microsoft Word merely slants Jameel Noori Kasheeda's italic runs — it does not swap glyphs the way InPage does.

**Gate G2 result: PASS (2026-07-10) — `font-swap` mechanism works.** The kasheeda forms are a **named style ("Kasheeda") inside the "Jameel Noori Nastaleeq" family** (Regular / Kasheeda), selected by FONT, not by the italic property. Applying the Kasheeda face to a run widens **only** the segments that have a designed kasheeda variant — verified: in `كہہ رہے تھے اشقياء` only `كہہ` elongated (the initial ک stretches). Jameel is therefore reclassified **`mechanism:"font-swap"`** (NOT whitespace, NOT the dead italic-run):
- Base face cs name: `Jameel Noori Nastaleeq`; wider face cs name: `Jameel Noori Nastaleeq Kasheeda`.
- The Task-4/5 **fasl subset-selection carries over unchanged** — only the measurement (base-font width vs Kasheeda-font width per fasl) and the emission (per-run `<w:rFonts w:cs>` swap instead of `<w:i/>`) differ. Fasls with no kasheeda variant measure equal width → zero gain → never swapped, so the "not every combo has a kasheeda form" reality is handled by the analyzer, not special-cased.
- Both faces must be installed on the reader's machine; both `.ttf`s loaded into the WebView (private `@font-face`, gitignored) so the canvas analyzer can measure them.
- **Open on-device risk (Task-8 manual check):** confirm `w:cs="Jameel Noori Nastaleeq Kasheeda"` resolves to the wider face in the generated `.docx` (Word style-name-within-family resolution can be finicky); if it does not, inspect the font's name table for the correct full/family string to emit.

The mechanism enum is now **`tatweel | font-swap | whitespace`** (`italic-run` retired).

## Findings that reshape scope (verified in code)

- **The `nastaliq` mode is hard-coded to a non-kashida font** in five places: `word-html.js:142` (CSS stack), `taskpane.js:256` (`previewFontFamily`), `word-tabstop.js:68` and `word-html.js:1093,1230,1271` (`<w:rFonts w:cs="Noto Nastaliq Urdu"/>`), and the canvas fallback name `"Noto Nastaliq Urdu"` at `taskpane.js:864,1001,1069`. Adding fonts by copy-pasting these literals does not scale and can silently desync preview from Word output.
- **The mode chooser is a flat `<select>`** (`taskpane.html:46` `<option value="nastaliq">`). It has no notion that a font may forbid a justification mode.
- **The justify pipeline assumes tatweel.** `justifySelection` runs `probeFont → calibrate → justifyLine → insertText` (per guided-justification spec `taskpane.js:1229,1248`, calibrate/overwrite at `1234–1251`) with a single-run `cell.body.insertText()`. There is no code path that emits **multiple runs with per-run italic**, which the Jameel mechanism requires.
- **A bundled-font precedent already exists.** `FatemiMaqala` is an `@font-face` in `taskpane.css:4`; `AL-KANZ` at `:19`. The Custom Font Uploader (`font-store.js`/`AshaarFontStore`, spec 2026-07-09) established that a font loaded from **bytes** via `FontFace` is measurable in the sandboxed WebView — the exact property our bundled fonts need for canvas measurement.
- **Guided-justification is building a native Word-kashida path.** §3 of `2026-07-10-guided-justification-ux-design.md` emits `lowKashida/mediumKashida/highKashida` (Arabic) and `distribute` (non-Arabic) with an auto trailing `<w:br/>`. That path is font-agnostic (uses the reader's installed font) and becomes the natural fallback for Mehr and the whitespace mechanism for Gulzar/Noto.

## Key insight

Make the **font** — not the mode — own the mechanism. A small font registry tags each font with `mechanism ∈ {tatweel, italic-run, whitespace}`. The registry then (a) supplies preview CSS + Word run name, (b) declares whether the font is bundled, (c) **filters the guided-justification mode chooser** so a font is never offered a mode it can't honor, and (d) selects the justify strategy at run time. Adding a future Nastaliq font becomes one registry entry, not five literal edits.

---

## Components

### 1. Font registry — `src/taskpane/fonts.js` (`AshaarFonts`)

New UMD module (repo pattern: Node for tests, browser for the pane). Single source of truth; replaces the five scattered literals.

```js
AshaarFonts.LIST = {
  mehr:   { id:"mehr",   label:"Mehr Nastaliq",
            css:"'Mehr Nastaliq Web', serif", wordName:"Mehr Nastaliq Web",
            mechanism:"tatweel",   bundled:true,  file:"MehrNastaliqWeb.woff2",
            // Beta 2.0 whitelist: tatweel is legal ONLY on a join leading INTO these letters.
            tatweelRules:{
              version:"beta-2.0",
              medialInto: ["ب","پ","ت","ٹ","ث","س","ش","ف","ک","گ"],
              finalInto:  ["ب","پ","ت","ٹ","ث","ف","ک","گ"],   // س ش drop out word-final
              probeMayExpand:true } },
  jameel: { id:"jameel", label:"Jameel Noori Kasheeda",
            css:"'Jameel Noori Nastaleeq Kasheeda','Jameel Noori Nastaleeq',serif",
            wordName:"Jameel Noori Nastaleeq",           // italic run triggers kasheeda
            mechanism:"italic-run", bundled:true, private:true,
            file:"JameelNooriNastaleeqKasheeda.ttf", italicFile:"…-italic.ttf" },
  gulzar: { id:"gulzar", label:"Gulzar",
            css:"'Gulzar', serif", wordName:"Gulzar",
            mechanism:"whitespace", bundled:true, file:"Gulzar-Regular.woff2" },
  noto:   { id:"noto",   label:"Noto Nastaliq Urdu",
            css:"'Noto Nastaliq Urdu', serif", wordName:"Noto Nastaliq Urdu",
            mechanism:"whitespace", bundled:false },
  // "document" / "arabic-serif" existing modes remain, mechanism:"whitespace"
};
```

**Pure helpers (Node-testable):** `get(id)`, `mechanismOf(id)`, `wordNameOf(id)`, `cssFamilyOf(id)`, `modesFor(id)` (→ which §5 modes are valid — see §3), `allBundled()`.

**Callers refactored to read the registry** (removing the literals): `previewFontFamily` (`taskpane.js:256`), `fontFamilyStyle` (`word-html.js:140`), the three `<w:rFonts>` sites, the canvas-fallback name sites, and the `<select>` population (§ below).

### 2. Bundling & preview loading

- Vendor the font files into `assets/fonts/`: **Mehr** (from the `mehr` npm package, CC-BY-SA — add attribution to `LICENSE`/README), **Gulzar** (Google Fonts, OFL — add OFL notice), **Jameel Kasheeda** `.ttf` (owner-supplied; `private:true` so it is git-ignored / excluded from the public Pages deploy — see Distribution).
- Add `@font-face` blocks in `taskpane.css` mirroring the FatemiMaqala precedent. **Jameel needs two faces under one family** — a `font-style:normal` and a `font-style:italic` src — so both the canvas (`ctx.font = "italic …"`) and Word (`<w:i/>`) resolve the kasheeda glyphs. If Jameel ships as a single TTF whose italic bit swaps glyphs, the italic `@font-face` points at the same file with `font-style:italic`; if elongation lives in a separate italic file, point at that. **This is verified by the §Spike.**
- Bundled fonts are registered for measurement exactly like the uploader path: either the plain `@font-face` (sufficient for CSS + canvas) or, if we want a single code path, `AshaarFontStore.loadFont(family, bytes)` at Office-ready. Prefer plain `@font-face` for the three bundled families (simplest); reuse `AshaarFontStore` only if a face needs byte-loading.

### 3. Mechanism drives the guided-justification mode chooser (§5 interlock)

`AshaarFonts.modesFor(id)` returns the justification modes valid for the selected font; the guided-justification §5 chooser renders only those, and defaults to the recommended one:

| Font mechanism | Engine (tatweel) | Space out words | Let Word fill it | Leave as typed | Default |
|---|---|---|---|---|---|
| `tatweel` (Mehr) | ✅ **recommended** | ✅ | ✅ (`highKashida` fallback) | ✅ | Engine |
| `italic-run` (Jameel) | ❌ (no tatweel glyphs) | ✅ | ⚠️ only if reader has Kasheeda build | ✅ | **Italic-run** (new, §4) |
| `whitespace` (Gulzar, Noto) | ❌ **disabled** (breaks shaping) | ✅ **recommended** | ✅ (`distribute` only) | ✅ | Space out words |

- Selecting a `whitespace` font **disables** the engine/expressive slider and shows a one-line why: *"Gulzar has no stretch letters — fill by spacing instead."*
- Selecting Jameel adds an **"Elongate (Kasheeda)"** mode backed by §4.
- This is additive to the guided-justification §5 spec: that spec defines the four modes; this spec adds a `modesFor(font)` filter in front of them.

### 4. New module `src/taskpane/kashida-italic.js` (`AshaarKashidaItalic`) — Jameel

The italic-run strategy. Jameel elongation is a **discrete subset-selection**: choose which spans to italicize so the line's width approaches the target.

**Pure (Node-testable):**
- `selectItalicRuns(spans, widthsNormal, widthsItalic, targetPx) → { runs:[{text,italic}], fill, reason }`
  Dynamic-programming / greedy 0-1 selection over spans: maximize filled width ≤ target (with the last legal step allowed to cross, mirroring how the tatweel binary-search picks the max that fits). Returns the ordered run list to emit, the achieved fill ratio, and a `reason` when it underfills (e.g. `"no elongatable spans"`, `"discrete steps overshoot"`), feeding the guided-justification result panel (§1 there).
- `splitSpans(text) → spans[]` (pure). **Spans are connected segments (fasl / piece-of-Arabic-word), NOT words and NOT characters** — confirmed against real Jameel Kasheeda use: the italic→kasheeda swap operates on a whole connected segment; you cannot italicize a single character mid-cluster. Split the text at each **non-joining letter boundary** (after ا أ إ آ د ذ ر ز ژ ڑ و ؤ ے and after whitespace), so each span is one joined cluster. Adjacent spans that carry no elongatable join collapse into their neighbor's run to avoid needless run fragmentation. This is the finest legal granularity, giving the subset-selector many small levers for smooth fill while respecting the font's selection unit.

**Browser-only:**
- `measureSpans(text, ctx) → { spans, widthsNormal, widthsItalic }` — sets `ctx.font` normal vs `italic <family>` and measures each span. Reuses the same WebView `@font-face`-loaded ctx the tatweel path depends on (same concern flagged at `taskpane.js:1166`).

**Emission (add-in layer):** replace the single `cell.body.insertText()` with `cell.body.insertOoxml(runsToOoxml(runs))`, where italic runs carry `<w:i/>` in `<w:rPr>` alongside the Jameel `<w:rFonts>`. The OOXML run builder in `word-html.js` already emits `<w:rPr>` runs; extend it to accept a per-run italic flag.

### 5. Justify dispatch

In `justifySelection` (and the qaseeda/apply-to-all path), branch on `AshaarFonts.mechanismOf(fontId)`:
- **`tatweel`** → existing `probeFont → calibrate → justifyLine → insertText`. Slots are the **intersection** of `buildSlots()`'s generic legal joins with the font's `tatweelRules` whitelist: for Mehr, keep a candidate slot only if the letter *following* the tatweel is in `medialInto` (mid-word) or `finalInto` (word-final). `probeFont` runs as a cross-check and, when `probeMayExpand`, may *add* slots it empirically measures as elongating (so a post-beta Mehr with fuller coverage unlocks automatically without a code change). Primary path for Mehr.
- **`italic-run`** → §4: `measureSpans → selectItalicRuns → insertOoxml`. No `probeFont` (tatweel probe is meaningless here).
- **`whitespace`** → **no glyph manipulation**; delegate to the guided-justification §3 Word path with `distribute` (or `w:jc="both"`), set on the misra paragraphs. Hard guard: assert we never call `justifyLine` for a `whitespace` font (prevents Gulzar shaping failures).

### 6. Distribution / reader-end (per decision: document the requirement)

- Runs carry `wordName` from the registry so Word uses the right face; Jameel italic runs additionally carry `<w:i/>`.
- The add-in **cannot install fonts on the reader's machine nor embed them into the `.docx` via Office.js**. The pane shows a per-font note: *"Readers need **{font}** installed to see this correctly."* For Jameel: *"…and specifically the **Kasheeda** build, or italic runs won't elongate."*
- **Public-deploy guard for Jameel:** `private:true` fonts are excluded from the GitHub Pages deploy (git-ignored under `assets/fonts/` or filtered by the deploy step). Only Mehr + Gulzar (freely licensed) ship publicly; Jameel is bundled for the owner's local/private install.

---

## Architecture summary

```
src/taskpane/fonts.js          NEW  AshaarFonts registry (pure + browser); single source of truth
src/taskpane/kashida-italic.js NEW  AshaarKashidaItalic: splitSpans, selectItalicRuns (pure) + measureSpans (browser)
src/taskpane/taskpane.js       EDIT read registry (preview, canvas fallback, mode chooser filter); dispatch by mechanism
src/taskpane/word-html.js      EDIT fontFamilyStyle + <w:rFonts> from registry; run builder accepts per-run italic
src/taskpane/word-tabstop.js   EDIT <w:rFonts> from registry
src/taskpane/taskpane.html     EDIT populate font <select> from registry; add fonts.js + kashida-italic.js to loader
src/taskpane/taskpane.css      EDIT @font-face for Mehr, Gulzar, Jameel (normal + italic)
assets/fonts/                  ADD  Mehr (CC-BY-SA), Gulzar (OFL), Jameel Kasheeda (private)
```

No edits to `src/vendor/*` (vendoring preserved). The italic-run strategy lives in the add-in layer because it depends on Word run formatting and a font-packaging quirk, not on a browser-generic engine primitive.

## Testing

- **Unit (Node, `tests/fonts.test.js`):** `AshaarFonts.mechanismOf/modesFor/wordNameOf` for each font; assert Gulzar/Noto exclude the engine mode; assert `wordName` matches the `<w:rFonts>` a caller emits.
- **Unit (Node, `tests/kashida-italic.test.js`):** `splitSpans` tokenization; `selectItalicRuns` on synthetic width arrays — picks the max-fill subset ≤ target, reports `reason` when no span is elongatable, handles the all-fit and none-fit edges. Pure, deterministic, no canvas.
- **Manual (in Word, checklist):**
  1. **Spike first (blocking):** does setting a run italic in Jameel Kasheeda actually swap to elongated glyphs in Word (Mac + Windows)? Sample doc with normal vs italic Jameel runs.
  2. **Mehr whitelist verification (before shipping):** render tatweel between each documented pair (into ب پ ت ٹ ث س ش ف ک گ medial; ب پ ت ٹ ث ف ک گ final) in the *bundled* Mehr and confirm each actually elongates; confirm a non-whitelisted join (e.g. into ر/د/ه) does **not**. Reconcile any drift between the blog and the shipped `.woff2`, and check whether the bundled version is still Beta 2.0 or a later build (which may widen coverage). Then: engine justify fills lines; probe agrees with the whitelist; ragged→clean.
  3. Gulzar: engine mode is disabled in the chooser; "Space out words" fills; injecting tatweels is never attempted.
  4. Reader-end: open the `.docx` on a machine without the font → graceful fallback + the pane note was accurate.

## Suggested build order (each its own implementation plan)

1. **§Spike — Jameel Word-italic swap (blocking gate).** Verify in Word before building §4. If Word does **not** honor the italic→kasheeda swap, `italic-run` is dead: Jameel degrades to a `whitespace`-mechanism selectable font and §4 is cut. Do this first.
2. **§1 Font registry + refactor the five literals** — pure, testable, no behavior change; unblocks everything.
3. **§2 Bundling + `@font-face` + preview** — Mehr & Gulzar render in preview and Word; Jameel gated on the spike.
4. **§5 dispatch + §3 mode-chooser interlock** — mechanism selects strategy; whitespace guard; Gulzar disables engine. Mehr's tatweel path works end-to-end here (reuses existing engine).
5. **§4 `kashida-italic.js` + insertOoxml italic runs** — only if the spike passed; the one genuinely new algorithm.
6. **§6 reader-end notes + public-deploy guard for Jameel.**

## Out of scope

- Changing the vendored `ashaar-js` engine (tatweel algorithm, probe scorer) — reused as-is.
- The guided-justification components themselves (§1–§7 of that spec); this only adds a `modesFor(font)` filter in front of §5 and reuses §3's Word path for `whitespace`.
- Per-run measurement of *mixed* fonts within one misra (that is the run-aware project); Jameel italic runs are single-font, italic-toggled only.
- Additional Nastaliq fonts (Awami/Nafees/Alvi/Fajer) — the registry makes them one-entry additions later; none add a new mechanism (Awami = whitespace via Graphite, unavailable in Word anyway).
- Embedding fonts into the `.docx` (Office.js can't; would need a separate post-process tool — explicitly deferred per the reader-end decision).

## Decisions locked during brainstorming

- Scope: **full kashida wiring, both mechanisms** (+ Gulzar whitespace).
- Architecture: **A1 — font-descriptor registry + add-in-layer strategy** (not upstream into vendored ashaar-js).
- Jameel file: **bundle privately** (with a public-deploy guard).
- Reader end: **document the requirement** (no embed/post-process this round).
- Mehr primary path: **canvas-tatweel engine**; Word-native kashida (§3) is its no-font-loaded fallback.

## Open questions for spec review

- ~~**Mehr `tatweelRules`:** do we need an explicit slot-override table, or is `probeFont` sufficient?~~ **Resolved:** explicit whitelist (Beta 2.0 letters above), authoritative for this version; probe cross-checks and may expand post-beta. A **verification task** confirms the whitelist against the actual bundled `.woff2` in Word before shipping (fonts and blog can drift).
- ~~**Jameel span granularity:** word vs sub-word?~~ **Resolved (owner has used Jameel):** sub-word IS supported, at the **connected-segment (fasl/PAW)** unit — split at non-joining letters, never mid-cluster. `splitSpans` implements this. (Word-level remains a trivial fallback only if the spike shows per-segment italic runs shape badly at boundaries.)
- Whether the **Noto** default should be retired entirely in favor of Gulzar as the default `whitespace` Nastaliq, or kept for continuity.
