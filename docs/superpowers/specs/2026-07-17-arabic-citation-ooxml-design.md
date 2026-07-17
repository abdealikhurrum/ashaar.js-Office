# Arabic citation OOXML rendering — font + italic (design)

**Date:** 2026-07-17
**Status:** Approved design, pre-implementation
**Track:** Follow-on to SP-A (citation locators + editing + tagging). Fixes two Arabic-rendering
defects found in live Word testing.

## Problem (confirmed in live Word + via the engine)

Citations insert via `insertHtml` + `alignment = right`. Two defects for Arabic:

1. **Wrong font.** The inserted citation carries no font, so it inherits the footnote's default
   (Times New Roman). `insertHtml` cannot set the **complex-script (Arabic) font** — CSS
   `font-family` only targets the Latin face; the Arabic face is the OOXML `w:cs` attribute.
2. **Squares on Arabic italic.** citeproc italicizes titles — for an Arabic book it emits
   `القاضي النعمان, <i>كتاب الهمة</i> (دار المعارف, 1996).`. Times New Roman *Italic* has no
   Arabic glyphs, so the italicized Arabic title renders as tofu squares. (Arabic typography does
   not use italics; proper Arabic citation styles drop it.)

The earlier `dir="rtl"` block wrap (committed) already fixed the run-level bidi (punctuation
ordering), and the Styles-tab "RTL document setup" fixed the footnote-number side. This spec
addresses only the font + italic.

## Decision (settled in brainstorming)

- **Render Arabic citations via `Range.insertOoxml()`** (the codebase's existing RTL-paragraph
  mechanism, as in `word-html.js`/`word-tabstop.js`) instead of `insertHtml`, so we control the
  complex-script font, italic, and bidi per run.
- **Font:** reuse the **document's complex-script font** — read the `Ashaar Normal` paragraph
  style's `font.nameBidirectional` at insert time (that style is created by the Styles-tab RTL
  setup); fall back to the selection/target range's `font.nameBidirectional`, then to a default
  constant if neither is available.
- **Italic:** **suppress italic on Arabic runs** (`<w:rPr>` without `<w:i/>`/`<w:iCs/>`), keep it
  on Latin runs (a Latin-titled source cited in an Arabic doc still italicizes correctly).
- **Scope:** notes (footnote/endnote) and the bibliography use OOXML. **Inline** citations stay
  on `insertHtml` (they live in the surrounding paragraph's flow — a block OOXML paragraph would
  be wrong there); inline keeps today's behavior.
- **LTR (en-US) citations are unchanged** — they keep the existing `insertHtml` path. OOXML is the
  Arabic-only path.

## Scope

**In scope**
- `cite-word.js`: a pure `htmlToOoxmlRuns(html, {csFont})` converter — tokenizes the sanitized
  citeproc HTML (whitelist `i/b/em/strong/span/sup/sub/br`), tracks italic/bold/superscript state
  and Arabic-vs-Latin per character run, and emits OOXML `<w:r>` runs: Arabic runs get
  `<w:rtl/>` + `<w:rFonts w:cs="<csFont>"/>` and **no** italic; Latin runs keep `<w:i/>`/`<w:b/>`
  as tagged. Plus `buildRtlParagraphOoxml(runsXml)` → `<w:p><w:pPr><w:bidi/><w:jc w:val="right"/></w:pPr>…</w:p>`
  and a FlatOPC package wrapper for `insertOoxml` (reuse the `word-tabstop.js` wrapper pattern).
- `cite-pane.js`: for Arabic locale, `insertCitation` (footnote/endnote) and `insertBibliography`
  read the document CS font (Office.js) then insert via `range.insertOoxml(pkg, replace/after)`;
  wrap the range in the content control with the SP-A tag as today. Inline + all LTR paths keep
  `insertHtml`.
- node-`assert` tests for the pure converter.

**Explicitly deferred**
- Changing the citeproc styles themselves (the Fatemi CSL variants that would drop italic at the
  style level) — SP-4.
- Any non-citation Arabic font handling (that is the Styles feature's job).
- SP-B (persist/restore) and SP-C (refresh) — unchanged by this.

## Architecture & components

### `cite-word.js` — pure HTML→OOXML converter (Node-testable)
- `htmlToOoxmlRuns(html, opts)` → OOXML run XML string. `opts.csFont` = the complex-script font
  name. Algorithm: reuse the tokenizer shape from `wrapRtlRuns` (tags vs chars); maintain a
  format stack for `i/b/sup`; group consecutive chars of the same script (Arabic vs Latin/neutral)
  into runs; per run emit `<w:r><w:rPr>[<w:rtl/>][<w:rFonts w:cs="…"/> if Arabic][<w:b/>][<w:i/> if
  Latin][<w:vertAlign w:val="superscript"/>]</w:rPr><w:t xml:space="preserve">escaped</w:t></w:r>`.
  Arabic runs never get `<w:i/>`. XML-escape text (`& < >`). `<br>` → separate handling or a
  `<w:br/>`.
- `buildCitationOoxml({html, csFont})` → a FlatOPC package string wrapping one RTL paragraph of
  the converted runs, ready for `Range.insertOoxml(pkg, Word.InsertLocation.replace)`.
- Keep `sanitize`/`wrapRtlRuns`/`buildNotePayload`/tag helpers unchanged (still used by the LTR
  path + preview).

### `cite-pane.js` — Arabic insert via OOXML
- Read the document CS font inside `Word.run` before building: try the `Ashaar Normal` style
  (`ctx.document.getStyles()` / `getByNameOrNullObject("Ashaar Normal").font.load("nameBidirectional")`),
  fall back to the target range's `font.load("nameBidirectional")`, then a default constant
  (e.g. a known Arabic UI font). One extra `ctx.sync()`.
- `insertCitation` (Arabic, footnote/endnote): `note.body.getRange().insertOoxml(buildCitationOoxml({html, csFont}), replace)`, then tag the returned range's content control. Inline (Arabic or not) and all en-US paths keep `insertHtml`.
- `insertBibliography` (Arabic): `insertOoxml(...)` at `after`; tag as today. en-US keeps `insertHtml`.
- Preview (pane, browser) is unchanged — it stays HTML with `dir="rtl"`; OOXML is Word-only.

## Risks / feasibility (verify first, mirroring SP-2's translator-name check)

- **`insertOoxml` into a footnote/endnote body** must be confirmed live (SP-A uses `insertHtml`
  there; `insertOoxml` on a note-body range is the load-bearing assumption). First implementation
  task verifies this in Word before the converter work is relied on. If unsupported, fallback:
  insert OOXML at the document range and move — but the expectation is it works (Range.insertOoxml
  is a general Range method).
- **Reading the document CS font** — if `Ashaar Normal` is absent (user didn't run RTL setup),
  the range/`nameBidirectional` fallback + default constant keep it working (just not matching a
  document font that was never set).
- **FlatOPC package correctness** — reuse the vetted `word-tabstop.js` wrapper; malformed OOXML
  silently fails `insertOoxml`, so the converter output is validated in tests against the known
  paragraph/run schema order.

## Testing (node-`assert`, added to `npm test`)

- `cite-word.test.js` — `htmlToOoxmlRuns`:
  - Arabic run → contains `<w:rtl/>` + `<w:rFonts w:cs="<font>"/>` and NO `<w:i/>` even when the
    input wraps it in `<i>…</i>` (the squares fix).
  - Latin `<i>` run → keeps `<w:i/>`, no `w:cs`, no `<w:rtl/>`.
  - Mixed Arabic+Latin citation → distinct runs, each with the right props.
  - XML-escapes `&`/`<`/`>` in text; `<sup>` → `<w:vertAlign w:val="superscript"/>`.
  - `buildCitationOoxml` wraps a `<w:p><w:pPr><w:bidi/>…` paragraph and a FlatOPC envelope.
- Manual Word checklist: Arabic footnote/bibliography now render in the document's Arabic font,
  with the (previously square) title showing correctly and non-italic; en-US citations unchanged.

## Success criteria

Inserting an Arabic footnote/bibliography renders the text in the document's complex-script font
(not Times New Roman), with the title shown correctly and **not** italicized (no tofu squares),
while keeping the RTL paragraph + correct punctuation. English citations are byte-for-byte
unchanged (still `insertHtml`). `npm test` passes including the converter tests.
