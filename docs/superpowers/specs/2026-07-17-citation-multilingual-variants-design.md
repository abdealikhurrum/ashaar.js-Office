# SP-3 — Multilingual variant model (design)

**Date:** 2026-07-17
**Status:** Approved design, pre-implementation
**Sub-project:** SP-3 of the Annotation & Citation initiative (see SP-1 core design for program context)
**Branch:** `feat/citation`

## Summary

Give the citation subsystem the ability to render an item's **original-script (Arabic),
romanized (transliteration), or both** variants, sourced from the user's real reference library.
The citeproc CSL-M *rendering* side already exists (`cite-engine.js` consumes `langPrefs` +
citeproc's `multi` model, wired in SP-1). SP-3 adds the two pieces SP-1 explicitly deferred:

1. **Sourcing** — a pure parser that reads variant data out of the two conventions the user's
   library uses and normalizes it into the `multi` shape the engine already reads.
2. **Policy UI** — a pane selector (Original / Romanized / Both) that chooses which variant slot
   renders, persisted in the citation tag and re-applied on Refresh, exactly like style/locale.

Nothing is offloaded to Zotero's internal citation pipeline (see "Why we keep our own engine").

## Decisions fixed in brainstorming (2026-07-17)

- **Variant source = both conventions.** The parser normalizes whichever is present per item:
  - **Juris-M mlzsync** blob in the CSL-JSON `note` field (what the library has today), and
  - **Cite Non-English (CNE)** `cne-*` key-value lines in the Zotero `Extra` field (cleaner,
    actively maintained authoring UI; the user may migrate items to it gradually).
  - When **both** are present on one item, **CNE wins** (newer/maintained convention).
  - When **neither** is present, the item's real fields are used unchanged (current behavior).
- **Output form = pane-selectable**, per insertion: **Original (ar) / Romanized / Both**.
- **Granularity = single global policy** for v1 — one selector applies the same policy to all
  segments (names, titles, publisher, …). A per-segment (names vs titles) split is deferred (YAGNI).
- **Refresh uses the pane's current** variant policy (bulk re-format), consistent with how SP-C
  already re-reads style/locale from the pane rather than freezing them per citation.
- **Rendering stays in our own client-side engine** (see below).

### Why we keep our own engine (CNE does not offload rendering)

CNE (`github.com/boan-anbo/cite-non-english`) stores variants as `cne-*` lines in Extra and, at
citation time **inside Zotero**, patches Zotero's pipeline to rewrite the CSL-JSON sent to citeproc
(building the native `multi` model for names; injecting `cne-*` CSL variables for simple fields that
only CNE's own bundled styles reference). It renders through Zotero's **native** word-processor /
Quick-Copy integration and exposes **no external API / CAYW**; its **Better BibTeX integration is
not yet implemented**. Our add-in reaches Zotero *only* through Better BibTeX (CAYW + the "Better
CSL JSON" translator), so CNE cannot feed our pipeline. Even if it could, offloading rendering would
discard everything SP-1/2/A/B/C already built and Word-verified: RTL/Arabic **OOXML** insertion,
the Fatemi custom styles, in-place **Refresh**, content-control **tagging**, footnote/endnote
insertion. So CNE (and Juris-M) are treated purely as upstream *authoring/storage* tools; SP-3
consumes their stored data and our own engine renders.

## The data (real example)

The user's "Better CSL JSON" export carries the Arabic as the real fields and the Juris-M variants
in the `note` field as a length-prefixed mlzsync blob:

```json
{
  "id": "AldAEy…4", "type": "book", "language": "ar",
  "title": "‫عيون الأخبار ج/4‬",
  "author": [{ "literal": "الداعي الأجل سيدنا إدريس عماد الدينؓ" }],
  "note": "mlzsync1:0215{\"type\":\"book\",\"multifields\":{\"main\":{},\"_keys\":{\"title\":{\"en\":\"‫Uyun al-Akhbar Vol. 4‬\"}}},\"multicreators\":{\"0\":{\"_key\":{\"en\":{\"lastName\":\"al-Dai al-Ajal Syedna Idris Imaduddin RA\",\"firstName\":\"\"}},\"fieldMode\":1}}}"
}
```

Notes that shape the parser:
- `mlzsync1:` + a **4-digit zero-padded length** + JSON. `multifields._keys[field][tag]` holds field
  variants; `multicreators[i]._key[tag]` holds creator variants; `fieldMode:1` = literal/institutional.
- The variant is tagged **`en`** but its content is **transliteration** ("Uyun al-Akhbar…"), so `en`
  is registered as the *transliteration* slot (not translation). Easily reconfigurable per tag.
- Strings carry embedded bidi control chars (U+202B RLE / U+202C PDF) that the parser strips from
  variant values.
- mlzsync creator fields are `lastName`/`firstName` (Zotero names) → map to CSL `family`/`given`;
  `fieldMode:1` → `literal`.

CNE's convention (for the same item, were it authored in CNE):
```
Extra:
  cne-title-romanized: Uyun al-Akhbar Vol. 4
  cne-author-0-last-romanized: al-Dai al-Ajal Syedna Idris Imaduddin RA
```
CNE keys: `cne-<field>-<original|romanized|translated>` for simple fields
(title, container-title, publisher, …) and `cne-<creator>-<index>-<last|first>-<variant>` for
creators (e.g. `cne-author-0-last-romanized`).

## Architecture & components

### New module: `src/taskpane/cite-variants.js` (UMD, Node-testable, pure — no Office.js)

The entire sourcing layer. Functions:

- `parseMlzsync(note)` → `null` if no `mlzsync1:` prefix; else inflate the length-prefixed JSON and
  return a **normalized** intermediate:
  ```
  { fields:   { <cslField>: { <tag>: "value" } },
    creators: { <cslCreatorVar>: { <index>: { <tag>: {family?,given?,literal?} } } } }
  ```
  Maps mlzsync `multifields._keys` → `fields`; `multicreators` → `creators` (lastName/firstName →
  family/given; fieldMode 1 → literal). Strips bidi control chars from values.
- `parseCne(extra)` → same normalized intermediate from `cne-*` Extra lines (or `null` if none).
  Groups `cne-<field>-<variant>` and `cne-<creator>-<index>-<part>-<variant>`; maps CNE variant
  names to tags (`romanized`→translit tag, `translated`→translation tag, `original`→ the real
  script, used only if the real field is empty).
- `applyVariantsToItem(item)` → returns a **new** item (does not mutate input). Chooses source:
  CNE if `parseCne` non-null, else mlzsync. Writes:
  - `item.multi = { main: {...}, _keys: { <field>: { <tag>: value } } }`
  - for each creator in the CSL creator array at index *i*:
    `item[creatorVar][i].multi = { _key: { <tag>: { family|given|literal } } }`
  When neither source present, returns the item unchanged.
- `enrichItemMap(items)` → maps `applyVariantsToItem` over an id→item map. This is the single call
  site used by both the fixture load and SP-2's live fetch.
- `LANG_TAGS` (or an options arg) declares which tags are transliteration vs translation, so the
  pane can register them with the engine via the existing `langPrefs.translit` / `langPrefs.translat`
  hooks. Default: `translit = ["en", ...cne romanized tag]`, `translat = [cne translated tag]`.

### Wiring the source into the item map

`cite-pane.js` builds `cache.items` at load (fixture) and SP-2 populates it from live Zotero via
`fetchCslJson`. Both paths route the raw item map through `CiteVariants.enrichItemMap(...)` **before**
it reaches `buildEngine`, so the engine sees `multi`-enriched items and needs **no change**.

### Pane: the Variant selector (`taskpane.html` / `cite-pane.js` / `taskpane.css`)

- New `#cite-variant` dropdown alongside `#cite-style` / `#cite-locale`: **Original (ar)** (value
  `orig`, default), **Romanized** (`translit`), **Both** (`both`).
- `currentVariant()` reads it (default `orig`), mirroring `currentStyleFile()` / `currentLang()`.
- A pure mapper `variantToLangPrefs(variant, tags)` → the `langPrefs` object passed to
  `CiteEngine.build`. Presets:
  - `orig`   → no langPrefs override (real fields render; current behavior).
  - `translit` → all segment prefs = `["translit"]` (fallback to real field when a variant is absent).
  - `both`   → all segment prefs = `["orig","translit"]` (orig + romanized; best-effort layout).
  Segment keys set uniformly: `persons`, `institutions`, `titles`, `journals`, `publishers`,
  `places`, `number`, `title-short` (per citeproc-js CSL-M `langPrefs` groups) + register
  `translit`/`translat` tag lists.
- `buildEngine(styleFile, lang)` gains the current variant's `langPrefs` (or `buildEngine` grows a
  third arg). Applies to preview, insert, and refresh alike.

### Persistence, tag schema & Refresh

- Citation tag payload **v:1 → v:2**: add `variant: "orig"|"translit"|"both"`.
  `parseCitationTag` migrates read-time: missing `variant` → `"orig"` (matches today's output).
  `buildCitationTag` / `buildBibliographyTag` write `variant` from the pane.
- `refreshCitations()` already reads `currentStyleFile()`/`currentLang()` from the pane; it also
  reads `currentVariant()` and passes its `langPrefs` into the per-CC re-render, and rewrites each
  tag with the current `variant`. Bulk re-format semantics, consistent with SP-C.

## Data flow

```
raw item map (fixture load OR SP-2 fetchCslJson)
  → CiteVariants.enrichItemMap()            // mlzsync/cne-* → item.multi._keys / creator.multi._key
  → cache.items
Cite tab: user picks style + locale + VARIANT (orig/translit/both)
  → variantToLangPrefs(variant) → langPrefs
  → CiteEngine.build({..., langPrefs})       // engine already consumes langPrefs
      → citation/bibliography HTML (variant-selected)
      → cite-word insert (footnote/endnote/inline/bib), tag payload v2 {..., variant}
Refresh: pane's current style+locale+variant re-applied to every tagged CC; tags rewritten v2
```

## Error handling

- Malformed mlzsync length prefix or JSON → `parseMlzsync` returns `null` (fall through to cne-* /
  real fields); never throws into the render path.
- CNE line with an unknown field/variant → ignored, other lines still parsed.
- A creator index in the variant data with no matching creator in the CSL array → skipped.
- Missing variant for the selected policy → citeproc `langPrefs` fallback renders the real field
  (no blank output).

## Testing (node-`assert`, added to the `npm test` chain)

- `tests/cite-variants.test.js` (new):
  - `parseMlzsync` on the user's **real exported item** → expected fields/creators + bidi stripped.
  - `parseCne` on a `cne-*` Extra sample → same normalized shape.
  - `applyVariantsToItem`: CNE-wins-when-both, mlzsync-only, cne-only, neither (passthrough),
    literal vs family/given name mapping.
  - `variantToLangPrefs`: orig/translit/both → expected langPrefs shape + tag registration.
- `tests/cite-engine.test.js` (extend): a variant-enriched item rendered under each policy
  (orig/translit/both) in `ar` and `en`, asserting the output HTML switches variants.
- `tests/cite-word.test.js` (extend): tag payload v2 round-trip (`variant` written + parsed) and
  v1→v2 read migration (missing `variant` → `"orig"`).
- Fixture: add one mlzsync item (the user's) and one cne-* item.

## Risks

- **"Both" layout is style-dependent** — citeproc's bracketed-secondary rendering relies on the
  style's multi affixes; may need per-style tuning. Mitigation: ship orig/translit as the reliable
  core; treat "both" as best-effort and cover it with a rendering test to catch regressions.
- **Tag registration semantics** — `en` currently means transliteration for this library; if some
  items genuinely carry translations under a tag, the translit/translat mapping needs revisiting
  (isolated in `LANG_TAGS`).
- **CNE format drift** — CNE is young; its `cne-*` key scheme could change. Mitigation: parser is
  pure + fully tested; a scheme change is a localized edit.
- **Asset version** — bump the pane asset version when this ships (repo practice).

## Success criteria

From the fixture and from the live Zotero library, a user can pick an item that has Arabic +
romanized variants, choose **Romanized** (or **Both**) in the pane, and insert a footnote /
bibliography that renders the selected variant — with the real Arabic fields as automatic fallback
where a variant is absent; the choice is stored in the citation tag and reproduced on Refresh;
items authored in either Juris-M (mlzsync) or CNE (`cne-*`) both work, CNE winning when both are
present; and `npm test` passes including the new variant/engine/word tests.
