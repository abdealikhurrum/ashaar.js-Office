# SP-3 — Multilingual variant model (design)

**Date:** 2026-07-17
**Status:** Approved design, pre-implementation
**Sub-project:** SP-3 of the Annotation & Citation initiative (see SP-1 core design for program context)
**Branch:** `feat/citation`

## Summary

Give the citation subsystem the ability to render an item's **original-script (Arabic),
romanized (transliteration), or both** variants, sourced from the user's real reference library —
and consolidate all multilingual data onto a single **maintained** convention: **Cite Non-English
(CNE)**, a Zotero plugin that stores variants as `cne-*` lines in the item's Extra field.

The citeproc CSL-M *rendering* side already exists (`cite-engine.js` consumes `langPrefs` +
citeproc's `multi` model, wired in SP-1). SP-3 has two components:

- **Component A — add-in variant model (the feature):** a pure parser that reads `cne-*` variant
  data and normalizes it into the `multi` shape the engine already reads, plus a pane selector
  (Original / Romanized / Both) that chooses which variant renders, persisted in the citation tag
  and re-applied on Refresh, exactly like style/locale.
- **Component B — one-time migration utility:** a dev-run script that converts the user's legacy
  **Juris-M mlzsync** blobs into **CNE `cne-*`** lines and writes them back into the Zotero library
  via the local API, so all multilinguality lives in the one maintained convention going forward.

Rendering stays in our own client-side engine (see "Why we keep our own engine").

## Decisions fixed in brainstorming (2026-07-17)

- **CNE is the single go-forward convention.** In Zotero 9 + Better BibTeX (no Juris-M installed),
  the legacy **mlzsync** blob in Extra is *frozen*: nothing reads, edits, or regenerates it. CNE is
  the only maintained way to author/maintain variants. Therefore the add-in's **runtime reads
  `cne-*` only**; mlzsync is handled once, by the migration utility, not on the render path.
- **CNE has no Juris-M importer** (confirmed against the CNE README) → we build the migration
  utility ourselves (Component B).
- **Output form = pane-selectable**, per insertion: **Original (ar) / Romanized / Both**.
- **Granularity = single global policy** for v1 — one selector applies the same policy to all
  segments (names, titles, publisher, …). A per-segment split is deferred (YAGNI).
- **Refresh uses the pane's current** variant policy (bulk re-format), consistent with how SP-C
  re-reads style/locale from the pane rather than freezing them per citation.
- **Rendering stays in our own client-side engine.**

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
insertion. So CNE is treated purely as an upstream *authoring/storage* tool; SP-3 consumes its
stored `cne-*` data and our own engine renders.

## The data

### Legacy source (input to the migration utility only)

The user's "Better CSL JSON" export carries Arabic as the real fields and the Juris-M variants as a
length-prefixed mlzsync blob in the Extra field (surfaced as `note` in CSL-JSON):

```json
{
  "id": "AldAEy…4", "type": "book", "language": "ar",
  "title": "‫عيون الأخبار ج/4‬",
  "author": [{ "literal": "الداعي الأجل سيدنا إدريس عماد الدينؓ" }],
  "note": "mlzsync1:0215{\"type\":\"book\",\"multifields\":{\"main\":{},\"_keys\":{\"title\":{\"en\":\"‫Uyun al-Akhbar Vol. 4‬\"}}},\"multicreators\":{\"0\":{\"_key\":{\"en\":{\"lastName\":\"al-Dai al-Ajal Syedna Idris Imaduddin RA\",\"firstName\":\"\"}},\"fieldMode\":1}}}"
}
```

mlzsync shape: `mlzsync1:` + **4-digit zero-padded length** + JSON; `multifields._keys[field][tag]`
holds field variants; `multicreators[i]._key[tag]` holds creator variants (`lastName`/`firstName`,
`fieldMode:1` = literal/institutional). Values carry embedded bidi control chars (U+202B RLE /
U+202C PDF). The variant tag is **`en`** but its content is **transliteration**, so it maps to the
CNE **romanized** slot (not translated).

### Go-forward source (what the add-in reads)

After migration, the same item's Extra carries CNE lines:
```
cne-title-romanized: Uyun al-Akhbar Vol. 4
cne-author-0-last-romanized: al-Dai al-Ajal Syedna Idris Imaduddin RA
```
CNE keys: `cne-<field>-<original|romanized|translated>` for simple fields
(title, container-title, publisher, …); `cne-<creator>-<index>-<last|first>-<variant>` for creators
(e.g. `cne-author-0-last-romanized`). In CSL-JSON these Extra lines surface in `note` (as mlzsync
did), which is where the add-in parser reads them.

> **Load-bearing verification (plan, early):** confirm Better BibTeX's "Better CSL JSON" passes
> `cne-*` Extra lines through to `note` unmodified (they are non-CSL keys, so expected — but BBT may
> treat some Extra lines as CSL-variable overrides). If BBT strips them, the add-in reads items via
> the Zotero **local API** CSL-JSON instead, or reads the raw `data.extra` field directly.

## Component A — add-in variant model

### New module: `src/taskpane/cite-variants.js` (UMD, Node-testable, pure — no Office.js)

- `parseCne(noteOrExtra)` → parses `cne-*` lines → normalized intermediate:
  ```
  { fields:   { <cslField>: { <tag>: "value" } },
    creators: { <cslCreatorVar>: { <index>: { <tag>: {family?,given?,literal?} } } } }
  ```
  Maps CNE variant names to tags: `romanized`→translit tag, `translated`→translation tag,
  `original`→used only if the real field is empty. Strips bidi control chars.
- `applyVariantsToItem(item)` → returns a **new** item (no mutation) with
  `item.multi = { main, _keys: { <field>: { <tag>: value } } }` and, per creator at index *i*,
  `item[creatorVar][i].multi = { _key: { <tag>: { family|given|literal } } }`. No `cne-*` present →
  item unchanged.
- `enrichItemMap(items)` → maps `applyVariantsToItem` over an id→item map. Single call site for
  both the fixture load and SP-2's live fetch.
- `variantToLangPrefs(variant, tags)` → the `langPrefs` object for `CiteEngine.build`:
  - `orig`     → no override (real fields render; current behavior).
  - `translit` → all segment prefs `["translit"]` (fallback to real field when a variant is absent).
  - `both`     → all segment prefs `["orig","translit"]` (orig + romanized; best-effort layout).
  Sets segment keys uniformly (citeproc-js CSL-M `langPrefs` groups: `persons`, `institutions`,
  `titles`, `journals`, `publishers`, `places`, `number`, `title-short`) and registers the
  translit/translat tag lists via the engine's existing `setLangTagsForCsl*` hooks.
- `parseMlzsync(note)` — a pure inflater kept here too, but **only the migration utility calls it**
  (not the runtime path). Reuses the same normalized intermediate; lastName/firstName→family/given,
  fieldMode 1→literal, bidi-strip. Co-locating it keeps all variant parsing + its tests in one
  module.

### Wiring the source into the item map

Both the fixture load and SP-2's `fetchCslJson` route the raw item map through
`CiteVariants.enrichItemMap(...)` **before** `buildEngine`, so the engine sees `multi`-enriched
items and needs **no change**.

### Pane: the Variant selector (`taskpane.html` / `cite-pane.js` / `taskpane.css`)

- New `#cite-variant` dropdown by `#cite-style` / `#cite-locale`: **Original (ar)** (`orig`, default),
  **Romanized** (`translit`), **Both** (`both`).
- `currentVariant()` reads it (default `orig`), mirroring `currentStyleFile()` / `currentLang()`.
- `buildEngine` gains the current variant's `langPrefs` (third arg or options); applies to preview,
  insert, and refresh alike.

### Persistence, tag schema & Refresh

- Citation tag payload **v:1 → v:2**: add `variant: "orig"|"translit"|"both"`.
  `parseCitationTag` migrates read-time: missing → `"orig"`. `buildCitationTag` /
  `buildBibliographyTag` write `variant` from the pane.
- `refreshCitations()` reads `currentVariant()` alongside style/locale, passes its `langPrefs` into
  the per-CC re-render, and rewrites each tag with the current `variant`. Bulk re-format, per SP-C.

## Component B — migration utility: `scripts/migrate-mlzsync-to-cne.mjs`

A one-time, dev-run Node script (not pane UI) that consolidates legacy data onto CNE.

- **Read:** enumerate library items via the Zotero **local API** (`localhost:23119/api/…`, no auth);
  for each, read the raw `data.extra` field.
- **Convert:** `parseMlzsync(extra)` → for each field/creator variant, emit the corresponding
  `cne-*` line(s). Mapping: mlzsync `en` field tag → `cne-<field>-romanized`; creator
  `_key.en.lastName`/`firstName` → `cne-<creator>-<i>-last-romanized` / `-first-romanized`. Strip
  bidi control chars. (The Arabic "original" already lives in the item's real fields; CNE reads that
  as the original, so `-original` lines are emitted only if a real field is empty — confirmed in the
  plan against CNE's read behavior.)
- **Write:** merge the new `cne-*` lines into `extra` (preserving all other Extra lines, including
  the mlzsync block) and **PATCH** the item via the local API with its current version
  (`If-Unmodified-Since-Version`; on 412, re-fetch + retry).
- **Safety (writing to the user's real library is hard to reverse):**
  - **Dry-run by default** — prints a per-item diff of the `extra` changes; writes nothing.
  - Requires explicit `--write` to PATCH.
  - **Backs up** first — exports the affected items (JSON) to `scratch`/a dated file before any write.
  - **Idempotent** — an item that already has the target `cne-*` lines is skipped (or overwritten
    only with `--force`); safe to re-run.
  - **Non-destructive** — leaves the mlzsync block in Extra unless `--strip-mlzsync` is passed
    (after the user has verified the migration).
  - Prints a summary (converted / skipped / failed) and never proceeds past a parse error on one
    item (that item is reported and skipped).

> **Load-bearing verification (plan, first gate):** confirm the user's Zotero build exposes local-API
> **write** (PATCH) support (read is universal; write is version-dependent). If unavailable, fall
> back to: emit an updated CSL-JSON/RDF the user re-imports, or emit per-item Extra text the user
> pastes. Component A does not depend on this — it is testable against fixtures regardless.

## Data flow

```
Migration (one-time):
  Zotero local API (data.extra) → parseMlzsync → cne-* lines → merge into extra → PATCH item

Runtime (the feature):
  raw item map (fixture load OR SP-2 fetchCslJson, cne-* in note/extra)
    → CiteVariants.enrichItemMap()          // cne-* → item.multi._keys / creator.multi._key
    → cache.items
  Cite tab: user picks style + locale + VARIANT (orig/translit/both)
    → variantToLangPrefs(variant) → langPrefs → CiteEngine.build({..., langPrefs})
      → citation/bibliography HTML (variant-selected)
      → cite-word insert; tag payload v2 {..., variant}
  Refresh: pane's current style+locale+variant re-applied to every tagged CC; tags rewritten v2
```

## Error handling

- No `cne-*` on an item → real fields render (no blank output; citeproc `langPrefs` fallback).
- Malformed `cne-*` line → ignored; other lines still parsed.
- Creator index with no matching CSL creator → skipped.
- Migration: parse error on one item → reported + skipped, run continues; PATCH 412 → re-fetch +
  retry once, else report + skip.

## Testing (node-`assert`, added to the `npm test` chain)

- `tests/cite-variants.test.js` (new):
  - `parseCne` on a `cne-*` sample → expected normalized shape; bidi stripped; unknown line ignored.
  - `applyVariantsToItem`: cne present (fields + literal/family/given creators), cne absent
    (passthrough), creator-index mismatch skip.
  - `variantToLangPrefs`: orig/translit/both → expected langPrefs + tag registration.
  - `parseMlzsync` on the user's **real exported item** → expected normalized shape (shared with
    Component B).
  - `mlzsyncToCneLines` (pure converter used by Component B) → expected `cne-*` lines for the real
    item; idempotency (re-emit is stable); bidi stripped.
- `tests/cite-engine.test.js` (extend): a variant-enriched item under each policy (orig/translit/
  both) in `ar` + `en`, asserting the output HTML switches variants.
- `tests/cite-word.test.js` (extend): tag payload v2 round-trip (`variant` written + parsed) and
  v1→v2 read migration (missing `variant` → `"orig"`).
- Fixture: add one cne-* item; keep the real mlzsync item as the migration-converter fixture.
- The migration script's I/O (local-API read/PATCH) is exercised via the manual verification
  checklist against the live library (dry-run diff reviewed before `--write`), not unit tests.

## Risks

- **Local-API write support** (Component B) is version-dependent — verified first; documented
  fallbacks (re-import / paste) if absent.
- **BBT passthrough of `cne-*`** to CSL-JSON `note` — verified early; fallback to local-API CSL-JSON
  or raw `data.extra`.
- **"Both" layout is style-dependent** — citeproc's bracketed-secondary rendering relies on style
  multi affixes; ship orig/translit as the reliable core, "both" best-effort + covered by a test.
- **CNE `-original` semantics** — whether CNE requires explicit `cne-*-original` or reads the real
  field as original; pinned in the plan (affects the migration mapping only).
- **CNE format drift** — CNE is young; the `cne-*` scheme could change. Parser is pure + tested; a
  scheme change is a localized edit.
- **Writing to the user's real library** — mitigated by dry-run default, backup, idempotency,
  non-destructive default, and manual diff review before `--write`.
- **Asset version** — bump the pane asset version when Component A ships (repo practice).

## Success criteria

- **Component B:** running the migration in dry-run prints a correct `cne-*` diff for the user's
  mlzsync items; with `--write`, the items' Extra fields gain the `cne-*` lines (verified in Zotero),
  idempotently and non-destructively.
- **Component A:** from the fixture and from the live (migrated) Zotero library, the user can pick an
  item with Arabic + romanized variants, choose **Romanized** (or **Both**) in the pane, and insert a
  footnote / bibliography rendering the selected variant — real Arabic fields as automatic fallback
  where a variant is absent; the choice is stored in the v2 citation tag and reproduced on Refresh.
- `npm test` passes including the new variant/engine/word tests.
