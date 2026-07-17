# SP-A — Citation locators + reference-list editing + citation tagging (design)

**Date:** 2026-07-17
**Status:** Approved design, pre-implementation
**Track:** First of a three-part "document-embedded, re-editable citations" initiative that
builds on SP-1 (citation engine) and SP-2 (live Zotero). This is **SP-A**; SP-B (persist &
restore the reference set in the document) and SP-C (refresh & re-format inserted citations)
are deferred to their own specs and depend on the citation tagging established here.

## Summary

Add three connected capabilities to the Cite tab:

1. **Per-citation locators** — page / chapter / section / verse + a value (`42`, `42–45`),
   entered per checked item, applying to *this insertion only* (a reference can be cited many
   times at different locators). Rendered by the existing CSL engine (localized term: "p." in
   English, the Arabic equivalent under the `ar` locale).
2. **Reference-list editing** — a `×` per item that removes a work from the working set (and
   therefore from the bibliography).
3. **Citation tagging (foundation)** — every inserted citation is wrapped in a content control
   tagged with a hardened, unambiguous encoding of `{style, locale, keys, locators}`, so SP-C can
   later find and refresh it. This wires in the citation-tagging follow-up deferred from SP-1 and
   replaces its delimiter-fragile `AshaarCite:<style>:<keys>` format.

Insertion remains **static** (no live-field refresh yet — that is SP-C). Everything inserted
persists in the document as ordinary Word content; the pane's working set is transient (SP-B adds
document persistence).

## Model (settled in brainstorming)

- The **Items list = the reference set = the bibliography source.** No hidden document state.
- **Checkbox** = "include this item in the citation I'm about to insert."
- **Locator** (type ∈ {page, chapter, section, verse} + value) sits on each *checked* item and
  applies to **this insertion only**.
- **`×` per row** = remove the work from the working set (`cache.items`) → gone from citations
  and the bibliography.
- **Insert citation** → footnote / endnote / inline of the checked items at their locators; the
  locator inputs clear after a successful insert (per-insertion), the item list is unchanged.
- **Insert bibliography** → all items currently in the list.

## Scope

**In scope**
- Locator UI (type select + value input) shown inline under each checked item; label set
  page / chapter / section / verse.
- `cite-engine.js`: `cite()` accepts locators and passes them to `makeCitationCluster`.
- `cite-word.js`: hardened citation tag encode/decode (`{v, style, locale, keys[{id,locator,label}]}`
  as base64-JSON) + bibliography tag `{style, locale}`.
- `cite-pane.js`: `×` remove, per-checked-item locator inputs, locator-aware
  preview/insert, tag write on insert, locator-input reset after insert.
- `cite-zotero.js`: `caywPick`/`parseCaywResult` return `[{citekey, locator, label}]`; best-effort
  pandoc-suffix locator parse; `addFromZotero` pre-fills + auto-checks located picks.
- `taskpane.html`/CSS markup; `ASHAAR_ASSET_VERSION` bump.
- node-`assert` tests for the pure logic (engine locators, tag round-trip, CAYW locator parse).

**Explicitly deferred**
- **SP-B**: persisting the reference set (CSL-JSON) in the document (Office custom XML part or
  `document.settings`) and restoring it on pane open.
- **SP-C**: a "Refresh" action that scans citation tags, re-runs the engine at the current
  style/locale, and rewrites each citation + rebuilds the bibliography; re-format on style change.
- Reading the citation tags back (SP-A only *writes* them; nothing consumes them until SP-C).
- Multilingual `langPrefs` (SP-3 of the original roadmap), Fatemi grouping (SP-4), Hijri (SP-5).

## Architecture & components

### `cite-engine.js` — locator-aware citations
- Change `cite(itemKeys)` → `cite(citationItems)` where `citationItems` is
  `[{ id, locator?, label? }]`. Map each to a citeproc citation item
  `{ id, locator, label }` (omit `locator`/`label` when absent) and pass to
  `makeCitationCluster`. citeproc renders the localized locator term from the active locale
  (en: "p."/"chap."/"sec."/"v."; ar: the Arabic terms already in `locales-ar.xml`).
- `bibliography()` is **unchanged** — it renders every item in the working set; `×` (which
  deletes from `cache.items`) is the whole mechanism for "remove from bibliography".
- `label` must be a CSL-valid locator term. The pane's four types map 1:1 to CSL terms
  `page`, `chapter`, `section`, `verse`.

### `cite-word.js` — hardened citation/bibliography tags
- `buildCitationTag({style, locale, items})` → `"AshaarCite:" + base64(JSON.stringify({
  v: 1, style, locale, keys: items.map(i => ({id:i.id, locator:i.locator||null, label:i.label||null})) }))`.
  base64 of a JSON payload removes the SP-1 delimiter-collision risk entirely (any `:`/`,` inside
  a style id, item key, or locator is safe).
- `parseCitationTag(tag)` → the decoded payload object, or `null` if the tag is not an
  `AshaarCite:` tag or fails to decode (defensive; SP-C will consume this).
- `buildBibliographyTag({style, locale})` → `"AshaarBib:" + base64(JSON.stringify({v:1, style, locale}))`
  (extends SP-1's static `"AshaarBibliography"` so SP-C can refresh it; keep a title of
  "Ashaar Bibliography").
- Use a UMD-safe base64 that works in both Node (tests) and the Word WebView (e.g. a small
  helper over `Buffer`/`btoa` with a UTF-8 step, since keys/locators may be non-ASCII).
- **Tag-length note:** OOXML content-control tags are length-bounded. A 1–3-key citation encodes
  well within the limit; a pathologically large multi-key citation could exceed it. SP-A only
  *writes* the tag (nothing reads it yet), so an over-long tag is non-fatal here; SP-B/SP-C will
  move to XML-part storage keyed by a short id if this proves limiting. Implementation logs a
  warning if a tag exceeds the limit rather than failing the insert.

### `cite-zotero.js` — locator capture from CAYW
- `parseCaywResult(text)` → `[{ citekey, locator, label }]` (was `string[]`). Parse the pandoc
  output:
  - Split multi-item clusters on `;` (outside is `[ ... ]`, which is stripped first).
  - Each item: leading `@citekey`, optional suffix after a comma = locator phrase.
  - Locator phrase → `{label, locator}`:
    `p.`/`pp.`/`page`/`pages` → `page`; `chap.`/`chapter` → `chapter`;
    `sec.`/`section`/`§` → `section`; `v.`/`vv.`/`verse` → `verse`.
    A bare number with no label → `page` (pandoc's default). Value = the remaining
    digits/range (`42`, `42-45`).
  - Unrecognized suffix → `{citekey}` with no locator (never throws; best-effort).
  - `""`/whitespace/nullish → `[]` (cancel), as today.
- `caywPick(fetchImpl?)` → resolves to the `[{citekey, locator, label}]` array (derives citekeys
  from `.citekey`). `fetchCslJson` still takes bare citekeys (unchanged) — callers map
  `.citekey` out.

### `cite-pane.js` — UI, preview, insert
- **Item row** (in `populateItems`): `☐ Title (citekey)  ×`. On check, an inline locator row
  appears beneath: `cite at: [type ▾] [value]` (type = page/chapter/section/verse). Unchecked
  rows show no locator row.
- **`removeItem(id)`**: delete `cache.items[id]`, rebuild the list, `renderPreview()`.
- **Citation builder**: replace `selectedIds()` with `selectedCitationItems()` → `[{id, locator,
  label}]` read from checked rows + their locator inputs (empty value ⇒ omit locator/label).
- **`renderPreview`**: `engine.cite(selectedCitationItems())` for the citation block;
  bibliography unchanged.
- **`insertCitation`**: build the HTML from `selectedCitationItems()`; wrap the inserted content
  control with `CiteWord.buildCitationTag({style, locale, items})`; after a successful insert,
  clear the locator inputs (per-insertion) and leave checkboxes/list intact.
- **`insertBibliography`**: tag the content control with `buildBibliographyTag({style, locale})`.
- **`addFromZotero`**: `caywPick()` → for each returned `{citekey, locator, label}`, fetch
  CSL-JSON by citekey (as today), add to the list; if a locator was captured, auto-check that
  item and pre-fill its locator inputs so a "Kitab, p. 42" pick lands ready to insert.
- `taskpane.html`: locator-row markup/template + `×` button; CSS for both. Bump
  `ASHAAR_ASSET_VERSION`.

## Data flow

```
Compose a citation:
  check item(s) → (optional) set [type ▾][value] per checked item
  → selectedCitationItems() = [{id, locator, label}]
  → engine.cite(items) → preview
  → Insert citation → note/inline HTML + content control tagged
       AshaarCite:base64({v,style,locale,keys:[{id,locator,label}]})
  → locator inputs cleared for the next citation

Add from Zotero:
  CAYW pick → parseCaywResult → [{citekey, locator, label}]
  → fetchCslJson(citekeys) → merge into list
  → located picks: auto-check + pre-fill locator inputs

Remove:  × → delete from cache.items → list + bibliography update
```

## Error handling

- Locator value empty ⇒ that item is cited with no locator (a label without a value is dropped);
  never errors.
- `×` on a checked item removes it and drops it from the pending citation.
- Unparseable CAYW locator ⇒ item added with no pre-filled locator.
- Over-long content-control tag ⇒ logged warning, insert still succeeds (tag unconsumed in SP-A).
- All existing SP-1/SP-2 error paths (offline Zotero fallback, insert failures) unchanged.

## Testing (node-`assert`, added to `npm test`)

- `cite-engine.test.js` — `cite([{id, locator:"42", label:"page"}])` renders "p. 42" under en and
  the Arabic locator term under `ar`; a two-item cluster each with its own locator; an item with
  no locator renders as today (regression).
- `cite-word.test.js` — `buildCitationTag`/`parseCitationTag` round-trip incl. keys/locators and a
  style id containing `:`; base64 survives non-ASCII (Arabic) values; `parseCitationTag` returns
  `null` on a non-Ashaar or corrupt tag; `buildBibliographyTag` round-trip.
- `cite-zotero.test.js` — `parseCaywResult` across `@key` (no locator), `[@key, p. 42]`→page/42,
  `[@key, pp. 42-45]`→page/42-45, `[@key, chap. 3]`→chapter/3, `[@key, sec. 2]`→section/2,
  `[@key, v. 7]`→verse/7, a bare-number suffix→page, an unrecognized suffix→no locator, and a
  multi-item `[@k1, p. 1; @k2]` cluster.
- Pane DOM behavior (× removal, locator inputs → citation, reset after insert) → browser smoke +
  the manual Word checklist.

## Success criteria

In the Cite tab: checking an item reveals a locator row; setting page/chapter/section/verse + a
value makes the preview and the inserted footnote read "…, p. 42" (Arabic term under the `ar`
locale); the same reference can be inserted repeatedly at different locators. `×` removes a work
from the list and the bibliography. A Zotero pick with a page set in the CAYW popup lands
checked with the locator pre-filled. Every inserted citation carries an `AshaarCite:` content
control tag encoding `{style, locale, keys, locators}` (verified by decoding it back). `npm test`
passes including the new assertions. (Restore-on-open and refresh are SP-B/SP-C.)
