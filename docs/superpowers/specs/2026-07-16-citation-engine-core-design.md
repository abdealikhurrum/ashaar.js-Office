# SP-1 — Citation Engine Core + Word Output (design)

**Date:** 2026-07-16
**Status:** Approved design, pre-implementation
**Sub-project:** SP-1 of the Annotation & Citation initiative (see "Program context" below)

## Summary

Add a citation subsystem to the Ashaar Word add-in that renders scholarly citations and
bibliographies — with correct Arabic/RTL output and Arabic locale terms — and inserts them
into the Word document as footnotes/endnotes and a bibliography block. The citation engine is
the unreleased **`Juris-M/citeproc-js` master run in CSL-M mode**, with the latest CSL locales
and styles pulled in as git submodules. SP-1 proves the whole vertical slice end-to-end using a
**CSL-JSON fixture** as its data source; the live Zotero connection is the next sub-project (SP-2).

## Program context

The overall goal is annotation + citation riding on top of the user's Zotero/Juris-M library,
using the latest CSL + citeproc-js for better Arabic support, delivered inside this Word add-in.
That is a multi-subsystem initiative and has been decomposed into sequential sub-projects, each
with its own spec → plan → implementation cycle:

- **SP-1 (this spec)** — Citation engine core + Word output (footnotes/endnotes + bibliography),
  RTL/Arabic rendering, Arabic locale terms, selectable styles.
- **SP-2** — Live Zotero connection: `server.mjs` proxy → Zotero 9 local API + Better BibTeX
  (JSON-RPC / CAYW), item picker, caching. Replaces the fixture with the real library.
- **SP-3** — Multilingual variant model: where original Arabic / transliteration / translation
  variants are authored and stored, mapped into citeproc's `multi` model.
- **SP-4** — Source classification & differentiated output: primary/secondary and
  Fatemi/non-Fatemi metadata → grouped/segmented bibliographies with distinct styling.
- **SP-5** — Hijri / dual (AH/CE) dates.
- **SP-6** — Richer annotations: annotated bibliography, in-text source notes, PDF/reading
  annotation import.

### Decisions already fixed (apply across the program)

- **Host / deliverable:** this Word add-in repo; citations are inserted into Word documents.
- **Data path (product):** live local connection to the reference manager (SP-2). SP-1 uses a
  CSL-JSON fixture only for its own dev/testing so the engine work is not blocked on SP-2.
  The mixed-content problem (HTTPS pane → HTTP `localhost:23119`) is solved in SP-2 by proxying
  through the add-in's own HTTPS `server.mjs` (same-origin pane → server; server → Zotero locally).
- **Engine:** `Juris-M/citeproc-js` master (CSL-M mode). Rationale: it is the only actively
  maintained engine with the multilingual (CSL-M) model that the Arabic-variant requirement needs;
  its master is ahead of any tagged/npm release (the "unreleased" code the user wants); it is pure
  JS with no WASM, matching this repo's no-build vendor/UMD architecture. `citeproc-rs` was
  rejected (stalled, no CSL-M, WASM weight). citeproc-js delegates RTL to the rendering surface
  (it removed internal RTL handling in 1.1.237), which suits Word's native bidi.
- **Locales:** latest from `citation-style-language/locales` as a submodule ("always latest").
- **Min platform:** WordApi 1.5 (Word 2021 / Microsoft 365 desktop) — required for the
  footnote/endnote API. Feature-detected at runtime.

### Open decision recorded for SP-3 (does NOT block SP-1 or SP-2)

The user runs **Zotero 9 + Better BibTeX**, not Juris-M. Plain Zotero has no native multilingual
`multi` model; Juris-M has it but has lost Better BibTeX support and lags Zotero versions. The
citeproc-js CSL-M engine consumes the `multi` model regardless of origin, so the only question is
where variant data comes from. SP-3 will choose among: (a) use Juris-M purely as a
variant-authoring tool; (b) adopt an Extra-field convention in Zotero 9 parsed into `multi`;
(c) an add-in sidecar store keyed by item. For SP-1, a `multi` item is included in the fixture
(hand-authored or from a small Juris-M export) purely to test CSL-M rendering.

## SP-1 scope

**In scope**

- Vendor `citeproc-js` (CSL-M), CSL locales, and a curated set of CSL styles as git submodules,
  synced to `src/vendor/` by a sync script, following the existing ashaar-js vendoring pattern.
- A client-side engine wrapper that runs citeproc in CSL-M mode and returns citation +
  bibliography HTML from loaded item data.
- A "Cite" tab in the task pane: load the fixture, list items, select one or more, choose output
  form (footnote / endnote / inline) and style, insert; live HTML preview in the pane.
- Insert into Word: footnotes/endnotes (empty note → HTML into the note body), and a
  bibliography block inside a content-control-tagged region.
- Correct RTL/Arabic rendering and Arabic (`ar`) locale terms.
- Selectable styles: stock **Chicago notes-and-bibliography** and **APA**, plus a **slot** for
  the Fatemi-aware custom variants (`chicago-notes-fatemi`, `apa-fatemi`).
- node-`assert` tests, added to the `npm test` chain, matching repo conventions.

**Explicitly deferred**

- Live Zotero connection (SP-2).
- Authoring/storage of multilingual variants (SP-3). SP-1 *consumes* `multi` variants when
  present in the fixture; it does not build the variant store.
- Fatemi/primary classification *data* (SP-4). The custom styles are authored in SP-1 but their
  Fatemi-specific behavior only activates once SP-4 supplies the classifying CSL variable; until
  then they render identically to their stock parents.
- Hijri dates (SP-5); annotated bibliography, in-text source notes, PDF annotation import (SP-6).
- Live-updating citation fields / full document re-sync (Zotero-Word-plugin style). SP-1 tags
  inserted citations with content controls carrying their cited item keys + style so a
  bibliography can be generated from the cited set, and so refresh/re-sync is *possible later*,
  but SP-1 does not implement re-sync.

## Architecture & components

The engine runs **client-side in the task pane**, not in `server.mjs`. Rationale: citeproc-js is
browser JavaScript; keeping it in the pane leaves `server.mjs` thin for its future SP-2 proxy
role, enables an in-pane live preview (mirroring the existing Ashaar render preview), and matches
the repo's `vendor → src/vendor → UMD` pattern.

### New vendored submodules (mirrors `vendor/ashaar-js`)

- `vendor/citeproc-js/` — `Juris-M/citeproc-js` (master, pinned commit).
- `vendor/csl-locales/` — `citation-style-language/locales` (master, pinned commit).
- `vendor/csl-styles/` — `citation-style-language/styles` (master, pinned commit; large repo —
  the sync script copies only the curated subset below, not the whole tree).

Added to `.gitmodules` alongside the existing `vendor/ashaar-js` and `vendor/font-fatemi` entries.

### `scripts/sync-citeproc-vendor.mjs`

Clone of `scripts/sync-ashaar-vendor.mjs`. Copies:

- citeproc bundle (`citeproc.js`, the raw browser bundle) → `src/vendor/citeproc.js`.
- Arabic + English locale XML (`locales-ar.xml`, `locales-en-US.xml`, and the locale that each
  bundled style declares) → `src/vendor/csl-locales/`.
- The curated style set → `src/vendor/csl-styles/`:
  `chicago-notes-bibliography.csl`, `apa.csl` (from the submodule),
  and the repo-owned custom styles copied from `src/styles/` (see below).
- Writes `src/vendor/CITEPROC_UPSTREAM_VERSION` (repo/branch/commit/date/version) for each
  submodule, exactly like `ASHAAR_UPSTREAM_VERSION`.

`package.json`: add `sync:citeproc` and `update:citeproc` scripts mirroring the ashaar ones.

### Repo-owned styles: `src/styles/`

The two Fatemi-aware custom styles live here (NOT in the submodule, since they are ours):

- `src/styles/chicago-notes-fatemi.csl` — derived from stock Chicago-notes.
- `src/styles/apa-fatemi.csl` — derived from stock APA.

Each keys its Fatemi / primary-vs-secondary behavior off a **CSL variable** (candidate: `genre`,
or a dedicated category variable) that **SP-4 will populate**. Until SP-4 supplies that data, the
conditional branches are inert and the styles render identically to their stock parents. This lets
all four styles ship and be selectable in SP-1 without blocking on SP-4.

### `src/taskpane/cite-engine.js` (UMD, Node-testable)

Wraps citeproc-js. Responsibilities:

- Build the citeproc `sys` object:
  - `retrieveLocale(lang)` — returns bundled locale XML text from `src/vendor/csl-locales/`.
  - `retrieveItem(id)` — returns the CSL-JSON item from the currently loaded item map.
- Instantiate `CSL.Engine` in **CSL-M mode** with a chosen style XML and locale.
- `formatCitation(itemKeys, opts)` → citeproc citation **HTML**.
- `formatBibliography()` → citeproc bibliography **HTML**.
- Configure CSL-M multilingual language preferences so that when a `multi` variant exists
  (original Arabic + transliteration/translation), the engine renders per the selected policy;
  when absent, the "real" field is the fallback.

Depends on: bundled citeproc + locale/style text. No Office.js dependency (so it is unit-testable
under Node).

### `src/taskpane/cite-word.js` (UMD, Node-testable for HTML→payload building)

Turns engine HTML into Word insertions:

- Footnote/endnote: `range.insertFootnote()` / `range.insertEndnote()` to create an **empty**
  note, then `note.body.getRange().insertHtml(sanitizedHtml)` to inject rich, RTL-capable content.
- Bibliography: `insertHtml(sanitizedHtml)` into a **content-control-tagged** block (mirrors the
  existing Ashaar content-control tag pattern) so the block is findable/regenerable later.
- RTL: when the active locale/item is Arabic, set the note/bibliography paragraph direction to
  RTL and select the `ar` locale for terms.
- HTML sanitization: reduce citeproc's HTML to the tag/attribute subset that Word's `insertHtml`
  reliably accepts (a known-safe whitelist). The pure sanitizer + payload-building functions are
  Node-testable; the actual `Word.run` insertion is exercised via the manual Word checklist.
- Citation tagging: each inserted citation's content control stores the cited item keys + active
  style, enabling bibliography-from-cited-set now and refresh/re-sync later.

### Cite tab (`taskpane.html` / `taskpane.js` / `taskpane.css`)

- Load the CSL-JSON fixture; render an item list.
- Select item(s); choose output form (footnote / endnote / inline) and style (dropdown:
  Chicago-notes, APA, Chicago-notes-Fatemi, APA-Fatemi); choose display locale (ar / en).
- Live HTML preview of the citation + bibliography (same spirit as the existing render preview).
- Insert button → `cite-word` → Word.
- Follows the existing IIFE + `bind()`-on-Office-ready structure.

## Data flow

```
CSL-JSON fixture (real BBT/Juris-M export + one hand-authored multi item) → item map
  → cite-engine (citeproc CSL-M; chosen style + ar/en locale)
      → citation HTML     → cite-word → empty footnote/endnote → sanitized insertHtml into note body (RTL-aware)
      → bibliography HTML  → cite-word → sanitized insertHtml into tagged bibliography content control
Each inserted citation's content control stores {citedItemKeys, style}
  → bibliography generated from the union of cited keys (refresh/re-sync deferred)
```

## Error handling

- Missing locale/style file → explicit pane error, not silent.
- citeproc throws on a malformed item → caught; that item flagged in the pane, others proceed.
- `insertHtml` rejects markup → sanitizer restricts to the Word-safe whitelist; if a construct is
  still lossy, documented fallback is hand-built OOXML (as `word-html.js` already does).
- WordApi < 1.5 (no footnote API) → feature-detected; note insertion disabled with a clear message;
  inline insertion still offered.

## Testing (node-`assert`, added to the `npm test` chain)

- `tests/cite-engine.test.js` — deterministic engine output: fixture in → expected citation and
  bibliography HTML out, across Chicago-notes and APA, in `ar` and `en`, including a `multi`-variant
  item to prove CSL-M multilingual rendering (original + transliteration).
- `tests/cite-word.test.js` — HTML→insertion-payload building, HTML sanitization to the Word-safe
  subset, and RTL paragraph flagging. Pure functions only; no Office.js.
- Submodule pin check in the spirit of `tests/vendor-version.test.js` — assert the citeproc/locale
  submodule versions recorded in `CITEPROC_UPSTREAM_VERSION` match expectations.
- A manual Word verification checklist (like the existing conversion/settings checklists) for the
  `Word.run` insertion paths (footnote body HTML, RTL direction, bibliography block).

## Risks

- **`insertHtml` fidelity** for citeproc's markup (nested spans, RTL runs). Mitigation: sanitizer +
  tests; fallback to hand-built OOXML if needed.
- **CSL-M API surface** differs subtly from plain citeproc. Mitigation: pinned submodule + engine
  tests guard against drift.
- **CSL styles submodule size** (`citation-style-language/styles` is large). Mitigation: the sync
  script copies only the curated subset, never the whole tree.
- **Asset version staleness** — per repo practice, the pane asset version must be bumped when pane
  code ships so installed users don't run stale JS.

## Success criteria

From the fixture, a user can: pick an Arabic item, pick Chicago-notes (or APA), and insert a
correctly-rendered RTL footnote (with Arabic locale terms) plus a bibliography entry into a Word
document; the same works in English; a `multi`-variant item renders its transliteration/translation
per CSL-M; all four styles are selectable (the two Fatemi variants render as their stock parents
until SP-4); and `npm test` passes including the new engine/word tests.
