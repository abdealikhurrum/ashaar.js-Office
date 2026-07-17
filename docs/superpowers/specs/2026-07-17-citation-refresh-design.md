# SP-C — Refresh & re-format inserted citations (design)

**Date:** 2026-07-17
**Status:** Approved design (decomposed from the 2026-07-17 persistence brainstorm), pre-implementation
**Track:** Third and final of the "document-embedded, re-editable citations" initiative. Follows
SP-A (locators + editing + **tagging**) and SP-B (persist/restore reference set). This is what the
`AshaarCite:` / `AshaarBib:` tags were laid down for — SP-A only *wrote* them; SP-C *reads* them.

## Summary

Add a **"Refresh citations"** action to the Cite tab. It scans the document for the citation +
bibliography content controls this add-in inserted (tagged `AshaarCite:` / `AshaarBib:`), re-runs
the CSL engine at the **currently-selected style + locale**, and rewrites each one in place —
preserving each citation's stored keys + locators. This makes "change the style and update the
whole document" possible without re-inserting anything. Also retires the now-vestigial
`buildBibliographyPayload.tag` field flagged in SP-A's final review.

## Grounding (MS Learn, 2026-07-17)

- `Word.Document.contentControls` gets the content controls in the **main document body** — it
  does **not** reach controls inside footnote/endnote stories.
- Footnote/endnote bodies **are** reachable: `Word.Body` exposes `getFootnoteBody(...)` /
  `getEndnoteBody(...)`, and the footnotes/endnotes collection (`Word.NoteItemCollection` via
  `body.footnotes` / `range.footnotes`) yields `Word.NoteItem`s whose `.body` is a `Word.Body`
  with its own `.contentControls`. **The exact member + whether our footnote CCs enumerate and
  rewrite is confirmed live in Word during implementation** (mirrors how SP-2 confirmed the BBT
  translator and the Arabic path confirmed `insertOoxml`-into-a-footnote). SP-C is built
  **defensively**: it refreshes every citation it can reach and reports counts, degrading
  gracefully if footnote enumeration yields nothing on a given Word build.
- `Word.ContentControl` supports `insertHtml` / `insertOoxml` (replace) and `tag` read/write.

## Decisions

- **Target style/locale = the pane's current selection.** Refresh applies the currently-chosen
  style + locale to *every* tagged citation (that is the "re-format the whole document to a new
  style" use case), and rewrites each `AshaarCite:` tag's `style`/`locale` to the new values
  (keys + locators preserved).
- **Citation body** is re-rendered from the tag's stored `keys` (`[{id,locator,label}]`) through
  the existing insert pipeline: `engine.cite(items)` → for Arabic, the OOXML path
  (`buildCitationParagraphOoxml` + `insertOoxml`); for LTR/inline, `insertHtml` of
  `wrapRtlRuns(sanitize(...))` — i.e. reuse SP-A/Arabic-OOXML rendering, keyed on the *new* locale.
- **Bibliography** (`AshaarBib:` CC) is rebuilt from the current **reference set** (`cache.items`)
  at the new style/locale — consistent with `insertBibliography`. (Union-of-actually-cited-keys is
  a possible future refinement, noted, not built.)
- **Items whose ids are no longer in `cache.items`**: the engine still needs the CSL-JSON to
  render them. Refresh builds its engine over `cache.items`; a citation referencing an id absent
  from the reference set is **skipped and counted as "unresolved"** (reported), not errored — the
  user can re-add it from Zotero. (SP-B restores the reference set on open, so the common case is
  that cited ids are present.)
- **Scope:** a single "Refresh citations" button. No automatic-on-style-change trigger (explicit
  button is clearer and avoids surprise rewrites); no live-field behavior beyond this.

## Architecture & components

### `cite-word.js` — a tiny pure helper (Node-testable)
- `citationItemsFromTag(parsed)` → `[{id, locator, label}]` from a `parseCitationTag` payload's
  `keys` (drops `null` locator/label). Pure; unit-tested. (parseCitationTag already exists.)
- Remove the vestigial `tag` field from `buildBibliographyPayload` (unused since SP-A wired
  `buildBibliographyTag`); update its test.

### `cite-pane.js` — the refresh orchestration (Office.js)
- **`refreshCitations()`** (new, wired to a `#cite-refresh` button). In one `Word.run`:
  1. `ensureAssets(currentStyleFile())` first (so `cache.styles`/locales/items are loaded), build
     one engine at the current style+locale over `cache.items`.
  2. Collect candidate content controls: `document.contentControls` (main body) **plus** each
     footnote/endnote body's `contentControls` (via the confirmed footnotes API), loaded with
     `tag`. Guard/try each footnote-enumeration step so an unsupported build just yields the
     main-body set.
  3. For each CC: `parseCitationTag(cc.tag)`:
     - `AshaarCite:` → `items = citationItemsFromTag(parsed)`; if any id ∉ `cache.items` → count
       unresolved, skip. Else re-render (Arabic OOXML vs LTR HTML, by the *current* locale) and
       `cc.insertOoxml/insertHtml(..., replace)`; set `cc.tag = buildCitationTag({style,locale,items})`.
     - `AshaarBib:` → rebuild from `cache.items` at current style/locale; rewrite body + `cc.tag = buildBibliographyTag({style,locale})`.
  4. `ctx.sync()` (batched). Report: "Refreshed N citation(s), K bibliograph(ies)" (+ "; U
     unresolved — re-add from Zotero" when U>0; + note if 0 footnote citations were reachable).
- Reuse the render helpers already in the pane (the Arabic-OOXML build + `readDocCsFont`, the LTR
  `wrapRtlRuns(sanitize(...))`), factored so `insertCitation` and `refreshCitations` share them
  rather than duplicating.
- Guard `typeof Word`/`CiteWord`; browser preview → the button reports "Refresh needs Word."

### `taskpane.html`
- Add a `#cite-refresh` button near the insert actions; add nothing else. Bump `ASHAAR_ASSET_VERSION`.

## Data flow
```
"Refresh citations" (current style/locale in the pane)
  → ensureAssets + build engine over cache.items @ current style/locale
  → enumerate CCs: document.contentControls + footnote/endnote bodies' contentControls
  → per CC: parseCitationTag →
       AshaarCite: → citationItemsFromTag → engine.cite(items) → rewrite body (OOXML if ar) + rewrite tag
       AshaarBib:  → engine.bibliography() from cache.items → rewrite body + rewrite tag
  → report counts (refreshed / bibliographies / unresolved / footnote-reachability)
```

## Error handling
- Unreadable/foreign CC tag (`parseCitationTag` → null) → skipped (not ours).
- Citation with an id missing from `cache.items` → counted "unresolved", skipped, surfaced.
- Footnote enumeration unsupported / empty on this Word build → refreshes main-body + bibliography,
  reports "0 footnote citations reached" (graceful; the live-verify gate).
- Any per-CC rewrite failure → caught, counted as failed, other CCs still processed; final status
  reports failures without throwing the whole run.
- No Word (browser) → "Refresh needs Word."

## Testing (node-`assert`, added to `npm test`)
- `cite-word.test.js` — `citationItemsFromTag`: a parsed tag with keys `[{id:"A",locator:"42",label:"page"},{id:"B",locator:null,label:null}]` → `[{id:"A",locator:"42",label:"page"},{id:"B"}]`; empty/missing keys → `[]`. Round-trips with `buildCitationTag`/`parseCitationTag`. `buildBibliographyPayload` no longer carries `tag`.
- The refresh orchestration is Office.js (not node-testable) → the **manual Word checklist** is the gate:
  insert a couple of citations (Arabic + en, footnote + inline) + a bibliography → change the style
  (e.g. Chicago → APA) → **Refresh** → all update in place to the new style, footnotes included,
  locators preserved; the report counts match; an unresolved id is reported not crashed.

## Success criteria
With citations + a bibliography in the document, changing the pane's style/locale and clicking
**Refresh citations** re-formats every reachable `AshaarCite:`/`AshaarBib:` control in place
(footnotes included) to the new style, preserving keys + locators, and reports accurate counts;
missing-reference and non-Ashaar controls are handled without error. `npm test` passes including
the `citationItemsFromTag` tests. (Footnote reachability is confirmed on the first live Word run.)
