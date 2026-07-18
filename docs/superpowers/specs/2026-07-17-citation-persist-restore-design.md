# SP-B — Persist & restore the citation reference set (design)

**Date:** 2026-07-17
**Status:** Approved design (decomposed from the 2026-07-17 persistence brainstorm), pre-implementation
**Track:** Second of the "document-embedded, re-editable citations" initiative. Follows SP-A
(locators + editing + tagging). SP-C (refresh/re-format) is separate and still deferred.

## Summary

Persist the Cite tab's working **reference set** (the `cache.items` id-keyed CSL-JSON map) in the
document so it survives save / close / reopen. Today `cache.items` is in-memory only, seeded from
the bundled `fixtures/cite-sample.json`; every reopen loses any Zotero-added items and `×`
removals. SP-B stores the reference set via **`Office.context.document.settings`** (per-add-in,
per-document) and restores it on pane load.

## Grounding (verified via MS Learn, 2026-07-17)

`Office.context.document.settings` — **Applications: Excel, PowerPoint, Word.** `set(name,value)`
/ `get(name)` / `remove(name)` operate on an in-memory property bag (values may be objects/strings,
serialized as JSON); `saveAsync(cb)` persists the bag into the document (the file itself is written
when the user next saves the doc). Settings auto-load with the `Document` object on open. `get`
returns `null` when absent. This is the standard "persist add-in state per document" mechanism —
no shared runtime required.

## Decisions

- **Storage:** one setting, key `AshaarCiteRefs`, holding a JSON string `{"v":1,"items":{…}}`
  (version-tagged for future migration). Not custom XML parts (heavier); settings is sufficient
  for a reference set of tens of items.
- **Restore-vs-fixture:** on pane load, if the document has a saved reference set, use it; **else**
  fall back to the bundled fixture (so a fresh document / bare browser still demos). The fixture is
  no longer force-loaded when saved refs exist.
- **Save triggers:** after the two mutations of `cache.items` — `addFromZotero` (merge) and
  `removeItem` (delete). Fire-and-forget with error surfaced to the pane status.
- **Scope:** only the reference set (the Items list / bibliography source). Inserted citations and
  their `AshaarCite:` tags are already in the document (SP-A); reading those tags back to
  reconstruct/refresh is **SP-C**, not SP-B.

## Architecture & components

### `src/taskpane/cite-store.js` (new, UMD, Node-testable)

Talks only to a `settings` object (default `Office.context.document.settings`), injected for tests.

- `REFS_KEY` = `"AshaarCiteRefs"`.
- **pure** `serializeRefs(items)` → `JSON.stringify({ v: 1, items: items || {} })`.
- **pure** `parseRefs(str)` → items map. `null`/`""`/malformed/`v !== 1`/missing `items` → `{}`.
  Never throws.
- `saveRefs(items, settingsImpl?)` → `Promise`. If no settings available (browser preview),
  resolve as a no-op. Else `settingsImpl.set(REFS_KEY, serializeRefs(items))` then
  `settingsImpl.saveAsync(cb)`; resolve on `Office.AsyncResultStatus.Succeeded`, reject with the
  error message otherwise. (Status compared defensively by value so tests need no Office global.)
- `loadRefs(settingsImpl?)` → `Promise<items map>`. If no settings, resolve `{}`. Else
  `parseRefs(settingsImpl.get(REFS_KEY))`. Never rejects (resolve `{}` on any error).
- `resolveSettings()` helper: returns `Office.context.document.settings` when the full chain
  exists, else `null` (so the pane and tests both work without Office).

### `src/taskpane/cite-pane.js` (wiring)

- **Restore on load:** in `ensureAssets`, replace the unconditional fixture load with: when
  `cache.items` is not yet set, first `CiteStore.loadRefs()`; if it returns a non-empty map, use it
  as `cache.items`; otherwise fetch the fixture as today. (Guard `typeof CiteStore === "undefined"`
  → fixture.)
- **Save on mutate:** at the end of `addFromZotero`'s merge and in `removeItem`, call
  `CiteStore.saveRefs(cache.items)` (fire-and-forget; `.catch` → non-fatal status hint, never
  blocks the UI). Guard `typeof CiteStore === "undefined"`.
- No behavior change when Office/settings is absent (browser preview): loadRefs → `{}` → fixture;
  saveRefs → no-op.

### `taskpane.html`

- Add `"./cite-store.js"` to `srcs` **before** `"./cite-pane.js"`; bump `ASHAAR_ASSET_VERSION`.

## Data flow

```
Pane load → ensureAssets → CiteStore.loadRefs() (settings.get AshaarCiteRefs → parseRefs)
   → non-empty? use as cache.items : fetch fixture
Add from Zotero / × remove → mutate cache.items → CiteStore.saveRefs(cache.items)
   (settings.set + saveAsync → persists into the doc; written to file on next doc save)
Reopen doc → settings auto-loaded → loadRefs restores the same reference set
```

## Error handling

- `saveRefs` failure (async status Failed) → pane status hint "Couldn't save your reference list to
  the document."; never throws into the UI flow.
- `loadRefs` any failure / absent / malformed → `{}` → fixture fallback (tab still works).
- Office/settings unavailable (browser) → load `{}` (fixture), save no-op — preview unaffected.

## Testing (node-`assert`, added to `npm test`)

- `tests/cite-store.test.js` with a fake settings object (`{ _bag, set, get, remove, saveAsync }`):
  - `serializeRefs`/`parseRefs` round-trip an items map; `parseRefs` returns `{}` for `null`, `""`,
    non-JSON, `{v:2,…}`, and JSON lacking `items`.
  - `saveRefs` calls `set(REFS_KEY, <string>)` + `saveAsync`, resolves on a Succeeded fake, rejects
    on a Failed fake.
  - `loadRefs` returns the stored map after a prior `saveRefs`; returns `{}` when the key is absent.
  - With no settings (pass `null` / omit): `saveRefs` resolves no-op, `loadRefs` resolves `{}`.
- Manual Word checklist: add Zotero items, `×` one, **save + close + reopen** the document → the
  Items list is restored (not the fixture); a brand-new document shows the fixture.

## Success criteria

After adding/removing references and saving the document, closing and reopening it restores the
same Items list + bibliography from the document (no Zotero round-trip needed to see them). A fresh
document or a bare browser still shows the fixture. `npm test` passes including the new tests.
