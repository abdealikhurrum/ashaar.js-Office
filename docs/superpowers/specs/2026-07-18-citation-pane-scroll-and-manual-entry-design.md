# Citation Pane Scrolling + Manual Entry — Design

**Date:** 2026-07-18
**Branch:** `feat/citation`
**Context:** Two citation-UI improvements requested after SP-4 shipped: the task pane runs out of
vertical space with no scroll, and there's no way to add a reference without Zotero.

## Summary

Two independent, small-to-medium features:

1. **Pane scrolling** — make the task pane scroll vertically (CSS only) so long content is reachable.
2. **Manual citation entry** — an inline "Add manually" form that builds a CSL-JSON item and adds it
   to the reference set, no Zotero required.

## Feature 1: Pane scrolling

**Problem:** `.pane` is a `display:flex; flex-direction:column; min-height:100vh` container with no
outer scroll region, so when a tab's content (notably the Cite tab) exceeds the task-pane iframe
height, the overflow is unreachable.

**Fix (CSS only, no JS):** allow the pane to scroll vertically. Set the scroll on the pane/body so
the whole tab content scrolls within the task-pane iframe:

- The pane container: keep `min-height:100vh` (so short content still fills the pane) but ensure the
  body/pane allows vertical overflow to scroll rather than clip — `overflow-y: auto` on the
  scrolling container, `overflow-x: hidden` to avoid a horizontal bar.
- Verify the fix doesn't break the existing inner scroll regions (e.g. the preview area already has
  `overflow:auto`) — the outer scroll is additive.

**Scope:** applies to all tabs (Table Input, Convert, Cite, Styles), not just Cite. No behavior
change beyond scrollability.

## Feature 2: Manual citation entry

**UI:** an **"Add manually"** button beside "Add from Zotero" in the Cite tab. It toggles an inline
form (not an Office dialog) within the tab.

**Form — Type selector + per-type fields:**

| Field | Book | Book chapter | Journal article | Webpage |
|-------|------|--------------|-----------------|---------|
| Title | ✓ | ✓ (chapter title) | ✓ (article title) | ✓ |
| Author(s) | ✓ | ✓ | ✓ | ✓ |
| Year | ✓ | ✓ | ✓ | ✓ |
| Publisher | ✓ | ✓ | | |
| Place | ✓ | ✓ | | |
| Container title | | ✓ (book title) | ✓ (journal) | |
| Editor(s) | | ✓ | | |
| Volume | | | ✓ | |
| Issue | | | ✓ | |
| Pages | | ✓ | ✓ | |
| URL | | | | ✓ |
| Accessed | | | | ✓ |

CSL types: Book → `book`, Book chapter → `chapter`, Journal article → `article-journal`,
Webpage → `webpage`.

**Author/Editor entry:** a textarea, **one name per line** as `Family, Given`. A line with **no
comma** is treated as an organization / single-name literal (CSL `{ literal: "…" }`). Parsed into
CSL name objects (`{ family, given }` or `{ literal }`). Empty lines ignored.

**Item build:** assemble a CSL-JSON item:
- `id`: generated, `"manual-<n>"` where `<n>` is a monotonic counter seeded to avoid collisions with
  existing `cache.items` ids.
- `type`: per the selector.
- Fields mapped: title→`title`, container→`container-title`, year→`issued: {date-parts:[[year]]}`,
  publisher→`publisher`, place→`publisher-place`, volume→`volume`, issue→`issue`, pages→`page`,
  url→`URL`, accessed→`accessed: {date-parts:[[…]]}`, authors→`author`, editors→`editor`.
- Only non-empty fields are included.

**Integration:**
- On submit: add the item to `cache.items`, `persistRefs()` (same doc-settings store Zotero items
  use), re-render the Items list (`populateItems`), check the new item, re-render the preview.
- The item is then citable and refreshable exactly like any other — no special-casing downstream.

**Multilingual variants:** out of scope for v1. Fields accept whatever the user types (Arabic or
Latin); Arabic renders correctly via the content-aware font path. No romanized-variant fields.

**Editing:** add-only for v1. Mistakes are fixed by removing (existing `×`) and re-adding.

**New module:** `src/taskpane/cite-manual.js` (UMD, Node-testable):
- `parseNames(text)` → `[{family, given} | {literal}]` (pure).
- `buildManualItem(values)` → CSL-JSON item (pure). `values` is a plain object of the form fields
  plus `type` and an `id`. Only non-empty fields emitted.
- `cite-pane.js` owns the DOM (form markup toggle, gathering field values, generating the id, calling
  `buildManualItem`, wiring into `cache.items`/persist/populate/preview).

Keeping the field→CSL mapping and name parsing in a pure module makes them unit-testable and keeps
the DOM code thin.

## Testing

- **`tests/cite-manual.test.js`** (Node assert, pure): `parseNames` (family+given, literal/no-comma,
  blank-line skipping, multiple); `buildManualItem` per type (correct CSL `type`, field mapping,
  year→date-parts, empty-field omission, author/editor arrays).
- **Manual (browser/Word):** Add manually → item appears in list, checked → preview renders →
  Insert citation + bibliography work → persists across the document (save/reopen) → an Arabic-title
  manual item renders in the correct font.

## Out of scope / deferred

- Pop-out web (Office Dialog API) window — rejected in favor of in-pane scrolling.
- Editing existing items in place; romanized/multilingual variant fields; full CSL type/field
  coverage (editors' roles, translators, edition, series, DOI/ISBN, etc.).
