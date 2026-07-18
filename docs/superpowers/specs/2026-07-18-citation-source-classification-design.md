# SP-4 — Source Classification & Nested-Section Bibliography — Design

**Date:** 2026-07-18
**Branch:** `feat/citation`
**Roadmap slot:** SP-4 (from `2026-07-16-citation-engine-core-design.md`): *"Source classification &
differentiated output: primary/secondary and Fatemi/non-Fatemi metadata → grouped/segmented
bibliographies with distinct styling."*

## Summary

Classify each reference by two independent axes read from **Zotero tags**, and render the
bibliography segmented into **nested headed subsections** (up to four). Selecting a `-fatemi` style
is the toggle that turns this on; stock styles render one flat bibliography exactly as today.

**Explicitly in scope:** tag-based classification, per-bucket bibliography rendering, localized
RTL-aware section headings, collapse-to-flat behavior, Refresh (SP-C) parity.

**Explicitly out of scope (decided during brainstorming):**
- Per-entry Fatemi styling (honorifics, distinct font/emphasis). Fatemi = subsection membership only.
- "Drop italic at the CSL level." The existing OOXML-insertion italic suppression stays as-is.
- Any change to inline / footnote / endnote citations. Classification is **bibliography-only**.

## Classification model

Two **independent** axes, each carried by a **prefixed** Zotero tag (prefixes chosen to avoid
collision with other library tags):

| Axis    | Tag present            | Meaning   | Absent (default) |
|---------|------------------------|-----------|------------------|
| Corpus  | `corpus:fatemi`        | Fatemi    | non-Fatemi       |
| Class   | `class:secondary`      | secondary | primary          |

- `class:primary` may also appear explicitly; it is equivalent to the default.
- An untagged item → **Primary / non-Fatemi**.
- Four buckets result: `primary·fatemi`, `primary·other`, `secondary·fatemi`, `secondary·other`.
- **Malformed / unknown values** (e.g. `class:tertiary`, unrelated tags) → ignored; the axis falls
  back to its default. Only the two prefixes are ever consulted.

## Architecture & data flow

```
citekeys ─┬─ /zotero/json-rpc  item.export "Better CSL JSON"   → CSL-JSON item map (as today)
          └─ /zotero/json-rpc  item.export <tag-carrying fmt>  → { citekey: [tags] }   (NEW)
                                                                          │
              cite-classify.js (NEW, pure UMD): tags → { corpus, class } bucket per citekey
                                                                          │
   cite-word / cite-pane: if the selected style id ends in a -fatemi variant →
       for each NON-EMPTY bucket, in fixed order:
           run the engine over that bucket's items → bibliography HTML
           prepend the localized heading
       concatenate the sections → insert (RTL path for ar, insertHtml for en)
```

Stock (non-`-fatemi`) style selected → single flat bibliography via the existing path, unchanged.

### Tag retrieval (feasibility)

"Better CSL JSON" (the translator the engine already fetches) drops Zotero tags. Recommended path:
a **second `item.export` over the same citekeys** through a tag-carrying translator (BetterBibTeX
JSON, or BibLaTeX `keywords`), parsed per citekey. This reuses the existing `/zotero/json-rpc`
proxy route — **no new proxy plumbing**. Fallback if that proves unworkable: the Zotero **local API**
(`localhost:23119/api/users/0/items`, already used read-only by the SP-3 migration tool), correlating
item keys to citekeys. The exact translator is confirmed live during implementation.

### Module boundaries

- **`src/taskpane/cite-classify.js`** (new, pure UMD, Node-testable): the only new module.
  - `bucketForTags(tags)` → `{ corpus: "fatemi"|"other", class: "primary"|"secondary", key }`.
  - `bucketItems(itemMapOrKeys, tagsByCitekey)` → ordered list of non-empty buckets, each
    `{ key, corpus, class, citekeys: [...] }`, in the fixed order below.
  - No I/O; consumes an already-fetched `{citekey: [tags]}` map.
- **`cite-zotero.js`**: gains a `fetchTags(citekeys)` I/O helper (mirrors `fetchCslJson`), talking
  only to `/zotero/json-rpc`.
- **`cite-word.js` / `cite-pane.js`**: the sectioning orchestration — detect `-fatemi` style, drive
  per-bucket engine runs, assemble headings + sections, route to the existing RTL/insertHtml paths.
- The inert `[genre]` branch in `src/styles/chicago-notes-fatemi.csl` and `src/styles/apa-fatemi.csl`
  is **removed** (and re-synced to `src/vendor/csl-styles/`), since classification no longer travels
  through a CSL variable. The two `-fatemi` styles otherwise remain byte-identical to their stock
  parents and serve purely as the toggle.

## Section headings, order & rendering

- **Fixed bucket order:** Primary·Fatemi → Primary·Other → Secondary·Fatemi → Secondary·Other.
- **Headings localized by the pane's display locale:**
  - en: `Primary Sources — Fatemi`, `Primary Sources — Other`, `Secondary Sources — Fatemi`,
    `Secondary Sources — Other`.
  - ar (initial wording — tweak during implementation):
    `المصادر الأساسية — الفاطمية`, `المصادر الأساسية — أخرى`,
    `المصادر الثانوية — الفاطمية`, `المصادر الثانوية — أخرى`.
- **Heading rendering:** a distinct **bold run** so it reads as a subheading without relying on Word
  paragraph styles. Arabic headings + Arabic bibliography sections insert via the existing OOXML RTL
  path (`<w:p><w:bidi/>` + Arabic runs, italic suppressed); en headings + sections via `insertHtml`.
- **Collapse rule:** if exactly **one** bucket is non-empty, render it **flat with no heading** —
  byte-for-byte the same as today's output. Headings appear only when **≥2** buckets are populated.

## SP-C (Refresh) interaction

The "Refresh citations" flow rewrites the `AshaarBib:` content control in place. It must call the
**same section-building path**, so:
- Refresh under a `-fatemi` style reproduces the nested sections.
- Refresh after switching to a stock style collapses back to flat.

The `AshaarBib:` tag already records the style, which tells refresh which path to take — **no new tag
fields required.** (The bibliography is re-derived from the current reference set + style at refresh
time, as it is today.)

## Error handling

- **Tag fetch fails / returns nothing** → every item defaults to Primary·non-Fatemi → a single bucket
  → collapse rule → **flat bibliography**. Classification is best-effort and never blocks insertion.
- **Malformed axis value** → treated as absent on that axis (default applies).
- **Unknown extra tags** → ignored.
- A citekey present in the CSL-JSON map but missing from the tags map → default bucket.

## Testing

- **`tests/cite-classify.test.js`** (Node assert, pure): tag arrays → buckets; defaults for
  untagged; `class:primary` ≡ default; malformed values fall back; prefix isolation (unrelated tags
  ignored); ordering of non-empty buckets; collapse to one bucket.
- **Sectioning tests**: given a bucketed item map, per-bucket engine runs yield the right ordered,
  non-empty sections; empty buckets skipped; single bucket → flat/no-heading.
- **Manual Word checklist** (new spec file): tag items in Zotero (`corpus:fatemi`, `class:secondary`),
  select Chicago (Fatemi), insert bibliography → verify nested sections, order, and RTL headings under
  ar locale; Refresh reproduces sections; switching to a stock style collapses to flat; tag-fetch
  failure degrades to flat.

## YAGNI / deferred

- Configurable heading text or per-bucket style overrides — fixed strings for now.
- A pane control decoupling grouping from style (rejected in favor of "Fatemi style = the toggle").
- Any third axis or "Unclassified" catch-all section (defaults absorb untagged items).
