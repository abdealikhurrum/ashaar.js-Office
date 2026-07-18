# SP-4 Source Classification — Manual Word Checklist

Prereqs: Zotero running with Better BibTeX; a few library items tagged with
`corpus:fatemi` and/or `class:secondary`; the rest untagged (default primary/non-Fatemi).

1. [ ] **Live tag fetch.** With a `-fatemi` style selected, add several items from Zotero
   spanning ≥2 buckets. Insert bibliography. Confirm it splits into headed subsections in the
   fixed order Primary·Fatemi → Primary·Other → Secondary·Fatemi → Secondary·Other, and empty
   buckets are skipped.
   - If headings do NOT appear: check the console — the BBT translator name may differ from
     `"BetterBibTeX JSON"`. Confirm the live name (Zotero → File → Export → translator list) and,
     if needed, update `buildTagsRequest` + `parseTagsResult` (+ their tests) or fall back to the
     Zotero local API path (see the design's feasibility note).
2. [ ] **Collapse rule.** Select items that all fall in ONE bucket (e.g. all untagged). Insert
   bibliography under a `-fatemi` style → confirm a single flat list with NO heading (identical
   to a stock-style bibliography).
3. [ ] **Stock style = flat.** Switch to Chicago (notes & bibliography) or APA. Insert
   bibliography → confirm one flat list, no headings, no tag fetch behavior change.
4. [ ] **Arabic RTL headings.** Set locale = ar, select a `-fatemi` style, insert a multi-bucket
   bibliography → confirm Arabic headings render bold, right-to-left, no tofu, and entries follow
   under each (Arabic titles upright, not italic).
5. [ ] **Refresh reproduces sections.** With a sectioned bibliography inserted, change nothing and
   click "Refresh citations" → confirm the bibliography CC re-renders with the same sections.
6. [ ] **Refresh collapses on style switch.** Switch from a `-fatemi` style to a stock style, click
   "Refresh citations" → confirm the bibliography collapses to one flat list in place.
7. [ ] **Tag-fetch failure degrades to flat.** Quit Zotero (or block the proxy), select a
   `-fatemi` style, insert bibliography → confirm it still inserts as a flat list (no crash, a
   status message is acceptable).
8. [ ] **Save/close/reopen.** Save the doc, close, reopen → confirm the sectioned bibliography
   text persists (it is static once inserted) and a subsequent Refresh still works.
