# Test Plan — Bring in Existing Poems (Adopt Table + Separators)

Branch: `feature/adopt-existing-table`. Reload the task pane (or re-sideload) so the
new JS loads — Office caches aggressively; the cache-buster is `?v=20260530-adopt`.

## Already automated (no manual action needed)

`npm test` covers the pure cores: `tests/table-adopt.test.js` (reconstruction) and
`tests/separators.test.js` (detection + normalization), plus the existing suites.
The items below are what those tests **can't** cover — Word behavior and UI wiring.

---

## A. Phase 1 — Adopt Existing Table  *(needs Word)*

1. **Plain couplet table (2 columns).** Cursor inside → Adopt Existing Table → table is
   replaced by a managed block; sadr on the right, ajuz on the left; gap looks right.
2. **Content control.** After adopt, click inside the new block → **Justify Selected Text**
   works (confirms it was wrapped as an "Ashaar Poem" content control).
3. **Mixed table** (full-width/merged solo rows + 2-col couplet rows) → solos reconstruct as
   centered single lines, couplets split correctly.
4. **Multi-misra table (3+ columns)** → reconstructs as a multi-misra row.
5. **Already-justified source** (cells contain kashidas) → adopted text is clean (no leftover
   tatweels), then re-justified per current settings.
6. **Gap/empty columns** (including a table produced by this add-in's own generator, which has
   interleaved gap columns) → gaps dropped; no empty/phantom misras.
7. **RTL correctness** → sadr lands on the right. If reversed, set **Reading direction = LTR**
   and re-run; confirm it flips.
8. **Review before replacing = ON** → recovered text loads in the editor + preview, and the
   **table is NOT changed** until you press **Replace Selection**.
9. **One-click default** → table is replaced immediately; **Ctrl+Z** restores the original table.
10. **Cursor not in a table** → friendly message "Place the cursor inside a table to adopt it."
11. **Scope = All tables in selection** → multiple selected tables combine into stanza-separated
    blocks; confirm result and that surrounding non-table content isn't lost unexpectedly.
12. **Multi-paragraph cell** → internal line breaks collapse to a single space.
13. **Empty/whitespace-only table** → message "didn't contain any text to adopt."
14. **Settings respected** → adopted block uses the current Font / Justification / Kashida
    strength / Middle gap / layout selections.

## B. Phase 2 — Separator flexibility / conversion  *(mostly in Word; some browser-checkable)*

15. **Paste dash-separated poem** (`بیت - بیت`) into the editor → auto-converts to `\`, note
    reads "Converted separators (dash)…", preview correct.
16. **Load Selection** of dash / tab / asterisk / wide-gap text from the document → converts;
    note names the detected separator.
17. **Tab-separated** paste → converts.
18. **Asterisk** (`m1 * m2`) → converts to `\` and inserts correctly.
19. **Wide gap (2+ spaces)** → converts; but **normal single-spaced prose is NOT mangled**.
20. **Hyphenated words** (`well-known`) → NOT split.
21. **Solo-only text** (no separators anywhere) → unchanged; inserts as solo lines.
22. **Manual override**: set *Split hemistichs on = Dash* → Apply → forces dash splitting.
23. **Custom separator**: choose *Custom…*, type e.g. `//` → Apply → splits on it (literal, not
    regex).
24. **Pair every 2 lines**: a file with one hemistich per line → tick the box → Apply → couplets;
    blank line = stanza break respected; an odd trailing line stays solo.
25. **Already-canonical `\`** text → no change and no spurious note.
26. **Structure preserved** → blank lines (stanzas) and `---` (poem breaks) survive conversion;
    insert produces multiple stanzas/poems correctly.
27. **After conversion** → Insert as Table / Insert as Paragraphs / Replace Selection all work.

## C. Regression — recent fixes this work touches  *(needs Word)*

28. **Stanza-grid independence**: insert a stanza of several solo lines + one couplet → the
    couplet's gap does NOT move when a solo line is long (the fixed-layout fix). Adopt produces
    exactly this shape, so verify on an adopted block too.
29. **Re-justification**: Justify a poem, then narrow the column or change the font, and Justify
    again → kashidas **reduce** (don't compound); justify to a wide then tight column returns to
    near-bare.
30. **Core insert paths** unaffected: Insert as Table, Insert as Paragraphs, Replace Selection.
31. **Templates** still work: Capture from Word, Apply, Export/Import JSON, Drop Grid.

## D. Cross-cutting

32. **No console errors** in the task pane (open DevTools on the pane).
33. **Fonts**: Document / Arabic-serif / Nastaliq each render and adopt/convert correctly.
34. **Undo** behaves for every document-changing action (Adopt, Insert, Replace, Justify).
35. **Cache**: confirm the new build actually loaded (Adopt Existing Table button + Import
    options panel are visible).

## Known limitations to confirm-as-expected (not bugs)

- Nested tables: adopt targets the outer table only.
- "All tables in selection" replaces the whole selection range with one block.
- Reading direction is default-RTL with manual override (no auto-detection from alignment).
- `pairLines` re-applied to already-paired text will re-pair — it's a one-shot import tool.
