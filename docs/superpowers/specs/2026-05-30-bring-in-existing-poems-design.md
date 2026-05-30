# Bring in Existing Poems — Design

**Date:** 2026-05-30
**Status:** Approved (pending spec review)

## Problem

Most users arrive with poems already in their Word document, in one of two shapes:

1. **Marked-up / loosely-delimited text** — e.g. `misra1 * misra2`, or hemistichs split
   by a dash, tab, or wide gap.
2. **An existing Word table** — hand-built or pasted from elsewhere.

Today the text path works only for the exact separators `\ * |`, and there is no way to
turn an existing table into a managed, re-justifiable Ashaar block. The objective is
**flexibility in separators and ease of conversion**: something that "just works" by
default, with advanced options when it doesn't.

## Scope & sequencing

Two related but independent subsystems, built in this order:

- **Phase 1 — Adopt Existing Table** (this spec's primary focus). Word-dependent; the new
  logic is a pure, unit-testable reconstruction function plus an Office.js adapter.
- **Phase 2 — Text/separator flexibility** (sketched here, detailed when reached). Auto-detect
  the hemistich separator on import and visibly normalize, with advanced overrides.

Both share one editor → preview → Insert/Replace pipeline.

## Chosen approach (Phase 1): Reconstruct to canonical source, then regenerate

Adopt reads a table's cells **in RTL reading order**, reconstructs the **canonical Ashaar
source** (`sadr \ ajuz`, solos bare, one table = one stanza), and feeds that source through
the **existing** `renderForWordOoxml` pipeline (fixed-grid layout, content-control wrapping,
justification, current pane options), replacing the original table in place.

Rationale: the only genuinely new logic is a **pure function** `tableModel → canonicalSource`,
fully testable in Node. Everything downstream is already built and tested. The recovered
source is transparent (shown in the editor, editable to fix mis-reads), and it sets up Phase 2
to reuse the same flow.

Rejected alternative — *direct OOXML remap* (build the normalized table directly from cells,
no source-text intermediate): preserves exact cell text but gives no editable recovered text,
duplicates mapping logic, is harder to test (no pure intermediate), and does not unify with the
text path.

---

## Phase 1 — Adopt Existing Table

### 1. User-facing behavior

A new **"Adopt Table"** action. With the cursor inside (or a selection over) a table of poetry:

- **Default (one-click):** read the table → reconstruct canonical source → load it into the
  editor + preview → **replace the table in place** with a clean, content-controlled,
  justifiable Ashaar block using the current pane settings (layout/font/justify/gap).
- **Safety net (honors "advanced if it doesn't"):**
  1. the recovered source stays in the editor — fix a mis-read by editing + Replace;
  2. native **Ctrl+Z** reverts the swap;
  3. an **Advanced → "Review before replacing"** toggle stops after preview so the user
     commits manually.

### 2. Reconstruction engine — the testable core

New UMD module **`src/taskpane/table-adopt.js`** exporting a pure function:

```
adoptTableToSource(rows, opts) -> canonicalSourceString
```

- `rows`: array of rows; each row an array of **cell text strings in visual right-to-left
  order** (sadr first). Produced by the Office.js adapter (§4); the function itself is
  Word-free and fully unit-testable.
- Per cell, before classifying:
  - **strip justification artifacts** — U+0640 tatweel, U+200A hair space, U+2009 thin space
    (same rule used elsewhere), so an already-justified source reconstructs cleanly;
  - collapse internal newlines to a single space; trim.
- Per row, by count of non-empty cells **K**:
  - `K = 0` → skip (e.g. an all-gap row);
  - `K = 1` → solo misra → emit the text as a bare line;
  - `K = 2` → couplet → emit `sadr \ ajuz` (sadr = first cell in RTL order);
  - `K ≥ 3` → multi-misra row → emit `m1 \ m2 \ m3 …` (the parser treats 3+ on one line as a
    multi-misra `row`).
- One table → one stanza; lines joined by `\n`.
- `opts.direction`: `"rtl"` (default) or `"ltr"`; `"ltr"` reverses each row before classifying.

### 3. Invocation & UI

- New button **"Adopt Table"** in Table Input mode, near *Capture from Word*.
- An **Advanced** disclosure (collapsed by default) containing:
  - **Review before replacing** (checkbox; default off → one-click replace);
  - **Reading direction** (Auto / RTL / LTR; default Auto → treat as RTL);
  - **Scope** (Table at cursor [default] / All tables in selection).

### 4. Office.js adapter & in-place replace (`taskpane.js`)

New `adoptTable()`:

1. Find target table(s): the cursor's enclosing table, or tables in the selection (mirrors
   `justifySelection`'s table-loading pattern).
2. Load rows → cells → `cell.body.text`; build a row/cell-text model; order each row per
   direction (default RTL → `cells.items[0]` = sadr).
3. Call `adoptTableToSource(model, {direction})` → source; set `input.value = source`;
   `renderPreview()`.
4. **Default:** select the table's range (`table.getRange().select()`) and invoke the
   **existing** replace path (`insertPoem(true)`), so generation reuses the tested OOXML path,
   the fixed-grid fix, justification, and the **Ashaar Poem content-control wrapping** for free.
5. **Review mode:** stop after step 3; the user presses *Replace Selection* to commit.

Net new Word code = table-reading + range selection. Generation/replace is reused.

### 5. Data flow

```
Word table
  -> [adapter] rows-of-cell-text (RTL)
  -> adoptTableToSource()
  -> canonical source
  -> editor + preview
  -> (existing) parse -> renderForWordOoxml -> Ashaar Poem content control
  -> replace table range
```

### 6. Error handling & edge cases

- Non-table selection → friendly message ("Place the cursor in a table to adopt it").
- Gap / empty / whitespace-only cells → ignored; all-empty rows → skipped.
- Merged full-width cell → K=1 → solo. Multi-paragraph cell → newlines collapsed to spaces.
- Already-justified cells → tatweels / micro-spaces stripped on read → clean reconstruction.
- Wrong RTL guess → Advanced direction override; recovered text is editable as a fallback.
- Nested tables → out of scope (documented); adopt the outer table only.
- Multi-table scope → each table becomes its own stanza/block; default scope is the single
  table at the cursor to keep the one-click action predictable.

### 7. Testing

- **Unit (Node, new `tests/table-adopt.test.js`)** — fixtures and asserted canonical output for:
  couplet (2-col), multi-misra (3-col), solo rows, mixed solo + couplet, gap columns,
  tatweel-laden cells (assert stripped), RTL vs LTR ordering, empty rows.
- **Manual Word checklist** — adopt a real couplet table → managed block; the review-mode path;
  Ctrl+Z revert; multi-table selection.

### 8. Module boundaries / files

- **New:** `src/taskpane/table-adopt.js` (pure `adoptTableToSource`, UMD, exported for tests);
  `tests/table-adopt.test.js`.
- **Edit:** `src/taskpane/taskpane.js` (`adoptTable()` + button / advanced wiring);
  `src/taskpane/taskpane.html` (button + advanced panel + load `table-adopt.js` with
  cache-buster); `src/taskpane/taskpane.css` (disclosure styling).
- `word-html.js` is untouched. Pure logic stays isolated and testable; Office.js stays in
  `taskpane.js`; OOXML generation stays in `word-html.js`.

---

## Phase 2 — Text / separator flexibility (sketch)

Detailed in a later iteration of this spec. Intent:

- On **Load Selection / paste**, auto-detect the dominant hemistich separator across the
  block. Candidates: existing `\ * |`, spaced dash (` - `), tab, em/en-dash, runs of 2+ spaces.
  Pick the candidate that splits the most lines into a consistent number of parts; require a
  confidence threshold to avoid false positives (e.g. a hyphen inside a word).
- When confident, **visibly normalize** the editor text to canonical `\` form and show a small
  note of what was detected. When not confident, leave text unchanged and surface the advanced
  panel.
- **Advanced:** "Split hemistichs on…" picker (Auto / `\` / `*` / `|` / dash / tab / 2+ spaces /
  custom), and a "Pair every 2 lines" option for files that store one hemistich per line with no
  separator.
- Reuses the same editor → preview → Insert/Replace pipeline; produces canonical text that the
  existing parser already understands.

## Open questions / deferred

- Reading-direction detection heuristic for Phase 1 is "default RTL + manual override"; a
  later refinement could infer direction from cell paragraph alignment.
- Batch adoption across many tables / whole document is available via the Scope option but is
  not the default.
