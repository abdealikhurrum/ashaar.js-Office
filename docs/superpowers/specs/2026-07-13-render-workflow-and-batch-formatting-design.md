# Render workflow & batch formatting — design

**Date:** 2026-07-13
**Status:** Draft for review
**Builds on:** `2026-07-12-unified-settings-design.md` (unified panel, shipped on `feat/guided-justification-ux`)

## Context

Word verification of the unified Settings panel surfaced workflow problems beyond bugs:
gap/cell edits cascade re-justification across every poem sharing a profile; there is no
explicit way to re-render one block after native text/font edits; per-cell and per-gap
formatting can only target one cell at a time; justification artifacts (tatweels,
micro-spaces) left in cells break later font recognition; and justification actively
hampers authors while they are still writing.

## Goals

1. Apply never touches more than the current block; profile-wide refresh only on explicit profile actions.
2. A Re-render button for the current block (text edits, font changes).
3. Batch formatting: apply cell/gap formatting to "this / this bandh / whole poem".
4. Cell-level visual formatting (fill + text color), same UX as gap decor minus symbol.
5. Capture natively-applied Word formatting into the pane → overrides → profile.
6. Every action's refresh cost is visible before it runs (poem rebuild vs re-justify).
7. Font recognition is never corrupted by justification artifacts.
8. Probe/calibration run once per font/poem, not on every Apply.

## Non-goals (explicitly deferred)

- **Light-apply** (in-place decor writes / per-cell re-justify with no whole-poem pass) —
  revisit if the cost labels (§6) reveal poem-level refresh is the pain; it is the natural
  next step after this cycle.
- **Draft mode** (document-level justification deferral) — cut as YAGNI (review decision
  2026-07-13). The existing `justifyMode:"none"` setting is the manual draft workflow:
  write with justification off, switch it on to publish. Revisit only if cost transparency
  plus Re-render still leave writing-flow pain.
- **Auto-capture** of native formatting on cursor move — would fight the provenance model.
- Hemistich markers / glyph styling (the rest of the deferred spacing-cell-styling spec).
- Borders or any Word formatting outside the modeled schema (fill, text color, symbol).

---

## 1. Cascade descope

**Today:** `justifySelection` on a profile-linked block delegates to `applyProfileToQaseeda`,
re-justifying every poem with that profile. A gap decor tweak re-renders the document.

**Design:** panel Apply (all four scopes) and the Re-render button operate on the current
block only. `justifySelection`'s profile delegation is removed from the Apply path; the
block's effective settings still come from the resolver (profile layer included), applied
to this block alone. `applyProfileToQaseeda` (all-blocks refresh) runs only from Assign,
Update, and Restore — the actions whose meaning is profile-wide.

**Expectation change (intended):** sibling poems no longer refresh when one poem is
tweaked; they refresh on Update.

## 2. Re-render button

Panel footer gains **Re-render** beside Apply, wired to the existing `reRender()`:
reconstructs the source from the block's cells (picking up native text edits), preserves
the representative font and size (picking up native font changes), carries the block's own
tag payload (profile/local/profileCache — already fixed), embeds the fresh control via
`wrapOoxmlControl`. Enabled only when the cursor is in a block; justifies per resolved
settings (with `justifyMode:"none"` it rebuilds unjustified — the manual draft workflow).

## 3. Apply-to-all toggle

Cell and Gap scope bodies gain a target selector: **This cell/gap · This bandh · Whole
poem** (default: this). Apply writes the pending override/decor to every matching key —
one `setTagOverride`/`setTagSlotDecor` pass per key into a single tag write, then one
block re-justify. Keys are enumerated from the block's cell map (content cells for cell
scope, spacing cells for gap scope). The per-key values written are the pane's pending
values; keys not covered by the toggle are untouched. ⟲-cleared fields clear on every
targeted key.

## 4. Cell formatting (fill + text color)

`overrides[cellKey]` gains `fill` (hex or "" = none) and `color` (hex or "" = inherit).
Cell scope UI adds the same color+checkbox controls the Gap scope uses (no symbol).
Rendered by the same rebuild path that applies gap decor (cell shading + run color),
normalizing Word's no-color/automatic quirks (known pattern, see
`office-js-word-constraints` memory). Resolver: `fill`/`color` resolve at cell scope only
(no profile/local layer for now — profile-level cell styling arrives with capture, §5,
via spacingDecor-style defaults if needed later; keep v1 to per-cell overrides).

## 5. Capture from Word

Cell and Gap scopes gain **Capture formatting**: reads the cursor cell's actual
`shadingColor`, `font.color`, and (gap only) its text as symbol; normalizes no-color
values; loads them into the pane as pending values (dirty dots). Persistence then flows
through the existing machinery: Apply → block overrides; "Set as default for all bandhs" /
Update → profile. Combined with §3: format one cell natively → Capture → target "whole
poem" → Apply → Update profile.

Capture reads only modeled properties. It never writes; it populates pending.

## 6. Refresh-cost transparency

**Problem:** the user cannot tell whether an action will rebuild the poem's tables
(slow, destructive-and-recreate), re-justify the poem in place (medium), or touch less.
They can't make workflow choices — e.g. batching edits before one Apply — without that.

**Design:** the footer states the blast radius before the click, computed from the
pending buffer and scope:

- Poem scope with structural pending (`gap`, `widthMode`, `widthPct`, `layoutMode`,
  `colWidthMode`) → "Apply — rebuilds poem tables".
- Poem scope, non-structural → "Apply — re-justifies poem".
- Bandh/cell/gap scope → "Apply — re-justifies poem" (v1 honesty: cell/gap applies still
  re-justify the whole block; if this label makes the cost hurt visible, light-apply is
  the deferred fix).
- Re-render button → "Re-render — rebuilds poem tables".

Rendered as a small caption under the footer (updates on every `refreshPanel`), plus the
same text in each button's tooltip. With `justifyMode:"none"` resolved, labels say
"…unjustified" so the manual draft workflow is self-explanatory.

## 7. Strip before font determination

**Problem (observed):** after a native font change to nastaliq, justification artifacts
(tatweels, micro-spaces) still carried the previous font. Per-word font recognition saw
mixed-font runs, failed, and whole words fell back to Arial.

**Rule:** every font-determination pass — per-run real-font dispatch, representative-font
capture, `descriptorForFontName` inputs, mechanism detection — must operate on content
with justification artifacts removed or ignored. Concretely:

- Pipelines that re-render (Apply, Re-render, Finalize) run `stripJustification`
  (existing: removes U+0640 tatweels, U+200A/U+2009 micro-spaces) over each cell **before**
  reading run fonts for detection.
- Where stripping before reading is not possible (read-only passes), the font reader
  filters out runs whose text consists solely of artifact characters before deciding a
  word's font.
- Node-testable: detection given artificial mixed runs (word-run in Jameel + tatweel-run
  in Arial) must resolve the word's font as Jameel.

## 8. Probe & calibration caching

**Problem (observed):** every Apply runs `AshaarTune.probeFont` (full pair-quality probe)
and a 50-iteration `calibrate` hill-climb from scratch (`taskpane.js:2552-2581`). The
autotune module was designed for bake-once reuse (`session.bake()` → recipe →
`loadRecipe`); the add-in never adopted it.

**Design (v1):**
- `probeFont` memoized per font family: in-memory map + `localStorage`
  (`ashaar:fontProbe:<family>`, stamped with an engine version — bust on vendor sync).
  Font metrics are machine-scoped, so localStorage (not document settings) is correct.
- `calibrate` memoized in-memory per (family, size, container-width bucket, texts hash):
  repeat applies of an unchanged poem skip the hill-climb. No persistence in v1 (texts
  change often enough that localStorage would mostly miss; the in-session hit rate is
  what kills the per-Apply cost).
- Cache bypass when the fonts strip registers/replaces a font (measurement basis changed).

**Deferred:** pre-baked recipes for bundled fonts (Jameel/Mehr/Noto) shipped as JSON via
`bake()`/`loadRecipe`, eliminating first-apply calibration on fresh machines.

## Interfaces & storage summary

| Change | Where |
|---|---|
| Apply/Re-render descoped to block | `taskpane.js` routing (no schema change) |
| `overrides[key].fill/color` | tag v3 payload (additive; parse guarantees unchanged) |
| Apply-target toggle | pane state only (not persisted) |
| Refresh-cost caption | pane state only (computed per refresh) |
| Probe cache | in-memory + localStorage `ashaar:fontProbe:<family>` (versioned) |
| Calibration memo | in-memory only |
| Capture | read-only Office.js loads → pending buffer |
| Strip-before-detect | render/justify pipelines + font-reader filter |

Canonical settings keys are unchanged; no new resolver layers. Tag writes stay behind
`parseContentControlTag`/setters.

## Error handling

- Batch writes (§3) build the full new tag in memory and write once — no partial tag state.
- Capture failures (no cell at cursor, API error) message and leave pending untouched.
- All tag writes on Apply paths stay on `withWordStrict` semantics (failures keep pending).

## Testing

- Node: override fill/color round-trip through tag parse/setters; batch key enumeration
  from cell maps; strip-before-detect on synthetic mixed runs; refresh-cost label
  computation (pending × scope → label) as a pure function.
- Browser (Playwright + mocked Office settings): toggle/capture/pending flows.
- Word (manual checklist addendum): cascade descope (sibling poems untouched), Re-render
  after native text+font edits, batch apply to whole poem, capture→apply→update round
  trip, cost captions match what actually happens, nastaliq font change no longer produces Arial
  words.

## User expectations (copy, shown in the pane)

- Footer caption examples: "Apply — rebuilds poem tables" / "Apply — re-justifies poem" /
  "Apply — re-justifies poem (unjustified: Justification is None)".
- Capture: "Read this cell's formatting into the pane — Apply to persist."
- Writing workflow note (docs/checklist, not UI): set Justification to None while
  drafting; switch it back and Apply (or Re-render) to publish.
