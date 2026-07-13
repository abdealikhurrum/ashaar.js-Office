# Unified Settings Panel — Design

**Date:** 2026-07-12
**Status:** Approved in brainstorming; pending spec review
**Branch context:** `feat/guided-justification-ux`

## Problem

Justification/layout settings are scattered across the task pane in four
positions with overlapping meanings:

1. **Shared controls** (top): justify mode, fill mode, kashida strength, misra
   gap, table width + auto-fit, font mode, font uploader. Used at insert time.
2. **Qaseeda profile panel** (bottom, collapsed): duplicates justify mode,
   fill mode, strength, width — applied to tagged blocks.
3. **Three floating cursor-context editors:** per-cell override, per-bandh
   width, gap decoration.
4. **Conversion tab:** `layout-mode` and `width-mode` selects that overlap the
   width settings above.

Consequences: the same setting exists in two places with different semantics;
the pane's gap slider is not synced from the active block, so Re-render
silently adopts a stale pane value; the profile schema's `gap` field exists
but is dead (never read or written); a profile edit clobbers deliberate
per-poem tweaks because full values are stored everywhere.

## Goals

- **One position** for all justification/layout settings.
- **Profiles selectable and assignable from that same position.**
- **Ad-hoc adjustments apply only to the current selection/block** and can be
  **saved as a new profile**.
- **Word-style overlay semantics:** a block follows its profile; ad-hoc
  changes are local deltas layered on top; profile edits propagate to
  everything except locally overridden settings.
- Justification must work on **plain selections** with no Ashaar block at all.

## Decisions (from brainstorming)

| Question | Decision |
|---|---|
| Profile vs ad-hoc relationship | **Overlay (Word-style):** block stays linked; ad-hoc change stored as local delta; profile edits update everything except overridden settings |
| Apply timing | **Explicit Apply button** — batch changes, one rebuild+justify |
| Fine-grained editors (cell/bandh/gap-decor) | **Absorbed** into the panel via a scope switcher |
| No Ashaar block at cursor | Target is **"Selection"** — justify arbitrary selected text without blocks |
| Font | **Not a panel setting.** Font comes from the document/selection. Font *uploader* (WebView measurement) sits adjacent and is prompted just-in-time |
| Layer targeting | **Scope chips decide; cursor only enables.** No auto-narrowing to Cell scope |
| Profile mutation | Never a side effect of Apply — only via explicit `Update "name"` button |

## Architecture

### The panel

One always-visible **Settings panel** at the top of the pane, replacing the
shared-controls section, the Qaseeda panel, and the three floating editors.
Three zones:

**1. Target header** — tracks the cursor:

- `Selection` — cursor in plain text; no scope chips.
- `Poem — Karbala` / `Poem — (no profile)` — inside an Ashaar block.
- Scope chips inside a block: `Poem | Bandh | Cell | Gap`. Default is always
  **Poem**. Cell/Gap chips are enabled only when the cursor is in a content /
  spacing cell (same detection as today's override editors). Moving the
  cursor never auto-switches scope; the user narrows explicitly by clicking a
  chip. Breadcrumb shows the narrowed target, e.g. `Poem — Karbala › Cell A2:3`.

**2. Profile row** — dropdown of all saved profiles + `(none)`, plus:

- `Assign` — links the profile to the current block (disabled for plain
  selections — nowhere to store the link; for a selection, picking a profile
  just loads its values into the controls as a preset).
- `Save as…` — snapshots the panel's resolved values into a new named
  profile; the current block (if any) is assigned to it and its local delta
  map cleared (the tweaks became the profile).
- `Update "name"` — appears only when panel values have drifted from the
  assigned profile; pushes current values into the stored profile and
  re-applies to every assigned block (each block's own local deltas survive).

**3. Settings body + Apply** — controls resolved for the current target.
Locally tweaked settings show a provenance dot with a per-setting reset (⟲);
resetting deletes the delta so the value falls back to the inherited layer.
`Revert to profile` clears the whole delta map (reads `Reset to defaults`
when no profile). One `Apply` button commits. Scope determines the visible
control set:

- **Poem:** justify mode, fill mode, strength, misra gap, unified width
  control (`auto-fit | fixed %`), collapsible Advanced (layout mode, misra
  width pt, font correction, debug colors).
- **Bandh:** misra width (pt).
- **Cell:** strength, target width (pt), cap lift (em) — each showing the
  inherited value when unset.
- **Gap:** decoration symbol, fill color, symbol color; plus the existing
  `Set as default for all bandhs` action (writes the profile's
  `spacingDecor`, gated behind the same explicit-mutation rule as
  `Update "name"`).

### Fonts strip (adjacent, not inside)

A collapsed `▸ Fonts` strip below the panel holds the existing uploader
(file picker, register-as name, verify, registered list). Its single job is
making fonts measurable in the WebView sandbox. On Apply, the pipeline
force-loads every distinct target font (existing behavior); if one fails to
load, Apply is interrupted with an inline prompt:

```
⚠ Can't measure "Jameel Noori Nastaleeq" — justification would be inaccurate.
[Add font file…]   [Continue anyway]
```

`Add font file…` expands the strip with the name pre-filled. `Continue
anyway` proceeds with fallback metrics (today's silent behavior, made
visible). Detection is per-Apply/per-target, so registered fonts never
prompt again.

## Data model

Three storage locations, all existing:

### 1. Block tag — payload v3

```js
{
  k: "ashaar-poem", v: 3,
  profile: "Karbala",           // "" = none (renames v2 `qaseeda`)
  local: { gap: 8 },            // ONLY user-touched block-level settings
  profileCache: { ... },        // snapshot of profile values as last applied
  // render facts (describe what's physically in the document):
  misraPattern: "paired", misraCount: 4, cells: [...], sourceHash: "…",
  // finer-scope data, unchanged from v2:
  overrides: { "A2:3": { strength: 9 } },  // per-cell
  widthPt: null,                            // per-bandh width
  slotDecor: { "A#1": { symbol: "؎" } },   // gap decoration
  runFonts: { ... }                         // per-word font packs
}
```

Preferences (gap, strength, justify mode, width…) appear in `local` only
when explicitly tweaked; absent means inherited. Render facts stay top-level.
`local` deltas exist **only at block level** — bandh (`widthPt`) and cell
(`overrides`) slots are already deltas by construction (absent = inherit) and
are not restructured.

### 2. Profile store (localStorage)

Existing `profiles.js` schema (`width`, `justify`, `gap`, `spacingDecor`,
`fontCorrections`, `derived`). The dead `gap` field is wired up. The `font`
field is dropped (font is never a setting).

### 3. App defaults

`defaultProfile()` in `profiles.js` — the bottom layer, unchanged.

### Resolver

One new pure function in `profiles.js`:

```js
resolveSettings({ payload, profileStore, scopeKey })
// → { values: { gap: 8, strength: 6, ... },
//     source: { gap: "local", strength: "profile", ... } }
```

Layering: **defaults → profile (live store, else `profileCache`) → block
`local` → bandh `widthPt` → cell override.** `values` drives the panel
display *and* every apply pipeline; `source` drives provenance dots. Pure and
Node-testable.

### v2 migration — read-time only

`parseContentControlTag` maps v2 payloads to v3 shape on read: stored
preference values become `local`, `qaseeda` becomes `profile`. The document
is untouched until the next Apply writes the tag back as v3. No bulk
migration pass. Malformed tags are treated as not-an-Ashaar-block (as today).
Unknown fields in a v3 tag round-trip untouched — setters modify only their
own keys.

## Data flow

### Panel fill (cursor moves)

The existing reflection loop (`reflectActiveContext`) keeps its detection
(enclosing block, tag parse, cell-at-cursor) but renders the panel from
`resolveSettings()` — values, dots, breadcrumb. Unapplied edits are held in a
**pending-deltas buffer**; reflection never overwrites a control being
edited (existing guard).

### Apply — routed by target

- **Plain selection:** panel values feed the existing selection-justify path
  (probe → calibrate → justify). Nothing persisted; one-shot tool.
- **Poem scope:** pending deltas are written into `local` (a delta set back
  to its inherited value is deleted, not stored). Rebuild the table structure
  only if a structural setting changed (gap, width, pattern); then justify
  with resolved values. Scoped to **this block only**, never profile siblings.
- **Bandh / Cell / Gap scope:** deltas go to `widthPt` / `overrides` /
  `slotDecor` via existing setters; block re-justifies.

On success the pending buffer clears and the tag gains an updated
`profileCache`. On failure the buffer is kept for retry.

### Profile actions

- **Assign:** sets `tag.profile`, applies; profile values show through
  wherever the block has no local tweaks.
- **Save as…:** resolved values → new profile; block assigned; `local`
  cleared.
- **Update "name":** panel values → stored profile; re-apply to every
  assigned block; each block resolves through its own `local`, so deliberate
  tweaks survive (the overlay payoff).

### Rebuild-skip fix

The apply pipeline's `sizeSig` (currently target width + source text) gains
the structural settings (gap, pattern) so a gap-only change reliably triggers
the table rebuild. The signature computation is extracted into a pure,
testable function.

The measurement engine (probe, calibrate, kashida insertion, run-font
preservation) is untouched — it receives resolved values instead of raw pane
reads.

## Disposition of existing UI

**Deleted (absorbed into the panel):**

- Top shared-controls section: justify mode, fill mode, strength, gap, table
  width + auto-fit → panel controls. `font-mode` dropdown removed entirely
  (insert uses the document font at the cursor; re-render already preserves
  fonts).
- Qaseeda panel: name box, width mode/pct, strength, Advanced group.
  `Assign block at cursor` → `Assign`; `Save & Apply to all` → `Update "name"`.
- Floating editors `cell-override`, `bandh-override`, `slot-decor` → the
  Cell / Bandh / Gap scope bodies.
- `Justify Selected Text` and `Re-render (keep font)` buttons → `Apply`.

**Consolidated:** Conversion tab's `layout-mode` and `width-mode` selects
move into the panel — layout mode under Advanced; one unified Width control
(`auto-fit | fixed %`) replaces the width-mode select, table-width slider,
and auto-fit checkbox.

**Moved:** the font uploader → the `▸ Fonts` strip + just-in-time prompt.

**Stays:**

- Table Input tab: bandh/misra counts, layout preset + spec/grid editor,
  Draw Table, Drop Grid, Adopt Existing Table, Templates. **Addition:** a
  `Replace Selection` button under `Adopt Existing Table`, sharing the
  Conversion tab button's handler (`insertPoem(true)`) and enablement/
  messaging — one handler, two buttons.
- Conversion tab: poem textarea, separator import options, preview, Insert
  as Table / Paragraphs / Replace Selection / Load Selection.
- Utilities `Reset (clear kashida & spaces)` and `Show cell structure` —
  small buttons under the Apply row.

Resulting pane order: **Settings panel → Fonts strip → mode tabs (structure
& text only) → utilities → message line.**

## Error handling

- **Missing profile (cross-machine):** profiles are per-machine
  (localStorage); documents travel. The resolver falls back to the tag's
  `profileCache` when the store lacks the profile; the panel shows
  `Karbala (not on this machine)` with a one-click **Restore profile from
  this poem** that recreates the store entry from the cache. The cache is
  fallback data only — never consulted when the store has the profile.
- **Unmeasurable font:** inline prompt (above); Apply blocked until the user
  chooses.
- **Legacy/malformed tags:** v2 migrates at read time; unparseable = not an
  Ashaar block.
- **Apply failure:** pending deltas kept, error shown, retry is one click.
  `_reflectBusy` guard prevents reflection clobbering during Apply.

## Testing

Node `assert` style (this repo's convention; no framework):

- `tests/profiles-resolve.test.js` (new): resolver layering order,
  provenance labels, delta add/remove, clear-falls-back-to-inherited,
  `profileCache` fallback, v2→v3 migration.
- `tests/word-html.test.js` (extend): v3 tag round-trip; each setter touches
  only its key; unknown-field preservation.
- Rebuild-skip signature (extracted pure function): same inputs → same sig;
  gap or pattern change → different sig.
- Panel state logic kept in pure functions
  (`panelStateFor(resolved, pendingDeltas)` → renderable struct) so it is
  Node-testable; DOM glue stays thin.
- **Manual Word checklist** (established pattern for Office.js behavior):
  assign/apply/tweak/revert on a two-poem document; profile edit propagates
  except overridden settings; cross-machine profile restore; plain-selection
  justify; unmeasurable-font prompt; legacy v2 document open-and-apply;
  adopt → Replace Selection round-trip from the Table Input tab.

## Out of scope

- Insert-time font selection UI (removed; revisit only if a real need
  emerges).
- Per-setting delta maps at bandh/cell level (existing single-purpose slots
  suffice).
- Bulk migration of v2 documents (read-time migration covers it).
- Any change to the measurement/justification engine.
