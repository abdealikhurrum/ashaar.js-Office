# Cell Configurations — Decomposition Roadmap

**Date:** 2026-07-11
**Status:** Living index — one spec (and one implementation session) per sub-project.

"Spacing and content cell configurations" was decomposed into two sequential sub-projects. They split along one axis: **where the config lives and whether it needs a per-cell editing UI.** Each gets its own design spec, implementation plan, and build session.

## Sub-project 1 — Bandh cell-map (content/spacing tag + positional labels)  *(implemented; Tasks 1–6 done, Task 7 manual-verify pending)*

- **Scope narrowed during brainstorming** to the minimal data model that makes table creation + justification user-friendly: a **content/spacing tag** and a **positional label within the bandh** (`A1/A2/B1…`) per cell, derived and **persisted at creation**.
- Broader "roles" and styling were explicitly pushed later (see backlog) — only gap/spacing-vs-text and the cell-id are in scope now.
- Covers the **structural tagging** capability (as the label + tag model); spacing/content *styling* moves to the backlog.
- **No dependency** on any parked feature — ships standalone.
- Spec: `docs/superpowers/specs/2026-07-11-bandh-cell-map-design.md` ✅ written.

## Sub-project 2 — Per-cell identity + select-cell → pane editing  *(follow-on)*

- Stable **per-cell id**; per-cell **overrides** (style variance **and** justification variance) stored on the Ashaar block, keyed by id.
- The pane **reflects the cell at the cursor** and edits its overrides — this IS the parked "pane reflects the active block's settings" feature.
- Covers the fourth capability: **per-cell justification override** (the §4 under-resolved-cell variance) plus per-cell style overrides.
- **Depends on SP1's** role model and on the parked pane-reflects-active-block UI.
- Spec: `docs/superpowers/specs/YYYY-MM-DD-per-cell-overrides-design.md` *(future session)*.

## Backlog — future features (keep on the board)

Requested / raised but deliberately out of SP1's scope. Each becomes its own spec + session when picked up. The SP1 cell-map (content/spacing tag + label + spacing `slot`) is the anchor these hang off.

- **Special gaps** — turn a spacing cell into a purposed slot:
  - **Hemistich symbols** (the `*`/refrain marker rendered in a gap).
  - **Bandh numbers** (stanza numbering in a gap/margin cell).
  - **Annotations** (notes attached to a gap slot).
  - **In-gap labels** (showing the cell label on the page — deferred; Word has no clean cell-text overlay).
- **Block-level roles + styling** — sadr/ajuz/solo/refrain classification; role-based content-cell text color/shade; spacing-cell fill glyphs / dot-leaders / shading. Extends the unused `profile.misraSymbol` / `profile.symbolColor` fields.
- **Per-cell overrides + select-cell→pane editing** — SP2 below (justification variance + style variance keyed by label; the parked "pane reflects the active block" UI). SP1's read-only pane view is its seed.

## Related prior specs

- `2026-07-11-poetry-justification-modes-design.md` §1 (content/spacing tagging introduced), §4 (per-cell variances → SP2), §7 (spacing-cell styling non-goal → SP1).
- `2026-05-31-qaseeda-profiles-design.md` (the profile model SP1 extends).
