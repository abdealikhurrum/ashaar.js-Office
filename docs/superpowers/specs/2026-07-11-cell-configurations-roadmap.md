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

## Sub-project 2 — Per-cell overrides + active-block sync  *(implemented; Tasks 1–4 done, Task 5 manual-verify pending)*

- Stable **per-cell id**; per-cell **overrides** (style variance **and** justification variance) stored on the Ashaar block, keyed by id.
- The pane **reflects the cell at the cursor** and edits its overrides — this IS the parked "pane reflects the active block's settings" feature.
- Covers the fourth capability: **per-cell justification override** (the §4 under-resolved-cell variance). Per-cell *style* overrides were dropped from SP2's MVP → backlog.
- MVP (approved): per-cell justify override (**strength / target width / cap lift**, keyed `"<tableIndex>:<label>"`, one cell) **+ full active-block sync** (pane reflects the block at the cursor). This absorbs the parked pane-reflects-active-block UI.
- **Depends on SP1's** label + `cells` map (the per-cell identity).
- Spec: `docs/superpowers/specs/2026-07-11-per-cell-overrides-design.md` ✅ written *(brainstormed; ready for planning)*.

## Sub-project 3 — Spacing-cell decorations (symbols + fill color)  *(spec written; ready for planning)*

- Put a **symbol** (hemistich / decorative glyph) and a **fill color** (+ symbol text color) into gap cells. Config = **both**: profile default by slot-position + per-slot override on the block tag; reuses SP2's editor/tag-writer/resolve pattern. Anchored on SP1's spacing `slot`.
- Spec: `docs/superpowers/specs/2026-07-11-spacing-cell-decor-design.md` ✅.
- **Fast-follow:** auto-numbering (bandh/verse counters in a slot) — own spec (numbering scheme).

## Backlog — future features (keep on the board)

Each becomes its own spec + session. The SP1 cell-map (content/spacing tag + label + spacing `slot`) is the anchor these hang off.

- **Auto-numbering** — bandh/verse counters rendered in a slot (which slot, per-bandh vs global, format `۱ ۲ ۳` / `1.`). SP3's fast-follow.
- **Annotations** — richer notes attached to a gap slot (beyond a static symbol).
- **In-gap labels** — showing the cell label on the page (deferred; Word has no clean cell-text overlay).
- **Block-level roles + content-cell styling** — sadr/ajuz/solo/refrain classification; role-based content-cell text color/shade; multi-glyph rules / dot-leaders. (SP3 covers spacing-cell symbol + fill; content-cell shading and repeated-fill leaders remain here.)

## Related prior specs

- `2026-07-11-poetry-justification-modes-design.md` §1 (content/spacing tagging introduced), §4 (per-cell variances → SP2), §7 (spacing-cell styling non-goal → SP1).
- `2026-05-31-qaseeda-profiles-design.md` (the profile model SP1 extends).
