# Kickoff prompt — Text conversion / find-replace subsystem

> Paste the block below into a fresh Claude Code session in the ashaar.js-Office
> repo to start this feature. It's the third subsystem deferred out of the
> 2026-07-16 Ashaar Styles brainstorm (the other two: annotations/citations,
> keyboard shortcuts). This is a *distinct* subsystem from the Styles tab — it
> should get its own brainstorm → spec → plan → implementation cycle.

---

I want to add a **text conversion / find-replace** feature to this Word add-in
(the Ashaar poetry/prose task pane). Please start with the `superpowers:brainstorming`
skill — this is a new subsystem and needs its own spec; don't jump to code.

**What it should do (from an earlier brainstorm, not yet detailed):**

1. **"Double-press" conversion** — convert to and from a "double-press" form.
   (I need to pin down exactly what this means with you — likely a
   doubled-character input convention that maps to specific letters/forms.
   Treat the precise definition as an open question to resolve first.)

2. **Symbol-to-letter conversions, bidirectional ("to and from"):**
   - Arabic semicolon (؛) ↔ cheh (چ)
   - brackets and mathematical symbols ↔ honorific abbreviations / expansions
     (e.g. "maza lemaza", "bhaisaheb", "behnsaheb", and similar). The full
     mapping table is TBD — I'll provide it during brainstorming.

3. **Find-replace-style behavior** — a UI to run these conversions over the
   document (or a selection), in either direction, the way Word's Find &
   Replace works but driven by these preset mappings.

**Open questions to work through in the brainstorm (don't assume answers):**
- Exact meaning/scope of "double-press."
- The complete, authoritative mapping table (symbols ↔ letters, and the
  honorific abbreviation ↔ expansion pairs) — and whether every mapping is
  reversible 1:1 or some are one-way.
- Live autocorrect (as you type) vs. an on-demand find-replace pass vs. both.
- Scope: whole document, current selection, or user-chosen range.
- Direction control: how the user picks "to" vs "from" (toggle? two buttons?
  per-mapping-category?).
- How it interacts with RTL / Arabic complex-script text and existing content
  (must not corrupt harakat, kashida artifacts, or content-control-tagged poem
  blocks).
- Preset "conversion sets" the user can enable/disable (semicolon↔cheh,
  honorifics, math/brackets) — one big pass or independently toggleable.

**Relevant context for whoever picks this up:**
- Project overview and architecture: `CLAUDE.md` (vanilla ES5/UMD, no build
  step; Office.js v1; UI in `src/taskpane/`).
- This feature was explicitly scoped OUT of the Styles spec — see the
  "Explicitly out of scope" section of
  `docs/superpowers/specs/2026-07-16-ashaar-styles-design.md`.
- Office.js gives you `Word.Body.search()` / `Range.search()` (with
  match-case/whole-word options) and range replacement, plus `Range.insertText`
  — check the Word JS API for Find/Replace capabilities and limits (especially
  around RTL/diacritics and search string length) as part of feasibility.
- Follow the same discipline the Styles feature used: a pure, node-testable
  data module for the mapping table + conversion logic (like `word-styles.js`),
  and a thin Office.js orchestration layer (like `styles-pane.js`), wired as its
  own task-pane mode/section. Keep `src/vendor/` untouched.

Start by exploring the repo, then ask me clarifying questions one at a time —
beginning with what "double-press" means and the mapping table.
