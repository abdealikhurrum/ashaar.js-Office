# Qaseeda Profiles — Implementation Status (resume point)

**Living doc.** Updated after each phase so a fresh session can resume mid-build
after a pause. Design spec: `2026-05-31-qaseeda-profiles-design.md`.

## Branch / commit trail

| Work | Branch | Status |
| --- | --- | --- |
| Kashida cap fix (engine) | `feature/layout-grid-mode` | ✅ pushed (submodule `47e6250`, parent `d6abf8b`) |
| P1 in-place auto-fit resize on Justify | `feature/layout-grid-mode` | ✅ pushed (`8320235`) |
| FatemiMaqala webfont bundling | `feature/layout-grid-mode` | ✅ pushed (`3ffc67d`) |
| **P2** profile model + store + tagging | `feature/qaseeda-p2-profiles` | ✅ pushed |
| **P3** apply/refresh engine | `feature/qaseeda-p3-apply` | ✅ this branch |
| **P4** profile UI + correction + warning | `feature/qaseeda-p4-ui` | ✅ this branch (colouring deferred) |

Each phase branches off the previous phase branch (stacked); none merged to `main` yet.

## What exists now

### Engine (vendored `ashaar-js` submodule, synced to `src/vendor`)
- `justifyLine` tatweel cap raised from `charCount` → `charCount*8` so long lines
  fill wide auto-fit columns (binary search still bounded by target). Regression
  test in `vendor/ashaar-js/test/justify.test.js`.

### Measurement accuracy
- `fontAvailable(name)` (taskpane.js) — canvas width-comparison resolver; Debug `res` column.
- FatemiMaqala vendored as `vendor/font-fatemi` submodule; TTF served from
  `assets/fonts/FatemiMaqala-Regular.ttf`; `@font-face { font-family: "FatemiMaqala" }`
  in `taskpane.css` (family = name-table nameID 1, matches what Word reports).
  `justifySelection` `await document.fonts.load(repFont)` before measuring.

### P2 — profile model + store + tagging (pure logic tested)
- **`src/taskpane/profiles.js`** (`AshaarProfiles`, UMD, tested in `tests/profiles.test.js`):
  - `defaultProfile(name)`, `normalizeProfile(p)`, `mergeProfile(base, partial)` (deep, non-mutating).
  - `applyFontCorrection(px, font, corrections)`.
  - `deriveSharedWidths(columns, {headroom, corrections})` → per-column target px
    (widest corrected measurement × headroom). **P3 consumes this.**
- **`word-html.js`** (tested in `tests/word-html.test.js`):
  - `contentControlTag` payload gained a `qaseeda` field.
  - `parseContentControlTag(tag)` → payload or null (guarantees string `qaseeda`).
  - `setTagQaseeda(tag, name)` → tag copy with only `qaseeda` replaced.
- **`taskpane.js`** Office plumbing (not node-testable; verify in Word):
  - `loadProfileStore()` / `saveProfileStore(store)` — document settings, key `"ashaar:profiles"`.
  - `getProfile(name)` / `putProfile(profile)` / `listProfileNames()`.
  - `setQaseedaOnSelection(name)` / `getQaseedaAtSelection()` — read/write the
    qaseeda name on the Ashaar Poem block at the cursor.
- profiles.js wired into `taskpane.html` (before word-html.js) and `package.json` test script.

## P3 — apply/refresh engine (done)

Implemented **additively** — `justifySelection` was deliberately left untouched (it
is the confirmed-working path; any P3 regression is isolated to this unmerged branch
and recoverable by reverting to P2). The new engine mirrors the proven P1 resize +
justify snippets rather than refactoring the shared function. **Verify in Word.**

- **profiles.js (tested):** `columnPointsFromContentPx(px, marginPt)` (content px →
  column width in points incl. both cell margins); `strengthToTargetFill(strength)`
  (0..24 → 0.90..1.0, matches word-html `sliderToFill`).
- **taskpane.js:**
  - `gatherQaseedaBlocks(context, name)` — all "Ashaar Poem" CCs whose tag qaseeda === name.
  - `applyProfileToQaseeda(name)` — loads every block's tables/cells (real font + width),
    and for `width.mode==="auto-fit"` (and WordApiDesktop 1.3) computes ONE shared table
    width = the largest a block's tightest cell needs (× headroom), capped at page width,
    then scales every block's columns to it; finally re-justifies every cell with the
    profile's params (`AshaarJustify.justifyLine`, targetFill from strength). Falls back to
    re-justify-only when resize isn't supported. Caches via `putProfile`.

**Refinements — done on `feature/qaseeda-refine` (verify in Word):**
- **Per-column width equality** — when every block of a qaseeda has the same column count,
  `applyProfileToQaseeda` now equalises PER COLUMN: column *j* width = max over blocks of
  (its width × that block's needed scale), and every block is set to the identical vector
  (capped at the page). Mixed-shape qaseedas fall back to equal total width + proportional
  columns. Replaces the earlier one-shared-total-width behaviour.
- **Hybrid auto-refresh on justify** — `justifySelection` now begins with a guard: if the
  cursor's block belongs to a qaseeda with a stored profile, it delegates to
  `applyProfileToQaseeda(name)` (sizes + justifies all the qaseeda's blocks together) and
  returns. Untagged blocks justify exactly as before, so the proven path is untouched.

**Still deferred:**
- **OOXML symbol/debug run-colouring** in the `word-html.js` generator. Intentionally left
  out: its UI fields were removed in the P4 simplification, and the colouring is cosmetic +
  unverifiable without a live .docx. Revisit with the colour UI re-added if wanted.
- The measure+justify loop still duplicates ~70 lines of `justifySelection`; consolidate into
  a shared `justifyWorkRange(context, range, opts, cfg)` helper once it can be verified live.

## P4 — profile UI + correction + warning (done; colouring deferred)

- **Profile panel** (`taskpane.html` `.qaseeda-panel`, styled in `taskpane.css`): qaseeda
  name (with a `<datalist>` of saved names), width auto-fit/fixed + %, justification
  type + strength, and under "Appearance & correction": gap, misra symbol + colour,
  debug tatweel/space colours, per-font correction (font + factor).
- **Wiring** (`taskpane.js`): `panelToProfile`/`profileToPanel`, `populateQaseedaNames`,
  `loadQaseedaIntoPanel` (typing a saved name loads it), `saveAndApplyQaseeda`
  ("Save & Apply to all" → `putProfile` then `applyProfileToQaseeda`),
  `assignBlockToQaseeda` ("Assign block at cursor" → `setQaseedaOnSelection`),
  `checkQaseedaFont` ("Check font at cursor" → `fontAvailable`, shows accurate/approximate).
- `options()` now carries the qaseeda name, so **new inserts are tagged** with it.
- **Per-font correction is live** in the apply engine: measured width × `fontCorrections`
  factor feeds the resize (via tested `AshaarProfiles.applyFontCorrection`).

**Deferred (cosmetic, not safely verifiable without a live Word session):**
- `word-html.js` generator **OOXML run-colouring** of the misra symbol and of inserted
  tatweels/spaces (debug colours). The profile *fields* are complete and stored; only the
  coloured rendering in the generator remains. Implement with a real .docx to check the
  run-level `<w:color>` output.

## Suggested merge / verify order
1. In Word: re-confirm Justify (cap fix + font res). 2. Tag two blocks with one qaseeda,
   "Save & Apply to all", confirm both get the same width + justification. 3. "Check font
   at cursor" on Arial (ok) and an unresolved font (warn). 4. Merge `feature/layout-grid-mode`
   first (P1+cap+font), then the qaseeda branches in order P2→P3→P4 (they are stacked).

## Resume instructions
- `npm test` must stay green (now 7 suites incl. `profiles`).
- Never edit `src/vendor/*` directly — edit `vendor/ashaar-js`, then `npm run sync:ashaar`.
- In-place resize is desktop-only (`WordApiDesktop 1.3`); always keep a no-resize fallback.
- If a change regresses, **revert** it (do not patch forward) per the standing instruction.
