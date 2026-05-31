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
| **P4** profile UI + colours + correction + warning | `feature/qaseeda-p4-ui` | ⏳ next |

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

**Scope notes / deferred refinements (do in P4 or later, with in-Word testing):**
- Shared sizing is **one table width per qaseeda** (proportional column scaling), not yet a
  per-grid-column absolute vector — `deriveSharedWidths` exists and is tested for when we
  move to per-column equality (needs gridSpan-aware column mapping).
- **Hybrid auto-refresh on justify** (re-apply to siblings when a block outgrows cached
  widths) is **not** wired into `justifySelection` yet — kept out to protect that path.
  Fold it into P4's "Apply to all" flow / a guarded justify hook once testable in Word.
- The measure+justify loop duplicates ~70 lines of `justifySelection`; consolidate into a
  shared `justifyWorkRange(context, range, opts, cfg)` helper once it can be verified live.

## P4 — profile UI + colours + correction + warning (pending)

- Profile panel (collapsible): pick/name a qaseeda; edit width auto/fixed, gap, misra
  symbol + colour, justify type/strength, debug tatweel/space colours, per-font correction.
  "Apply to all". Assign-qaseeda control on inserts. `res`/warning when font unresolved.
- `word-html.js` generator: misra-symbol colouring + debug colouring of inserted artifacts.

## Resume instructions
- `npm test` must stay green (now 7 suites incl. `profiles`).
- Never edit `src/vendor/*` directly — edit `vendor/ashaar-js`, then `npm run sync:ashaar`.
- In-place resize is desktop-only (`WordApiDesktop 1.3`); always keep a no-resize fallback.
- If a change regresses, **revert** it (do not patch forward) per the standing instruction.
