# Qaseeda Profiles — Implementation Status (resume point)

**Living doc.** Updated after each phase so a fresh session can resume mid-build
after a pause. Design spec: `2026-05-31-qaseeda-profiles-design.md`.

## Branch / commit trail

| Work | Branch | Status |
| --- | --- | --- |
| Kashida cap fix (engine) | `feature/layout-grid-mode` | ✅ pushed (submodule `47e6250`, parent `d6abf8b`) |
| P1 in-place auto-fit resize on Justify | `feature/layout-grid-mode` | ✅ pushed (`8320235`) |
| FatemiMaqala webfont bundling | `feature/layout-grid-mode` | ✅ pushed (`3ffc67d`) |
| **P2** profile model + store + tagging | `feature/qaseeda-p2-profiles` | ✅ this branch |
| **P3** apply/refresh engine | `feature/qaseeda-p3-apply` | ⏳ next |
| **P4** profile UI + colours + correction + warning | `feature/qaseeda-p4-ui` | ⏳ pending |

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

## P3 — apply/refresh engine (next)

Per design §"Apply / refresh engine":
1. `gatherQaseedaBlocks(name)` — find all content controls whose tag names that qaseeda
   (iterate `context.document.contentControls`, filter title "Ashaar Poem" +
   `parseContentControlTag(tag).qaseeda === name`).
2. `deriveProfileWidths(blocks, profile)` — measure each block's per-column natural
   widths (reuse the per-cell real-font canvas measurement from `justifySelection`),
   feed into `AshaarProfiles.deriveSharedWidths` (auto-fit) or use fixed %.
3. `applyProfileToQaseeda(name)` — resize every block's columns in place to the shared
   vector (reuse P1's `TableColumn.width` path, WordApiDesktop 1.3) and re-justify with
   the profile's params; cache `derived.colWidthVector` via `putProfile`.
4. **Hybrid trigger:** in `justifySelection`, if the block has a qaseeda and its content
   exceeds the cached widths, refresh the vector and re-apply to all the qaseeda's blocks.
- Pure width-vector→twips/span conversion math should be added to `profiles.js` and tested.

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
