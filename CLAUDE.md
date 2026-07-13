# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

A Microsoft Word task pane add-in for entering and formatting Arabic, Urdu, and Persian poetry with professional kashida (text justification) and flexible table layouts. Integrates Ashaar.js (a poetry formatter) as a vendored git submodule.

## Commands

```bash
# Run the add-in with debugging (requires Word installed)
npm start

# Start the HTTPS dev server only (port 3000)
npm run dev-server

# Run all tests
npm test

# Run a specific test file
node tests/word-html.test.js
node tests/word-tabstop.test.js

# Generate dev HTTPS certificates
npm run certs

# Sync vendored Ashaar.js files from the submodule
npm run sync:ashaar

# Update Ashaar.js submodule to latest, sync, and run tests
npm run update:ashaar
```

## Architecture

### Tech Stack
- Vanilla JavaScript (ES5/UMD modules, no build step, no transpilation)
- Office.js v1 API for Word document interactions
- Ashaar.js (git submodule at `vendor/ashaar-js/`, vendored copy at `src/vendor/`)
- Node.js assert for tests (no jest/mocha)

### Code Layers

**UI Layer** (`src/taskpane/taskpane.html`, `taskpane.js`, `taskpane.css`)
- IIFE-structured event handlers; `bind()` wires all controls on Office ready
- Two mode tabs: **Table Input** (draw blank tables first, type into cells) and **Conversion** (paste poetry, convert to Word)
- Live preview in taskpane via Ashaar.js native HTML rendering

**Word Interaction Layer**
- `src/taskpane/word-html.js` — `AshaarWord` class: helpers that build the OOXML/HTML table structures for parsed poetry. The orchestrating `insertNativeLayoutTables()` lives in `taskpane.js`, which calls these helpers and inserts the tables via `Word.run()` with content control tags.
- `src/taskpane/word-tabstop.js` — `AshaarTabStop` class: generates OOXML paragraphs with tab stops as an alternative to tables. `poemToOoxml()` builds OOXML XML strings from parsed poem structure.

**Vendor Layer** (`src/vendor/`)

Three modules synced from `vendor/ashaar-js/` submodule via `scripts/sync-ashaar-vendor.mjs`. Never edit `src/vendor/` directly.

- `ashaar.js` — Poetry parser (`Ashaar.parse()`), HTML renderer, and in-browser layout engine. `Ashaar.justifyEl(el, opts)` applies block-level justification to a rendered poem container in the DOM, balancing all misras to the width of the longest line.

- `ashaar-justify.js` — Low-level canvas-based kashida insertion engine. `AshaarJustify.justifyLine(text, targetWidth, ctx, params, fontProfile)` binary-searches for the maximum number of tatweels that fit within a target pixel width. `buildSlots()` ranks legal tatweel insertion positions; `spreadTatweels()` is a simpler count-based alternative. Used directly by both `word-html.js` and `ashaar-autotune.js`.

- `ashaar-autotune.js` (upstream: `ashaar-tune.js`) — Font analysis, visual calibration, and recipe deployment for optimal kashida justification. Three-phase workflow:
  1. **Probe** (`AshaarTune.probeFont({fontFamily, fontSize})`) — analyses a loaded font via canvas width-delta measurements to score which character pairs have designed kashida glyphs vs. generic fallback. Returns a `FontProfile` with per-pair quality scores.
  2. **Calibrate** (`AshaarTune.calibrate({texts, fontFamily, fontSize, containerWidth, mode, fontProfile, iterations})`) — runs a hill-climbing optimiser on sample poem texts, scoring candidates visually by rendering to an offscreen canvas and measuring ink-column density evenness and inter-line harmony. Returns a `CalibrationSession` with optimal `{targetFill, fontQualityBoost}` params; call `.bake()` to get a portable JSON recipe. Use `calibrateWidths()` for responsive multi-column recipes.
  3. **Deploy** (`AshaarTune.loadRecipe(recipe)`) — loads a baked JSON recipe and returns a `justifyEl` drop-in replacement that uses the calibrated params.
  - The `setScorer(fn)` extension point allows replacing the canvas scorer with an ML model.
  - Treat each Ashaar Poem content control as a single calibration unit: gather all misra texts from the poem's tables, calibrate once, then apply the same params to every cell.

### Two Main Workflows

**Table Input Mode:**
```
User sets layout → "Draw Table" → insertStructure()
  → AshaarWord.layoutTablesForTemplate() → Word.run() table with content controls

"Drop Grid" → insertBareGrid() → AshaarWord.generateBareGrid12Ooxml()
  → 12-column blank table with thin borders; user merges/reshapes cells natively in Word
  → "Capture from Word" → captureSelectedTableLayout()
    → reads cell columnWidths, infers 12-col spans, saves to localStorage as named template
  → "Apply" → applyTemplate() → AshaarWord.templateToOoxml() → inserts captured layout
  → "Export / Import JSON" → portable across documents
```

**Conversion Mode:**
```
Paste poetry text → renderPreview() (live Ashaar.js render in taskpane)
  → "Insert as Table"      → insertPoem() → insertNativeLayoutTables()
  → "Insert as Paragraphs" → insertTabStopPoem() → AshaarTabStop.poemToOoxml()
  → "Replace Selection"    → insertPoem(true) replaces selected Word text
```

**Settings Panel → Apply:**
```
Click inside an Ashaar Poem block (or on plain text) → onSelectionChanged()
  → reflectActiveContext() → AshaarProfiles.resolveSettings({payload, profileStore, scope})
    → resolved values (profile → local overrides → cell/bandh/gap scope)
  → settings-panel.js (AshaarPanel) — the pure panel-state module: panelStateFor()
    builds the render state (chips, provenance dots, footer) from resolved + pending;
    mergePending()/pendingToLocal() manage the in-pane pending-edit buffer (nothing
    written to the document until Apply)
  → user edits a field → pending buffer updated → refreshPanel() re-renders
  → "Apply" → applyPanel() routes by target + scope:
      plain selection            → justifySelection() (one-shot; nothing persisted)
      poem/bandh/cell/gap scope  → font-measurability gate (ensureFacesMeasurable,
                                    ok/continue/cancel — cancel keeps pending for retry)
                                 → tag write (setTagBandhWidth / setTagOverride /
                                    setTagSlotDecor, or the poem-scope rebuild path)
                                 → reRender() or justifySelection() re-renders in place
  → tag payload is v3: {profile, local, profileCache, cells, overrides, slotDecor, ...}
    profile = assigned profile name; local = this scope's override deltas;
    profileCache = last-applied profile snapshot (used to detect drift/local edits).
    v2 tags (pre-panel documents) migrate to v3 read-time in parseContentControlTag().
  → justifySelection() itself still drives AshaarTune.probeFont()/calibrate() and
    AshaarJustify.justifyLine() per cell, as before — the panel only changed how its
    inputs (justifyMode/strength/gap/width/etc.) are gathered and applied.
  → Re-render button (block scope) re-applies params on document edits; apply-to-all targets (cell/bandh/poem) propagate via nulls; Capture reads cell formatting (fill, text color) as read-only in the pane; refresh-cost caption distinguishes rebuild from re-justify; lineHeightPt and separatorPt (advanced section) govern line spacing and inter-poem gaps.
```

### Poetry Input Format
- `\` separates misras (hemistichs) within a bayt (couplet)
- `|` creates a single-misra line
- `*` marks a refrain
- Blank line separates stanzas
- `---` separates poems
- `%` prefix marks a refrain bayt

### Module Pattern
All library files (`word-html.js`, `word-tabstop.js`) use UMD (Universal Module Definition) for compatibility with both Node.js (tests) and the browser (Word WebView):
```js
(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) {
    module.exports = factory(require("../vendor/ashaar-justify"));
  } else {
    root.AshaarWord = factory(root.AshaarJustify);
  }
}(typeof globalThis !== "undefined" ? globalThis : this, function (AshaarJustify) { ... }));
```

### Submodule Workflow
```bash
# After cloning, initialize the submodule
git submodule update --init

# To update Ashaar.js to latest upstream
npm run update:ashaar
# This runs: git submodule update --remote, then sync:ashaar, then tests
```
