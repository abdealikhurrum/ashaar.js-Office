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
- `src/taskpane/word-html.js` — `AshaarWord` class: converts parsed poetry into HTML tables inserted via `Word.run()`. Core function is `insertNativeLayoutTables()` which creates Word tables with content control tags.
- `src/taskpane/word-tabstop.js` — `AshaarTabStop` class: generates OOXML paragraphs with tab stops as an alternative to tables. `poemToOoxml()` builds OOXML XML strings from parsed poem structure.

**Vendor Layer** (`src/vendor/`)
- `ashaar.js` — Poetry parser (`Ashaar.parse()`) and HTML renderer
- `ashaar-justify.js` — Canvas-based kashida justification engine (`AshaarJustify.justifyEl()`)
- Files are synced from `vendor/ashaar-js/` submodule via `scripts/sync-ashaar-vendor.mjs`; never edit `src/vendor/` directly

### Two Main Workflows

**Table Input Mode:**
```
User sets layout → "Draw Table" → insertStructure()
  → AshaarWord.layoutTablesForTemplate() → Word.run() table with content controls
```

**Conversion Mode:**
```
Paste poetry text → renderPreview() (live Ashaar.js render in taskpane)
  → "Insert as Table"      → insertPoem() → insertNativeLayoutTables()
  → "Insert as Paragraphs" → insertTabStopPoem() → AshaarTabStop.poemToOoxml()
  → "Replace Selection"    → insertPoem(true) replaces selected Word text
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
  if (typeof module !== 'undefined') module.exports = factory();
  else root.AshaarWord = factory();
}(this, function () { ... }));
```

### Submodule Workflow
```bash
# After cloning, initialize the submodule
git submodule update --init

# To update Ashaar.js to latest upstream
npm run update:ashaar
# This runs: git submodule update --remote, then sync:ashaar, then tests
```
