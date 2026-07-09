# Custom Font Uploader — Design

**Date:** 2026-07-09
**Status:** Approved, implementing
**Scope:** The font uploader only. Multi-font/styled *runs within a single misra* (per-run measurement + run-preserving write-back) is a **separate later project** and explicitly out of scope here.

## Problem

The kashida justification engine measures glyph widths on an offscreen `<canvas>` in the task-pane WebView, then Word renders with the real font. Clean lines require the canvas to measure the *same* font. Fonts are only measurable if the WebView can load them.

Today only **FatemiMaqala** is bundled as an `@font-face`. A user's own calligraphic font (e.g. **AL-KANZ**) is applied in their document but is **not resolvable in the pane** — WebKit (Mac Word's WKWebView) deliberately hides arbitrary *installed* fonts from CSS/canvas for anti-fingerprinting. So the canvas silently measures a fallback, the tatweel counts are computed for the wrong outlines, and lines come out ragged.

`Check font at cursor` already reports this: for AL-KANZ it says *"NOT resolvable here — metrics are approximate."*

## Key insight

A font loaded from **bytes** via the `FontFace` API is an *explicitly provided* web font — the same category as a bundled `@font-face url()` — and is therefore **not** subject to WebKit's installed-font privacy block. It is measurable in WebView2 (Windows), WKWebView (Mac), and browsers (Word on the web). So letting the user supply their own font file makes it measurable, without us redistributing anything.

## Goals / non-goals

- **Goal:** Let a user load a font file from their machine so justification measures the real outlines. Persist it so it's a one-time action per machine. No redistribution (the user brings the font).
- **Non-goal:** Multi-font / bold / size runs *within* a misra (deferred). Bundling fonts into the add-in. Any change to the vendored justify engine.

## User experience

A new **"Custom fonts"** area in the task pane:

1. File picker accepting `.ttf .otf .woff .woff2`.
2. On file select, auto-detect the font's family name from its `name` table and prefill an **editable "Register as" field**. This name **must match what Word reports** for the text (`range.font.name`), which is what the measurement path keys on — hence editable.
3. **Add** button → load + persist the font.
4. **Verify at cursor**: with the cursor in the styled text, compares the registered name against `range.font.name` (reusing the existing `checkQaseedaFont` probe) → ✓ match / ✗ mismatch, so the user can correct the name before relying on it.
5. **List** of loaded fonts, each with a remove (×).
6. Helper note: *"Loaded from your computer so kashida metrics match Word. Stored on this machine for future sessions — nothing is uploaded to a server."*

## Architecture

### New UMD module: `src/taskpane/font-store.js` (`AshaarFontStore`)

Follows the repo's UMD pattern (works in Node for tests and in the browser).

**Pure (testable in Node):**
- `parseNames(buffer) → { family, fullName, postScript } | null`
  Parses the sfnt (`TTF`/`OTF`) `name` table. Prefers nameID 16 (Typographic Family) then nameID 1 (Family); also returns nameID 4 (Full) and 6 (PostScript). Decodes Windows (platform 3, UTF-16BE) and Mac (platform 1, Latin-1) records, preferring Windows. Accepts `ArrayBuffer` / typed-array / `DataView`. Returns `null` for non-sfnt inputs (WOFF/WOFF2/TTC) — those are wrapped/compressed, so auto-detect falls back to the filename and the user edits the name.

**Browser-only (guarded by `typeof document` / `indexedDB`):**
- `loadFont(family, buffer) → Promise` — `new FontFace(family, buffer)` → `.load()` → `document.fonts.add(ff)`.
- IndexedDB (`ashaar-fonts` DB, `fonts` store, key = family): `saveFont(family, filename, buffer)`, `listFonts()`, `deleteFont(family)`.
- `registerAll() → Promise<Array<{family, filename}>>` — read all stored fonts, `loadFont` each into `document.fonts`. Called on Office-ready.
- `addUserFont(family, file) → Promise` — `file.arrayBuffer()` → `loadFont` → `saveFont`.

### Wiring: `src/taskpane/taskpane.js`
- On Office-ready, `await AshaarFontStore.registerAll()` **before** any justify/measurement can run, so stored fonts are already in `document.fonts`.
- UI handlers: file-select (auto-detect name via `parseNames`, prefill field), Add, Remove, Verify-at-cursor (reuse the cursor-name probe).
- **No change to the justify/measurement code** — it already `await document.fonts.load(repName)` and keys on the Word font name. A font registered under that name is picked up automatically.
- The existing per-font `fontCorrections` knob is untouched (still useful for fonts a user hasn't uploaded).

### `src/taskpane/taskpane.html`
- Add `font-store.js` to the ordered script loader list (before `taskpane.js`).
- Add the "Custom fonts" UI block.

## Persistence

IndexedDB `ashaar-fonts`, store `fonts`, records `{ family, filename, bytes(ArrayBuffer) }`, keyed by `family`. **Global for the machine** — loaded once, auto-re-registered on every future session and document until removed. (localStorage is already used for templates but is unsuited to binary font blobs; IndexedDB is the right store.)

## Error handling / edge cases

- Corrupt/invalid font → `FontFace.load()` rejects → surface an error, persist nothing.
- Duplicate family → overwrite (re-upload updates the stored bytes).
- No IndexedDB (private mode) → load for the session only + warn it won't persist.
- Undetectable name (WOFF/WOFF2/TTC) → prefill from filename; user confirms/edits.
- Name mismatch with Word → `Verify at cursor` flags it; measurement silently falls back if unmatched, so the verify step is the guard.
- Harvested/user-provided family names are only ever used as `FontFace` family strings, IndexedDB keys, and DOM text — never interpolated into shell/JSON.

## Testing

- **Unit (Node, `tests/font-store.test.js`):** `parseNames` against the repo's own `assets/fonts/FatemiMaqala-Regular.ttf` → `family === "FatemiMaqala"`; and `vendor/font-fatemi/alfatemi/AlFatemi-Regular.ttf` → `family === "AlFatemi"`. Assert `null` for a non-sfnt buffer. Deterministic, no external files. Added to the `npm test` script.
- **Manual (in Word, checklist provided):** load AL-KANZ → `Check font at cursor` flips to "resolves" → justify a poem → lines clean; reload the add-in → font still registered (persistence).

## Out of scope (next project)

Per-run measurement and run-preserving (OOXML) write-back so mixed fonts/sizes/bold within one misra justify correctly and keep their styling. Tracked separately.
