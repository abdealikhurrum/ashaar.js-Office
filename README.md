# Ashaar Poetry for Word

A Microsoft Word task pane add-in for entering Arabic, Urdu, and Persian poetry blocks with Ashaar.js formatting and kashida justification.

📖 **User guide:** [Markdown](docs/USER_GUIDE.md) · [Word (.docx)](docs/USER_GUIDE.docx) — a journey-based walkthrough of every mode and control. (Regenerate the `.docx` from the Markdown with `npm run make-user-guide`.)

## What is included

- Vendored Ashaar.js browser renderer in `src/vendor`.
- A Word task pane UI in `src/taskpane`.
- Word-friendly HTML insertion that turns Ashaar parsed bayts into RTL tables.
- Table layout presets: balanced, equal, compact, and stacked.
- Ashaar content controls around inserted poem blocks, tagged with layout settings and a source hash.
- A table-first workflow that draws blank Ashaar grids from ordered misras, so users can type into Word and then justify the finished table.
- A separate Ashaar.js conversion workflow for pasting marked-up poetry and inserting a converted Word table.
- Font mode defaults to the document font; Nastaliq is available as an explicit option, not the default.
- Selection tools for replacing selected text with a formatted poetry block or applying plain-text kashida justification.

## Poetry input

```text
دل ناداں تجھے ہوا کیا ہے \ آخر اس درد کی دوا کیا ہے

یا رب وہ نہ سمجھے ہیں نہ سمجھیں گے مری بات \
دے اور دل ان کو جو نہ دے مجھ کو زباں اور
```

Ashaar.js supports `\`, `*`, or `|` between hemistiches, blank lines between stanzas, `---` between poems, and `%` for refrain marking.

## Run locally

The manifest expects the add-in at `https://localhost:3000/src/taskpane/taskpane.html`.

```sh
npm run start
```

Then open `https://localhost:3000/src/taskpane/taskpane.html`. The repo includes a local self-signed certificate for development, so your browser or Word may ask you to trust it.

## Sideload in Word

Use `manifest.xml` as the Office add-in manifest. The add-in requests `ReadWriteDocument` so it can insert or replace selected document content.

## Deploying for testers

The add-in is hosted as a static site on GitHub Pages from the `main` branch (root path). Because the whole project is static (no build step) and the taskpane loads its assets via relative paths, the live files mirror the repo layout:

- Task pane: `https://abdealikhurrum.github.io/ashaar.js-Office/src/taskpane/taskpane.html`
- Distribution manifest: `https://abdealikhurrum.github.io/ashaar.js-Office/manifest.prod.xml`

`manifest.prod.xml` is the manifest to hand to testers. It points at the live GitHub Pages URL and uses its own `<Id>` GUID (distinct from `manifest.xml`) so it can coexist with the local-dev manifest on the same machine. `manifest.xml` stays pointed at `https://localhost:3000` for local development via `npm start`.

A `.nojekyll` file at the repo root tells Pages to serve files verbatim (no Jekyll processing). Pushing to `main` rebuilds the site automatically in about a minute.

### Sideload `manifest.prod.xml`

Testers download `manifest.prod.xml` from the URL above (right-click → Save As), then:

- **Word on the web:** New doc → Home → Add-ins → More Add-ins → My Add-ins → Upload My Add-in → select the manifest.
- **Word for Mac:** Save it into `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/` (create `wef` if missing), restart Word, then Home → Add-ins → (dropdown) → Shared Folder.
- **Word for Windows:** Put it in a shared folder, then File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs → add the folder's share path → Show in Menu → restart Word → Home → Add-ins → Shared Folder.

### Cache busting after updates

Files are served with the `?v=20260527-native-layout` query string in `manifest.prod.xml`. If a JS/CSS change doesn't appear after a deploy, bump that version string in `manifest.prod.xml` and re-distribute it to force testers' clients to refetch.

## Bundled fonts

Font files in `assets/fonts/` are redistributed under their respective open licenses; each `@font-face` block in `src/taskpane/taskpane.css` documents its font's provenance inline.

- **FatemiMaqala** — vendored via the `font-fatemi` submodule (`assets/fonts/FatemiMaqala-Regular.ttf`).
- **Mehr Nastaliq Web** (`assets/fonts/MehrNastaliqWeb.woff2`) — created by Muhammad Zeeshan Nasar following the calligraphy of Nasrullah Mehr, distributed via the [`mehr`](https://www.npmjs.com/package/mehr) npm package (font v2.0/Beta 2.0). Licensed under [Creative Commons Attribution-ShareAlike 4.0 (CC BY-SA 4.0)](https://creativecommons.org/licenses/by-sa/4.0/). The `mehr` package's JS/CSS wrapper is MIT; only the font file itself is CC BY-SA.
- **Gulzar** (`assets/fonts/Gulzar-Regular.woff2`) — copyright 2021 The Gulzar Project Authors (designers: Borna Izadpanah, Fiona Ross, Alice Savoie, Simon Cozens; [github.com/googlefonts/gulzar](https://github.com/googlefonts/gulzar)), from [Google Fonts](https://fonts.google.com/specimen/Gulzar). Licensed under the [SIL Open Font License 1.1 (OFL)](https://scripts.sil.org/OFL).
- **Jameel Noori Nastaleeq (base) + Jameel Noori Nastaleeq Kasheeda (wider)** — offered as a font-mode option but not bundled; both files are private/user-supplied. `taskpane.css` has private `@font-face` blocks for both (so the taskpane's canvas can measure them), but the files themselves are gitignored and never committed. Kashida fill uses the **font-swap** mechanism: Gate G2 (2026-07-10) found Word applies Jameel's elongated forms via a named "Kasheeda" style within the family, not via italic runs, so `justifySelection` measures each connected segment (fasl) in the base vs Kasheeda face and swaps the highest-gain fasls to the wider face until the misra fills the column (`src/taskpane/kashida-fontswap.js`, dispatched in `src/taskpane/taskpane.js`). Fill only works if the reader also has both faces installed — see the reader-requirements note below.

## Fonts & reader requirements

The add-in previews poetry in the taskpane and inserts it into a Word document; the fonts each side needs are not the same.

- **Bundled in the taskpane (preview only)**: Noto Nastaliq Urdu, Mehr Nastaliq Web (`assets/fonts/MehrNastaliqWeb.woff2`), Gulzar (`assets/fonts/Gulzar-Regular.woff2`), and FatemiMaqala are loaded via `@font-face` in `src/taskpane/taskpane.css` and render inside the pane for anyone, on any machine, with no install.
- **Word documents render with the document's own installed fonts.** The taskpane only writes a font *name* into the OOXML (`<w:rFonts w:cs="…">`); Word does not embed the font file. This means each font mode has a different reader-side requirement, shown as a note under the Font picker (`#font-install-note` in `src/taskpane/taskpane.html`, wired by `updateFontNote()` in `src/taskpane/taskpane.js`):
  - **Document default / Arabic serif** — no note; these rely on whatever font is already applied in the document, or a common system fallback.
  - **Mehr Nastaliq** — note shown; the reader must have "Mehr Nastaliq Web" installed on their machine to see the kashida-stretched glyphs correctly in Word.
  - **Jameel Noori Kasheeda** — note shown, and it names BOTH faces: the reader needs "Jameel Noori Nastaleeq" (Regular + Kasheeda) installed to see this correctly. Mechanism is `font-swap`: kashida fill measures each fasl in the base face vs the Kasheeda face and swaps the highest-gain fasls to the wider face until the misra fills the column (see `src/taskpane/kashida-fontswap.js`); it only fills visibly if the reader has both faces installed — otherwise Word substitutes a fallback font for the missing `w:cs` name and the swap has no visual effect.
  - **Gulzar** — no reader-install note today (its `mechanism` is `"whitespace"`, the same bucket as the no-install fonts), even though the taskpane preview does load a Gulzar font file. Readers without Gulzar installed will see Word's fallback font instead.
- **Jameel Noori Nastaleeq (base + Kasheeda) is private and is never deployed publicly.** The add-in is served as a static site from GitHub Pages, so anything committed to `assets/fonts/` is world-readable at a public URL. The Jameel font files' license does not permit that kind of public rehosting, so:
  - `assets/fonts/JameelNooriNastaleeq*` (matching both the base `-Regular.ttf` and the `Kasheeda.ttf` file) is listed in `.gitignore` — neither can be `git add`ed by accident, tracked in `main`, or published to Pages.
  - The font mode still appears in the picker; `taskpane.css` has private `@font-face` blocks for both faces so the taskpane's own canvas can measure them, but the actual `.ttf` files are meant to be dropped into `assets/fonts/` locally (or loaded via the in-app "Load a font from your computer" custom-font flow) for development and for testers who already hold a licensed copy — never checked in.
  - To confirm the guard after a deploy, request either file's Pages URL directly (e.g. `https://<user>.github.io/<repo>/assets/fonts/JameelNooriNastaleeqKasheeda.ttf`) and confirm it 404s.

## Development Notes

The preview pane uses Ashaar.js native HTML/CSS. Word insertion uses generated table HTML because Word's HTML importer preserves table layout more reliably than browser flex layout.

Ashaar.js upstream is tracked as a git submodule at `vendor/ashaar-js`. To refresh the vendored browser files from the pinned submodule commit:

```sh
npm run sync:ashaar
```

To fetch the latest upstream `master`, update the submodule checkout, sync the files, and run tests:

```sh
npm run update:ashaar
```

The currently synced upstream commit is recorded in `src/vendor/ASHAAR_UPSTREAM_VERSION`.

The test fixtures in `test-documents/` are generated with:

```sh
/Users/abdealikhurrum/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 scripts/generate_test_docs.py
```

Each fixture uses fixed Word table grids and wraps each Ashaar block in a rich text content control tagged with `ashaar:` metadata.
