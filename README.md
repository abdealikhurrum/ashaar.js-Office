# Ashaar Poetry for Word

A Microsoft Word task pane add-in for entering Arabic, Urdu, and Persian poetry blocks with Ashaar.js formatting and kashida justification.

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
