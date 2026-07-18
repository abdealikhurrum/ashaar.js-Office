# SP-3 Multilingual Variants — Manual Verification Checklist

Live checks that unit tests can't cover: the Zotero local-API migration and the Word render path.

## Component B — migration (live Zotero)

- [ ] Zotero open; confirm the local API is exposed and returns item JSON with `data.extra`:
      `curl "http://localhost:23119/api/users/0/items?limit=1&format=json"` → JSON array with `key`, `version`, `data.extra`.
      (If this 404s or returns HTML, the local API / write path is unavailable — fall back to the
      export/paste path noted in the design's Risks section.)
- [ ] `npm run migrate:cne` (dry-run) prints per-item `+ cne-*` diffs and `N item(s) to convert`,
      writing nothing. The `Uyun al-Akhbar` book shows:
      `cne-title-romanized: Uyun al-Akhbar Vol. 4` and
      `cne-author-0-last-romanized: al-Dai al-Ajal Syedna Idris Imaduddin RA`.
- [ ] `npm run migrate:cne -- --write` applies. Confirm `scratch/mlzsync-backup.json` exists and the
      Zotero item's Extra now shows the `cne-*` lines (the `mlzsync1:` block is still present).
- [ ] Re-run `npm run migrate:cne -- --write` → reports items as already migrated (idempotent;
      `converted=0`).
- [ ] (Optional) `--strip-mlzsync` removes the legacy blob after you've verified the migration.

## Component A — feature (live Word)

- [ ] Cite tab shows the **Variant** dropdown (Original (ar) / Romanized / Both), between Locale and
      Output form.
- [ ] "Add from Zotero" a migrated Arabic item; the preview updates without error.
- [ ] Variant = **Romanized** → the footnote/bibliography preview renders the romanized title +
      author (e.g. "Uyun al-Akhbar…", "al-Dai al-Ajal…").
- [ ] Variant = **Original** → renders Arabic script; Variant = **Both** → Arabic + romanized.
- [ ] Insert a footnote for the item; re-open the document / inspect the content-control tag →
      the `AshaarCite:` payload stores `variant` (v2).
- [ ] Change the Variant dropdown, click **Refresh citations** → previously inserted notes + the
      bibliography re-format in place to the new variant.
- [ ] An item with **no** variants still cites correctly (Arabic real fields) under every Variant
      setting (no blank output — langPrefs fallback).
- [ ] Bump check: `window.ASHAAR_ASSET_VERSION` is `20260717-cite-variants` (installed users get
      fresh JS).
