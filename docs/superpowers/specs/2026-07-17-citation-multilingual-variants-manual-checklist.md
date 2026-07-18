# SP-3 Multilingual Variants — Manual Verification Checklist

Live checks that unit tests can't cover: the Zotero local-API migration and the Word render path.

## Component B — migration (live Zotero)

**Enable the local API first:** Zotero → Settings → Advanced → "Allow other applications on this
computer to communicate with Zotero", then restart Zotero.

**Note:** the Zotero local API is currently **read-only** (PATCH returns `501 Not Implemented`), so
`--write` does not work on current Zotero. Use the `--emit-snippet` path, which converts in Node
(tested code) and applies via Zotero's own Run JavaScript (full write access).

- [ ] Local API reachable: `curl "http://localhost:23119/api/users/0/items?limit=1&format=json"`
      → JSON array with `key`, `version`, `data.extra` (not `403 "Local API is not enabled"`).
- [ ] `npm run migrate:cne` (dry-run) prints per-item `+ cne-*` diffs and `N item(s) to convert`.
      Direction check: Arabic-primary items emit `-romanized`; Latin-primary items whose variant is
      Arabic script emit `-original` (e.g. `Eat Not this Flesh` → `cne-author-0-last-original: سمونس`).
- [ ] `npm run migrate:cne -- --emit-snippet` writes `scratch/migrate-cne-snippet.js`.
- [ ] In Zotero: Tools → Developer → Run JavaScript → paste the snippet → Run →
      returns `cne-migrate: updated N, failed 0`.
- [ ] Spot-check an item's Extra in Zotero: the `cne-*` lines are present and the `mlzsync1:` block
      is still there (non-destructive).
- [ ] (If a future Zotero adds local-API write support: `npm run migrate:cne -- --write` becomes the
      one-command path, with `scratch/mlzsync-backup.json` written first.)

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
