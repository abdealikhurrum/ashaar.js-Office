# SP-2 — Live Zotero Connection (design)

**Date:** 2026-07-17
**Status:** Approved design, pre-implementation
**Sub-project:** SP-2 of the Annotation & Citation initiative (follows SP-1, see
`docs/superpowers/specs/2026-07-16-citation-engine-core-design.md`)

## Summary

Replace SP-1's CSL-JSON fixture with the user's **live Zotero 9 + Better BibTeX (BBT)**
library. Selection uses **CAYW** (Zotero's native cite-as-you-write picker); item data is
fetched as **CSL-JSON via BBT's JSON-RPC** and formatted by the **existing SP-1 engine** (so the
bidi/Arabic/CSL-M rendering all still apply). The Office WebView's HTTPS→HTTP-localhost
mixed-content problem is dissolved by proxying through the add-in's own `server.mjs`: the pane
talks same-origin HTTPS to the dev server, and the server talks HTTP to `localhost:23119`.

## Grounding (verified live on this machine, 2026-07-17)

- Zotero is running: `GET http://localhost:23119/connector/ping` → **200**.
- Zotero 9's **built-in local API is disabled**: `GET /api/` → **403 "Local API is not enabled"**
  (would require a manual toggle in Zotero settings). We therefore do **not** depend on it.
- **Better BibTeX is live and sufficient**:
  - `GET /better-bibtex/cayw?probe=probe` → **200 `ready`**
  - `POST /better-bibtex/json-rpc {method:"user.groups"}` → `{"result":[{"id":1,"name":"My Library"}]}`

So BBT (CAYW + JSON-RPC) is the working, richer path and the basis of this design.

## Decisions fixed (from brainstorming)

- **Target library:** Zotero 9 + Better BibTeX (the user has migrated off Juris-M).
  Consequence: the library carries **no MLZ `multi` multilingual variants** — that gap is SP-3's
  problem, not SP-2's. SP-2 passes whatever CSL-JSON BBT returns straight through.
- **Data path:** BBT endpoints via a `server.mjs` reverse proxy (not the Zotero built-in API,
  which is off; not the Zotero Web API, which drops Arabic fidelity and needs the cloud).
- **Selection UX:** CAYW (Zotero's native picker). The pane requests selection; Zotero pops its
  own quick-search bar; the picked citekeys return to the pane, and **our SP-1 engine formats
  them** (CAYW is used for selection only, not for formatting).
- **Offline behavior:** when Zotero isn't reachable, the tab **falls back to the bundled
  fixture** so it still works for demo/testing.

## Scope

**In scope**
- `server.mjs` `/zotero/*` reverse proxy to `localhost:23119` (ping, CAYW, JSON-RPC).
- `src/taskpane/cite-zotero.js` — client module: availability ping, CAYW pick → citekeys,
  fetch CSL-JSON via BBT `item.export`, and an in-memory citekey→CSL-JSON cache.
- Cite-tab wiring: an **"Add from Zotero"** action that picks via CAYW, fetches CSL-JSON, feeds
  the existing `CiteEngine`, and inserts through the existing SP-1 path (footnote/endnote/
  bibliography, with the bidi + Arabic-punctuation treatment).
- Availability handling + fixture fallback; node-`assert` tests for the pure logic.

**Explicitly deferred**
- Multilingual variant sourcing and the recovered Arabic `langPrefs` preset (persons/institutions/
  places→translit, titles/journals/publishers→orig) — SP-3. SP-2 does not synthesize `multi`.
- Source classification / Fatemi grouping (SP-4), Hijri dates (SP-5), richer annotations (SP-6).
- Recreating the `jm-chicago-fullnote-bibliography-Arabic` (Jāmiʿa) style + preset — parked,
  pending the user's referencing-guidelines document.
- Wiring `citationTag` into inserted citations + hardening its delimiter — carried over from
  SP-1's final review as an SP-2-adjacent follow-up; may be folded in but is tracked separately.

## Architecture & components

The engine still runs client-side; `server.mjs` gains only a thin reverse-proxy responsibility.

### `server.mjs` — `/zotero/*` reverse proxy (added before static serving)

`server.mjs` is currently a ~47-line static HTTPS file server (`createServer` → map URL to file →
stream). Add a branch at the top of the request handler: if the request path starts with
`/zotero/`, proxy to Zotero; otherwise fall through to the existing static logic. The proxy makes
a Node `http` request to `http://localhost:23119` and pipes the response back. Routes:

- `GET /zotero/ping` → `http://localhost:23119/connector/ping` — availability probe.
- `GET /zotero/cayw?<qs>` → `http://localhost:23119/better-bibtex/cayw?<qs>` with
  `format=citekeys` (and `probe=probe` supported). **Long-poll**: BBT does not respond until the
  user finishes (or cancels) picking in Zotero, so the proxy must use a long (effectively
  unbounded) timeout and stream the eventual response back.
- `POST /zotero/json-rpc` → `http://localhost:23119/better-bibtex/json-rpc` — body forwarded
  verbatim (used for `item.export` → CSL-JSON, and any `item.search`/introspection).

The route decision (`is this path a `/zotero/*` proxy target?`) is factored into a small **pure
function** so it is unit-testable without a live socket. If `localhost:23119` is unreachable, the
proxy responds `502` with a small JSON error body the pane can surface.

### `src/taskpane/cite-zotero.js` (UMD, Node-testable core)

Talks only to the same-origin dev server (`fetch("/zotero/…")`), never directly to `:23119`.

- `ping(fetchImpl?)` → `Promise<boolean>` — true when `/zotero/ping` returns ok.
- `caywPick(fetchImpl?)` → `Promise<string[]>` — GET `/zotero/cayw?format=citekeys`, parse the
  response into an array of citekeys (empty array when the user cancels).
- `fetchCslJson(citekeys, fetchImpl?)` → `Promise<Object>` — POST `/zotero/json-rpc` with
  `{jsonrpc:"2.0", method:"item.export", params:[citekeys, <CSL-JSON translator>], id:1}`, parse
  `result` into a CSL-JSON **items map keyed by `id`** (the shape `CiteEngine.build({items})`
  expects), consulting/populating an in-memory `citekey → item` cache.
- The `fetchImpl` parameter (defaulting to global `fetch`) is the injection seam that makes the
  request-building and response-parsing unit-testable with a fake fetch — no live Zotero in tests.

Request-building and response-parsing are pure and separately exported (e.g. `buildExportRequest`,
`parseExportResult`, `parseCaywResult`) so tests assert on them directly.

### Cite-tab wiring (`cite-pane.js`)

- An **"Add from Zotero"** button. On click: set status "Picking in Zotero…", `caywPick()`, then
  `fetchCslJson()`, merge the returned items into the tab's working item set, re-render the item
  list + preview, and let the existing Insert path run.
- On tab-show: `ping()`. If false, show the "Start Zotero (with Better BibTeX)…" status and keep
  the bundled fixture loaded so the tab still functions.

## Data flow

```
"Add from Zotero"
  → GET /zotero/cayw?format=citekeys        (pane → our HTTPS server, same origin)
      → server → GET :23119/better-bibtex/cayw   (Zotero pops native picker; long-poll)
      → user selects → citekeys returned
  → POST /zotero/json-rpc {item.export, [citekeys, csljson]}
      → server → :23119 BBT → CSL-JSON
  → items map (cached) → CiteEngine.build({items}) → preview
  → Insert → existing SP-1 footnote/endnote/bibliography path (bidi + Arabic punctuation)
```

## Error handling / availability

- Zotero down / proxy 502 → pane status "Start Zotero (with Better BibTeX) to cite from your
  library"; fixture remains available.
- CAYW returns empty (cancelled) → no-op, status cleared.
- JSON-RPC error object in the response → surfaced in the pane status.
- CAYW long-poll: generous server timeout; pane shows an in-progress status so it doesn't look
  hung.

## Testing (node-`assert`, added to the `npm test` chain)

- `tests/cite-zotero.test.js` — inject a fake `fetch`: `buildExportRequest` produces the correct
  JSON-RPC body; `parseExportResult` turns a representative BBT response into an id-keyed items
  map; `parseCaywResult` extracts citekeys (incl. the empty/cancel case); cache returns the cached
  item without a second fetch.
- `tests/server-proxy.test.js` (or fold into an existing server test) — the pure route-decision
  function: `/zotero/ping|cayw|json-rpc` → proxy target; everything else → static.
- Manual checklist (needs Zotero running): CAYW picks a real item → CSL-JSON fetched → footnote
  inserts with correct (Arabic where applicable) formatting; Zotero-off path shows the fallback.

## Risks

- **CAYW response format**: `format=citekeys` shape must be parsed robustly (delimiter/empty
  cases) — covered by `parseCaywResult` tests; if citekeys prove insufficient, the fallback is
  CAYW `format=translate&translator=<csljson>` to get item data in one call (documented
  alternative, not the default).
- **BBT `item.export` translator name** for CSL-JSON must be confirmed against the live BBT during
  implementation (BBT exposes a CSL-JSON translator; the exact identifier is verified in the first
  task, mirroring how SP-1 discovered the CSL-M pref method from the vendored source).
- **Long-poll proxy** must not be killed by a default socket timeout — set explicitly.

## Success criteria

With Zotero 9 + BBT running: clicking "Add from Zotero" pops Zotero's picker, and the selected
item(s) render in the pane preview and insert into Word as a correctly-formatted footnote +
bibliography (Arabic items get the bidi/Arabic-comma treatment) — all without enabling Zotero's
built-in local API and without any mixed-content/CORS error. With Zotero closed, the tab shows the
fallback message and still works off the fixture. `npm test` passes including the new tests.
