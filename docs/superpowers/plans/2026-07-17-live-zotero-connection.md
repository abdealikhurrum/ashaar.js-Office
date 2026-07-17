# SP-2 — Live Zotero Connection (implementation plan)

**Date:** 2026-07-17
**Design spec:** `docs/superpowers/specs/2026-07-17-live-zotero-connection-design.md` (approved)
**Builds on:** SP-1 citation engine (complete) — `cite-engine.js`, `cite-word.js`, `cite-pane.js`.

Replace SP-1's static fixture-only citation flow with a live path to the user's
**Zotero 9 + Better BibTeX (BBT)** library: select via CAYW, fetch CSL-JSON via BBT
JSON-RPC, format with the existing SP-1 engine, insert via the existing SP-1 Word path.
The Office WebView HTTPS→HTTP-localhost mixed-content problem is dissolved by proxying
`/zotero/*` through the add-in's own `server.mjs`.

## Live-verified facts (probed against the running Zotero on this machine, 2026-07-17)

These are ground truth — use them verbatim; do not re-guess them.

- Zotero base URL: `http://localhost:23119`.
- `GET /connector/ping` → `200` (availability probe).
- `GET /better-bibtex/cayw?probe=probe` → `200`, body is the plain text `ready`.
- JSON-RPC endpoint: `POST /better-bibtex/json-rpc`, `Content-Type: application/json`.
  - Success response: `{"jsonrpc":"2.0","result":<...>,"id":1}`.
  - Error response: `{"jsonrpc":"2.0","error":{"code":<int>,"message":"<str>"},"id":null}`
    (note: `id` is `null` on error).
- **CSL-JSON translator name is `"Better CSL JSON"`.** Stock `"CSL JSON"` does NOT
  resolve (`error.message`: `getTranslatorId: "CSL JSON" could not be resolved to a
  translator`). "Better CSL JSON" sets each item's `id` **equal to its citekey**, which is
  exactly what `CiteEngine`'s `retrieveItem(id)` / `cite([citekeys])` require.
- `item.export` call shape:
  `{"jsonrpc":"2.0","method":"item.export","params":[["<citekey>",...],"Better CSL JSON"],"id":1}`.
  - The `result` field is a **JSON *string*** (not a parsed object), e.g.:
    `"[\n  {\"id\":\"YaumulMabasUyun\",\"citation-key\":\"YaumulMabasUyun\",\"title\":\"Yaumul Mabas - Uyun 1\",\"type\":\"document\"}\n]\n"`.
    The parser must `JSON.parse` the string, then reduce the resulting array to an
    `id`-keyed object map (the shape `CiteEngine.build({items})` expects).
- `item.search("<term>")` returns `{result:[{id,"citation-key",title,library,citekey}, ...]}`
  (used only for introspection; not on the SP-2 hot path).

## Global Constraints

Every task must honor these. The reviewer's attention lens copies from here.

- **No new runtime dependencies.** Node built-ins only in `server.mjs` (`node:http` for the
  proxy). No npm packages. No build step (the repo has none).
- **Zotero base URL is `http://localhost:23119`** and appears in exactly one place
  (`zotero-proxy.js`), not scattered as string literals.
- **The pane never talks to `:23119` directly** — `cite-zotero.js` only ever calls
  `fetch("/zotero/…")` (same-origin HTTPS). Only `server.mjs` reaches `:23119`.
- **CSL-JSON translator string is exactly `"Better CSL JSON"`** wherever `item.export` is built.
- **Fixture fallback is preserved.** When Zotero is unreachable, the Cite tab keeps working
  off the bundled `fixtures/cite-sample.json` exactly as SP-1 does today — the live path is
  strictly additive. No SP-1 behavior regresses.
- **Pure logic is separated from I/O and unit-tested with `node`'s `assert`** (no live socket,
  no live Zotero, no DOM in tests). I/O seams take an injectable `fetchImpl` (default: global
  `fetch`); the proxy's route decision is a pure function.
- **UMD/module conventions match the codebase**: `.js` files are CommonJS (no `"type":"module"`
  in `package.json`); `server.mjs` is ESM. Client modules that also run under `node` tests use
  the same UMD pattern as `cite-engine.js` (`module.exports` in node, `window.X` in browser).
- **Tests are added to the `npm test` chain** in `package.json`.
- **Explicitly OUT of scope for SP-2** (do not build): multilingual `multi`/`langPrefs`
  synthesis (SP-3), Fatemi source classification (SP-4), Hijri dates (SP-5), the Jāmiʿa style,
  and the SP-1 citation-tagging follow-up (`citationTag` on inserted citations + delimiter
  hardening) — that is tracked separately per the design spec and must NOT be folded in here.

---

## Task 1: `server.mjs` `/zotero/*` reverse proxy + pure route function

**Goal:** Add a `/zotero/*` reverse proxy to `server.mjs` so the pane can reach Zotero
same-origin over HTTPS while the server speaks HTTP to `localhost:23119`. Factor the
route-decision into a pure, separately-testable CommonJS module.

**Files:**
- **Create `zotero-proxy.js`** (repo root, next to `server.mjs`; CommonJS —
  `module.exports = { ... }`). Pure, no sockets, no `node:http`. Exports:
  - `ZOTERO_BASE` = `"http://localhost:23119"` (the single source of the base URL).
  - `zoteroProxyTarget(pathname)` → the upstream **pathname** to proxy to, or `null` when the
    request is not a Zotero proxy target (so `server.mjs` falls through to static serving).
    Mapping (match on the pathname only; the caller re-attaches the original query string):
    - `/zotero/ping`      → `/connector/ping`
    - `/zotero/cayw`      → `/better-bibtex/cayw`
    - `/zotero/json-rpc`  → `/better-bibtex/json-rpc`
    - anything else (incl. `/zotero/` alone, `/zotero/unknown`, `/src/...`, `/`) → `null`.
    Keep it a clean lookup (exact-match table). Do not prefix-match `/zotero/` loosely — an
    unknown `/zotero/*` path must return `null`, not a partial proxy.
- **Edit `server.mjs`**: `import zoteroProxy from "./zotero-proxy.js";` (default import of the
  CJS module — Node supports this). At the **top of the request handler**, before the existing
  static-file logic:
  1. Parse the request pathname (`new URL(req.url, "https://localhost:" + port).pathname`).
  2. `const target = zoteroProxy.zoteroProxyTarget(pathname);`
  3. If `target === null`, fall through to the existing static logic (unchanged).
  4. Otherwise proxy: build the upstream URL = `ZOTERO_BASE + target + search` (preserve the
     original query string), issue a `node:http` request with `req.method` and forward the
     `content-type` header; pipe the request body (`req.pipe(upstreamReq)`) for POST; on the
     upstream response, copy its status code + `content-type` and `upstreamRes.pipe(res)`.
  5. **Long-poll:** CAYW does not respond until the user finishes/cancels the picker. Do NOT
     impose a short timeout on the upstream request — set no timeout (or an effectively
     unbounded one) so the eventual response streams back. (The other two routes respond fast;
     a single no-timeout path is simplest and correct for all three.)
  6. **Unreachable Zotero:** on the upstream request's `"error"` event (e.g. `ECONNREFUSED`
     when Zotero is closed), respond `502` with JSON body
     `{"error":"zotero-unreachable","detail":"<err.message>"}` and
     `Content-Type: application/json`. Guard against double-writing headers if the response
     already started.
  - Keep the existing static path, cert loading, MIME map, and `filePathFor` exactly as they
    are. The proxy branch is purely additive at the top of the handler.

**Tests — create `tests/server-proxy.test.js`** (`require("../zotero-proxy.js")`, `node:assert`):
- `zoteroProxyTarget("/zotero/ping")` === `"/connector/ping"`.
- `zoteroProxyTarget("/zotero/cayw")` === `"/better-bibtex/cayw"`.
- `zoteroProxyTarget("/zotero/json-rpc")` === `"/better-bibtex/json-rpc"`.
- `zoteroProxyTarget("/zotero/")`, `"/zotero/unknown"`, `"/src/taskpane/taskpane.html"`, `"/"`
  each === `null`.
- `ZOTERO_BASE` === `"http://localhost:23119"`.
- Add `node tests/server-proxy.test.js` to the `test` script in `package.json`.

**Acceptance:** `node tests/server-proxy.test.js` passes; `server.mjs` still boots and serves
static files unchanged; the route decision is covered without opening a socket. (Live proxy
behavior against Zotero is on the manual checklist — not unit-testable.)

---

## Task 2: `src/taskpane/cite-zotero.js` client module + tests

**Goal:** A UMD client module that talks only to the same-origin `/zotero/*` proxy, with all
request-building and response-parsing factored into pure, exported functions so tests run with
a fake `fetch` and no live Zotero.

**Files:**
- **Create `src/taskpane/cite-zotero.js`** — UMD, matching `cite-engine.js`'s wrapper
  (`module.exports` in node; `root.CiteZotero = ...` in browser). No DOM, no Office.js. Exports:

  Pure (no I/O — separately exported and directly unit-tested):
  - `buildExportRequest(citekeys)` → the JSON-RPC **payload object**
    `{ jsonrpc: "2.0", method: "item.export", params: [citekeys, "Better CSL JSON"], id: 1 }`.
    (The caller `JSON.stringify`s it for the POST body.)
  - `parseExportResult(rpcResponse, citekeys)` → an `id`-keyed items map.
    - If `rpcResponse.error` is present, `throw new Error(rpcResponse.error.message)`.
    - `rpcResponse.result` is a JSON **string** → `JSON.parse` it → expect an array of CSL-JSON
      objects → reduce to `{ [item.id]: item }`.
    - Defensive: if a requested citekey from `citekeys` is absent from the resulting map keys
      but the array has a matching-index entry, also key that entry by the requested citekey
      (guards against any translator that doesn't set `id`=citekey). With "Better CSL JSON"
      the primary `item.id` keying already matches; this is a belt-and-suspenders fallback.
  - `parseCaywResult(text)` → `string[]` of citekeys.
    - `""` / whitespace-only / nullish → `[]` (user cancelled).
    - Otherwise: strip a single pair of surrounding `{}` or `[]` if present; strip a leading
      `@` from each token; split on `[\s,;]+`; trim; drop empties.

  I/O (each takes optional `fetchImpl`, default = global `fetch`; talk only to `/zotero/…`):
  - `ping(fetchImpl?)` → `Promise<boolean>` — `GET /zotero/ping`; resolve `true` iff `res.ok`;
    resolve `false` on any thrown error (never reject).
  - `caywPick(fetchImpl?)` → `Promise<string[]>` — `GET /zotero/cayw?format=citekeys`, read
    `res.text()`, return `parseCaywResult(text)`.
  - `fetchCslJson(citekeys, fetchImpl?)` → `Promise<Object>` (id-keyed map). Consult an
    in-memory `citekey → item` cache: only POST for citekeys not already cached; POST
    `/zotero/json-rpc` with `JSON.stringify(buildExportRequest(missing))` and
    `Content-Type: application/json`; parse via `parseExportResult(await res.json(), missing)`;
    merge new items into the cache; return a map containing every requested citekey's item.
  - `clearCache()` → resets the in-memory cache (used by tests; harmless in prod).

**Tests — create `tests/cite-zotero.test.js`** (`require("../src/taskpane/cite-zotero.js")`,
`node:assert`; inject a fake `fetch`):
- `buildExportRequest(["YaumulMabasUyun"])` deep-equals
  `{jsonrpc:"2.0",method:"item.export",params:[["YaumulMabasUyun"],"Better CSL JSON"],id:1}`.
- `parseExportResult` turns a representative live-shaped response
  `{jsonrpc:"2.0",result:"[\n  {\"id\":\"YaumulMabasUyun\",\"citation-key\":\"YaumulMabasUyun\",\"title\":\"Yaumul Mabas - Uyun 1\",\"type\":\"document\"}\n]\n",id:1}`
  into `{ YaumulMabasUyun: {id:"YaumulMabasUyun", ...} }`.
- `parseExportResult` throws with the RPC message on
  `{jsonrpc:"2.0",error:{code:-32603,message:"boom"},id:null}`.
- `parseCaywResult`: `"YaumulMabasUyun,IsraaWalMiraaj"` → `["YaumulMabasUyun","IsraaWalMiraaj"]`;
  `""` → `[]`; a braced/`@`-prefixed variant (e.g. `"{@YaumulMabasUyun}"`) → `["YaumulMabasUyun"]`.
- **Cache:** a fake `fetch` counter proves the 2nd `fetchCslJson([sameKey])` does **not** issue
  a second network call (served from cache); a mixed call (one cached + one new) fetches only
  the new key.
- `ping` returns `true` for `{ok:true}` fake and `false` when the fake `fetch` throws.
- Add `node tests/cite-zotero.test.js` to the `test` script in `package.json`.

**Acceptance:** `node tests/cite-zotero.test.js` passes; no DOM/Office.js/`:23119` references in
the module (`grep`-clean); request/response shapes match the live-verified facts above.

---

## Task 3: Cite-tab wiring — "Add from Zotero"

**Goal:** Wire the live path into the existing Cite tab: an **"Add from Zotero"** action that
picks via CAYW, fetches CSL-JSON, merges into the tab's working item set, and re-renders +
inserts through the existing SP-1 path. Availability ping on tab-show with fixture fallback.

**Interfaces from earlier tasks (do not re-derive):**
- `CiteZotero.ping()`, `CiteZotero.caywPick()`, `CiteZotero.fetchCslJson(citekeys)` — from Task 2.
- Existing `cite-pane.js` internals you will extend (already present): `cache.items` (the
  id-keyed items map, seeded from `fixtures/cite-sample.json`), `itemsPopulated` (a one-shot
  flag), `populateItems()`, `renderPreview()`, `selectedIds()`, `setStatus(msg, warn)`,
  `bind()`, `onTabShown()`. `CiteEngine.build({items})` reads `cache.items`.

**Files:**
- **Edit `src/taskpane/cite-pane.js`:**
  - Add an **"Add from Zotero"** handler (`#cite-add-zotero`):
    1. `setStatus("Picking in Zotero…")`.
    2. `CiteZotero.caywPick()`. If it returns `[]` (cancelled) → `setStatus("")` and return
       (no-op).
    3. `CiteZotero.fetchCslJson(citekeys)` → merge the returned items into `cache.items`
       (initialize `cache.items` to `{}` first if it is null). 
    4. Re-render the item list so the new items appear **and are checked**: reset the one-shot
       guard (`itemsPopulated = false`) and re-run `populateItems()`, then check the freshly
       added citekeys' checkboxes; call `renderPreview()`.
    5. `setStatus("Added " + citekeys.length + " item(s) from Zotero.")` (or clear).
    6. On any error (proxy 502 / JSON-RPC error / network) → `setStatus("Start Zotero (with
       Better BibTeX) to cite from your library.", true)`; leave the fixture items intact.
  - Extend `onTabShown()` (or `bind()`): call `CiteZotero.ping()` once; if `false`, show the
    same "Start Zotero (with Better BibTeX)…" hint (non-fatal, `warn`) while keeping the
    fixture loaded so the tab still functions. If `true`, do not disturb the existing preview.
    Guard for `typeof CiteZotero === "undefined"` so the pane still works if the script is
    absent.
  - **Do not** change the Insert/Insert-bibliography paths — merged Zotero items flow through
    the existing `selectedIds()` → `CiteEngine.build({items:cache.items})` → SP-1 insertion
    unchanged (bidi + Arabic-punctuation treatment already applied there).
- **Edit `src/taskpane/taskpane.html`:**
  - Add the button to the citation actions section (near `#cite-insert`):
    `<button id="cite-add-zotero" type="button" class="button--secondary">Add from Zotero</button>`.
    Place it so the picking action reads before Insert (e.g. a small actions row above or within
    the existing `actions` block); match existing button classes/markup.
  - Add `"./cite-zotero.js"` to the `srcs` array **before** `"./cite-pane.js"` (so
    `window.CiteZotero` exists when the pane binds).
  - Bump `window.ASHAAR_ASSET_VERSION` to `"20260717-cite-zotero"` (installed users must not
    run stale JS after this pane-code change).
  - Optionally update the tab's intro `<p>` to mention "or add from your Zotero library" —
    keep it short and honest.

**Tests:** No new node test file — this task is DOM/Office.js wiring (the pure logic is covered
by Tasks 1–2). Verification is the **manual checklist** below plus a browser smoke check that
the tab still loads with zero console errors and the fixture path is unbroken (the
implementer runs the existing `npm test` to confirm no regression, and does a Playwright/browser
load if available).

**Acceptance:** With the dev server running, the Cite tab shows an "Add from Zotero" button;
`npm test` still passes (no regressions); the fixture-only flow is unchanged when Zotero is
absent; the tab loads with no console errors.

---

## Manual checklist (needs Word + the running Zotero+BBT — user-run, not a code gate)

1. `npm start`; open the Cite tab. With Zotero running, click **Add from Zotero** → Zotero's
   native CAYW picker pops → pick a real item (e.g. `YaumulMabasUyun`) → it appears in the item
   list (checked) and in the live preview.
2. Insert as footnote → the citation inserts into Word with correct formatting; an Arabic-title
   item gets the bidi + Arabic-comma treatment (SP-1 path).
3. Insert bibliography → tagged content control inserts correctly.
4. Close Zotero → reopen the tab → status shows "Start Zotero (with Better BibTeX) to cite from
   your library"; the fixture flow still works end-to-end.
5. No mixed-content / CORS errors in the WebView console at any point.

## Success criteria (from the design spec)

Clicking "Add from Zotero" with Zotero+BBT running pops the native picker; selected items render
in the pane preview and insert into Word as correctly-formatted footnote + bibliography (Arabic
items get the bidi/Arabic-comma treatment) — with no mixed-content/CORS error and without
enabling Zotero's built-in local API. With Zotero closed, the tab shows the fallback and still
works off the fixture. `npm test` passes including the new tests.
