# Unified Settings Panel Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** One Settings panel replaces all scattered justification/layout controls; blocks follow profiles with Word-style local-delta overlays (tag payload v3), resolved through one pure resolver.

**Architecture:** Pure logic lands first (resolver in `profiles.js`, tag v3 in `word-html.js`, panel-state in a new `settings-panel.js`) with Node tests; then the HTML restructure; then the DOM glue in `taskpane.js` (reflection → resolver → panel; Apply routed by target). Old documents migrate at read time — v2 tags parse into v3 shape, nothing rewrites until the next Apply.

**Tech Stack:** Vanilla ES5, UMD modules, Office.js v1, Node `assert` tests (no framework), no build step.

**Spec:** `docs/superpowers/specs/2026-07-12-unified-settings-design.md` — read it before starting any task.

## Global Constraints

- ES5 only (`var`, `function`), UMD module pattern matching existing files. No transpilation, no new dependencies.
- Never edit `src/vendor/` — it is synced from the submodule.
- Tests are plain Node scripts run via the `npm test` chain in `package.json`; new test files must be appended to that chain.
- All tag payloads live behind `AshaarWord.parseContentControlTag` / setters — no ad-hoc JSON parsing of tags anywhere else.
- Canonical settings keys (used by resolver, panel, tags — exact spelling): `justifyMode`, `fillMode`, `strength`, `gap`, `widthMode`, `widthPct`, `misraWidthPt`, `layoutMode`, `colWidthMode`, `capEm`, `fontCorrections`, `debugColors`.
- Office.js constraints that already bit this project (see memory/CLAUDE.md): insertOoxml-replace on CC ranges throws — insert-after-then-delete; table resize must rebuild OOXML, never `columns.setWidth`.

---

### Task 1: Canonical settings + resolver in `profiles.js`

**Files:**
- Modify: `src/taskpane/profiles.js`
- Create: `tests/profiles-resolve.test.js`
- Modify: `package.json` (append test to chain)

**Interfaces:**
- Consumes: existing `defaultProfile`, `normalizeProfile`, `normalizeStrength`, `normalizeFillMode`.
- Produces (all exported from `AshaarProfiles`):
  - `defaultSettings()` → full canonical values object (the "defaults" layer)
  - `settingsFromProfile(profile)` → partial canonical object (only keys a profile owns)
  - `profileFromSettings(name, values)` → profile-schema object (inverse mapping, used by Save-as and Restore)
  - `resolveSettings({payload, profileStore, scope})` → `{values, source, profileName, profileMissing, usedCache}` where `scope` is `{level: "poem"|"bandh"|"cell"|"gap", key?: string}` and `source[k]` ∈ `"default"|"profile"|"local"|"bandh"|"cell"`.

- [ ] **Step 1: Write the failing test**

Create `tests/profiles-resolve.test.js`:

```js
const assert = require("assert");
const {
  defaultProfile,
  defaultSettings,
  settingsFromProfile,
  profileFromSettings,
  resolveSettings,
} = require("../src/taskpane/profiles");

// ── defaultSettings ──────────────────────────────────────────────────────────
{
  const d = defaultSettings();
  assert.equal(d.justifyMode, "kashida");
  assert.equal(d.fillMode, "natural-fit");
  assert.equal(d.strength, 6);
  assert.equal(d.gap, 4);
  assert.equal(d.widthMode, "auto-fit");
  assert.equal(d.widthPct, 50);
  assert.equal(d.misraWidthPt, null);
  assert.equal(d.layoutMode, "balanced");
  assert.equal(d.colWidthMode, "optimized");
  assert.equal(d.capEm, 0.28);
  assert.deepEqual(d.fontCorrections, {});
  assert.deepEqual(d.debugColors, { tatweel: "", space: "" });
  // Fresh object every call — no shared mutable state.
  assert.notStrictEqual(defaultSettings(), d);
}

// ── settingsFromProfile: only keys the profile owns, canonical spelling ──────
{
  const prof = defaultProfile("Karbala");
  prof.justify.strength = 9;
  prof.justify.widthPt = 320;
  prof.gap = 8;
  prof.width.mode = "fixed";
  prof.width.pct = 80;
  const s = settingsFromProfile(prof);
  assert.equal(s.strength, 9);
  assert.equal(s.misraWidthPt, 320);
  assert.equal(s.gap, 8);
  assert.equal(s.widthMode, "fixed");
  assert.equal(s.widthPct, 80);
  assert.equal(s.justifyMode, "kashida");
  assert.equal(s.fillMode, "natural-fit");
  // Profiles do NOT own layoutMode / colWidthMode / capEm.
  assert.equal("layoutMode" in s, false);
  assert.equal("colWidthMode" in s, false);
  assert.equal("capEm" in s, false);
}

// ── profileFromSettings: inverse mapping round-trips ─────────────────────────
{
  const values = defaultSettings();
  values.strength = 9;
  values.gap = 8;
  values.widthMode = "fixed";
  values.widthPct = 80;
  values.misraWidthPt = 320;
  const p = profileFromSettings("Karbala", values);
  assert.equal(p.name, "Karbala");
  assert.equal(p.justify.strength, 9);
  assert.equal(p.justify.widthPt, 320);
  assert.equal(p.gap, 8);
  assert.equal(p.width.mode, "fixed");
  assert.equal(p.width.pct, 80);
  const back = settingsFromProfile(p);
  assert.equal(back.strength, 9);
  assert.equal(back.gap, 8);
}

// ── resolveSettings: layering defaults → profile → local ────────────────────
{
  const store = { Karbala: profileFromSettings("Karbala", Object.assign(defaultSettings(), { strength: 9, widthPct: 80 })) };
  const payload = { profile: "Karbala", local: { gap: 8 } };
  const r = resolveSettings({ payload, profileStore: store, scope: { level: "poem" } });
  assert.equal(r.values.gap, 8, "local wins");
  assert.equal(r.source.gap, "local");
  assert.equal(r.values.strength, 9, "profile shows through");
  assert.equal(r.source.strength, "profile");
  assert.equal(r.values.justifyMode, "kashida", "default shows through");
  assert.equal(r.source.justifyMode, "default");
  assert.equal(r.profileName, "Karbala");
  assert.equal(r.profileMissing, false);
  assert.equal(r.usedCache, false);
}

// ── No payload at all (plain selection): pure defaults ───────────────────────
{
  const r = resolveSettings({ payload: null, profileStore: {}, scope: { level: "poem" } });
  assert.equal(r.values.gap, 4);
  assert.equal(r.source.gap, "default");
  assert.equal(r.profileName, "");
}

// ── profileCache fallback when store lacks the profile ───────────────────────
{
  const payload = {
    profile: "Karbala",
    local: {},
    profileCache: { strength: 9, gap: 8 },
  };
  const r = resolveSettings({ payload, profileStore: {}, scope: { level: "poem" } });
  assert.equal(r.profileMissing, true);
  assert.equal(r.usedCache, true);
  assert.equal(r.values.strength, 9, "cache supplies profile layer");
  assert.equal(r.source.strength, "profile");
  // Live store wins over cache when present.
  const store = { Karbala: profileFromSettings("Karbala", Object.assign(defaultSettings(), { strength: 3 })) };
  const r2 = resolveSettings({ payload, profileStore: store, scope: { level: "poem" } });
  assert.equal(r2.values.strength, 3, "store beats cache");
  assert.equal(r2.usedCache, false);
}

// ── Bandh scope: payload.widthPt overlays misraWidthPt ───────────────────────
{
  const payload = { profile: "", local: { misraWidthPt: 300 }, widthPt: 350 };
  const poem = resolveSettings({ payload, profileStore: {}, scope: { level: "poem" } });
  assert.equal(poem.values.misraWidthPt, 300, "poem scope ignores bandh widthPt");
  const bandh = resolveSettings({ payload, profileStore: {}, scope: { level: "bandh" } });
  assert.equal(bandh.values.misraWidthPt, 350);
  assert.equal(bandh.source.misraWidthPt, "bandh");
}

// ── Cell scope: overrides win over everything ────────────────────────────────
{
  const payload = {
    profile: "",
    local: { strength: 4 },
    widthPt: 350,
    overrides: { "A2:3": { strength: 9, widthPt: 400, capEm: 0.5 } },
  };
  const r = resolveSettings({ payload, profileStore: {}, scope: { level: "cell", key: "A2:3" } });
  assert.equal(r.values.strength, 9);
  assert.equal(r.source.strength, "cell");
  assert.equal(r.values.misraWidthPt, 400);
  assert.equal(r.source.misraWidthPt, "cell");
  assert.equal(r.values.capEm, 0.5);
  // Another cell key inherits the bandh/local layers instead.
  const other = resolveSettings({ payload, profileStore: {}, scope: { level: "cell", key: "A2:4" } });
  assert.equal(other.values.strength, 4);
  assert.equal(other.source.strength, "local");
  assert.equal(other.values.misraWidthPt, 350);
  assert.equal(other.source.misraWidthPt, "bandh");
}

// ── Deleting a local delta falls back to the layer below ─────────────────────
{
  const store = { K: profileFromSettings("K", Object.assign(defaultSettings(), { gap: 6 })) };
  const withLocal = resolveSettings({ payload: { profile: "K", local: { gap: 8 } }, profileStore: store, scope: { level: "poem" } });
  assert.equal(withLocal.values.gap, 8);
  const cleared = resolveSettings({ payload: { profile: "K", local: {} }, profileStore: store, scope: { level: "poem" } });
  assert.equal(cleared.values.gap, 6, "clear falls back to profile, not default");
  assert.equal(cleared.source.gap, "profile");
}

// ── defaultProfile no longer carries a font field ────────────────────────────
{
  const p = defaultProfile("X");
  assert.equal("font" in p, false, "font is not a profile setting");
}

console.log("profiles-resolve tests passed");
```

- [ ] **Step 2: Run test to verify it fails**

Run: `node tests/profiles-resolve.test.js`
Expected: FAIL — `defaultSettings is not a function` (TypeError on destructure/call).

- [ ] **Step 3: Implement in `profiles.js`**

In `src/taskpane/profiles.js`:

(a) Delete the line `font: "",` from `defaultProfile` (line 34) and remove `"font"` from the `keys` array in `mergeProfile` (line 49 becomes `var keys = ["name", "gap", "misraSymbol", "symbolColor"];`).

(b) Insert before the `return {` export block:

```js
  // ── Canonical settings (unified panel) ────────────────────────────────────
  // One flat shape shared by the resolver, the panel, and tag `local` maps.
  // Layering: defaults → profile → block local → bandh widthPt → cell override.

  function defaultSettings() {
    return {
      justifyMode: "kashida",     // "kashida" | "css" | "spacing" | "none"
      fillMode: "natural-fit",    // "natural-fit" | "cell-fit" | "adaptive"
      strength: 6,                // 1..10
      gap: 4,                     // middle-gap grid columns, 0..20
      widthMode: "auto-fit",      // "auto-fit" | "fixed"
      widthPct: 50,               // 25..100 (only meaningful when fixed)
      misraWidthPt: null,         // explicit fill target; null = computed
      layoutMode: "balanced",     // "balanced"|"equal"|"compact"|"stacked"|"auto"
      colWidthMode: "optimized",  // "optimized" | "fixed" (column-width strategy)
      capEm: 0.28,                // residual spacing cap (cell scope)
      fontCorrections: {},
      debugColors: { tatweel: "", space: "" },
    };
  }

  // Keys a profile owns. layoutMode/colWidthMode/capEm are block/cell-level
  // preferences with no profile layer.
  function settingsFromProfile(profile) {
    var p = normalizeProfile(profile || {});
    var out = {
      justifyMode: p.justify.mode,
      fillMode: normalizeFillMode(p.justify.fillMode),
      strength: normalizeStrength(p.justify.strength),
      gap: Number(p.gap),
      widthMode: p.width.mode === "fixed" ? "fixed" : "auto-fit",
      widthPct: Number(p.width.pct),
    };
    if (p.justify.widthPt != null) out.misraWidthPt = p.justify.widthPt;
    if (p.fontCorrections && Object.keys(p.fontCorrections).length) out.fontCorrections = p.fontCorrections;
    if (p.debugColors && (p.debugColors.tatweel || p.debugColors.space)) out.debugColors = p.debugColors;
    return out;
  }

  // Inverse of settingsFromProfile: canonical values → profile-schema object.
  function profileFromSettings(name, values) {
    var v = values || {};
    var p = defaultProfile(name);
    if (v.justifyMode != null) p.justify.mode = v.justifyMode;
    if (v.fillMode != null) p.justify.fillMode = normalizeFillMode(v.fillMode);
    if (v.strength != null) p.justify.strength = normalizeStrength(v.strength);
    if (v.misraWidthPt !== undefined) p.justify.widthPt = v.misraWidthPt;
    if (v.gap != null) p.gap = Number(v.gap);
    if (v.widthMode != null) p.width.mode = v.widthMode === "fixed" ? "fixed" : "auto-fit";
    if (v.widthPct != null) p.width.pct = Number(v.widthPct);
    if (v.fontCorrections) p.fontCorrections = v.fontCorrections;
    if (v.debugColors) p.debugColors = v.debugColors;
    return p;
  }

  // Resolve the effective settings for a target.
  //   payload:      parsed v3 tag payload (or null for a plain selection)
  //   profileStore: { name: profile } (localStorage contents)
  //   scope:        { level: "poem"|"bandh"|"cell"|"gap", key?: "A2:3" }
  // Returns { values, source, profileName, profileMissing, usedCache }.
  function resolveSettings(args) {
    args = args || {};
    var payload = args.payload || null;
    var store = args.profileStore || {};
    var scope = args.scope || { level: "poem" };
    var values = defaultSettings();
    var source = {};
    Object.keys(values).forEach(function (k) { source[k] = "default"; });

    var profileName = (payload && typeof payload.profile === "string") ? payload.profile : "";
    var prof = profileName ? store[profileName] : null;
    var profileMissing = !!(profileName && !prof);
    var usedCache = false;
    var layer = null;
    if (prof) layer = settingsFromProfile(prof);
    else if (profileMissing && payload && isObj(payload.profileCache)) { layer = payload.profileCache; usedCache = true; }
    if (layer) {
      Object.keys(layer).forEach(function (k) {
        if (k in values) { values[k] = layer[k]; source[k] = "profile"; }
      });
    }

    var local = (payload && isObj(payload.local)) ? payload.local : {};
    Object.keys(local).forEach(function (k) {
      if (k in values) { values[k] = local[k]; source[k] = "local"; }
    });

    if (scope.level === "bandh" || scope.level === "cell") {
      if (payload && typeof payload.widthPt === "number" && payload.widthPt > 0) {
        values.misraWidthPt = payload.widthPt; source.misraWidthPt = "bandh";
      }
    }
    if (scope.level === "cell" && scope.key && payload && isObj(payload.overrides) && isObj(payload.overrides[scope.key])) {
      var ov = payload.overrides[scope.key];
      if (ov.strength != null) { values.strength = ov.strength; source.strength = "cell"; }
      if (ov.widthPt != null) { values.misraWidthPt = ov.widthPt; source.misraWidthPt = "cell"; }
      if (ov.capEm != null) { values.capEm = ov.capEm; source.capEm = "cell"; }
    }

    return { values: values, source: source, profileName: profileName, profileMissing: profileMissing, usedCache: usedCache };
  }
```

(c) Add to the export object:

```js
    defaultSettings: defaultSettings,
    settingsFromProfile: settingsFromProfile,
    profileFromSettings: profileFromSettings,
    resolveSettings: resolveSettings,
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `node tests/profiles-resolve.test.js`
Expected: `profiles-resolve tests passed`

Run: `node tests/profiles.test.js`
Expected: passes (nothing in it asserts a `font` field — verified). If it fails on `mergeProfile` key count, re-check step 3(a).

Also run: `grep -n "profile\.font" src/taskpane/taskpane.js`
Expected: one hit at line ~797 (`var fallbackName = profile.font || "Times New Roman";`). Change that line to:

```js
    var fallbackName = "Times New Roman";
```

(The representative-font capture from document cells supersedes it; the profile no longer stores a font.)

- [ ] **Step 5: Add the test to the npm chain and commit**

In `package.json`, append to the `test` script: `&& node tests/profiles-resolve.test.js`

```bash
npm test   # full chain must pass
git add src/taskpane/profiles.js tests/profiles-resolve.test.js package.json src/taskpane/taskpane.js
git commit -m "feat(profiles): canonical settings shape + layered resolver"
```

---

### Task 2: Tag payload v3 — write, parse, migrate, setters

**Files:**
- Modify: `src/taskpane/word-html.js` (functions `contentControlTag` ~line 892, `parseContentControlTag` ~line 921, export block ~line 1810)
- Modify: `tests/word-html.test.js` (append)

**Interfaces:**
- Consumes: Task 1's canonical settings keys (tag `local` maps hold canonical keys only).
- Produces (exported from `AshaarWord`):
  - `contentControlTag(text, opts, cellPatterns)` — now writes v3: `{k, v:3, profile, local, profileCache, misraPattern, misraCount, sourceHash, cells?}`. `opts.profile` (string), `opts.local` (object), `opts.profileCache` (object|null) are taken verbatim; `opts.misraPattern`/`opts.misraCount` as before.
  - `parseContentControlTag(tag)` — returns v3-shaped payload for BOTH v2 and v3 tags (read-time migration). Guarantees: `profile` string, `local` object, `profileCache` object|null, plus existing guarantees (`qaseeda` kept as deprecated alias of `profile` so untouched call sites keep working until Task 10 removes them; `cells`, `overrides`, `slotDecor`, `runFonts`, `widthPt`).
  - `setTagProfile(tag, name)` — sets `profile` (replaces `setTagQaseeda`, which becomes an alias calling it).
  - `setTagLocal(tag, local)` — REPLACES the whole `local` map (full-replace is idempotent; callers compute the new map).
  - `setTagProfileCache(tag, cache)` — sets/clears `profileCache`.

- [ ] **Step 1: Write the failing tests**

Append to `tests/word-html.test.js`:

```js
// ── Tag payload v3 ───────────────────────────────────────────────────────────

// Writer emits v3 with profile/local/profileCache.
{
  const tag = AshaarWord.contentControlTag("متن", {
    profile: "Karbala",
    local: { gap: 8, strength: 9 },
    profileCache: { gap: 6 },
    misraPattern: "paired",
    misraCount: 4,
  });
  const p = AshaarWord.parseContentControlTag(tag);
  assert.equal(p.v, 3);
  assert.equal(p.profile, "Karbala");
  assert.deepEqual(p.local, { gap: 8, strength: 9 });
  assert.deepEqual(p.profileCache, { gap: 6 });
  assert.equal(p.misraPattern, "paired");
  assert.equal(p.misraCount, 4);
  assert.equal(p.qaseeda, "Karbala", "deprecated alias mirrors profile");
}

// Writer defaults: no profile/local → empty string / empty object.
{
  const p = AshaarWord.parseContentControlTag(AshaarWord.contentControlTag("متن", {}));
  assert.equal(p.profile, "");
  assert.deepEqual(p.local, {});
  assert.equal(p.profileCache, null);
}

// v2 read-time migration: stored preferences become local deltas (canonical
// keys), qaseeda becomes profile, fontMode is dropped, render facts survive.
{
  const v2payload = {
    k: "ashaar-poem", v: 2,
    layoutMode: "equal", widthMode: "fixed", justifyMode: "spacing",
    tatweelCount: 9, gapWidth: 8, misraPattern: "paired", misraCount: 4,
    fontMode: "jameel", tableWidthPct: 80, qaseeda: "Karbala",
    sourceHash: "abc123",
    overrides: { "A2:3": { strength: 5 } },
    widthPt: 350,
    slotDecor: { "A#1": { symbol: "؎" } },
  };
  const v2tag = "ashaar:" + encodeURIComponent(JSON.stringify(v2payload));
  const p = AshaarWord.parseContentControlTag(v2tag);
  assert.equal(p.v, 3, "migrated shape");
  assert.equal(p.profile, "Karbala");
  assert.equal(p.local.layoutMode, "equal");
  assert.equal(p.local.colWidthMode, "fixed", "v2 widthMode → colWidthMode");
  assert.equal(p.local.justifyMode, "spacing");
  assert.equal(p.local.strength, 9, "v2 tatweelCount → strength");
  assert.equal(p.local.gap, 8, "v2 gapWidth → gap");
  assert.equal(p.local.widthPct, 80, "v2 tableWidthPct → widthPct");
  assert.equal("fontMode" in p.local, false, "fontMode dropped");
  assert.equal(p.misraPattern, "paired");
  assert.equal(p.misraCount, 4);
  assert.equal(p.sourceHash, "abc123");
  assert.deepEqual(p.overrides, { "A2:3": { strength: 5 } });
  assert.equal(p.widthPt, 350);
  assert.deepEqual(p.slotDecor, { "A#1": { symbol: "؎" } });
}

// Setters touch only their own key; unknown fields round-trip untouched.
{
  const base = AshaarWord.contentControlTag("متن", { profile: "K", local: { gap: 8 } });
  // Simulate a future-version field.
  const withFuture = (() => {
    const raw = JSON.parse(decodeURIComponent(base.slice("ashaar:".length)));
    raw.futureField = { keep: true };
    return "ashaar:" + encodeURIComponent(JSON.stringify(raw));
  })();

  const t1 = AshaarWord.setTagProfile(withFuture, "Najaf");
  const p1 = AshaarWord.parseContentControlTag(t1);
  assert.equal(p1.profile, "Najaf");
  assert.deepEqual(p1.local, { gap: 8 }, "local untouched");
  assert.deepEqual(JSON.parse(decodeURIComponent(t1.slice("ashaar:".length))).futureField, { keep: true });

  const t2 = AshaarWord.setTagLocal(t1, { strength: 9 });
  const p2 = AshaarWord.parseContentControlTag(t2);
  assert.deepEqual(p2.local, { strength: 9 }, "full replace");
  assert.equal(p2.profile, "Najaf", "profile untouched");

  const t3 = AshaarWord.setTagProfileCache(t2, { gap: 6, strength: 3 });
  assert.deepEqual(AshaarWord.parseContentControlTag(t3).profileCache, { gap: 6, strength: 3 });
  const t4 = AshaarWord.setTagProfileCache(t3, null);
  assert.equal(AshaarWord.parseContentControlTag(t4).profileCache, null);

  // setTagQaseeda alias still works and writes `profile`.
  const t5 = AshaarWord.setTagQaseeda(t4, "Alias");
  assert.equal(AshaarWord.parseContentControlTag(t5).profile, "Alias");
}

// Existing setters (override/slot-decor/bandh-width/run-fonts) still work on
// migrated v2 tags — they parse → mutate → re-encode, so the write is v3.
{
  const v2tag = "ashaar:" + encodeURIComponent(JSON.stringify({ k: "ashaar-poem", v: 2, gapWidth: 8, qaseeda: "K" }));
  const out = AshaarWord.setTagBandhWidth(v2tag, 300);
  const p = AshaarWord.parseContentControlTag(out);
  assert.equal(p.widthPt, 300);
  assert.equal(p.v, 3, "any setter write upgrades to v3");
  assert.equal(p.local.gap, 8, "migrated local survives the setter");
}
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `node tests/word-html.test.js`
Expected: FAIL at the first new block (`p.v` is 2, `p.profile` undefined).

- [ ] **Step 3: Implement in `word-html.js`**

(a) Replace the body of `contentControlTag` (keep the hash loop):

```js
  function contentControlTag(text, opts, cellPatterns) {
    opts = opts || {};
    var source = String(text || "");
    var hash = 0;
    for (var i = 0; i < source.length; i++) {
      hash = ((hash << 5) - hash + source.charCodeAt(i)) | 0;
    }
    var payload = {
      k: "ashaar-poem",
      v: 3,
      profile: typeof opts.profile === "string" ? opts.profile : "",
      local: (opts.local && typeof opts.local === "object") ? opts.local : {},
      profileCache: (opts.profileCache && typeof opts.profileCache === "object") ? opts.profileCache : null,
      misraPattern: opts.misraPattern || "paired",
      misraCount: Number(opts.misraCount || 4),
      sourceHash: (hash >>> 0).toString(16)
    };
    if (cellPatterns && cellPatterns.length) payload.cells = cellPatterns;
    return "ashaar:" + encodeURIComponent(JSON.stringify(payload));
  }
```

(b) In `parseContentControlTag`, after the existing guarantees and before `return payload;`, add the migration + v3 guarantees:

```js
      // v2 → v3 read-time migration: stored preferences become local deltas
      // (canonical keys), qaseeda becomes profile, fontMode is dropped. The
      // document is not touched until the next write re-encodes as v3.
      if (payload.v !== 3) {
        var local = {};
        if (payload.justifyMode != null) local.justifyMode = payload.justifyMode;
        if (Number(payload.tatweelCount) > 0) local.strength = Number(payload.tatweelCount);
        if (payload.gapWidth != null) local.gap = Number(payload.gapWidth);
        if (payload.tableWidthPct != null) local.widthPct = Number(payload.tableWidthPct);
        if (payload.layoutMode != null) local.layoutMode = payload.layoutMode;
        if (payload.widthMode != null) local.colWidthMode = payload.widthMode;
        if (payload.fillMode != null) local.fillMode = payload.fillMode;
        payload.local = local;
        payload.profile = payload.qaseeda || "";
        payload.v = 3;
        delete payload.justifyMode; delete payload.tatweelCount; delete payload.gapWidth;
        delete payload.tableWidthPct; delete payload.layoutMode; delete payload.widthMode;
        delete payload.fillMode; delete payload.fontMode;
      }
      if (typeof payload.profile !== "string") payload.profile = "";
      payload.local = (payload.local && typeof payload.local === "object") ? payload.local : {};
      payload.profileCache = (payload.profileCache && typeof payload.profileCache === "object") ? payload.profileCache : null;
      payload.qaseeda = payload.profile; // deprecated alias, removed in cleanup task
```

Note: the existing line `if (typeof payload.qaseeda !== "string") payload.qaseeda = "";` becomes redundant — delete it.

(c) Add the new setters next to `setTagQaseeda` and re-point the alias:

```js
  // Return a copy of an "ashaar:" tag with only its profile name replaced.
  function setTagProfile(tag, name) {
    var payload = parseContentControlTag(tag);
    if (!payload) return tag;
    payload.profile = name || "";
    payload.qaseeda = payload.profile;
    return "ashaar:" + encodeURIComponent(JSON.stringify(payload));
  }

  // Deprecated alias (v2 name) — remove with the cleanup task.
  function setTagQaseeda(tag, name) { return setTagProfile(tag, name); }

  // Return a copy of an "ashaar:" tag with the local delta map REPLACED.
  // Callers compute the full new map (delete-by-omission).
  function setTagLocal(tag, local) {
    var payload = parseContentControlTag(tag);
    if (!payload) return tag;
    payload.local = (local && typeof local === "object") ? local : {};
    return "ashaar:" + encodeURIComponent(JSON.stringify(payload));
  }

  // Return a copy of an "ashaar:" tag with the profile snapshot set/cleared.
  function setTagProfileCache(tag, cache) {
    var payload = parseContentControlTag(tag);
    if (!payload) return tag;
    payload.profileCache = (cache && typeof cache === "object") ? cache : null;
    return "ashaar:" + encodeURIComponent(JSON.stringify(payload));
  }
```

Delete the old `setTagQaseeda` body (lines ~940-945).

(d) Export: add `setTagProfile`, `setTagLocal`, `setTagProfileCache` to the return block (keep `setTagQaseeda`).

- [ ] **Step 4: Run tests to verify they pass**

Run: `node tests/word-html.test.js`
Expected: all pass, including pre-existing assertions (renderForWord etc. don't read the tag).

Run: `npm test`
Expected: full chain passes. If `bandh-cell-map.test.js` or `cell-overrides.test.js` construct v2 tags and assert on parse results, update those assertions to the v3 shape (`payload.local.gap` instead of `payload.gapWidth`, `payload.profile` instead of `payload.qaseeda` — the `qaseeda` alias keeps most passing).

- [ ] **Step 5: Commit**

```bash
git add src/taskpane/word-html.js tests/word-html.test.js
git commit -m "feat(word-html): tag payload v3 — profile + local deltas + profileCache, v2 read-time migration"
```

---

### Task 3: Extract and fix the rebuild-skip signature

**Files:**
- Modify: `src/taskpane/word-html.js` (new pure function + export)
- Modify: `src/taskpane/taskpane.js:1219-1224` (use it)
- Modify: `tests/word-html.test.js` (append)

**Interfaces:**
- Produces: `AshaarWord.applySizeSignature({targetTwips, sources, gap, misraPattern})` → string. Same inputs → same string; any input change → different string.
- Consumes: called from `applyProfileToQaseeda` pass 1 (taskpane.js) with the block's resolved gap and pattern.

- [ ] **Step 1: Write the failing test**

Append to `tests/word-html.test.js`:

```js
// ── applySizeSignature ───────────────────────────────────────────────────────
{
  const base = { targetTwips: 9360, sources: ["متن الف", "متن ب"], gap: 4, misraPattern: "paired" };
  const sig = AshaarWord.applySizeSignature(base);
  assert.equal(typeof sig, "string");
  assert.equal(AshaarWord.applySizeSignature(base), sig, "deterministic");
  assert.notEqual(AshaarWord.applySizeSignature(Object.assign({}, base, { gap: 8 })), sig, "gap changes sig");
  assert.notEqual(AshaarWord.applySizeSignature(Object.assign({}, base, { misraPattern: "single" })), sig, "pattern changes sig");
  assert.notEqual(AshaarWord.applySizeSignature(Object.assign({}, base, { targetTwips: 9000 })), sig, "width changes sig");
  assert.notEqual(AshaarWord.applySizeSignature(Object.assign({}, base, { sources: ["متن الف"] })), sig, "source changes sig");
}
```

- [ ] **Step 2: Run test to verify it fails**

Run: `node tests/word-html.test.js` — Expected: FAIL, `applySizeSignature is not a function`.

- [ ] **Step 3: Implement**

In `word-html.js` (near `contentControlTag`), add + export:

```js
  // Rebuild-skip signature for the apply pipeline: identical signature ⇒ the
  // tables are already sized/shaped correctly and the destructive rebuild can
  // be skipped. MUST include every structural input (a gap-only change used to
  // slip through when the sig was width+source only).
  function applySizeSignature(args) {
    args = args || {};
    var parts = [String(args.targetTwips || 0), String(args.gap != null ? args.gap : ""), String(args.misraPattern || "")];
    (args.sources || []).forEach(function (s) {
      var h = 0; s = String(s || "");
      for (var i = 0; i < s.length; i++) h = ((h << 5) - h + s.charCodeAt(i)) | 0;
      parts.push((h >>> 0).toString(16));
    });
    return parts.join("|");
  }
```

In `taskpane.js`, replace lines 1219-1223 (the `srcSig` computation and `sizeSig = targetTwips + "|" + srcSig;`) with:

```js
        sizeSig = AshaarWord.applySizeSignature({
          targetTwips: targetTwips,
          sources: cap.blockInfos.map(function (b) { return b.source; }),
          // Structural inputs: any block's effective gap/pattern participates.
          gap: cap.blockInfos.map(function (b) { return Number((b.payload.local || {}).gap != null ? b.payload.local.gap : ""); }).join(","),
          misraPattern: cap.blockInfos.map(function (b) { return b.payload.misraPattern || ""; }).join(","),
        });
```

(After Task 7 the gap component switches to the resolver's effective value; this interim form already fixes the gap-only-change skip.)

- [ ] **Step 4: Run tests**

Run: `node tests/word-html.test.js && node tests/profiles.test.js` — Expected: pass.
Run: `npm test` — Expected: pass.

- [ ] **Step 5: Commit**

```bash
git add src/taskpane/word-html.js src/taskpane/taskpane.js tests/word-html.test.js
git commit -m "feat(apply): extracted rebuild-skip signature; gap/pattern now trigger rebuilds"
```

---

### Task 4: Panel-state module (`settings-panel.js`, pure)

**Files:**
- Create: `src/taskpane/settings-panel.js`
- Create: `tests/settings-panel.test.js`
- Modify: `package.json` (append test), `src/taskpane/taskpane.html` (script include — done properly in Task 5; here only the module + tests)

**Interfaces:**
- Consumes: `resolveSettings` result shape from Task 1.
- Produces (UMD global `AshaarPanel`):
  - `SCOPE_FIELDS` — `{poem: [...], bandh: [...], cell: [...], gap: []}` canonical keys shown per scope (gap scope's decoration editor is not settings-keyed; its fields ride separately).
  - `panelStateFor({resolved, pending, target})` → `{header, chips, controls, profileRow, footer}` (full shape in the test below).
  - `mergePending(pending, key, value)` → new pending object (`{set:{}, clear:[]}` shape).
  - `pendingToLocal(local, pending, scopeKeys)` → the NEW full local map to write (applies sets, removes clears; only canonical keys).

- [ ] **Step 1: Write the failing test**

Create `tests/settings-panel.test.js`:

```js
const assert = require("assert");
const AshaarPanel = require("../src/taskpane/settings-panel");
const { resolveSettings, profileFromSettings, defaultSettings } = require("../src/taskpane/profiles");

// ── panelStateFor: header + chips + controls with provenance ────────────────
{
  const store = { Karbala: profileFromSettings("Karbala", Object.assign(defaultSettings(), { strength: 9 })) };
  const payload = { profile: "Karbala", local: { gap: 8 } };
  const resolved = resolveSettings({ payload, profileStore: store, scope: { level: "poem" } });
  const st = AshaarPanel.panelStateFor({
    resolved,
    pending: { set: {}, clear: [] },
    target: { kind: "block", scope: { level: "poem" }, cellEnabled: false, gapEnabled: false },
  });
  assert.equal(st.header.title, "Poem — Karbala");
  assert.deepEqual(st.chips, [
    { level: "poem", enabled: true, active: true },
    { level: "bandh", enabled: true, active: false },
    { level: "cell", enabled: false, active: false },
    { level: "gap", enabled: false, active: false },
  ]);
  const gap = st.controls.find((c) => c.key === "gap");
  assert.equal(gap.value, 8);
  assert.equal(gap.source, "local");
  assert.equal(gap.dirty, false);
  const strength = st.controls.find((c) => c.key === "strength");
  assert.equal(strength.value, 9);
  assert.equal(strength.source, "profile");
  assert.equal(st.profileRow.name, "Karbala");
  assert.equal(st.profileRow.assignEnabled, true);
  assert.equal(st.profileRow.updateVisible, false, "no pending edits → no drift");
  assert.equal(st.footer.revertLabel, "Revert to profile");
}

// ── Pending edits show as dirty and surface the Update action ────────────────
{
  const resolved = resolveSettings({ payload: { profile: "Karbala", local: {} }, profileStore: { Karbala: profileFromSettings("Karbala", defaultSettings()) }, scope: { level: "poem" } });
  const st = AshaarPanel.panelStateFor({
    resolved,
    pending: { set: { strength: 3 }, clear: [] },
    target: { kind: "block", scope: { level: "poem" }, cellEnabled: false, gapEnabled: false },
  });
  const strength = st.controls.find((c) => c.key === "strength");
  assert.equal(strength.value, 3, "pending value shown");
  assert.equal(strength.dirty, true);
  assert.equal(st.profileRow.updateVisible, true, "drift → Update visible");
}

// ── Plain selection target ───────────────────────────────────────────────────
{
  const resolved = resolveSettings({ payload: null, profileStore: {}, scope: { level: "poem" } });
  const st = AshaarPanel.panelStateFor({
    resolved,
    pending: { set: {}, clear: [] },
    target: { kind: "selection", scope: { level: "poem" }, cellEnabled: false, gapEnabled: false },
  });
  assert.equal(st.header.title, "Selection");
  assert.equal(st.chips.length, 0, "no chips outside a block");
  assert.equal(st.profileRow.assignEnabled, false, "nowhere to store the link");
  assert.equal(st.footer.revertLabel, "Reset to defaults");
}

// ── Missing profile surfaces the restore affordance ──────────────────────────
{
  const resolved = resolveSettings({ payload: { profile: "Karbala", local: {}, profileCache: { strength: 9 } }, profileStore: {}, scope: { level: "poem" } });
  const st = AshaarPanel.panelStateFor({
    resolved,
    pending: { set: {}, clear: [] },
    target: { kind: "block", scope: { level: "poem" }, cellEnabled: false, gapEnabled: false },
  });
  assert.equal(st.profileRow.missing, true);
  assert.equal(st.header.title, "Poem — Karbala (not on this machine)");
}

// ── Scope narrows the control set ────────────────────────────────────────────
{
  const resolved = resolveSettings({ payload: { profile: "", local: {}, overrides: {} }, profileStore: {}, scope: { level: "cell", key: "A2:3" } });
  const st = AshaarPanel.panelStateFor({
    resolved,
    pending: { set: {}, clear: [] },
    target: { kind: "block", scope: { level: "cell", key: "A2:3" }, cellEnabled: true, gapEnabled: false, cellLabel: "A2:3" },
  });
  assert.equal(st.header.title, "Poem › Cell A2:3");
  assert.deepEqual(st.controls.map((c) => c.key), ["strength", "misraWidthPt", "capEm"]);
}

// ── mergePending / pendingToLocal ────────────────────────────────────────────
{
  let p = { set: {}, clear: [] };
  p = AshaarPanel.mergePending(p, "gap", 8);
  assert.deepEqual(p, { set: { gap: 8 }, clear: [] });
  p = AshaarPanel.mergePending(p, "gap", null); // ⟲ reset marks a clear
  assert.deepEqual(p, { set: {}, clear: ["gap"] });
  p = AshaarPanel.mergePending(p, "strength", 9);

  const local = { gap: 8, widthPct: 80 };
  const next = AshaarPanel.pendingToLocal(local, p, AshaarPanel.SCOPE_FIELDS.poem);
  assert.deepEqual(next, { widthPct: 80, strength: 9 }, "clear deletes, set adds, untouched survives");
  // Inputs not mutated.
  assert.deepEqual(local, { gap: 8, widthPct: 80 });
}

console.log("settings-panel tests passed");
```

- [ ] **Step 2: Run test to verify it fails**

Run: `node tests/settings-panel.test.js` — Expected: FAIL, cannot find module.

- [ ] **Step 3: Implement `src/taskpane/settings-panel.js`**

```js
/**
 * settings-panel.js — pure panel-state logic for the unified Settings panel.
 *
 * Computes what the panel should show (values, provenance, dirty flags,
 * visible actions) from a resolveSettings() result + the pending-edits buffer
 * + the cursor target. NO DOM and NO Office.js — taskpane.js renders this.
 * See docs/superpowers/specs/2026-07-12-unified-settings-design.md.
 */
(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) module.exports = factory();
  else root.AshaarPanel = factory();
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  // Canonical keys shown per scope. Gap scope edits slotDecor (not
  // settings-keyed), so its list is empty here.
  var SCOPE_FIELDS = {
    poem: ["justifyMode", "fillMode", "strength", "gap", "widthMode", "widthPct",
           "layoutMode", "colWidthMode", "misraWidthPt", "fontCorrections", "debugColors"],
    bandh: ["misraWidthPt"],
    cell: ["strength", "misraWidthPt", "capEm"],
    gap: [],
  };

  var LEVELS = ["poem", "bandh", "cell", "gap"];

  function scopeTitle(target, resolved) {
    if (target.kind !== "block") return "Selection";
    var name = resolved.profileName
      ? resolved.profileName + (resolved.profileMissing ? " (not on this machine)" : "")
      : "(no profile)";
    var head = "Poem — " + name;
    if (!resolved.profileName) head = "Poem";
    var lvl = target.scope.level;
    if (lvl === "bandh") return head + " › Bandh" + (target.bandhLabel ? " " + target.bandhLabel : "");
    if (lvl === "cell") return head + " › Cell " + (target.cellLabel || target.scope.key || "");
    if (lvl === "gap") return head + " › Gap " + (target.gapLabel || target.scope.key || "");
    return head;
  }

  function panelStateFor(args) {
    var resolved = args.resolved;
    var pending = args.pending || { set: {}, clear: [] };
    var target = args.target || { kind: "selection", scope: { level: "poem" } };
    var level = target.scope.level || "poem";

    var chips = target.kind === "block" ? LEVELS.map(function (l) {
      return {
        level: l,
        enabled: l === "poem" || l === "bandh" ? true : (l === "cell" ? !!target.cellEnabled : !!target.gapEnabled),
        active: l === level,
      };
    }) : [];

    var controls = SCOPE_FIELDS[level].map(function (key) {
      var dirty = (key in pending.set) || pending.clear.indexOf(key) !== -1;
      var value = (key in pending.set) ? pending.set[key]
        : (pending.clear.indexOf(key) !== -1 ? inheritedValue(resolved, key) : resolved.values[key]);
      return { key: key, value: value, source: resolved.source[key], dirty: dirty };
    });

    var anyDirty = Object.keys(pending.set).length > 0 || pending.clear.length > 0;

    return {
      header: { title: scopeTitle(target, resolved) },
      chips: chips,
      controls: controls,
      profileRow: {
        name: resolved.profileName,
        missing: resolved.profileMissing,
        assignEnabled: target.kind === "block",
        updateVisible: !!resolved.profileName && !resolved.profileMissing && anyDirty && level === "poem",
        restoreVisible: resolved.profileMissing && resolved.usedCache,
      },
      footer: {
        applyEnabled: true,
        revertLabel: resolved.profileName ? "Revert to profile" : "Reset to defaults",
      },
    };
  }

  // What the value would be if the local delta for `key` were removed:
  // profile layer if it set the key, else the default. (Bandh/cell layers are
  // not consulted — clears happen at the scope that owns the key.)
  function inheritedValue(resolved, key) {
    // resolved.source tells where the CURRENT value came from; for a clear we
    // need the layer below "local". The resolver exposes this indirectly: a
    // key whose source is "local" inherits the profile value when the profile
    // owns it, else the default. panelStateFor callers pass resolvedBase for
    // exactness; this fallback recomputation is display-only (Apply re-resolves).
    if (resolved.inherited && (key in resolved.inherited)) return resolved.inherited[key];
    return resolved.values[key];
  }

  function mergePending(pending, key, value) {
    var set = {}, clear = pending.clear.slice();
    Object.keys(pending.set).forEach(function (k) { set[k] = pending.set[k]; });
    var ci = clear.indexOf(key);
    if (value === null || value === undefined) {   // ⟲ reset → clear the delta
      delete set[key];
      if (ci === -1) clear.push(key);
    } else {
      set[key] = value;
      if (ci !== -1) clear.splice(ci, 1);
    }
    return { set: set, clear: clear };
  }

  // Compute the NEW full local map: apply sets, drop clears; only keys legal
  // for the scope. Never mutates inputs.
  function pendingToLocal(local, pending, scopeKeys) {
    var out = {};
    Object.keys(local || {}).forEach(function (k) {
      if (pending.clear.indexOf(k) === -1) out[k] = local[k];
    });
    Object.keys(pending.set).forEach(function (k) {
      if (!scopeKeys || scopeKeys.indexOf(k) !== -1) out[k] = pending.set[k];
    });
    return out;
  }

  return {
    SCOPE_FIELDS: SCOPE_FIELDS,
    panelStateFor: panelStateFor,
    mergePending: mergePending,
    pendingToLocal: pendingToLocal,
  };
}));
```

Note the `resolved.inherited` hook: extend `resolveSettings` (Task 1 file) with an `inherited` map — after building the profile layer, snapshot `values` into `inherited` **before** applying `local`:

```js
    var inherited = {};
    Object.keys(values).forEach(function (k) { inherited[k] = values[k]; });
    // ... existing local/bandh/cell layering ...
    return { values: values, source: source, inherited: inherited, profileName: profileName, profileMissing: profileMissing, usedCache: usedCache };
```

Add to `tests/profiles-resolve.test.js`:

```js
// ── inherited: the layer below local, for per-setting reset display ──────────
{
  const store = { K: profileFromSettings("K", Object.assign(defaultSettings(), { gap: 6 })) };
  const r = resolveSettings({ payload: { profile: "K", local: { gap: 8 } }, profileStore: store, scope: { level: "poem" } });
  assert.equal(r.values.gap, 8);
  assert.equal(r.inherited.gap, 6, "what gap falls back to on clear");
}
```

- [ ] **Step 4: Run tests**

Run: `node tests/settings-panel.test.js && node tests/profiles-resolve.test.js`
Expected: both pass.

- [ ] **Step 5: Add to npm chain and commit**

Append `&& node tests/settings-panel.test.js` to the `test` script.

```bash
npm test
git add src/taskpane/settings-panel.js tests/settings-panel.test.js tests/profiles-resolve.test.js src/taskpane/profiles.js package.json
git commit -m "feat(panel): pure panel-state module — scope fields, provenance, pending buffer"
```

---

### Task 5: HTML restructure

**Files:**
- Modify: `src/taskpane/taskpane.html`
- Modify: `src/taskpane/taskpane.css` (new classes; reuse existing `.cell-override`, `.qaseeda-font-status`, `.field`, `.actions` styles where they fit)

**Interfaces:**
- Produces: the DOM ids consumed by Tasks 6-9 (exact list below). No JS changes in this task — the pane will render but old handlers reference removed ids; that is acceptable mid-plan ONLY if taskpane.js guards are added in the same commit, so this task also stubs the removed-id references (step 3).

- [ ] **Step 1: Replace the shared-controls section (taskpane.html lines 40-87) with the Settings panel + Fonts strip**

```html
      <section id="settings-panel" class="controls settings-panel" aria-label="Settings">
        <div class="sp-header">
          <div id="sp-target" class="sp-target">Selection</div>
          <div id="sp-chips" class="sp-chips" hidden>
            <button id="sp-chip-poem" type="button" class="sp-chip is-active" data-level="poem">Poem</button>
            <button id="sp-chip-bandh" type="button" class="sp-chip" data-level="bandh">Bandh</button>
            <button id="sp-chip-cell" type="button" class="sp-chip" data-level="cell" disabled>Cell</button>
            <button id="sp-chip-gap" type="button" class="sp-chip" data-level="gap" disabled>Gap</button>
          </div>
        </div>

        <div class="sp-profile-row">
          <label for="sp-profile">Profile</label>
          <select id="sp-profile"></select>
          <button id="sp-profile-assign" type="button" class="button--secondary">Assign</button>
          <button id="sp-profile-saveas" type="button" class="button--secondary">Save as…</button>
          <button id="sp-profile-update" type="button" hidden>Update</button>
          <button id="sp-profile-restore" type="button" hidden>Restore profile from this poem</button>
        </div>

        <div id="sp-body-poem" class="sp-body">
          <div class="field"><label for="sp-justify-mode">Justification <span class="sp-src" data-key="justifyMode"></span></label>
            <select id="sp-justify-mode" data-key="justifyMode">
              <option value="kashida" selected>Kashida</option>
              <option value="css">Word justify</option>
              <option value="spacing">Spacing preview</option>
              <option value="none">None</option>
            </select></div>
          <div class="field"><label for="sp-fill-mode">Fill mode <span class="sp-src" data-key="fillMode"></span></label>
            <select id="sp-fill-mode" data-key="fillMode">
              <option value="natural-fit" selected>Natural-fit (harmony)</option>
              <option value="cell-fit">Cell-fit (precise)</option>
              <option value="adaptive">Adaptive</option>
            </select></div>
          <div class="field"><label for="sp-strength">Kashida strength <span id="sp-strength-value">6</span> <span class="sp-src" data-key="strength"></span></label>
            <input id="sp-strength" data-key="strength" type="range" min="1" max="10" value="6"></div>
          <div class="field"><label for="sp-gap">Misra gap <span class="sp-src" data-key="gap"></span></label>
            <input id="sp-gap" data-key="gap" type="number" min="0" max="20" value="4"></div>
          <div class="field"><label for="sp-width-mode">Width <span class="sp-src" data-key="widthMode"></span></label>
            <select id="sp-width-mode" data-key="widthMode">
              <option value="auto-fit" selected>Auto-fit</option>
              <option value="fixed">Fixed %</option>
            </select>
            <input id="sp-width-pct" data-key="widthPct" type="number" min="25" max="100" step="5" value="50"></div>
          <details class="sp-advanced"><summary>Advanced</summary>
            <div class="field"><label for="sp-layout-mode">Layout <span class="sp-src" data-key="layoutMode"></span></label>
              <select id="sp-layout-mode" data-key="layoutMode">
                <option value="balanced" selected>Balanced table</option>
                <option value="equal">Equal table</option>
                <option value="compact">Compact table</option>
                <option value="stacked">Stacked</option>
                <option value="auto">Auto preview</option>
              </select></div>
            <div class="field"><label for="sp-col-width-mode">Column widths <span class="sp-src" data-key="colWidthMode"></span></label>
              <select id="sp-col-width-mode" data-key="colWidthMode">
                <option value="optimized" selected>Optimized</option>
                <option value="fixed">Fixed</option>
              </select></div>
            <div class="field"><label for="sp-misra-width">Misra width (pt) <span class="sp-src" data-key="misraWidthPt"></span></label>
              <input id="sp-misra-width" data-key="misraWidthPt" type="number" min="1" placeholder="auto (harmony)"></div>
            <div class="field"><label for="sp-corr-font">Font correction</label>
              <input id="sp-corr-font" type="text" class="template-name-input" placeholder="font name">
              <input id="sp-corr-factor" type="number" min="0.5" max="2" step="0.01" value="1"></div>
            <div class="field"><label>Debug: color tatweels</label>
              <input id="sp-debug-tatweel" type="color" value="#a7352a">
              <label class="adopt-check"><input type="checkbox" id="sp-debug-tatweel-on"> on</label></div>
            <div class="field"><label>Debug: color micro-spaces</label>
              <input id="sp-debug-space" type="color" value="#1f6f68">
              <label class="adopt-check"><input type="checkbox" id="sp-debug-space-on"> on</label></div>
          </details>
        </div>

        <div id="sp-body-bandh" class="sp-body" hidden>
          <div class="field"><label for="sp-bandh-width">Misra width (pt) <span class="sp-src" data-key="misraWidthPt"></span></label>
            <input id="sp-bandh-width" data-key="misraWidthPt" type="number" min="1" placeholder="inherit"></div>
        </div>

        <div id="sp-body-cell" class="sp-body" hidden>
          <div class="field"><label for="sp-cell-strength">Strength <span class="sp-src" data-key="strength"></span></label>
            <input id="sp-cell-strength" data-key="strength" type="number" min="1" max="10" placeholder="inherit"></div>
          <div class="field"><label for="sp-cell-width">Target width (pt) <span class="sp-src" data-key="misraWidthPt"></span></label>
            <input id="sp-cell-width" data-key="misraWidthPt" type="number" min="1" placeholder="inherit"></div>
          <div class="field"><label for="sp-cell-cap">Cap lift (em) <span class="sp-src" data-key="capEm"></span></label>
            <input id="sp-cell-cap" data-key="capEm" type="number" min="0" step="0.01" placeholder="0.28"></div>
        </div>

        <div id="sp-body-gap" class="sp-body" hidden>
          <div class="field"><label for="sp-gap-symbol">Symbol</label>
            <input id="sp-gap-symbol" type="text" placeholder="e.g. ؎ (blank = none)"></div>
          <div class="field"><label for="sp-gap-fill">Fill color</label>
            <input id="sp-gap-fill" type="color" value="#f5f0e0">
            <label class="adopt-check"><input type="checkbox" id="sp-gap-fill-on"> on</label></div>
          <div class="field"><label for="sp-gap-color">Symbol color</label>
            <input id="sp-gap-color" type="color" value="#a7352a"></div>
          <button id="sp-gap-default" type="button" class="button--secondary">Set as default for all bandhs</button>
        </div>

        <div id="sp-font-prompt" class="sp-font-prompt" hidden>
          ⚠ Can't measure "<span id="sp-font-prompt-name"></span>" — justification would be inaccurate.
          <button id="sp-font-prompt-add" type="button">Add font file…</button>
          <button id="sp-font-prompt-continue" type="button" class="button--secondary">Continue anyway</button>
        </div>

        <div class="sp-footer">
          <button id="sp-revert" type="button" class="button--secondary">Reset to defaults</button>
          <button id="sp-apply" type="button">Apply</button>
        </div>
      </section>

      <details id="fonts-strip" class="fonts-strip">
        <summary>Fonts</summary>
        <div class="field font-store">
          <input id="font-upload" type="file" accept=".ttf,.otf,.woff,.woff2">
          <input id="font-upload-name" type="text" class="template-name-input" placeholder="Register as (name Word uses)">
          <div class="font-store-actions">
            <button id="font-upload-add" type="button" class="button--secondary">Add font</button>
            <button id="font-upload-verify" type="button" class="button--secondary">Verify at cursor</button>
          </div>
          <div id="font-upload-status" class="qaseeda-font-status" aria-live="polite"></div>
          <ul id="font-list" class="font-list"></ul>
          <p class="adopt-hint">Load a font from your computer so kashida metrics match Word. Stored on this machine — nothing is uploaded to a server.</p>
        </div>
      </details>
```

Each `.sp-src` span renders the provenance dot + reset: Task 6 fills it with `•` (clickable, title = "local — click to reset") when source is `local`/dirty, empty otherwise.

- [ ] **Step 2: Delete replaced sections and add the adopt-panel button**

- Delete the Conversion tab's `layout-mode` and `width-mode` fields (the whole `<section class="controls" aria-label="Conversion controls">`, lines 192-210).
- Delete the `Justify Selected Text` and `Re-render (keep font)` buttons (lines 251-252); keep `Reset (clear kashida & spaces)` and `Show cell structure` in that actions section.
- Delete the `cell-override`, `bandh-override`, and `slot-decor` divs (lines 256-277).
- Delete the whole `qaseeda-panel` details block (lines 280-348) EXCEPT keep the `<datalist id="qaseeda-names">` idea — the new profile select `#sp-profile` replaces it entirely, so delete the datalist too.
- In the adopt panel (after `<button id="adopt-table" ...>` line 135) add:

```html
          <button id="adopt-replace-selection" type="button" class="button--secondary">Replace Selection</button>
```

- Add `<script src="settings-panel.js"></script>` before the `taskpane.js` script tag (match the existing script-include pattern at the bottom of the file).

- [ ] **Step 3: Stub removed-id references so the pane still boots**

`taskpane.js` line 21 etc. call `document.getElementById` for removed ids (`gap-width`, `justify-mode`, `tatweel-count`, `table-width`, `auto-fit-width`, `font-mode`, `justify-fill-mode`, `layout-mode`, `width-mode`, and all `qaseeda-*` / `cell-ov-*` / `bandh-ov-*` / `slot-decor-*` ids). Until Tasks 6-8 rewire them, add ONE guard shim near the top of the IIFE (after the getElementById block):

```js
  // TRANSITIONAL (removed in cleanup task): old controls were replaced by the
  // Settings panel; give dead references an inert element so bind() survives.
  function elOrStub(el) {
    return el || { value: "", checked: false, hidden: true, textContent: "",
      addEventListener: function () {}, appendChild: function () {},
      style: {}, options: [], innerHTML: "" };
  }
```

and wrap every assignment whose id was removed, e.g. `var gapWidth = elOrStub(document.getElementById("gap-width"));`. Do this for ALL removed ids (grep each old id; the full list is in this step's first paragraph).

- [ ] **Step 4: Add CSS**

Append to `taskpane.css`:

```css
/* ── Unified settings panel ─────────────────────────────────────────────── */
.settings-panel { border-bottom: 1px solid var(--border, #d8d2c4); padding-bottom: 8px; }
.sp-header { display: flex; align-items: center; justify-content: space-between; gap: 8px; }
.sp-target { font-weight: 600; }
.sp-chips { display: flex; gap: 4px; }
.sp-chip { padding: 2px 8px; border-radius: 10px; border: 1px solid var(--border, #d8d2c4); background: transparent; cursor: pointer; }
.sp-chip.is-active { background: var(--accent, #1f6f68); color: #fff; }
.sp-chip:disabled { opacity: 0.4; cursor: default; }
.sp-profile-row { display: flex; align-items: center; gap: 6px; margin: 8px 0; flex-wrap: wrap; }
.sp-src { color: var(--accent, #1f6f68); cursor: pointer; font-weight: 700; }
.sp-footer { display: flex; justify-content: space-between; margin-top: 8px; }
.sp-font-prompt { background: #fdf3e7; border: 1px solid #e8c9a0; padding: 8px; margin-top: 8px; }
.fonts-strip { margin: 8px 0; }
```

- [ ] **Step 5: Verify boot and commit**

Run: `npm test` — Expected: pass (tests don't load the HTML).
Run: `npm run dev-server` and open `https://localhost:3000/taskpane.html` in a browser — Expected: pane renders, no console errors (stub shim absorbs dead references), panel visible with all four bodies (three hidden).

```bash
git add src/taskpane/taskpane.html src/taskpane/taskpane.css src/taskpane/taskpane.js
git commit -m "feat(ui): settings panel + fonts strip markup; retire scattered controls (transitional stubs)"
```

---

### Task 6: Panel DOM glue — render, pending buffer, chips, reflection

**Files:**
- Modify: `src/taskpane/taskpane.js` (reflection block ~lines 508-620, `bind()`)

**Interfaces:**
- Consumes: `AshaarPanel.panelStateFor/mergePending`, `AshaarProfiles.resolveSettings`, `AshaarWord.parseContentControlTag`, existing `reflectActiveContext` cursor detection, existing `loadProfileStore()`.
- Produces (module-internal, used by Tasks 7-8): 
  - `_panel = { pending: {set:{},clear:[]}, scopeLevel: "poem", target: null, resolved: null }` — the panel's single state object.
  - `refreshPanel()` — re-resolves and re-renders the panel from `_panel.target`.
  - `panelValues()` — `{...resolved.values, ...pending.set}` minus cleared keys → the effective values Apply uses.

- [ ] **Step 1: Add panel state + render**

Add after the `_activeSlot` declaration (~line 514):

```js
  // ── Unified settings panel state ──────────────────────────────────────────
  var _panel = {
    pending: { set: {}, clear: [] },
    scopeLevel: "poem",
    target: null,      // { kind:"block"|"selection", cc?, payload?, scope, cellEnabled, gapEnabled, cellLabel?, gapKey? }
    resolved: null,
  };

  var SP_BODIES = { poem: "sp-body-poem", bandh: "sp-body-bandh", cell: "sp-body-cell", gap: "sp-body-gap" };

  function panelValues() {
    var out = {};
    var base = _panel.resolved ? _panel.resolved.values : AshaarProfiles.defaultSettings();
    Object.keys(base).forEach(function (k) { out[k] = base[k]; });
    _panel.pending.clear.forEach(function (k) {
      if (_panel.resolved && _panel.resolved.inherited) out[k] = _panel.resolved.inherited[k];
    });
    Object.keys(_panel.pending.set).forEach(function (k) { out[k] = _panel.pending.set[k]; });
    return out;
  }

  function refreshPanel() {
    var target = _panel.target || { kind: "selection", scope: { level: "poem" } };
    var scope = { level: _panel.scopeLevel, key: target.kind === "block"
      ? (_panel.scopeLevel === "cell" ? target.cellLabel : _panel.scopeLevel === "gap" ? target.gapKey : undefined)
      : undefined };
    _panel.resolved = AshaarProfiles.resolveSettings({
      payload: target.kind === "block" ? target.payload : null,
      profileStore: loadProfileStore(),
      scope: scope,
    });
    var st = AshaarPanel.panelStateFor({ resolved: _panel.resolved, pending: _panel.pending,
      target: { kind: target.kind, scope: scope, cellEnabled: !!target.cellEnabled,
                gapEnabled: !!target.gapEnabled, cellLabel: target.cellLabel, gapLabel: target.gapKey } });
    renderPanel(st);
  }

  function renderPanel(st) {
    document.getElementById("sp-target").textContent = st.header.title;
    var chipsWrap = document.getElementById("sp-chips");
    chipsWrap.hidden = st.chips.length === 0;
    st.chips.forEach(function (c) {
      var el = document.getElementById("sp-chip-" + c.level);
      el.disabled = !c.enabled;
      el.classList.toggle("is-active", c.active);
    });
    Object.keys(SP_BODIES).forEach(function (lvl) {
      document.getElementById(SP_BODIES[lvl]).hidden = lvl !== _panel.scopeLevel;
    });
    // Values + provenance dots. Controls carry data-key; skip ones the user is
    // mid-editing (focused).
    st.controls.forEach(function (c) {
      var body = document.getElementById(SP_BODIES[_panel.scopeLevel]);
      var input = body.querySelector('[data-key="' + c.key + '"]');
      if (input && document.activeElement !== input) {
        input.value = c.value == null ? "" : c.value;
        if (c.key === "strength") document.getElementById("sp-strength-value").textContent = String(c.value);
      }
      var src = body.querySelector('.sp-src[data-key="' + c.key + '"]');
      if (src) {
        src.textContent = (c.source === "local" || c.dirty) ? "•" : "";
        src.title = c.dirty ? "edited — Apply to commit; click to reset"
          : c.source === "local" ? "local tweak — click to reset to inherited" : "";
      }
    });
    // Profile row.
    var sel = document.getElementById("sp-profile");
    var names = listProfileNames();
    sel.innerHTML = "<option value=\"\">(none)</option>" + names.map(function (n) {
      return "<option value=\"" + String(n).replace(/"/g, "&quot;") + "\">" + String(n) + "</option>";
    }).join("");
    sel.value = st.profileRow.missing ? "" : st.profileRow.name;
    document.getElementById("sp-profile-assign").disabled = !st.profileRow.assignEnabled;
    document.getElementById("sp-profile-update").hidden = !st.profileRow.updateVisible;
    document.getElementById("sp-profile-update").textContent = "Update \"" + st.profileRow.name + "\"";
    document.getElementById("sp-profile-restore").hidden = !st.profileRow.restoreVisible;
    document.getElementById("sp-revert").textContent = st.footer.revertLabel;
  }
```

- [ ] **Step 2: Rewire reflection to feed `_panel.target`**

In `reflectActiveContext` (line 528) replace the `syncBlockControls`/`_lastBlockTag` logic: after computing `isBlock` and `payload`, and after `reflectActiveCell` has resolved the cell/gap context (reuse its existing detection of `_activeOvKey`, `_activeDecorKey`, `_activeSlot`), set:

```js
        _panel.target = isBlock
          ? { kind: "block", cc: cc, payload: payload, tag: cc.tag,
              cellEnabled: !!_activeOvKey, gapEnabled: !!_activeDecorKey,
              cellLabel: _activeOvKey, gapKey: _activeDecorKey,
              scope: { level: _panel.scopeLevel } }
          : { kind: "selection", scope: { level: "poem" } };
        if (!isBlock) _panel.scopeLevel = "poem";
        // A new block target drops stale pending edits; same block keeps them.
        if (cc.tag !== _lastBlockTag) { _panel.pending = { set: {}, clear: [] }; _lastBlockTag = isBlock ? cc.tag : null; }
        refreshPanel();
```

Delete `syncBlockControls` (lines 516-525) entirely. Keep `reflectActiveCell`'s detection logic but delete its DOM writes to the removed editors (`editor.hidden`, `bandhEl`, value population) — it now only computes `_activeOvKey`/`_activeDecorKey`/`_activeSlot`.

- [ ] **Step 3: Wire panel inputs and chips in `bind()`**

Add to `bind()`:

```js
    // Settings panel: every data-key control feeds the pending buffer.
    document.querySelectorAll("#settings-panel [data-key]").forEach(function (input) {
      input.addEventListener("change", function () {
        var key = input.getAttribute("data-key");
        var raw = input.value;
        var val = input.type === "number" || input.type === "range"
          ? (raw === "" ? null : Number(raw)) : raw;
        _panel.pending = AshaarPanel.mergePending(_panel.pending, key, val);
        refreshPanel();
      });
    });
    // Provenance dots double as per-setting reset.
    document.querySelectorAll("#settings-panel .sp-src").forEach(function (span) {
      span.addEventListener("click", function () {
        _panel.pending = AshaarPanel.mergePending(_panel.pending, span.getAttribute("data-key"), null);
        refreshPanel();
      });
    });
    ["poem", "bandh", "cell", "gap"].forEach(function (lvl) {
      document.getElementById("sp-chip-" + lvl).addEventListener("click", function () {
        _panel.scopeLevel = lvl;
        _panel.pending = { set: {}, clear: [] }; // scope switch discards unapplied edits
        refreshPanel();
      });
    });
    document.getElementById("sp-revert").addEventListener("click", revertToProfile);
    document.getElementById("sp-apply").addEventListener("click", applyPanel);
    document.getElementById("sp-profile-assign").addEventListener("click", assignProfile);
    document.getElementById("sp-profile-saveas").addEventListener("click", saveAsProfile);
    document.getElementById("sp-profile-update").addEventListener("click", updateProfile);
    document.getElementById("sp-profile-restore").addEventListener("click", restoreProfileFromPoem);
    document.getElementById("adopt-replace-selection").addEventListener("click", function () { insertPoem(true); });
```

(`applyPanel`, `assignProfile`, `saveAsProfile`, `updateProfile`, `restoreProfileFromPoem`, `revertToProfile` are Task 7/8 — add empty `async function` declarations now so bind() resolves:  `async function applyPanel() { setMessage("Apply lands in the next task."); }` etc.)

- [ ] **Step 4: Verify in dev server**

Run: `npm run dev-server`; open the pane in a browser (no Word):
- Panel shows `Selection`, defaults visible.
- Changing gap shows the dirty dot; clicking the dot clears it.
- Chips hidden (no block detectable outside Word).
Run: `npm test` — Expected: pass.

- [ ] **Step 5: Commit**

```bash
git add src/taskpane/taskpane.js
git commit -m "feat(panel): DOM glue — resolver-driven render, pending buffer, scope chips"
```

---

### Task 7: Apply routing

**Files:**
- Modify: `src/taskpane/taskpane.js`

**Interfaces:**
- Consumes: `_panel`, `panelValues()`, `AshaarPanel.pendingToLocal`, `AshaarWord.setTagLocal/setTagOverride/setTagBandhWidth/setTagSlotDecor/setTagProfileCache`, existing `justifySelection()` (line ~2400s), `reRender()` internals, `applyProfileToQaseeda` machinery, `insertPoem`.
- Produces: `applyPanel()` — the single Apply entry; `options()` rewired to read `panelValues()`.

- [ ] **Step 1: Rewire `options()`**

`options()` (currently reads the deleted controls at ~line 250-260 and 3230s) becomes a thin adapter so ALL existing insert paths keep working:

```js
  function options() {
    var v = panelValues();
    return {
      layoutMode: v.layoutMode,
      widthMode: v.colWidthMode,
      justifyMode: v.justifyMode,
      fillMode: v.fillMode,
      tatweelCount: Number(v.strength || 6),
      gapWidth: Number(v.gap != null ? v.gap : 4),
      misraPattern: currentMisraPattern(),      // keep existing computation
      misraCount: Number(misraCount.value || 4), // still a Table-Input control
      fontMode: "document",
      tableWidthPct: v.widthMode === "fixed" ? Number(v.widthPct || 50) : 100,
      autoFitWidth: v.widthMode !== "fixed",
      // v3 tag fields for fresh inserts:
      profile: _panel.resolved ? _panel.resolved.profileName : "",
      local: AshaarPanel.pendingToLocal({}, _panel.pending, AshaarPanel.SCOPE_FIELDS.poem),
    };
  }
```

Adjust the two `contentControlTag(...)` insert-time call sites (grep `contentControlTag(source, opts)` / `contentControlTag("grid"` etc. — grid/template tags are NOT poem tags and keep their existing opts) so poem inserts pass `opts.profile`/`opts.local` through (Task 2's writer reads them).

- [ ] **Step 2: Implement `applyPanel()` routing**

```js
  // One Apply: route by target and scope. Writes deltas to the owning tag
  // slot, then re-renders/justifies. Pending clears only on success.
  async function applyPanel() {
    var target = _panel.target;
    var values = panelValues();
    try {
      if (!target || target.kind !== "block") {
        // Plain selection: one-shot justify with panel values; nothing persisted.
        await justifySelection();   // justifySelection reads options() → panelValues()
        _panel.pending = { set: {}, clear: [] };
        refreshPanel();
        return;
      }
      if (_panel.scopeLevel === "poem") {
        await applyPoemScope(target, values);
      } else if (_panel.scopeLevel === "bandh") {
        await withWord(async function (context) {
          var cc = await findBlockAt(context);           // helper below
          cc.tag = AshaarWord.setTagBandhWidth(cc.tag, values.misraWidthPt || 0);
          await context.sync();
        });
        await reapplyBlock();                            // re-justify in place
      } else if (_panel.scopeLevel === "cell") {
        await withWord(async function (context) {
          var cc = await findBlockAt(context);
          cc.tag = AshaarWord.setTagOverride(cc.tag, target.cellLabel, {
            strength: dirtyOrNull("strength"), widthPt: dirtyOrNull("misraWidthPt"), capEm: dirtyOrNull("capEm"),
          });
          await context.sync();
        });
        await reapplyBlock();
      } else if (_panel.scopeLevel === "gap") {
        await withWord(async function (context) {
          var cc = await findBlockAt(context);
          cc.tag = AshaarWord.setTagSlotDecor(cc.tag, target.gapKey, {
            symbol: document.getElementById("sp-gap-symbol").value,
            fill: document.getElementById("sp-gap-fill-on").checked ? document.getElementById("sp-gap-fill").value : "",
            color: document.getElementById("sp-gap-color").value,
          });
          await context.sync();
        });
        await reapplyBlock();
      }
      _panel.pending = { set: {}, clear: [] };
      _lastBlockTag = null;   // force reflection to re-read the updated tag
      await reflectActiveContext();
      setMessage("Applied.");
    } catch (e) {
      // Keep pending for retry (spec: apply failure keeps edits).
      setMessage("Apply failed: " + (e && e.message ? e.message : e));
    }
  }

  function dirtyOrNull(key) {
    return (key in _panel.pending.set) ? _panel.pending.set[key]
      : (_panel.resolved && _panel.resolved.source[key] === "cell" ? _panel.resolved.values[key] : null);
  }

  // Locate the enclosing Ashaar Poem control at the cursor (throws if none).
  async function findBlockAt(context) {
    var sel = context.document.getSelection();
    var cc = sel.parentContentControlOrNullObject;
    cc.load("title,tag");
    await context.sync();
    if (cc.isNullObject || cc.title !== "Ashaar Poem") throw new Error("Click inside an Ashaar poem first.");
    return cc;
  }
```

- [ ] **Step 3: Implement `applyPoemScope` on the existing pipelines**

```js
  // Poem scope: persist local deltas, then rebuild-if-structural + justify.
  // Reuses reRender()'s bare-rebuild + justifySelection() fill, both of which
  // now read options() → panelValues(), i.e. the resolved values.
  async function applyPoemScope(target, values) {
    var structuralDirty = ["gap", "widthMode", "widthPct", "layoutMode", "colWidthMode"].some(function (k) {
      return (k in _panel.pending.set) || _panel.pending.clear.indexOf(k) !== -1;
    });
    await withWord(async function (context) {
      var cc = await findBlockAt(context);
      var payload = AshaarWord.parseContentControlTag(cc.tag);
      var newLocal = AshaarPanel.pendingToLocal(payload.local, _panel.pending, AshaarPanel.SCOPE_FIELDS.poem);
      var tag = AshaarWord.setTagLocal(cc.tag, newLocal);
      // Snapshot the profile layer for cross-machine portability.
      var prof = payload.profile ? loadProfileStore()[payload.profile] : null;
      if (prof) tag = AshaarWord.setTagProfileCache(tag, AshaarProfiles.settingsFromProfile(prof));
      cc.tag = tag;
      await context.sync();
    });
    if (structuralDirty) await reRender();     // bare rebuild + in-place justify
    else await justifySelection();             // justify only — no destructive rebuild
  }

  // Re-justify the block in place (non-structural scopes).
  async function reapplyBlock() { await justifySelection(); }
```

Also update the two `p.gapWidth` reads at `taskpane.js:893` and `taskpane.js:1241`: both are inside per-block loops with `p = blk.payload` — replace with resolver reads:

```js
      var eff = AshaarProfiles.resolveSettings({ payload: p, profileStore: loadProfileStore(), scope: { level: "poem" } }).values;
      var geomOpts = { gapWidth: eff.gap, layoutMode: eff.layoutMode };
```

and in the `renderOpts` at 1239-1247: `gapWidth: eff.gap, layoutMode: eff.layoutMode, misraPattern: p.misraPattern || "paired", misraCount: Number(p.misraCount || 4), fontMode: "document", tatweelCount: 0, justifyMode: "none"` (compute `eff` once per block). Then update Task 3's interim `sizeSig` gap component to use `eff.gap` per block.

- [ ] **Step 4: Test in Word (manual, scripted checklist)**

Run: `npm start` (opens Word with the test doc). Verify:
1. Insert a poem from Conversion tab → block renders; tag is v3 (Debug: `Show cell structure` still works).
2. Click inside → header `Poem`; set gap 8 → dot appears → Apply → table rebuilds with wider gap; click out/in → gap shows 8 with local dot.
3. Cell chip → strength 9 → Apply → only that cell fattens.
4. Plain paragraph text selected → header `Selection` → Apply → text justifies, no tag created.

Run: `npm test` — Expected: pass.

- [ ] **Step 5: Commit**

```bash
git add src/taskpane/taskpane.js
git commit -m "feat(apply): panel Apply routed by target/scope through resolver-fed pipelines"
```

---

### Task 8: Profile actions — Assign, Save as, Update, Restore, Revert

**Files:**
- Modify: `src/taskpane/taskpane.js`

**Interfaces:**
- Consumes: `putProfile`, `getProfile`, `loadProfileStore`, `applyProfileToQaseeda` (existing, ~line 1142), `AshaarProfiles.profileFromSettings/settingsFromProfile`, `AshaarWord.setTagProfile/setTagLocal/setTagProfileCache`.
- Produces: the four handlers stubbed in Task 6 plus `revertToProfile`.

- [ ] **Step 1: Implement the handlers**

```js
  // Assign: link the block to the selected profile; local tweaks survive.
  async function assignProfile() {
    var name = document.getElementById("sp-profile").value;
    await withWord(async function (context) {
      var cc = await findBlockAt(context);
      cc.tag = AshaarWord.setTagProfile(cc.tag, name);
      await context.sync();
    });
    if (name) await applyProfileToQaseeda(name);
    _lastBlockTag = null; await reflectActiveContext();
    setMessage(name ? "Assigned to \"" + name + "\"." : "Profile link removed.");
  }

  // Save as…: panel's resolved+pending values → new profile; block assigned;
  // local cleared (the tweaks just became the profile).
  async function saveAsProfile() {
    var name = (prompt("Save current settings as profile:") || "").trim();
    if (!name) return;
    var values = panelValues();
    await putProfile(AshaarProfiles.profileFromSettings(name, values));
    if (_panel.target && _panel.target.kind === "block") {
      await withWord(async function (context) {
        var cc = await findBlockAt(context);
        var tag = AshaarWord.setTagProfile(cc.tag, name);
        tag = AshaarWord.setTagLocal(tag, {});
        tag = AshaarWord.setTagProfileCache(tag, AshaarProfiles.settingsFromProfile(AshaarProfiles.profileFromSettings(name, values)));
        cc.tag = tag;
        await context.sync();
      });
      await applyProfileToQaseeda(name);
    }
    _panel.pending = { set: {}, clear: [] };
    _lastBlockTag = null; await reflectActiveContext();
    setMessage("Profile \"" + name + "\" saved.");
  }

  // Update "name": push panel values into the stored profile; re-apply to all
  // its blocks (each block's own local map survives via the resolver).
  async function updateProfile() {
    var name = _panel.resolved ? _panel.resolved.profileName : "";
    if (!name) return;
    await putProfile(AshaarProfiles.profileFromSettings(name, panelValues()));
    _panel.pending = { set: {}, clear: [] };
    await applyProfileToQaseeda(name);
    _lastBlockTag = null; await reflectActiveContext();
    setMessage("Profile \"" + name + "\" updated — all its poems refreshed.");
  }

  // Restore a missing profile from the tag's cached snapshot.
  async function restoreProfileFromPoem() {
    var t = _panel.target;
    if (!t || t.kind !== "block" || !t.payload || !t.payload.profileCache) return;
    var name = t.payload.profile;
    await putProfile(AshaarProfiles.profileFromSettings(name, t.payload.profileCache));
    _lastBlockTag = null; await reflectActiveContext();
    setMessage("Profile \"" + name + "\" restored from this poem.");
  }

  // Revert to profile / Reset to defaults: clear the whole local map.
  async function revertToProfile() {
    _panel.pending = { set: {}, clear: [] };
    if (_panel.target && _panel.target.kind === "block") {
      await withWord(async function (context) {
        var cc = await findBlockAt(context);
        cc.tag = AshaarWord.setTagLocal(cc.tag, {});
        await context.sync();
      });
      await reRender();     // structure may change (gap/width may fall back)
      _lastBlockTag = null; await reflectActiveContext();
    } else {
      refreshPanel();
    }
    setMessage("Reverted.");
  }
```

Replace the Task 6 stubs with these bodies. Also update `applyProfileToQaseeda`'s internals: `gatherQaseedaBlocks(context, name)` matches blocks by `payload.qaseeda` — thanks to the parse alias it already sees v3 `profile`; no change needed yet (alias removed in Task 10, at which point switch it to `payload.profile`).

- [ ] **Step 2: Also wire `sp-gap-default`**

The Gap body's `Set as default for all bandhs` button writes the decoration into the assigned profile's `spacingDecor` (explicit-mutation rule — it edits the profile, so reuse the existing `slot-decor-save-profile` handler logic; grep `slot-decor-save-profile` and move that handler body to the new id, reading from the `sp-gap-*` inputs).

- [ ] **Step 3: Test in Word (manual)**

1. Two poems, both assigned "Test". Tweak gap on poem A → Apply.
2. Poem B → change strength → `Update "Test"` → both poems change strength; A keeps its gap.
3. `Save as…` from A with a tweak → new profile; A's dot disappears (local cleared).
4. Delete the profile from localStorage via DevTools (`localStorage` key used by `loadProfileStore`) → click in poem → header shows `(not on this machine)` → Restore → header heals.
5. `Revert to profile` on A → gap falls back.

Run: `npm test` — Expected: pass.

- [ ] **Step 4: Commit**

```bash
git add src/taskpane/taskpane.js
git commit -m "feat(profiles): assign/save-as/update/restore/revert wired to the panel"
```

---

### Task 9: Just-in-time font measurement prompt

**Files:**
- Modify: `src/taskpane/taskpane.js` (the face-loading block at ~line 918-930 inside `captureQaseedaTables`/apply path, and `justifySelection`'s font-load step)

**Interfaces:**
- Consumes: `document.fonts.load/check`, the `#sp-font-prompt*` DOM from Task 5, `AshaarFontStore` (existing uploader module).
- Produces: `ensureFacesMeasurable(faceNames)` → `Promise<"ok"|"continue"|"cancel">`.

- [ ] **Step 1: Implement the gate**

```js
  // Force-load every face; when one is invisible to the WebView (Word renders
  // it but canvas can't measure it) surface the inline prompt. Resolves "ok"
  // when all faces measure, "continue" when the user accepts fallback metrics,
  // "cancel" when they go add the font.
  async function ensureFacesMeasurable(faceNames) {
    var missing = [];
    for (var i = 0; i < faceNames.length; i++) {
      var f = faceNames[i];
      try { await document.fonts.load('16pt "' + f + '"'); } catch (e) {}
      if (!document.fonts.check('16pt "' + f + '"')) missing.push(f);
    }
    if (!missing.length) return "ok";
    return await new Promise(function (resolve) {
      var box = document.getElementById("sp-font-prompt");
      document.getElementById("sp-font-prompt-name").textContent = missing.join('", "');
      box.hidden = false;
      document.getElementById("sp-font-prompt-add").onclick = function () {
        box.hidden = true;
        var strip = document.getElementById("fonts-strip");
        strip.open = true;
        document.getElementById("font-upload-name").value = missing[0];
        resolve("cancel");
      };
      document.getElementById("sp-font-prompt-continue").onclick = function () {
        box.hidden = true;
        resolve("continue");
      };
    });
  }
```

- [ ] **Step 2: Gate the two apply paths**

In `applyPanel()` (before routing) and inside `applyProfileToQaseeda` (after the distinct-face set is built at ~line 918 — the `faceSet` object exists there): collect the distinct face names, call `ensureFacesMeasurable(names)`; on `"cancel"` abort with `setMessage("Add the font, then Apply again.")`; on `"continue"` proceed. For `applyPanel`'s plain-selection route, read the selection's `font.name` inside the existing `justifySelection` capture and gate on that single face.

- [ ] **Step 3: Test in Word (manual)**

1. Poem in a registered/system font → Apply → no prompt.
2. Poem in a Word-only font not visible to the WebView (per memory `font-measurement-model`: use a sandbox-inaccessible font) → Apply → prompt appears; `Add font file…` opens the strip pre-filled; after registering via the uploader, Apply again → no prompt.
3. `Continue anyway` → apply proceeds.

Run: `npm test` — Expected: pass.

- [ ] **Step 4: Commit**

```bash
git add src/taskpane/taskpane.js
git commit -m "feat(fonts): just-in-time unmeasurable-font prompt gating Apply"
```

---

### Task 10: Cleanup, alias removal, docs, manual checklist

**Files:**
- Modify: `src/taskpane/taskpane.js`, `src/taskpane/word-html.js`, `src/taskpane/taskpane.html`, `CLAUDE.md`
- Create: `docs/superpowers/specs/2026-07-12-unified-settings-manual-checklist.md`

- [ ] **Step 1: Remove the transitional stubs and dead code**

- Delete `elOrStub` and every stubbed variable that no longer has a reader (grep each: `gapWidth`, `justifyMode` (the old select var), `tatweelCount`, `tableWidth`, `autoFitWidth`, `fontMode`, `justifyFillMode`, `layoutMode` var, `widthMode` var, all `qaseeda*` vars, `cellOv*`, `bandhOv*`, `slotDecor*` vars). Template capture (line ~3236-3239) reads `fontMode.value`/`justifyMode.value` etc. — switch it to `panelValues()` fields (`justifyMode: v.justifyMode, tatweelCount: v.strength, gapWidth: v.gap`, drop `fontMode`).
- Delete `panelToProfile`/`profileToPanel`/`loadQaseedaIntoPanel`/`saveAndApplyQaseeda`/`assignBlockToQaseeda`/`populateQaseedaNames` (lines ~1676-1744) and their bind() wiring.
- Delete the old `justify-selection`/`re-render` button bindings (the functions stay — Apply calls them).
- In `word-html.js`: delete the `setTagQaseeda` alias + export and the `payload.qaseeda = payload.profile` alias line; switch `gatherQaseedaBlocks` (taskpane.js) and any remaining `.qaseeda` readers (grep `\.qaseeda`) to `.profile`. Update the affected assertions in `tests/word-html.test.js` (drop the alias test).
- `grep -rn "elOrStub\|qaseeda-" src/taskpane/*.js src/taskpane/*.html` — Expected: no hits (CSS class names may remain).

- [ ] **Step 2: Run everything**

Run: `npm test` — Expected: full chain passes.
Run: `npm run dev-server` + browser smoke: pane boots, no console errors.

- [ ] **Step 3: Write the manual checklist doc**

Create `docs/superpowers/specs/2026-07-12-unified-settings-manual-checklist.md`:

```markdown
# Unified settings panel — manual Word verification

Run with `npm start` against test-documents/marsiya-test.docx.

- [ ] Fresh insert from Conversion tab creates a v3-tagged block; panel header shows `Poem` on click-in.
- [ ] Gap 8 → Apply on poem A only; poem B untouched; dot shown on A's gap; ⟲ + Apply reverts.
- [ ] Assign both poems to "Test"; Update "Test" changes both EXCEPT A's tweaked gap.
- [ ] Save as… from a tweaked poem clears its dots and creates the profile.
- [ ] Delete profile in DevTools localStorage → header `(not on this machine)` → Restore heals it.
- [ ] Plain-text selection: header `Selection`, Apply justifies, no content control created.
- [ ] Cell chip: strength override applies to one cell; Clear via ⟲ + Apply falls back.
- [ ] Gap chip: symbol + colors render in the spacing cell; `Set as default for all bandhs` propagates.
- [ ] Unmeasurable font prompts once; after registering via Fonts strip it never prompts again.
- [ ] Legacy document (v2 tags, from main branch): opens, panel reflects old settings as local dots, Apply upgrades tag to v3, nothing visually moves except requested changes.
- [ ] Adopt Existing Table → Replace Selection (Table Input tab) round-trips.
- [ ] Apply twice in a row is idempotent (Debug dump: no growing nat/target/nSp/segs).
```

- [ ] **Step 4: Update CLAUDE.md**

In the Architecture section, replace the "Justify Selected Text" flow block with the panel flow (Settings panel → resolver → Apply routes by target), note tag payload v3 (`profile`/`local`/`profileCache`), and note that `settings-panel.js` is the pure panel-state module.

- [ ] **Step 5: Final commit**

```bash
git add -A
git commit -m "chore: retire v2 aliases and dead controls; manual checklist + docs"
```

---

## Self-review notes (already applied)

- Spec coverage: panel/chips (T5-6), overlay+resolver (T1), tag v3+migration (T2), sizeSig fix (T3+T7), Apply routing incl. plain selection (T7), profile actions + profileCache/restore (T8), font prompt (T9), UI disposition + adopt Replace Selection (T5) , cleanup+checklist (T10). Gap-decor profile default: T8 step 2.
- Type consistency: canonical keys identical across T1 resolver, T2 migration, T4 SCOPE_FIELDS, T6 data-key attributes, T7 options() mapping.
- Known judgment calls for the implementer: `misraCount` stays a Table-Input control (structure, not a setting); scope-switch discards unapplied edits (simplest safe rule); `prompt()` for Save-as name is acceptable v1 UX.
