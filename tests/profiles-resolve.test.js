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

// ── Provenance heuristic (accepted decision): profile value equal to the
// default keeps source "default"; only differing values claim "profile" ─────
{
  const values = Object.assign(defaultSettings(), { strength: 9 });
  // gap stays at the default (4) but the profile still emits it.
  const store = { K: profileFromSettings("K", values) };
  const r = resolveSettings({ payload: { profile: "K", local: {} }, profileStore: store, scope: { level: "poem" } });
  assert.equal(r.values.strength, 9);
  assert.equal(r.source.strength, "profile", "differing value claims profile");
  assert.equal(r.values.gap, 4);
  assert.equal(r.source.gap, "default", "profile value equal to default stays default");
  assert.equal(r.values.justifyMode, "kashida");
  assert.equal(r.source.justifyMode, "default", "profile-owned justifyMode at default value stays default");
}

// ── inherited: the layer below local, for per-setting reset display ──────────
{
  const store = { K: profileFromSettings("K", Object.assign(defaultSettings(), { gap: 6 })) };
  const r = resolveSettings({ payload: { profile: "K", local: { gap: 8 } }, profileStore: store, scope: { level: "poem" } });
  assert.equal(r.values.gap, 8);
  assert.equal(r.inherited.gap, 6, "what gap falls back to on clear");
}

// ── §9 vertical rhythm keys ──────────────────────────────────────────────────
{
  const d = defaultSettings();
  assert.strictEqual(d.lineHeightPt, null, "line height defaults to Word auto");
  assert.strictEqual(d.separatorPt, 1, "separator defaults to 1pt");

  const values = Object.assign(defaultSettings(), { lineHeightPt: 24, separatorPt: 6 });
  const p = profileFromSettings("V", values);
  assert.strictEqual(p.lineHeightPt, 24);
  assert.strictEqual(p.separatorPt, 6);
  const back = settingsFromProfile(p);
  assert.strictEqual(back.lineHeightPt, 24);
  assert.strictEqual(back.separatorPt, 6);

  // Layer through the resolver like any canonical key.
  const store = { V: p };
  const r = resolveSettings({ payload: { profile: "V", local: { separatorPt: 2 } }, profileStore: store, scope: { level: "poem" } });
  assert.strictEqual(r.values.lineHeightPt, 24, "profile layer");
  assert.strictEqual(r.values.separatorPt, 2, "local wins");
  assert.strictEqual(r.inherited.separatorPt, 6, "inherited = profile layer");
}

console.log("profiles-resolve tests passed");
