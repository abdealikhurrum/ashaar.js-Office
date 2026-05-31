const assert = require("assert");
const {
  defaultProfile,
  normalizeProfile,
  mergeProfile,
  applyFontCorrection,
  deriveSharedWidths,
} = require("../src/taskpane/profiles");

// ── defaultProfile ──────────────────────────────────────────────────────────

{
  const p = defaultProfile("Karbala");
  assert.equal(p.name, "Karbala", "carries the name");
  assert.equal(p.width.mode, "auto-fit", "width defaults to auto-fit");
  assert.equal(typeof p.width.pct, "number", "width has a numeric pct");
  assert.equal(p.justify.mode, "kashida", "justify defaults to kashida");
  assert.equal(typeof p.justify.strength, "number", "justify has a numeric strength");
  assert.equal(typeof p.gap, "number", "gap is numeric");
  assert.deepEqual(p.fontCorrections, {}, "no font corrections by default");
  assert.ok(p.derived && typeof p.derived === "object", "derived bucket exists");
  assert.equal(p.derived.colWidthVector, null, "no derived widths yet");
}

// ── normalizeProfile (fill missing fields, keep provided) ─────────────────────

{
  const p = normalizeProfile({ name: "X", justify: { strength: 24 } });
  assert.equal(p.name, "X", "keeps name");
  assert.equal(p.justify.strength, 24, "keeps provided strength");
  assert.equal(p.justify.mode, "kashida", "fills missing justify.mode");
  assert.equal(p.width.mode, "auto-fit", "fills missing width block");
  assert.deepEqual(p.fontCorrections, {}, "fills missing fontCorrections");
}

{
  // A profile with no name still normalizes (name stays empty string, not undefined).
  const p = normalizeProfile({});
  assert.equal(p.name, "", "missing name becomes empty string");
}

// ── mergeProfile (deep-merge overrides onto a base) ───────────────────────────

{
  const base = defaultProfile("Q");
  const merged = mergeProfile(base, { width: { pct: 80 }, justify: { mode: "spacing" } });
  assert.equal(merged.width.pct, 80, "overrides width.pct");
  assert.equal(merged.width.mode, "auto-fit", "keeps un-overridden width.mode");
  assert.equal(merged.justify.mode, "spacing", "overrides justify.mode");
  assert.equal(merged.name, "Q", "keeps base name");
  // Base must not be mutated.
  assert.equal(base.width.pct !== 80 || base.justify.mode === "kashida", true, "does not mutate base justify");
  assert.equal(base.justify.mode, "kashida", "base unchanged");
}

// ── applyFontCorrection ───────────────────────────────────────────────────────

{
  assert.ok(Math.abs(applyFontCorrection(100, "Arial", { Arial: 1.1 }) - 110) < 1e-6, "applies the per-font factor");
  assert.equal(applyFontCorrection(100, "Unknown", { Arial: 1.1 }), 100, "unknown font => factor 1");
  assert.equal(applyFontCorrection(100, "Arial", null), 100, "no corrections => factor 1");
}

// ── deriveSharedWidths (widest corrected measurement per column × headroom) ────

{
  const columns = [
    [{ px: 100, font: "A" }, { px: 120, font: "B" }], // column 0 across two blocks
    [{ px: 50, font: "A" }],                          // column 1
  ];
  const out = deriveSharedWidths(columns, { headroom: 1.1, corrections: { B: 1.2 } });
  // col0: max(100*1, 120*1.2=144) = 144 ; *1.1 = 158.4
  assert.ok(Math.abs(out[0] - 158.4) < 1e-6, "column 0 = widest corrected × headroom");
  assert.ok(Math.abs(out[1] - 55) < 1e-6, "column 1 = 50 × headroom");
}

{
  // Defaults: headroom 1, no corrections.
  const out = deriveSharedWidths([[{ px: 200, font: "A" }]], {});
  assert.equal(out[0], 200, "no headroom/corrections => raw max");
}

{
  // Empty column yields 0 (no measurements).
  const out = deriveSharedWidths([[], [{ px: 30, font: "A" }]], { headroom: 1 });
  assert.equal(out[0], 0, "empty column => 0");
  assert.equal(out[1], 30, "populated column => its max");
}

console.log("profiles tests passed");
