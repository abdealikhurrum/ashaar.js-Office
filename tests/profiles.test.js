const assert = require("assert");
const {
  defaultProfile,
  normalizeProfile,
  mergeProfile,
  applyFontCorrection,
  deriveSharedWidths,
  columnPointsFromContentPx,
  strengthToTargetFill,
  normalizeStrength,
  normalizeFillMode,
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

// ── columnPointsFromContentPx (content px → column width in points) ───────────

{
  // 192 px content @ 96dpi = 144 pt; + 2×5.76pt cell margins = 155.52 pt.
  assert.ok(Math.abs(columnPointsFromContentPx(192, 5.76) - 155.52) < 1e-6, "px→points incl. both margins");
  assert.ok(Math.abs(columnPointsFromContentPx(96, 0) - 72) < 1e-6, "no margin => pure px→pt");
}

// ── strengthToTargetFill (kashida-strength slider → targetFill, 1-10 domain) ──

{
  // strength domain is now 1–10
  assert.ok(Math.abs(strengthToTargetFill(1) - 0.90) < 1e-9, "strength 1 => 0.90 fill");
  assert.ok(Math.abs(strengthToTargetFill(10) - 1.0) < 1e-9, "strength 10 => 1.0 fill");
}

// ── normalizeStrength (plain clamp to [1,10], no legacy remap) ────────────────

{
  // normalizeStrength is a plain clamp to [1,10] (no legacy remap)
  assert.strictEqual(normalizeStrength(6), 6, "in-range => unchanged");
  assert.strictEqual(normalizeStrength(0), 1, "clamps low");
  assert.strictEqual(normalizeStrength(24), 10, "clamps high");
  assert.strictEqual(normalizeStrength(undefined), 1, "undefined => 1");
}

// ── fillMode: default natural-fit, normalized, preserved through merge ────────
{
  const p = defaultProfile("Karbala");
  assert.equal(p.justify.fillMode, "natural-fit", "defaults to natural-fit");
}
assert.equal(normalizeFillMode("cell-fit"), "cell-fit");
assert.equal(normalizeFillMode("natural-fit"), "natural-fit");
assert.equal(normalizeFillMode("bogus"), "natural-fit", "unknown => default");
assert.equal(normalizeFillMode(undefined), "natural-fit");
{
  // A stored profile carrying cell-fit survives normalizeProfile.
  const stored = { name: "Q", justify: { mode: "kashida", strength: 7, fillMode: "cell-fit" } };
  const n = normalizeProfile(stored);
  assert.equal(n.justify.fillMode, "cell-fit", "stored fillMode preserved");
}

// ── spacingDecor bucket ──────────────────────────────────────────────────────
{
  assert.deepStrictEqual(defaultProfile("Q").spacingDecor, {}, "defaults to empty");
  const n = normalizeProfile({ name: "Q", spacingDecor: { "A#1": { symbol: "؎" } } });
  assert.deepStrictEqual(n.spacingDecor, { "A#1": { symbol: "؎" } }, "carried through merge");
}

console.log("profiles tests passed");
