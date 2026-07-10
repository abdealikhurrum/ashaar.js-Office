"use strict";
const assert = require("assert");
const AshaarFonts = require("../src/taskpane/fonts");

// mechanism tags
assert.strictEqual(AshaarFonts.mechanismOf("mehr"), "tatweel");
assert.strictEqual(AshaarFonts.mechanismOf("jameel"), "font-swap");
assert.strictEqual(AshaarFonts.mechanismOf("gulzar"), "whitespace");
assert.strictEqual(AshaarFonts.mechanismOf("noto"), "whitespace");

// reader-install note: Mehr and Jameel need a reader-side font install;
// Gulzar does not (Gate G2 2026-07-10 reclassified Jameel to font-swap,
// but it still needs the note decoupled from mechanism via `readerNote`)
assert.strictEqual(AshaarFonts.get("mehr").readerNote, true);
assert.strictEqual(AshaarFonts.get("jameel").readerNote, true);
assert.ok(!AshaarFonts.get("gulzar").readerNote);
// unknown / plain modes default to whitespace (never the tatweel engine)
assert.strictEqual(AshaarFonts.mechanismOf("document"), "whitespace");
assert.strictEqual(AshaarFonts.mechanismOf("nope"), "whitespace");

// mechanismForFontName: resolve the mechanism from a run's ACTUAL Word font
// name (per-run justify dispatch). Registry fonts match by their wordName /
// kasheedaName; anything unrecognised — arbitrary Arabic fonts (Fatemi Maqala),
// Latin defaults, empty/null — resolves to "generic" (the tatweel engine), NOT
// "whitespace", so those runs still kashida instead of being forced to spacing.
assert.strictEqual(AshaarFonts.mechanismForFontName("Mehr Nastaliq Web"), "tatweel");
assert.strictEqual(AshaarFonts.mechanismForFontName("Jameel Noori Nastaleeq"), "font-swap");         // base face
assert.strictEqual(AshaarFonts.mechanismForFontName("Jameel Noori Nastaleeq Kasheeda"), "font-swap"); // wide face
assert.strictEqual(AshaarFonts.mechanismForFontName("Noto Nastaliq Urdu"), "whitespace");
assert.strictEqual(AshaarFonts.mechanismForFontName("Gulzar"), "whitespace");
assert.strictEqual(AshaarFonts.mechanismForFontName("Scheherazade New"), "whitespace");
assert.strictEqual(AshaarFonts.mechanismForFontName("Fatemi Maqala"), "generic");
assert.strictEqual(AshaarFonts.mechanismForFontName("Times New Roman"), "generic");
assert.strictEqual(AshaarFonts.mechanismForFontName(""), "generic");
assert.strictEqual(AshaarFonts.mechanismForFontName(null), "generic");
assert.strictEqual(AshaarFonts.mechanismForFontName("  Gulzar  "), "whitespace"); // trims whitespace

// Word cs names line up with what callers emit
assert.strictEqual(AshaarFonts.wordNameOf("mehr"), "Mehr Nastaliq Web");
assert.strictEqual(AshaarFonts.wordNameOf("gulzar"), "Gulzar");
assert.strictEqual(AshaarFonts.wordNameOf("jameel"), "Jameel Noori Nastaleeq"); // base face
assert.strictEqual(AshaarFonts.wordNameOf("document"), null);

// kasheeda (wider, font-swap target) names
assert.strictEqual(AshaarFonts.kasheedaNameOf("jameel"), "Jameel Noori Nastaleeq Kasheeda");
assert.strictEqual(AshaarFonts.kasheedaNameOf("mehr"), null);

// css families
assert.ok(/Mehr Nastaliq Web/.test(AshaarFonts.cssFamilyOf("mehr")));
assert.strictEqual(AshaarFonts.cssFamilyOf("document"), null);

// Mehr whitelist verbatim — allowed tatweel letters by shaping form
const r = AshaarFonts.tatweelRulesOf("mehr");
assert.deepStrictEqual(r.isolatedInto, ["ب","پ","ت","ٹ","ث","س","ش","ف","ک","گ"]);
assert.deepStrictEqual(r.finalInto,    ["ب","پ","ت","ٹ","ث","ف","ک","گ"]);
assert.strictEqual(AshaarFonts.tatweelRulesOf("gulzar"), null);

console.log("fonts tests passed");
