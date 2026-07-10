"use strict";
const assert = require("assert");
const AshaarFonts = require("../src/taskpane/fonts");

// mechanism tags
assert.strictEqual(AshaarFonts.mechanismOf("mehr"), "tatweel");
assert.strictEqual(AshaarFonts.mechanismOf("jameel"), "italic-run");
assert.strictEqual(AshaarFonts.mechanismOf("gulzar"), "whitespace");
assert.strictEqual(AshaarFonts.mechanismOf("noto"), "whitespace");
// unknown / plain modes default to whitespace (never the tatweel engine)
assert.strictEqual(AshaarFonts.mechanismOf("document"), "whitespace");
assert.strictEqual(AshaarFonts.mechanismOf("nope"), "whitespace");

// Word cs names line up with what callers emit
assert.strictEqual(AshaarFonts.wordNameOf("mehr"), "Mehr Nastaliq Web");
assert.strictEqual(AshaarFonts.wordNameOf("gulzar"), "Gulzar");
assert.strictEqual(AshaarFonts.wordNameOf("jameel"), "Jameel Noori Nastaleeq");
assert.strictEqual(AshaarFonts.wordNameOf("document"), null);

// css families
assert.ok(/Mehr Nastaliq Web/.test(AshaarFonts.cssFamilyOf("mehr")));
assert.strictEqual(AshaarFonts.cssFamilyOf("document"), null);

// Mehr whitelist verbatim
const r = AshaarFonts.tatweelRulesOf("mehr");
assert.deepStrictEqual(r.medialInto, ["ب","پ","ت","ٹ","ث","س","ش","ف","ک","گ"]);
assert.deepStrictEqual(r.finalInto,  ["ب","پ","ت","ٹ","ث","ف","ک","گ"]);
assert.strictEqual(AshaarFonts.tatweelRulesOf("gulzar"), null);

console.log("fonts tests passed");
