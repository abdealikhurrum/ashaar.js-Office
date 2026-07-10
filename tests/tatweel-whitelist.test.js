"use strict";
const assert = require("assert");
const AshaarTatweel = require("../src/taskpane/tatweel-whitelist");
const AshaarFonts = require("../src/taskpane/fonts");

const rules = AshaarFonts.tatweelRulesOf("mehr");

// "کتاب" = ک-ت-ا-ب. Pairs: کت (next ت ✓ allowed), تا (next ا ✗ block), اب (next ب ✓ allowed).
const t = AshaarTatweel.buildPriorityTable("کتاب", rules);
assert.ok(!t["کت"] || !t["کت"].blocked, "into ت allowed");
assert.ok(t["تا"] && t["تا"].blocked === true, "into ا blocked");
assert.ok(!t["اب"] || !t["اب"].blocked, "into ب allowed");

// join into ر is never whitelisted → blocked
const t2 = AshaarTatweel.buildPriorityTable("در", rules);
assert.ok(t2["در"] && t2["در"].blocked === true, "into ر blocked");

console.log("tatweel-whitelist tests passed");
