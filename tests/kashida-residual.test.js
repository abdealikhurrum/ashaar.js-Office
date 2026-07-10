"use strict";
var assert = require("assert");
var AshaarResidual = require("../src/taskpane/kashida-residual");

// zero / invalid inputs → 0
assert.strictEqual(AshaarResidual.capMicroSpaces(0, 5, 2, 16), 0);      // no residual
assert.strictEqual(AshaarResidual.capMicroSpaces(100, 0, 2, 16), 0);    // no gaps
assert.strictEqual(AshaarResidual.capMicroSpaces(100, 5, 0, 16), 0);    // no glyph width

// under the cap: fill the residual exactly (rounded to glyph count)
// cap = 0.28*16*4 = 17.92px; residual 10 < cap → round(10/2) = 5
assert.strictEqual(AshaarResidual.capMicroSpaces(10, 4, 2, 16), 5);

// cap binds: residual 100 > cap 17.92 → round(17.92/2) = round(8.96) = 9
assert.strictEqual(AshaarResidual.capMicroSpaces(100, 4, 2, 16), 9);

// explicit capEm honored: cap = 0.5*16*4 = 32px → round(min(100,32)/2)=16
assert.strictEqual(AshaarResidual.capMicroSpaces(100, 4, 2, 16, 0.5), 16);

console.log("kashida-residual tests passed");
