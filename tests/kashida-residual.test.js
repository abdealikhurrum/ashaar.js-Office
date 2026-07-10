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

// injectSpaceRuns — hair-spaces appended to the " " runs only
(function () {
  var H = AshaarResidual.HAIR_SPACE;
  var runs = [{ text: "كہہ", swap: true }, { text: " ", swap: false },
              { text: "رہے", swap: false }, { text: " ", swap: false },
              { text: "تھے", swap: false }];
  var out = AshaarResidual.injectSpaceRuns(runs, 3);      // 2 space runs, n=3
  assert.strictEqual(out.length, 5);
  assert.strictEqual(out[0].text, "كہہ");                 // words untouched
  assert.strictEqual(out[0].swap, true);                  // swap flag preserved
  assert.strictEqual(out[1].text, " " + H + H);           // first gap: remainder -> 2
  assert.strictEqual(out[3].text, " " + H);               // second gap: 1
  assert.strictEqual(runs[1].text, " ");                  // original NOT mutated
  assert.strictEqual(H, " ");                        // HAIR_SPACE is U+200A, not ASCII space

  // n <= 0 -> text unchanged (structural copy, new array)
  var out0 = AshaarResidual.injectSpaceRuns(runs, 0);
  assert.strictEqual(out0[1].text, " ");
  assert.notStrictEqual(out0, runs);

  // no space runs -> unchanged
  var solid = [{ text: "ابجد", swap: false }];
  assert.strictEqual(AshaarResidual.injectSpaceRuns(solid, 5)[0].text, "ابجد");

  // custom spaceChar honored (distinct from HAIR_SPACE so it truly discriminates)
  var TH = " "; // thin space
  var out2 = AshaarResidual.injectSpaceRuns(runs, 2, TH);
  assert.strictEqual(out2[1].text, " " + TH);
  assert.strictEqual(out2[3].text, " " + TH);
})();

console.log("kashida-residual tests passed");
