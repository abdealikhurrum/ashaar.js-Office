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

// spreadResidualSpacing — pixel-exact per-gap w:spacing twips instead of
// injected glyphs. 1px = 15 twips (96px/in vs 1440tw/in).
(function () {
  // Basic split: word gaps become their own single-space runs carrying twips.
  // residual 10px over 1 gap → 150tw on the gap.
  var one = AshaarResidual.spreadResidualSpacing(
    [{ text: "اب جد", csName: "X", asciiName: "B", sizePt: 12, color: "FF0000" }], 10, 16, 1.0);
  assert.strictEqual(one.runs.length, 3);
  assert.strictEqual(one.runs[0].text, "اب");
  assert.strictEqual(one.runs[1].text, " ");
  assert.strictEqual(one.runs[1].spacingTwips, 150);
  assert.strictEqual(one.runs[2].text, "جد");
  assert.strictEqual(one.gaps, 1);
  assert.ok(Math.abs(one.appliedPx - 10) < 0.01);
  // style props copied onto every split run
  one.runs.forEach(function (r) {
    assert.strictEqual(r.csName, "X");
    assert.strictEqual(r.asciiName, "B");
    assert.strictEqual(r.sizePt, 12);
    assert.strictEqual(r.color, "FF0000");
  });
  assert.strictEqual(one.runs[0].spacingTwips, undefined); // words never spaced

  // Even distribution with twip remainder: 1.1px over 3 gaps → 17tw → 6,6,5.
  var three = AshaarResidual.spreadResidualSpacing(
    [{ text: "ا ب ج د" }], 1.1, 16, 1.0);
  assert.strictEqual(three.gaps, 3);
  var tw = three.runs.filter(function (r) { return r.text === " "; })
    .map(function (r) { return r.spacingTwips; });
  assert.deepEqual(tw, [6, 6, 5]);

  // Cap binds: residual 100px, 1 gap, sizePx 16, capEm 0.28 → 4.48px → 67tw.
  var capped = AshaarResidual.spreadResidualSpacing([{ text: "ا ب" }], 100, 16, 0.28);
  assert.strictEqual(capped.runs[1].spacingTwips, 67);
  assert.ok(Math.abs(capped.appliedPx - 67 / 15) < 0.01);

  // A standalone " " run (inter-segment space) is a gap too.
  var seg = AshaarResidual.spreadResidualSpacing(
    [{ text: "اب", csName: "X" }, { text: " ", csName: "X" }, { text: "جد", csName: "Y" }], 2, 16, 1.0);
  assert.strictEqual(seg.gaps, 1);
  assert.strictEqual(seg.runs[1].spacingTwips, 30);

  // No positive residual → runs copied unchanged (no split, no twips).
  var none = AshaarResidual.spreadResidualSpacing([{ text: "اب جد" }], 0, 16, 1.0);
  assert.strictEqual(none.runs.length, 1);
  assert.strictEqual(none.runs[0].text, "اب جد");
  assert.strictEqual(none.appliedPx, 0);

  // No gaps → nothing to spread.
  var solid = AshaarResidual.spreadResidualSpacing([{ text: "ابجد" }], 10, 16, 1.0);
  assert.strictEqual(solid.appliedPx, 0);
  assert.strictEqual(solid.runs[0].text, "ابجد");

  // Input never mutated.
  var src = [{ text: "اب جد" }];
  AshaarResidual.spreadResidualSpacing(src, 10, 16, 1.0);
  assert.strictEqual(src.length, 1);
  assert.strictEqual(src[0].text, "اب جد");
})();

console.log("kashida-residual tests passed");
