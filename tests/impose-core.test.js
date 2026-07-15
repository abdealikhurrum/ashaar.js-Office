const assert = require("assert");
const { plan, facePlacements, drawParams, selectFaces } = require("../src/taskpane/impose-core");

// plan({pageCount, scheme, direction}) → {paddedCount, blanks, sheets}
// sheets[i] = {front: [left, right], back: [left, right]} with 1-based reader
// page numbers into the padded sequence. Numbers > pageCount are padding blanks.

// ── Saddle (single signature), RTL — the Awraq default ─────────────────────

// 8 pages, RTL: cover (page 1) sits on the LEFT half of the outer front face
// (spine folds on the right, booklet opens right-to-left).
{
  const p = plan({ pageCount: 8, scheme: "saddle", direction: "rtl" });
  assert.equal(p.paddedCount, 8);
  assert.deepEqual(p.blanks, []);
  assert.equal(p.sheets.length, 2, "8 pages = 2 sheets");
  assert.deepEqual(p.sheets[0].front, [1, 8], "outer front: 1 left, 8 right");
  assert.deepEqual(p.sheets[0].back, [7, 2], "outer back: 7 left, 2 right");
  assert.deepEqual(p.sheets[1].front, [3, 6]);
  assert.deepEqual(p.sheets[1].back, [5, 4]);
}

// Saddle LTR is the mirror image (standard Western booklet).
{
  const p = plan({ pageCount: 8, scheme: "saddle", direction: "ltr" });
  assert.deepEqual(p.sheets[0].front, [8, 1], "LTR outer front: 8 left, 1 right");
  assert.deepEqual(p.sheets[0].back, [2, 7]);
  assert.deepEqual(p.sheets[1].front, [6, 3]);
  assert.deepEqual(p.sheets[1].back, [4, 5]);
}

// Non-multiple-of-4 page counts pad with blanks at the end of reading order.
{
  const p = plan({ pageCount: 5, scheme: "saddle", direction: "rtl" });
  assert.equal(p.paddedCount, 8);
  assert.deepEqual(p.blanks, [6, 7, 8], "pages 6-8 are padding");
  assert.deepEqual(p.sheets[0].front, [1, 8]);
}

// Minimum booklet: 1 page pads to a single sheet.
{
  const p = plan({ pageCount: 1, scheme: "saddle", direction: "rtl" });
  assert.equal(p.paddedCount, 4);
  assert.equal(p.sheets.length, 1);
  assert.deepEqual(p.sheets[0].front, [1, 4]);
  assert.deepEqual(p.sheets[0].back, [3, 2]);
}

// Every padded page appears exactly once (16-side booklet, the classic Awraq).
{
  const p = plan({ pageCount: 16, scheme: "saddle", direction: "rtl" });
  const seen = p.sheets.flatMap((s) => s.front.concat(s.back)).sort((a, b) => a - b);
  assert.deepEqual(seen, Array.from({ length: 16 }, (_, i) => i + 1));
  assert.equal(p.sheets.length, 4, "16 sides = 4 sheets");
}

// ── Per-sheet 4-side quire (paper-saving small booklet) ─────────────────────

// Each sheet folds independently and holds 4 consecutive pages.
{
  const p = plan({ pageCount: 8, scheme: "quire4", direction: "rtl" });
  assert.equal(p.sheets.length, 2);
  assert.deepEqual(p.sheets[0].front, [1, 4]);
  assert.deepEqual(p.sheets[0].back, [3, 2]);
  assert.deepEqual(p.sheets[1].front, [5, 8], "second quire is pages 5-8");
  assert.deepEqual(p.sheets[1].back, [7, 6]);
}

// LTR quire mirror.
{
  const p = plan({ pageCount: 4, scheme: "quire4", direction: "ltr" });
  assert.deepEqual(p.sheets[0].front, [4, 1]);
  assert.deepEqual(p.sheets[0].back, [2, 3]);
}

// A 4-page booklet is identical under both schemes.
{
  const a = plan({ pageCount: 4, scheme: "saddle", direction: "rtl" });
  const b = plan({ pageCount: 4, scheme: "quire4", direction: "rtl" });
  assert.deepEqual(a.sheets, b.sheets);
}

// ── facePlacements: print-ready geometry per face ───────────────────────────
// Faces come in print order (sheet0 front, sheet0 back, sheet1 front, ...).
// Each face: {width, height, slots:[{page, x, y, rotateDeg}]}, x/y = bottom-left
// of the slot. Sheet is auto-sized: two source pages side by side.
// flip:'short' assumes the printer flips around the vertical edge of the
// landscape sheet (no rotation needed); flip:'long' pre-rotates back faces 180°.

{
  const faces = facePlacements({
    pageCount: 8, scheme: "saddle", direction: "rtl",
    srcWidth: 100, srcHeight: 200, flip: "short",
  });
  assert.equal(faces.length, 4, "2 sheets = 4 faces");
  assert.equal(faces[0].width, 200, "sheet width = 2 x source width");
  assert.equal(faces[0].height, 200);
  assert.deepEqual(faces[0].slots, [
    { page: 1, x: 0, y: 0, rotateDeg: 0 },
    { page: 8, x: 100, y: 0, rotateDeg: 0 },
  ], "outer front: 1 left, 8 right");
  assert.deepEqual(faces[1].slots, [
    { page: 7, x: 0, y: 0, rotateDeg: 0 },
    { page: 2, x: 100, y: 0, rotateDeg: 0 },
  ], "outer back, short-edge flip: no rotation");
  assert.deepEqual(faces[2].slots.map((s) => s.page), [3, 6]);
}

// Long-edge duplex: back faces are pre-rotated 180° (pages swap halves and
// flip upside down); front faces are untouched.
{
  const faces = facePlacements({
    pageCount: 8, scheme: "saddle", direction: "rtl",
    srcWidth: 100, srcHeight: 200, flip: "long",
  });
  assert.deepEqual(faces[0].slots.map((s) => s.rotateDeg), [0, 0], "front untouched");
  assert.deepEqual(faces[1].slots, [
    { page: 2, x: 0, y: 0, rotateDeg: 180 },
    { page: 7, x: 100, y: 0, rotateDeg: 180 },
  ], "back face rotated as a whole: [7,2] becomes [2,7] upside down");
}

// Padding blanks surface as page:null so the assembly layer leaves them empty.
{
  const faces = facePlacements({
    pageCount: 5, scheme: "saddle", direction: "rtl",
    srcWidth: 100, srcHeight: 200, flip: "short",
  });
  const pages = faces.flatMap((f) => f.slots.map((s) => s.page));
  assert.deepEqual(pages, [1, null, null, 2, 3, null, 5, 4], "6,7,8 are blank");
}

// ── drawParams: pdf-lib draw origin for a slot ──────────────────────────────
// pdf-lib's drawPage rotates around the given (x,y); a 180° slot must be drawn
// from the opposite corner so the content lands inside the slot rect.
{
  assert.deepEqual(
    drawParams({ page: 3, x: 100, y: 0, rotateDeg: 0 }, 100, 200),
    { x: 100, y: 0, rotateDeg: 0 }
  );
  assert.deepEqual(
    drawParams({ page: 3, x: 100, y: 0, rotateDeg: 180 }, 100, 200),
    { x: 200, y: 200, rotateDeg: 180 },
    "180° draws from opposite corner (x+w, y+h)"
  );
}

// ── selectFaces: pick fronts/backs for manual (non-duplex) printing ─────────
// Faces alternate front, back, front, back. 'fronts'/'backs' filter them;
// reverse=true flips the order of the selected faces for printers that stack
// output in reverse.
{
  const faces = ["f1", "b1", "f2", "b2", "f3", "b3"];
  assert.deepEqual(selectFaces(faces, "all"), faces);
  assert.deepEqual(selectFaces(faces, "fronts"), ["f1", "f2", "f3"]);
  assert.deepEqual(selectFaces(faces, "backs"), ["b1", "b2", "b3"]);
  assert.deepEqual(selectFaces(faces, "backs", true), ["b3", "b2", "b1"]);
  assert.deepEqual(selectFaces(faces, "fronts", true), ["f3", "f2", "f1"]);
  assert.throws(() => selectFaces(faces, "sides"), /which/);
}

// ── Validation ──────────────────────────────────────────────────────────────

assert.throws(() => plan({ pageCount: 0, scheme: "saddle", direction: "rtl" }), /pageCount/);
assert.throws(() => plan({ pageCount: 8, scheme: "spiral", direction: "rtl" }), /scheme/);
assert.throws(() => plan({ pageCount: 8, scheme: "saddle", direction: "up" }), /direction/);

console.log("impose-core tests passed");
