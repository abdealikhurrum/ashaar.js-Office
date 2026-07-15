const assert = require("assert");
const { plan, facePlacements, drawParams, selectFaces, PAPER_SIZES } = require("../src/taskpane/impose-core");

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

// ── Mini-quire (cut into quarters & staple) ─────────────────────────────────
// front/back are flat 2x2 grids in physical row-major order:
// [top-left, top-right, bottom-left, bottom-right]. Each grid cell is a
// leaf: front = recto, back = verso (no fold — the leaf is a flat cut card).

{
  const p = plan({ pageCount: 8, scheme: "miniquire8", direction: "rtl" });
  assert.equal(p.paddedCount, 8);
  assert.deepEqual(p.blanks, []);
  assert.equal(p.sheets.length, 1, "8 pages = 1 sheet");
  // RTL reading order: top-right, top-left, bottom-right, bottom-left.
  assert.deepEqual(p.sheets[0].front, [3, 1, 7, 5], "[TL,TR,BL,BR] recto");
  assert.deepEqual(p.sheets[0].back, [4, 2, 8, 6], "[TL,TR,BL,BR] verso");
}

// LTR reading order: top-left, top-right, bottom-left, bottom-right.
{
  const p = plan({ pageCount: 8, scheme: "miniquire8", direction: "ltr" });
  assert.deepEqual(p.sheets[0].front, [1, 3, 5, 7]);
  assert.deepEqual(p.sheets[0].back, [2, 4, 6, 8]);
}

// 16 pages = 2 independent mini-quire sheets.
{
  const p = plan({ pageCount: 16, scheme: "miniquire8", direction: "rtl" });
  assert.equal(p.sheets.length, 2);
  assert.deepEqual(p.sheets[1].front, [11, 9, 15, 13]);
  assert.deepEqual(p.sheets[1].back, [12, 10, 16, 14]);
  const seen = p.sheets.flatMap((s) => s.front.concat(s.back)).sort((a, b) => a - b);
  assert.deepEqual(seen, Array.from({ length: 16 }, (_, i) => i + 1));
}

// Non-multiple-of-8 pads with blanks.
{
  const p = plan({ pageCount: 5, scheme: "miniquire8", direction: "rtl" });
  assert.equal(p.paddedCount, 8);
  assert.deepEqual(p.blanks, [6, 7, 8]);
}

// ── Zine fold (single sheet, single-sided, no staples) ──────────────────────
// front is a flat 2x4 grid, row-major, row 0 = top. back is null: the
// technique is inherently single-sided (verified against a working
// reference implementation; see the comment above GRID/ZINE_LTR).

{
  const p = plan({ pageCount: 8, scheme: "zinefold8", direction: "ltr" });
  assert.equal(p.sheets.length, 1);
  assert.strictEqual(p.sheets[0].back, null);
  assert.deepEqual(p.sheets[0].front, [
    { page: 8, rotate180: true }, { page: 1, rotate180: true },
    { page: 2, rotate180: true }, { page: 7, rotate180: true },
    { page: 6, rotate180: false }, { page: 3, rotate180: false },
    { page: 4, rotate180: false }, { page: 5, rotate180: false },
  ]);
}

// RTL mirrors the grid left-right, keeping each row's rotation intact.
{
  const p = plan({ pageCount: 8, scheme: "zinefold8", direction: "rtl" });
  assert.deepEqual(p.sheets[0].front, [
    { page: 7, rotate180: true }, { page: 2, rotate180: true },
    { page: 1, rotate180: true }, { page: 8, rotate180: true },
    { page: 5, rotate180: false }, { page: 4, rotate180: false },
    { page: 3, rotate180: false }, { page: 6, rotate180: false },
  ]);
}

// Every page 1-8 appears exactly once, and pages carry through a second
// independent sheet with a +8 offset.
{
  const p = plan({ pageCount: 16, scheme: "zinefold8", direction: "ltr" });
  assert.equal(p.sheets.length, 2);
  const seen = p.sheets.flatMap((s) => s.front.map((e) => e.page)).sort((a, b) => a - b);
  assert.deepEqual(seen, Array.from({ length: 16 }, (_, i) => i + 1));
}

// Non-multiple-of-8 pads with blanks.
{
  const p = plan({ pageCount: 3, scheme: "zinefold8", direction: "ltr" });
  assert.equal(p.paddedCount, 8);
  assert.deepEqual(p.blanks, [4, 5, 6, 7, 8]);
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

// ── facePlacements: mini-quire (2x2 grid per side, cut & staple) ────────────
// Slots are laid out row-major, row 0 = top; y counts up from the bottom of
// the sheet (pdf-lib convention), so top-row slots get the larger y.

{
  const faces = facePlacements({
    pageCount: 8, scheme: "miniquire8", direction: "rtl",
    srcWidth: 100, srcHeight: 150, flip: "short",
  });
  assert.equal(faces.length, 2, "1 sheet = 2 faces (front+back)");
  assert.equal(faces[0].width, 200, "2 cols x source width");
  assert.equal(faces[0].height, 300, "2 rows x source height");
  assert.deepEqual(faces[0].slots, [
    { page: 3, x: 0, y: 150, rotateDeg: 0 },
    { page: 1, x: 100, y: 150, rotateDeg: 0 },
    { page: 7, x: 0, y: 0, rotateDeg: 0 },
    { page: 5, x: 100, y: 0, rotateDeg: 0 },
  ], "front: TL,TR,BL,BR = 3,1,7,5");
  assert.deepEqual(faces[1].slots, [
    { page: 4, x: 0, y: 150, rotateDeg: 0 },
    { page: 2, x: 100, y: 150, rotateDeg: 0 },
    { page: 8, x: 0, y: 0, rotateDeg: 0 },
    { page: 6, x: 100, y: 0, rotateDeg: 0 },
  ], "back, short-edge flip: verso lands directly behind its recto");
}

// Long-edge duplex generalizes per row: columns swap within each row, whole
// back face rotates 180°, exactly like the 1-row schemes.
{
  const faces = facePlacements({
    pageCount: 8, scheme: "miniquire8", direction: "rtl",
    srcWidth: 100, srcHeight: 150, flip: "long",
  });
  assert.deepEqual(faces[1].slots, [
    { page: 2, x: 0, y: 150, rotateDeg: 180 },
    { page: 4, x: 100, y: 150, rotateDeg: 180 },
    { page: 6, x: 0, y: 0, rotateDeg: 180 },
    { page: 8, x: 100, y: 0, rotateDeg: 180 },
  ]);
}

// ── facePlacements: zine fold (2x4 grid, single-sided) ──────────────────────
// No back face is emitted at all — one face per sheet — and rotation comes
// from the scheme's own baked-in table, not the flip setting.

{
  const faces = facePlacements({
    pageCount: 8, scheme: "zinefold8", direction: "ltr",
    srcWidth: 100, srcHeight: 150, flip: "short",
  });
  assert.equal(faces.length, 1, "1 sheet = 1 face (single-sided, no back)");
  assert.equal(faces[0].width, 400, "4 cols x source width");
  assert.equal(faces[0].height, 300, "2 rows x source height");
  assert.deepEqual(faces[0].slots, [
    { page: 8, x: 0, y: 150, rotateDeg: 180 },
    { page: 1, x: 100, y: 150, rotateDeg: 180 },
    { page: 2, x: 200, y: 150, rotateDeg: 180 },
    { page: 7, x: 300, y: 150, rotateDeg: 180 },
    { page: 6, x: 0, y: 0, rotateDeg: 0 },
    { page: 3, x: 100, y: 0, rotateDeg: 0 },
    { page: 4, x: 200, y: 0, rotateDeg: 0 },
    { page: 5, x: 300, y: 0, rotateDeg: 0 },
  ]);
}

// ── facePlacements: paperSize scale-to-fit ───────────────────────────────────
// With opts.paperSize, the sheet is fixed at that size and source content is
// scaled (preserving aspect) and centered into each grid cell, instead of
// the sheet auto-sizing from the source page. Slots gain width/height.

{
  // saddle: 1x2 grid. Source is narrower than its cell → centered horizontally.
  const faces = facePlacements({
    pageCount: 8, scheme: "saddle", direction: "rtl",
    srcWidth: 90, srcHeight: 150, flip: "short",
    paperSize: { width: 400, height: 200 },
  });
  assert.equal(faces[0].width, 400, "sheet is fixed at the target paper width");
  assert.equal(faces[0].height, 200);
  assert.deepEqual(faces[0].slots, [
    { page: 1, x: 40, y: 0, rotateDeg: 0, width: 120, height: 200 },
    { page: 8, x: 240, y: 0, rotateDeg: 0, width: 120, height: 200 },
  ], "each 200x200 cell fits a 90x150 source at scale 4/3 (120x200), centered horizontally");
}

// miniquire8: 2x2 grid. Source is wider than its cell → centered vertically,
// and centering must respect each row's own vertical band.
{
  const faces = facePlacements({
    pageCount: 8, scheme: "miniquire8", direction: "ltr",
    srcWidth: 200, srcHeight: 90, flip: "short",
    paperSize: { width: 400, height: 400 },
  });
  assert.equal(faces[0].width, 400);
  assert.equal(faces[0].height, 400);
  assert.deepEqual(faces[0].slots, [
    { page: 1, x: 0, y: 255, rotateDeg: 0, width: 200, height: 90 },
    { page: 3, x: 200, y: 255, rotateDeg: 0, width: 200, height: 90 },
    { page: 5, x: 0, y: 55, rotateDeg: 0, width: 200, height: 90 },
    { page: 7, x: 200, y: 55, rotateDeg: 0, width: 200, height: 90 },
  ], "top row's band is y:200-400, bottom row's is y:0-200; content centered in each");
}

// A named PAPER_SIZES key resolves the same way as an explicit object.
{
  const named = facePlacements({
    pageCount: 8, scheme: "saddle", direction: "rtl",
    srcWidth: 90, srcHeight: 150, flip: "short", paperSize: "a4",
  });
  assert.equal(Math.round(named[0].width), Math.round(PAPER_SIZES.a4.width));
  assert.equal(Math.round(named[0].height), Math.round(PAPER_SIZES.a4.height));
}
assert.deepEqual(Object.keys(PAPER_SIZES).sort(), ["a4", "a5", "halfLetter", "letter"]);

// Without paperSize, slots keep their original shape (no width/height keys) —
// existing callers see byte-identical output to before this feature.
{
  const faces = facePlacements({
    pageCount: 8, scheme: "saddle", direction: "rtl",
    srcWidth: 100, srcHeight: 200, flip: "short",
  });
  assert.deepEqual(Object.keys(faces[0].slots[0]).sort(), ["page", "rotateDeg", "x", "y"]);
}

assert.throws(() => facePlacements({
  pageCount: 8, scheme: "saddle", direction: "rtl",
  srcWidth: 100, srcHeight: 200, flip: "short", paperSize: "tabloid",
}), /paperSize/);

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
