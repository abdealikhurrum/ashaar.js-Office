const assert = require("assert");
const PDFLib = require("pdf-lib");
const { imposePdf, makeTestSheet, makeTestBooklet } = require("../src/taskpane/impose-pdf");

async function makeSourcePdf(pageCount, width, height) {
  const doc = await PDFLib.PDFDocument.create();
  for (let i = 0; i < pageCount; i++) {
    const page = doc.addPage([width, height]);
    page.drawText(String(i + 1), { x: width / 2, y: height / 2, size: 24 });
  }
  return doc.save();
}

(async () => {
  // 5-page portrait source → saddle RTL: pads to 8, 2 sheets, 4 printed faces.
  const src = await makeSourcePdf(5, 100, 200);
  const out = await imposePdf(src, {
    scheme: "saddle",
    direction: "rtl",
    flip: "short",
    pdfLib: PDFLib,
  });
  assert.ok(out instanceof Uint8Array, "returns bytes");

  const doc = await PDFLib.PDFDocument.load(out);
  assert.equal(doc.getPageCount(), 4, "4 faces");
  const first = doc.getPage(0);
  assert.equal(first.getWidth(), 200, "landscape sheet, 2x source width");
  assert.equal(first.getHeight(), 200);

  // Long-edge flip variant also round-trips.
  const outLong = await imposePdf(src, {
    scheme: "quire4",
    direction: "rtl",
    flip: "long",
    pdfLib: PDFLib,
  });
  const docLong = await PDFLib.PDFDocument.load(outLong);
  assert.equal(docLong.getPageCount(), 4, "quire4: 2 sheets = 4 faces");

  // Manual duplex: fronts-only / backs-only exports (5 pages → 2 sheets).
  const fronts = await imposePdf(src, {
    scheme: "saddle", direction: "rtl", flip: "short",
    faces: "fronts", pdfLib: PDFLib,
  });
  assert.equal((await PDFLib.PDFDocument.load(fronts)).getPageCount(), 2, "2 front faces");
  const backs = await imposePdf(src, {
    scheme: "saddle", direction: "rtl", flip: "short",
    faces: "backs", reverse: true, pdfLib: PDFLib,
  });
  assert.equal((await PDFLib.PDFDocument.load(backs)).getPageCount(), 2, "2 back faces");

  // Printer test sheet: one duplex sheet (2 faces) at the given sheet size.
  // The back carries both an upright and a 180°-rotated caption so whichever
  // reads upright after printing names the printer's duplex flip setting.
  const test = await makeTestSheet({ pdfLib: PDFLib, width: 842, height: 595 });
  const testDoc = await PDFLib.PDFDocument.load(test);
  assert.equal(testDoc.getPageCount(), 2, "test sheet = front + back");
  assert.equal(testDoc.getPage(0).getWidth(), 842);
  assert.equal(testDoc.getPage(1).getHeight(), 595);

  // Multi-signature test booklet: reader-order source pages carrying fold
  // instructions; the caller imposes it like any real document. sheets=3
  // → 12 pages. Defaults to A5 portrait.
  for (const scheme of ["saddle", "quire4"]) {
    for (const direction of ["rtl", "ltr"]) {
      const booklet = await makeTestBooklet({
        pdfLib: PDFLib, scheme: scheme, direction: direction, sheets: 3,
      });
      const doc = await PDFLib.PDFDocument.load(booklet);
      assert.equal(doc.getPageCount(), 12, scheme + "/" + direction + ": 3 sheets = 12 pages");
      assert.equal(Math.round(doc.getPage(0).getWidth()), 420, "A5 portrait default");
    }
  }
  // 8-pages-per-sheet schemes: sheets=3 → 24 pages.
  for (const scheme of ["miniquire8", "zinefold8"]) {
    for (const direction of ["rtl", "ltr"]) {
      const booklet = await makeTestBooklet({
        pdfLib: PDFLib, scheme: scheme, direction: direction, sheets: 3,
      });
      const doc = await PDFLib.PDFDocument.load(booklet);
      assert.equal(doc.getPageCount(), 24, scheme + "/" + direction + ": 3 sheets = 24 pages");
    }
  }
  await assert.rejects(
    makeTestBooklet({ pdfLib: PDFLib, scheme: "spiral", direction: "rtl", sheets: 2 }),
    /scheme/
  );

  // miniquire8: duplex sheets carry dashed cut lines at both center creases.
  {
    const miniSrc = await makeSourcePdf(8, 100, 150);
    const out = await imposePdf(miniSrc, {
      scheme: "miniquire8", direction: "rtl", flip: "short", pdfLib: PDFLib,
    });
    const doc = await PDFLib.PDFDocument.load(out);
    assert.equal(doc.getPageCount(), 2, "1 sheet = front+back faces");
    assert.equal(doc.getPage(0).getWidth(), 200, "2 cols x source width");
    assert.equal(doc.getPage(0).getHeight(), 300, "2 rows x source height");
  }

  // zinefold8: single-sided — one output page per sheet, no back face,
  // and manual duplex options ("faces") are ignored (nothing to split).
  {
    const zineSrc = await makeSourcePdf(8, 100, 150);
    const out = await imposePdf(zineSrc, {
      scheme: "zinefold8", direction: "ltr", pdfLib: PDFLib,
    });
    const doc = await PDFLib.PDFDocument.load(out);
    assert.equal(doc.getPageCount(), 1, "single-sided: 1 sheet = 1 face");
    assert.equal(doc.getPage(0).getWidth(), 400, "4 cols x source width");

    const outFronts = await imposePdf(zineSrc, {
      scheme: "zinefold8", direction: "ltr", faces: "fronts", pdfLib: PDFLib,
    });
    assert.equal((await PDFLib.PDFDocument.load(outFronts)).getPageCount(), 1,
      '"fronts" is ignored for a single-sided scheme');
  }

  // paperSize: the printed sheet matches the chosen paper regardless of the
  // source PDF's own page size — this is what fixes an oversized zinefold8
  // sheet (4 columns x an unscaled source page could dwarf any home printer).
  for (const scheme of ["saddle", "miniquire8", "zinefold8"]) {
    const src = await makeSourcePdf(scheme === "zinefold8" ? 8 : 8, 300, 500);
    const out = await imposePdf(src, {
      scheme: scheme, direction: "rtl", flip: "short", paperSize: "a4", pdfLib: PDFLib,
    });
    const doc = await PDFLib.PDFDocument.load(out);
    const sheet = doc.getPage(0);
    assert.equal(Math.round(sheet.getWidth()), 842, scheme + ": sheet width matches A4 landscape, not the source page");
    assert.equal(Math.round(sheet.getHeight()), 595, scheme + ": sheet height matches A4 landscape");
  }

  // An unknown paperSize name fails clearly rather than silently ignoring it.
  await assert.rejects(
    imposePdf(await makeSourcePdf(4, 300, 500), {
      scheme: "saddle", direction: "rtl", flip: "short", paperSize: "tabloid", pdfLib: PDFLib,
    }),
    /paperSize/
  );

  // makeTestSheet also accepts a named paperSize.
  {
    const test = await makeTestSheet({ pdfLib: PDFLib, paperSize: "letter" });
    const doc = await PDFLib.PDFDocument.load(test);
    assert.equal(Math.round(doc.getPage(0).getWidth()), 792, "letter landscape width");
    assert.equal(Math.round(doc.getPage(0).getHeight()), 612, "letter landscape height");
  }

  // End-to-end: the zinefold8 test booklet, imposed with a chosen paper size,
  // comes out at exactly that size (the bug this feature fixes).
  {
    const source = await makeTestBooklet({ pdfLib: PDFLib, scheme: "zinefold8", direction: "rtl", sheets: 1 });
    const out = await imposePdf(source, { scheme: "zinefold8", direction: "rtl", paperSize: "halfLetter", pdfLib: PDFLib });
    const doc = await PDFLib.PDFDocument.load(out);
    assert.equal(Math.round(doc.getPage(0).getWidth()), 612, "half-letter landscape width");
    assert.equal(Math.round(doc.getPage(0).getHeight()), 396, "half-letter landscape height");
  }

  console.log("impose-pdf tests passed");
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
