const assert = require("assert");
const PDFLib = require("pdf-lib");
const { imposePdf, makeTestSheet } = require("../src/taskpane/impose-pdf");

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

  console.log("impose-pdf tests passed");
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
