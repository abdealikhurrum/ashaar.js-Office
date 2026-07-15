const assert = require("assert");
const PDFLib = require("pdf-lib");
const { imposePdf } = require("../src/taskpane/impose-pdf");

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

  console.log("impose-pdf tests passed");
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
