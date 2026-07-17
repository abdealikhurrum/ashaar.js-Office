"use strict";
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const CiteEngine = require("../src/taskpane/cite-engine");

const V = path.join(__dirname, "..", "src", "vendor");
const read = (p) => fs.readFileSync(path.join(V, p), "utf8");
const items = JSON.parse(fs.readFileSync(path.join(__dirname, "fixtures", "cite-sample.json"), "utf8"));
const locales = { "us": read("csl-locales/locales-en-US.xml"), "en-US": read("csl-locales/locales-en-US.xml"), "ar": read("csl-locales/locales-ar.xml") };

const chicago = CiteEngine.build({ styleXml: read("csl-styles/chicago-notes-bibliography.csl"), locales, items, lang: "en-US" });
const cite = chicago.cite(["en-book"]);
assert.match(cite, /Daftary/, "citation names the author");
assert.match(cite, /Fatimid Empire/, "citation names the title");

const bib = chicago.bibliography();
assert.match(bib, /Daftary/);
assert.match(bib, /Edinburgh/);

const apa = CiteEngine.build({ styleXml: read("csl-styles/apa.csl"), locales, items, lang: "en-US" });
assert.match(apa.cite(["en-article"]), /Halm/);
assert.match(apa.bibliography(), /2001/);
console.log("cite-engine (en) test passed");

const chicagoAr = CiteEngine.build({ styleXml: read("csl-styles/chicago-notes-bibliography.csl"), locales, items, lang: "ar" });
assert.strictEqual(chicagoAr.isRTL(), true, "Arabic lang flags RTL");
const arBib = chicagoAr.bibliography();
assert.match(arBib, /دعائم الإسلام/, "Arabic title renders");
// Arabic locale term for edition/etc. — assert an Arabic letter appears in generated chrome, not just the data.
assert.ok(/[؀-ۿ]/.test(arBib), "bibliography contains Arabic script");

// Locale-chrome proof: a Latin-only item rendered under the ar locale can
// only contain Arabic script if the engine emits Arabic locale terms.
const arChrome = CiteEngine.build({
  styleXml: read("csl-styles/chicago-notes-bibliography.csl"),
  locales, items: { "en-article": items["en-article"] }, lang: "ar"
});
const chromeBib = arChrome.bibliography();
assert.ok(/[؀-ۿ]/.test(chromeBib),
  "Latin-only item under ar locale must contain Arabic locale chrome (proves retrieveLocale('ar') wiring)");
console.log("cite-engine (ar) test passed");
