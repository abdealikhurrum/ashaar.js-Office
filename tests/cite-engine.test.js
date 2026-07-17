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

// --- Task 4: CSL-M multilingual variant rendering ---
// The fixture carries an Arabic title + author that ALSO hold a Latin
// transliteration variant (multi._keys.title["ar-Latn"] / author.multi._key["ar-Latn"]).
// With langPrefs selecting the "translit" slot, the transliteration must render.
const multi = JSON.parse(fs.readFileSync(path.join(__dirname, "fixtures", "cite-multi.json"), "utf8"));
const mlEngine = CiteEngine.build({
  styleXml: read("csl-styles/chicago-notes-bibliography.csl"),
  locales, items: multi, lang: "en-US",
  langPrefs: {
    persons: ["translit"],
    titles: ["translit"],
    translit: ["ar-Latn"]
  }
});
const mlBib = mlEngine.bibliography();
// "Nuʿm" is a fragment of the author transliteration "al-Nuʿmān"; it appears
// nowhere in the Arabic-script primary data nor in en-US locale chrome.
assert.match(mlBib, /Nuʿm/, "author transliteration variant rendered");
// "Islām" is a fragment of the title transliteration "Daʿāʾim al-Islām".
assert.match(mlBib, /Islām/, "title transliteration variant rendered");
console.log("cite-engine (multilingual) test passed");

// --- Task 6: repo-owned Fatemi-aware styles must be inert until SP-4 ---
// The fixture items carry no `genre`, so the Fatemi styles' inert <choose>
// branch never fires and their bibliography output must equal the stock parent.
const parent = CiteEngine.build({ styleXml: read("csl-styles/chicago-notes-bibliography.csl"), locales, items, lang: "en-US" });
const fatemi = CiteEngine.build({ styleXml: read("csl-styles/chicago-notes-fatemi.csl"), locales, items, lang: "en-US" });
assert.strictEqual(fatemi.bibliography(), parent.bibliography(), "Fatemi style equals parent when genre absent");
console.log("cite-engine (fatemi parity) test passed");

const apaParent = CiteEngine.build({ styleXml: read("csl-styles/apa.csl"), locales, items, lang: "en-US" });
const apaFatemi = CiteEngine.build({ styleXml: read("csl-styles/apa-fatemi.csl"), locales, items, lang: "en-US" });
assert.strictEqual(apaFatemi.bibliography(), apaParent.bibliography(), "APA Fatemi style equals parent when genre absent");
console.log("cite-engine (apa fatemi parity) test passed");

// --- locators (SP-A Task 1) ---
// Reuse the existing engines: `chicago` (en-US) and `chicagoAr` (ar), and a
// real fixture id from `items`.
const enEngine = chicago;
const arEngine = chicagoAr;
const someId = "en-book";
const withPage = enEngine.cite([{ id: someId, locator: "42", label: "page" }]);
const noLoc = enEngine.cite([{ id: someId }]);
assert.ok(withPage.indexOf("42") !== -1, "page locator value appears in the citation");
assert.ok(noLoc.indexOf("42") === -1, "no locator ⇒ value absent (locator plumbing is real)");
// label-term plumbing: a chapter locator renders the localized 'chap.' term (en)
const withChap = enEngine.cite([{ id: someId, locator: "3", label: "chapter" }]);
assert.ok(/chap/i.test(withChap), "chapter label renders the 'chap.' term (en)");
// bare-string back-compat still works
assert.ok(typeof enEngine.cite([someId]) === "string", "cite() still accepts bare id strings");
// locale-independent plumbing: value also present under the ar engine
const arWithPage = arEngine.cite([{ id: someId, locator: "42", label: "page" }]);
assert.ok(arWithPage.indexOf("42") !== -1, "locator value appears under the ar locale too");
console.log("cite locators test passed");
