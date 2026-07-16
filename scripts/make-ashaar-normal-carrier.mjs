#!/usr/bin/env node
// Builds the "Ashaar Normal carrier" — a tiny .docx whose styles.xml defines an
// "Ashaar Normal" paragraph style WITH `<w:bidi/>` + `<w:rtl/>` (true RTL
// reading order), which Office.js's object model cannot set on a style. The
// add-in imports this carrier via document.insertFileFromBase64({importStyles})
// to merge that bidi-carrying style into the open document, then sets fonts/
// size on it with the normal Style API and deletes the carrier's sentinel
// paragraph. See src/taskpane/styles-pane.js (importAshaarNormalBidi).
//
// This emits a committed UMD module (src/taskpane/ashaar-normal-carrier.js) with
// the base64 + sentinel, so the browser needs no zip library. Regenerate with:
//   node scripts/make-ashaar-normal-carrier.mjs   (or: npm run make-normal-carrier)

import { createRequire } from "module";
import { fileURLToPath } from "url";
import path from "path";
import fs from "fs";

const require = createRequire(import.meta.url);
const AdmZip = require("adm-zip");

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const OUT_JS = path.join(__dirname, "..", "src", "taskpane", "ashaar-normal-carrier.js");

// Unique, ASCII-only marker so the add-in can find and delete exactly the
// paragraph this carrier inserts (via body.search), never user content.
const SENTINEL = "ZZASHAARNORMALCARRIERZZ";

var stylesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  // The whole point: Ashaar Normal carries paragraph bidi + run rtl, which the
  // Office.js Style API can't set. basedOn Normal keeps it a body style.
  '<w:style w:type="paragraph" w:customStyle="1" w:styleId="AshaarNormal">' +
  '<w:name w:val="Ashaar Normal"/><w:basedOn w:val="Normal"/><w:unhideWhenUsed/>' +
  '<w:pPr><w:bidi/><w:jc w:val="right"/></w:pPr><w:rPr><w:rtl/></w:rPr>' +
  '</w:style>' +
  // Built-in footnote styles redefined RTL, so importStyles merges right-to-left
  // reading order into the target's footnote text + reference marker (the object
  // model can't set their bidi either). styleId/name match Word's built-ins so
  // the merge lands on them. NOTE: the footnote SEPARATOR line is a special
  // footnote in footnotes.xml, NOT a style — importStyles can't touch it.
  '<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="Footnote Text"/>' +
  '<w:basedOn w:val="Normal"/><w:pPr><w:bidi/><w:jc w:val="right"/></w:pPr><w:rPr><w:rtl/></w:rPr>' +
  '</w:style>' +
  '<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="Footnote Reference"/>' +
  '<w:rPr><w:rtl/><w:vertAlign w:val="superscript"/></w:rPr>' +
  '</w:style>' +
  '</w:styles>';

var documentXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
  '<w:p><w:r><w:t xml:space="preserve">' + SENTINEL + '</w:t></w:r></w:p>' +
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>' +
  '</w:body></w:document>';

var contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
  '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
  '</Types>';

var rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
  '</Relationships>';

var docRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
  '</Relationships>';

var zip = new AdmZip();
zip.addFile("[Content_Types].xml", Buffer.from(contentTypes, "utf8"));
zip.addFile("_rels/.rels", Buffer.from(rels, "utf8"));
zip.addFile("word/document.xml", Buffer.from(documentXml, "utf8"));
zip.addFile("word/styles.xml", Buffer.from(stylesXml, "utf8"));
zip.addFile("word/_rels/document.xml.rels", Buffer.from(docRels, "utf8"));
var b64 = zip.toBuffer().toString("base64");

var js = "/* GENERATED by scripts/make-ashaar-normal-carrier.mjs — do not edit by hand.\n" +
  " * A base64 .docx carrying an \"Ashaar Normal\" paragraph style with true RTL\n" +
  " * (w:bidi + w:rtl) for import via insertFileFromBase64({importStyles:true}).\n" +
  " * Regenerate: npm run make-normal-carrier */\n" +
  "(function (root, factory) {\n" +
  "  if (typeof module !== \"undefined\" && module.exports) module.exports = factory();\n" +
  "  else root.AshaarNormalCarrier = factory();\n" +
  "}(typeof globalThis !== \"undefined\" ? globalThis : this, function () {\n" +
  "  \"use strict\";\n" +
  "  return { SENTINEL: " + JSON.stringify(SENTINEL) + ", base64: " + JSON.stringify(b64) + " };\n" +
  "}));\n";

fs.writeFileSync(OUT_JS, js, "utf8");
console.log("Wrote " + OUT_JS + " (" + js.length + " bytes; base64 " + b64.length + " chars)");
