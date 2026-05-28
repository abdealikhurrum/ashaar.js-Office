#!/usr/bin/env node
// Generates test-documents/marsiya-test.docx with the marsiya poem pre-inserted as Word tables.
// Run: node scripts/make-test-doc.mjs
import { createRequire } from "module";
import { fileURLToPath } from "url";
import path from "path";
import fs from "fs";

const require = createRequire(import.meta.url);
const AdmZip = require("adm-zip");
const Ashaar = require("../src/vendor/ashaar.js");
const AshaarWord = require("../src/taskpane/word-html.js");

const POEM = [
  "اهلِ بيت تھے عجب \\ اٗفقِ دعوت كے شٗہٗب \\ هو گئے قربان سب \\",
  "خالي لشكر هو گيا \\",
  "هائے كربلاء والو \\ هائے كربلاء والو",
  "---",
  "شاه كا اكبر جواں \\ رَنْ ميں اٰيا پہلواں \\ شير اٰيا هے يہاں \\",
  "كہہ رهے تھے اشقياء \\",
  "هائے كربلاء والو \\ هائے كربلاء والو"
].join("\n");

const opts = { justifyMode: "none", gapWidth: 1 };
const TEXT_WIDTH_TWIPS = 9360; // 6.5 inch at 1440 twips/inch

const body = AshaarWord.renderForWordOoxml(POEM, opts, Ashaar, TEXT_WIDTH_TWIPS);

const wns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const documentXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:document ' + wns + '><w:body>' + body + '<w:sectPr/></w:body></w:document>';

const contentTypes =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/word/document.xml"' +
  ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
  '</Types>';

const relsRoot =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1"' +
  ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"' +
  ' Target="word/document.xml"/>' +
  '</Relationships>';

const relsWord =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>';

const zip = new AdmZip();
zip.addFile("[Content_Types].xml", Buffer.from(contentTypes, "utf8"));
zip.addFile("_rels/.rels",          Buffer.from(relsRoot,     "utf8"));
zip.addFile("word/document.xml",    Buffer.from(documentXml,  "utf8"));
zip.addFile("word/_rels/document.xml.rels", Buffer.from(relsWord, "utf8"));

const dir  = path.join(path.dirname(fileURLToPath(import.meta.url)), "..", "test-documents");
const out  = path.join(dir, "marsiya-test.docx");
fs.mkdirSync(dir, { recursive: true });
zip.writeZip(out);
console.log("Created", out);
