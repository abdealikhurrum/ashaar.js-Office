"use strict";
var assert = require("assert");
var I = require("../src/taskpane/cite-import");

// sniffFormat — by filename extension
assert.strictEqual(I.sniffFormat("", "refs.json"), "csljson");
assert.strictEqual(I.sniffFormat("", "refs.bib"), "bibtex");
assert.strictEqual(I.sniffFormat("", "refs.bibtex"), "bibtex");
assert.strictEqual(I.sniffFormat("", "refs.ris"), "ris");
// sniffFormat — by content
assert.strictEqual(I.sniffFormat('[{"id":"a"}]'), "csljson");
assert.strictEqual(I.sniffFormat('  {"id":"a"}'), "csljson");
assert.strictEqual(I.sniffFormat("TY  - JOUR\nAU  - Halm, H.\nER  -"), "ris");
assert.strictEqual(I.sniffFormat("@book{key,\n title={X}\n}"), "bibtex");
assert.strictEqual(I.sniffFormat("just some prose"), null);
// extension wins over content
assert.strictEqual(I.sniffFormat("@book{k}", "x.json"), "csljson");

// parseImport — CSL-JSON array passes through, ids preserved
var arr = I.parseImport('[{"id":"k1","type":"book","title":"A"},{"id":"k2","type":"article-journal","title":"B"}]', "csljson");
assert.strictEqual(arr.length, 2);
assert.strictEqual(arr[0].id, "k1");
assert.strictEqual(arr[1].title, "B");
// single object -> wrapped in an array
var one = I.parseImport('{"id":"solo","type":"book","title":"S"}', "csljson");
assert.strictEqual(one.length, 1);
assert.strictEqual(one[0].id, "solo");
// missing id -> generated, non-empty, unique
var noId = I.parseImport('[{"type":"book","title":"X"},{"type":"book","title":"Y"}]', "csljson");
assert.ok(noId[0].id && noId[1].id && noId[0].id !== noId[1].id, "generated ids present and distinct");
// format auto-sniffed when not passed
assert.strictEqual(I.parseImport('[{"id":"z","title":"Z"}]').length, 1);
// invalid JSON -> throws
assert.throws(function () { I.parseImport("{not json", "csljson"); });
// BibTeX / RIS -> explicit "coming soon" (MVP), not a silent empty
assert.throws(function () { I.parseImport("@book{k}", "bibtex"); }, /BibTeX|RIS|CSL JSON/i);
assert.throws(function () { I.parseImport("TY  - JOUR\nER  -", "ris"); }, /BibTeX|RIS|CSL JSON/i);
// unrecognized -> throws
assert.throws(function () { I.parseImport("hello", null); }, /Unrecognized|CSL JSON/i);

console.log("cite-import.test.js passed");
