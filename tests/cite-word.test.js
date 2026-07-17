"use strict";
const assert = require("assert");
const CiteWord = require("../src/taskpane/cite-word");

// sanitize: keep italics, drop class/style attrs, unwrap unknown tags
const dirty = '<div class="csl-entry"><span style="font-variant:small-caps">Daftary</span>, <i>The Fatimid Empire</i><a href="x">link</a></div>';
const clean = CiteWord.sanitize(dirty);
assert.ok(!/class=/.test(clean) && !/style=/.test(clean), "attributes stripped");
assert.ok(!/<div/.test(clean) && !/<a /.test(clean), "unknown tags unwrapped");
assert.match(clean, /<i>The Fatimid Empire<\/i>/, "italics preserved");
assert.match(clean, /Daftary/, "text preserved");

// note payload: RTL direction
const note = CiteWord.buildNotePayload({ html: dirty, rtl: true });
assert.strictEqual(note.direction, "Rtl");
assert.ok(!/style=/.test(note.html), "note html is sanitized");

// bibliography payload: default tag
const bib = CiteWord.buildBibliographyPayload({ html: "<div>x</div>", rtl: false });
assert.strictEqual(bib.tag, "AshaarBibliography");
assert.strictEqual(bib.direction, "Ltr");

// citation tag round-trip
const tag = CiteWord.citationTag(["en-book", "en-article"], "chicago-notes-bibliography");
assert.deepStrictEqual(CiteWord.parseCitationTag(tag), { style: "chicago-notes-bibliography", itemKeys: ["en-book", "en-article"] });
assert.strictEqual(CiteWord.parseCitationTag("Nope"), null);
console.log("cite-word test passed");
