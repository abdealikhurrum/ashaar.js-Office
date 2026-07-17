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
// --- bidi run wrapping: give Word directional guidance for neutral punctuation ---
// citeproc emits plain mixed-direction text with no dir info; in an LTR paragraph
// Word's bidi algorithm mis-places the neutral (),.-, around Arabic runs. wrapRtlRuns
// wraps each maximal Arabic run in <span dir="rtl"> so the punctuation resolves RTL.

// Pure English: untouched (no Arabic present at all)
const enOnly = 'Farhad Daftary, <i>The Fatimid Empire</i> (Edinburgh University Press, 2018).';
assert.strictEqual(CiteWord.wrapRtlRuns(enOnly), enOnly, "pure-LTR citation is unchanged");

// Pure Arabic citation: whole thing wrapped once; <i> preserved inside; no interior span
const arOnly = 'القاضي النعمان, <i>دعائم الإسلام</i> (دار المعارف, 1951).';
const arWrapped = CiteWord.wrapRtlRuns(arOnly);
assert.strictEqual(arWrapped.indexOf('<span dir="rtl">'), 0, "arabic citation opens with an rtl span");
assert.ok(/<\/span>$/.test(arWrapped), "arabic citation closes the rtl span at the end");
assert.strictEqual((arWrapped.match(/<span dir="rtl">/g) || []).length, 1, "exactly one rtl span");
assert.ok(arWrapped.indexOf('<i>دعائم الإسلام</i>') !== -1, "italic tags preserved inside the span");

// Mixed cluster (English + Arabic in one note): English prefix stays OUTSIDE the span,
// the Arabic clause (incl its parens/comma/1951/period) is wrapped.
const mixed = 'Farhad Daftary, <i>The Fatimid Empire</i> (Edinburgh University Press, 2018); القاضي النعمان, <i>دعائم الإسلام</i> (دار المعارف, 1951).';
const mixedWrapped = CiteWord.wrapRtlRuns(mixed);
const spanIdx = mixedWrapped.indexOf('<span dir="rtl">');
assert.ok(spanIdx > 0, "mixed: rtl span starts after the english prefix");
assert.ok(mixedWrapped.slice(0, spanIdx).indexOf('2018)') !== -1, "english '2018)' stays before/outside the span");
assert.strictEqual(mixedWrapped.slice(0, spanIdx).indexOf('القاضي'), -1, "no arabic before the span");
assert.strictEqual((mixedWrapped.match(/<span dir="rtl">/g) || []).length, 1, "mixed: exactly one rtl span");
assert.ok(/<span dir="rtl">[\s\S]*1951[\s\S]*<\/span>/.test(mixedWrapped), "arabic clause incl 1951 is inside the span");

// Arabic-punctuation localization: ASCII comma/semicolon INSIDE an Arabic run become
// their Arabic forms (،/؛); Latin runs keep ASCII punctuation.
assert.ok(arWrapped.indexOf("،") !== -1, "arabic run uses the Arabic comma ،");
assert.strictEqual(arWrapped.indexOf(","), -1, "no ASCII comma remains in a pure-Arabic citation");
// Mixed: the English 'Daftary,' keeps its ASCII comma; the Arabic clause uses ،
assert.ok(mixedWrapped.slice(0, spanIdx).indexOf("Daftary,") !== -1, "english ASCII comma preserved outside the span");
assert.ok(mixedWrapped.slice(spanIdx).indexOf("،") !== -1, "arabic comma used inside the span");
// Pure English is still untouched (no Arabic run to localize)
assert.strictEqual(CiteWord.wrapRtlRuns(enOnly).indexOf("،"), -1, "pure-LTR citation gains no Arabic punctuation");

// buildNotePayload / buildBibliographyPayload apply the wrap after sanitize
const notePayload = CiteWord.buildNotePayload({ html: arOnly, rtl: true });
assert.ok(notePayload.html.indexOf('<span dir="rtl">') !== -1, "note payload html is bidi-wrapped");
console.log("cite-word test passed");
