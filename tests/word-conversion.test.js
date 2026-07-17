const assert = require("assert");
const C = require("../src/taskpane/word-conversion");
const T = C.TATWEEL;

// ── table shape ──
{
  assert.deepEqual(Object.keys(C.DIRECTIONS).sort(), ["TO_LEGACY", "TO_MODERN"]);
  assert.ok(Array.isArray(C.MAPPINGS) && C.MAPPINGS.length >= 10);
  const ids = C.MAPPINGS.map(m => m.id);
  assert.equal(new Set(ids).size, ids.length, "ids are unique");
  C.MAPPINGS.forEach(m => {
    assert.ok(["letter", "mark", "symbol"].includes(m.category), m.id + " category");
    assert.equal(typeof m.legacy, "string");
    assert.equal(typeof m.modern, "string");
    assert.equal(typeof m.wholeWord, "boolean");
    assert.equal(typeof m.lossy, "boolean");
  });
  const seen = C.MAPPINGS.find(m => m.id === "seen-baariye");
  assert.equal(seen.legacy, "سس"); assert.equal(seen.modern, "ے");
}

// ── letter tier: simple contiguous substitution both ways ──
{
  assert.equal(C.convert("حح", C.DIRECTIONS.TO_MODERN), "چ");
  assert.equal(C.convert("چ", C.DIRECTIONS.TO_LEGACY), "حح");
  assert.equal(C.convert("كك گگ", C.DIRECTIONS.TO_MODERN), "گ گگ",
    "كك→گ; an already-modern گگ is left alone");
}

// ── kashida-escape: genuine double letter round-trips ──
{
  // TO_MODERN: escaped double (سـس) becomes a genuine double seen (سس),
  // while a plain double (سس) becomes ے.
  assert.equal(C.convert("سس", C.DIRECTIONS.TO_MODERN), "ے");
  assert.equal(C.convert("س" + T + "س", C.DIRECTIONS.TO_MODERN), "سس");
  // TO_LEGACY: ے becomes سس (contiguous), while a genuine double seen (سس)
  // is protected with a tatweel so the old font won't merge it.
  assert.equal(C.convert("ے", C.DIRECTIONS.TO_LEGACY), "سس");
  assert.equal(C.convert("سس", C.DIRECTIONS.TO_LEGACY), "س" + T + "س");
  // Round-trip a token containing a genuine double seen.
  const word = "بسس";
  const toLegacy = C.convert(word, C.DIRECTIONS.TO_LEGACY);
  assert.equal(toLegacy, "ب" + "س" + T + "س");
  assert.equal(C.convert(toLegacy, C.DIRECTIONS.TO_MODERN), word, "round-trips");
}

// ── whole-word: چھے ⇄ ؛ ──
{
  assert.equal(C.convert("؛", C.DIRECTIONS.TO_MODERN), "چھے");
  assert.equal(C.convert("چھے", C.DIRECTIONS.TO_LEGACY), "؛", "standalone word collapses");
  assert.equal(C.convert("چھے دن", C.DIRECTIONS.TO_LEGACY).split(" ")[0], "؛",
    "چھے as a whitespace-delimited word still collapses");
  // Inside a larger word the semicolon collapse is suppressed; the constituent
  // چ and ے still convert to their own double-press forms (حح … سس).
  assert.equal(C.convert("اچھے", C.DIRECTIONS.TO_LEGACY).indexOf("؛"), -1,
    "چھے inside a larger word does NOT collapse to ؛");
}

// ── enabledIds filters which rows run ──
{
  assert.equal(C.convert("حح كك", C.DIRECTIONS.TO_MODERN, ["cheh-hah"]), "چ كك",
    "only the enabled row converts");
}

// ── buildOperations ordering: TO_MODERN puts contiguous before escape-drop ──
{
  const ops = C.buildOperations(C.DIRECTIONS.TO_MODERN, ["seen-baariye"]);
  const iContig = ops.findIndex(o => o.find === "سس");
  const iEscape = ops.findIndex(o => o.find === "س" + T + "س");
  assert.ok(iContig >= 0 && iEscape >= 0 && iContig < iEscape,
    "contiguous سس→ے must run before سـس→سس");
}
// ── buildOperations ordering: TO_LEGACY protects doubles before ے→سس ──
{
  const ops = C.buildOperations(C.DIRECTIONS.TO_LEGACY, ["seen-baariye"]);
  const iProtect = ops.findIndex(o => o.find === "سس");
  const iSub = ops.findIndex(o => o.find === "ے");
  assert.ok(iProtect >= 0 && iSub >= 0 && iProtect < iSub,
    "protect سس→سـس must run before ے→سس");
}

// ── groupsForUi ──
{
  const groups = C.groupsForUi();
  assert.ok(groups.every(g => g.category && Array.isArray(g.rows)));
  assert.ok(groups.some(g => g.category === "letter"));
}

console.log("word-conversion.test.js: all assertions passed");
