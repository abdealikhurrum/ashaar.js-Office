const assert = require("assert");
const { normalizeSeparators, detectSeparator } = require("../src/taskpane/separators");

function norm(text, opts) { return normalizeSeparators(text, opts).text; }

// --- auto-detect + normalize to canonical "\" ---

// Spaced dash between hemistichs
assert.equal(norm("بیت اول - بیت دوم\nسوم - چہارم"), "بیت اول \\ بیت دوم\nسوم \\ چہارم",
  "spaced dash becomes \\");
assert.equal(normalizeSeparators("a - b\nc - d").detected, "dash", "detects dash");

// Tab-separated
assert.equal(norm("a\tb\nc\td"), "a \\ b\nc \\ d", "tab becomes \\");

// Asterisk / pipe are normalized to canonical "\"
assert.equal(norm("a * b\nc * d"), "a \\ b\nc \\ d", "asterisk becomes \\");
assert.equal(norm("a | b\nc | d"), "a \\ b\nc \\ d", "pipe becomes \\");

// Runs of 2+ spaces (wide gap) when dominant
assert.equal(norm("a    b\nc    d"), "a \\ b\nc \\ d", "wide-gap becomes \\");

// Multi-misra line (3 parts)
assert.equal(norm("a - b - c"), "a \\ b \\ c", "3-part dash line becomes multi-misra");

// Already canonical "\" is left unchanged
assert.equal(norm("a \\ b\nc \\ d"), "a \\ b\nc \\ d", "canonical text unchanged");
assert.equal(normalizeSeparators("a \\ b").changed, false, "canonical text not marked changed");

// --- structure preservation ---

// Blank line (stanza) and --- (poem) markers are preserved verbatim
assert.equal(norm("a - b\n\n---\nc - d"), "a \\ b\n\n---\nc \\ d", "stanza/poem markers preserved");

// --- false-positive guards ---

// A hyphenated word (no surrounding spaces) is NOT a separator
assert.equal(norm("well-known truth\nfar-flung land"), "well-known truth\nfar-flung land",
  "intra-word hyphen is not a separator");
assert.equal(normalizeSeparators("well-known truth").detected, null, "no separator detected for hyphenated words");

// Solo lines (no separator at all) are left alone
assert.equal(norm("single misra line\nanother solo line"), "single misra line\nanother solo line",
  "solo lines unchanged");

// Single spaces never trigger the wide-gap rule
assert.equal(normalizeSeparators("word one\nword two").detected, null, "single spaces are not a separator");

// --- explicit overrides ---

// Custom literal separator
assert.equal(norm("a / b\nc / d", { separator: "custom", customPattern: "/" }), "a \\ b\nc \\ d",
  "custom separator applied");

// Forcing a specific separator
assert.equal(norm("a . b", { separator: "custom", customPattern: "." }), "a \\ b", "custom dot, literal not regex");

// --- pair-every-2-lines (for one-hemistich-per-line files) ---

assert.equal(norm("m1\nm2\nm3\nm4", { pairLines: true }), "m1 \\ m2\nm3 \\ m4", "pairs consecutive lines");
assert.equal(norm("m1\nm2\n\nm3\nm4", { pairLines: true }), "m1 \\ m2\n\nm3 \\ m4",
  "pairing respects stanza breaks");
assert.equal(norm("m1\nm2\nm3", { pairLines: true }), "m1 \\ m2\nm3", "odd leftover stays solo");

// --- edges ---
assert.equal(norm(""), "", "empty input");
assert.equal(normalizeSeparators("").detected, null, "empty input detects nothing");

console.log("separators tests passed");
