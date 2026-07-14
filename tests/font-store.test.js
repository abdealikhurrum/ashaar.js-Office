"use strict";

// Unit tests for AshaarFontStore.parseNames — the pure sfnt name-table parser
// used to auto-detect an uploaded font's family name. Browser-only pieces
// (FontFace / IndexedDB) are verified manually in Word.

const assert = require("assert");
const fs = require("fs");
const path = require("path");

const AshaarFontStore = require("../src/taskpane/font-store");

const root = path.join(__dirname, "..");

function read(rel) {
  return fs.readFileSync(path.join(root, rel));
}

// Fatemi Maqala (bundled; family renamed with a space in the 2026-07 font update)
const fatemi = AshaarFontStore.parseNames(read("assets/fonts/FatemiMaqala-Regular.ttf"));
assert.ok(fatemi, "parseNames returns a result for a TTF");
assert.equal(fatemi.family, "Fatemi Maqala", "detects Fatemi Maqala family name");

// AlFatemi (vendored) → family "AlFatemi"
const alfatemi = AshaarFontStore.parseNames(read("vendor/font-fatemi/alfatemi/AlFatemi-Regular.ttf"));
assert.ok(alfatemi, "parseNames returns a result for the AlFatemi TTF");
assert.equal(alfatemi.family, "AlFatemi", "detects AlFatemi family name");

// Accepts ArrayBuffer as well as Buffer/Uint8Array
const buf = read("assets/fonts/FatemiMaqala-Regular.ttf");
const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
assert.equal(AshaarFontStore.parseNames(ab).family, "Fatemi Maqala", "accepts ArrayBuffer input");

// Non-sfnt input (WOFF/garbage) → null, so callers fall back to the filename
assert.equal(AshaarFontStore.parseNames(Buffer.from("wOFFnot a real font")), null,
  "returns null for non-sfnt input");
assert.equal(AshaarFontStore.parseNames(Buffer.from([0, 1, 2])), null,
  "returns null for a too-short buffer");

console.log("font-store tests passed");
