"use strict";
const assert = require("assert");
const fs = require("fs");
const path = require("path");

const dest = path.join(__dirname, "..", "src", "vendor");
for (const rel of ["citeproc.js", "csl-locales/locales-ar.xml", "csl-locales/locales-en-US.xml",
                   "csl-styles/chicago-notes-bibliography.csl", "csl-styles/apa.csl",
                   "CITEPROC_UPSTREAM_VERSION"]) {
  assert.ok(fs.existsSync(path.join(dest, rel)), `missing vendored asset: ${rel}`);
}

const stamp = fs.readFileSync(path.join(dest, "CITEPROC_UPSTREAM_VERSION"), "utf8");
assert.ok(/^commit=[0-9a-f]{7,}/m.test(stamp), "stamp must have a commit= line");

const CSL = require(path.join(dest, "citeproc.js"));
assert.strictEqual(typeof CSL.Engine, "function", "citeproc bundle must expose CSL.Engine");
console.log("citeproc-vendor test passed");
