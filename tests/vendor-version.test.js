"use strict";
// Final review M1: taskpane.js hand-mirrors src/vendor/ASHAAR_UPSTREAM_VERSION
// (a build-time text stamp written by scripts/sync-ashaar-vendor.mjs — never
// loaded into the browser, see the comment at taskpane.js's ASHAAR_UPSTREAM_VERSION
// declaration) because this vanilla/no-build-step add-in has no fetch()/XHR
// path to read a sibling asset at runtime. That means `npm run update:ashaar`
// moving the vendor pointer WITHOUT a matching manual bump of the taskpane.js
// constant silently serves stale probe/calibration cache keys forever (the
// engine-build id never changes, so AshaarTuneCache never busts). This test
// is the cheapest insurance against that drift: it reads the authoritative
// vendor stamp and asserts the hand-mirrored constant still matches its
// commit prefix.
//
// taskpane.js is not require()-able in Node (it assumes Office.js/DOM
// globals at module scope), so the constant is regex-extracted from the
// file's source text — the same "read source as text" approach fonts.test.js
// uses to byte-guard a source-level invariant.
const assert = require("assert");
const fs = require("fs");
const path = require("path");

const vendorStampPath = path.join(__dirname, "..", "src", "vendor", "ASHAAR_UPSTREAM_VERSION");
const taskpanePath = path.join(__dirname, "..", "src", "taskpane", "taskpane.js");

const vendorStamp = fs.readFileSync(vendorStampPath, "utf8");
const commitMatch = vendorStamp.match(/^commit=([0-9a-f]+)/m);
assert.ok(commitMatch, "src/vendor/ASHAAR_UPSTREAM_VERSION must have a commit= line");
const fullCommit = commitMatch[1];
const shortCommit = fullCommit.slice(0, 8);

const taskpaneSrc = fs.readFileSync(taskpanePath, "utf8");
const constMatch = taskpaneSrc.match(/var\s+ASHAAR_UPSTREAM_VERSION\s*=\s*"([0-9a-f]+)"/);
assert.ok(constMatch, "taskpane.js must declare `var ASHAAR_UPSTREAM_VERSION = \"...\";`");
const mirrored = constMatch[1];

assert.strictEqual(
  mirrored,
  shortCommit,
  "taskpane.js's hand-mirrored ASHAAR_UPSTREAM_VERSION (\"" + mirrored + "\") is stale — " +
  "src/vendor/ASHAAR_UPSTREAM_VERSION now points at commit " + fullCommit +
  " (short: \"" + shortCommit + "\"). Run `npm run update:ashaar` and update the " +
  "ASHAAR_UPSTREAM_VERSION constant in src/taskpane/taskpane.js to match."
);

console.log("vendor-version.test.js OK");
