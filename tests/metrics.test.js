const assert = require("assert");
const M = require("../src/taskpane/metrics");
let t = 0; const now = () => t;
const run = M.startRun("apply", now);
t = 5; run.phase("tag write");
t = 25; run.phase("justify");
t = 100; run.end();
const r = run.report();
assert.strictEqual(r.label, "apply");
assert.strictEqual(r.totalMs, 100);
assert.deepStrictEqual(r.phases, [
  { name: "start", ms: 5 }, { name: "tag write", ms: 20 }, { name: "justify", ms: 75 },
]);
console.log("metrics tests passed");
