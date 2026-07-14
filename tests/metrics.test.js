const assert = require("assert");
const M = require("../src/taskpane/metrics");
let t = 0; const now = () => t;

// Correct semantics: phase(name) labels the segment that just ENDED —
// from the previous boundary (startRun or the last phase() call) up to now.
const run = M.startRun("apply", now);
t = 5; run.phase("tag write");
t = 25; run.phase("justify");
t = 100; run.end();
const r = run.report();
assert.strictEqual(r.label, "apply");
assert.strictEqual(r.totalMs, 100);
assert.deepStrictEqual(r.phases, [
  { name: "tag write", ms: 5 }, { name: "justify", ms: 20 }, { name: "(tail)", ms: 75 },
]);

// No tail when end() coincides with the last phase() boundary — nothing
// elapsed after the last phase() call, so no synthetic "(tail)" segment.
t = 0;
const run2 = M.startRun("apply", now);
t = 10; run2.phase("only phase");
t = 10; run2.end();
const r2 = run2.report();
assert.deepStrictEqual(r2.phases, [{ name: "only phase", ms: 10 }]);
assert.strictEqual(r2.totalMs, 10);

// A run with no phase() calls at all before end() reports just a tail.
t = 0;
const run3 = M.startRun("apply", now);
t = 50; run3.end();
const r3 = run3.report();
assert.deepStrictEqual(r3.phases, [{ name: "(tail)", ms: 50 }]);

console.log("metrics tests passed");
