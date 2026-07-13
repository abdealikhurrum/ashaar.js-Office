const assert = require("assert");
const TC = require("../src/taskpane/tune-cache");

assert.strictEqual(TC.probeKey("Jameel", "v7"), "probe|Jameel|v7");
assert.strictEqual(TC.probeKey("Jameel", "v8") === TC.probeKey("Jameel", "v7"), false, "engine bump busts");

const k1 = TC.calibKey("Jameel", 14, 412, ["الف", "ب"]);
assert.strictEqual(TC.calibKey("Jameel", 14, 420, ["الف", "ب"]), k1, "same 25px bucket");
assert.notStrictEqual(TC.calibKey("Jameel", 14, 460, ["الف", "ب"]), k1, "different bucket");
assert.notStrictEqual(TC.calibKey("Jameel", 14, 412, ["الف", "ج"]), k1, "texts hash differs");

// storage-backed probes, in-memory calib
const mem = {};
const storage = { getItem: k => (k in mem ? mem[k] : null), setItem: (k, v) => { mem[k] = v; }, removeItem: k => { delete mem[k]; } };
const c = TC.makeCache(storage);
assert.strictEqual(c.getProbe("probe|J|v7"), null);
c.putProbe("probe|J|v7", { pairs: { AB: 1 } });
assert.deepStrictEqual(c.getProbe("probe|J|v7"), { pairs: { AB: 1 } });
assert.deepStrictEqual(TC.makeCache(storage).getProbe("probe|J|v7"), { pairs: { AB: 1 } }, "survives new instance via storage");
c.putCalib("k", { targetFill: 0.9 });
assert.deepStrictEqual(c.getCalib("k"), { targetFill: 0.9 });
assert.strictEqual(TC.makeCache(storage).getCalib("k"), null, "calib is in-memory only");
c.bustAll();
assert.strictEqual(c.getProbe("probe|J|v7"), null, "bust clears storage too");

console.log("tune-cache tests passed");
