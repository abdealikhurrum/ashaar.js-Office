const assert = require("assert");
const AshaarPanel = require("../src/taskpane/settings-panel");
const { resolveSettings, profileFromSettings, defaultSettings } = require("../src/taskpane/profiles");

// ── panelStateFor: header + chips + controls with provenance ────────────────
{
  const store = { Karbala: profileFromSettings("Karbala", Object.assign(defaultSettings(), { strength: 9 })) };
  const payload = { profile: "Karbala", local: { gap: 8 } };
  const resolved = resolveSettings({ payload, profileStore: store, scope: { level: "poem" } });
  const st = AshaarPanel.panelStateFor({
    resolved,
    pending: { set: {}, clear: [] },
    target: { kind: "block", scope: { level: "poem" }, cellEnabled: false, gapEnabled: false },
  });
  assert.equal(st.header.title, "Poem — Karbala");
  assert.deepEqual(st.chips, [
    { level: "poem", enabled: true, active: true },
    { level: "bandh", enabled: true, active: false },
    { level: "cell", enabled: false, active: false },
    { level: "gap", enabled: false, active: false },
  ]);
  const gap = st.controls.find((c) => c.key === "gap");
  assert.equal(gap.value, 8);
  assert.equal(gap.source, "local");
  assert.equal(gap.dirty, false);
  const strength = st.controls.find((c) => c.key === "strength");
  assert.equal(strength.value, 9);
  assert.equal(strength.source, "profile");
  assert.equal(st.profileRow.name, "Karbala");
  assert.equal(st.profileRow.assignEnabled, true);
  assert.equal(st.profileRow.updateVisible, false, "no pending edits → no drift");
  assert.equal(st.footer.revertLabel, "Revert to profile");
}

// ── Pending edits show as dirty and surface the Update action ────────────────
{
  const resolved = resolveSettings({ payload: { profile: "Karbala", local: {} }, profileStore: { Karbala: profileFromSettings("Karbala", defaultSettings()) }, scope: { level: "poem" } });
  const st = AshaarPanel.panelStateFor({
    resolved,
    pending: { set: { strength: 3 }, clear: [] },
    target: { kind: "block", scope: { level: "poem" }, cellEnabled: false, gapEnabled: false },
  });
  const strength = st.controls.find((c) => c.key === "strength");
  assert.equal(strength.value, 3, "pending value shown");
  assert.equal(strength.dirty, true);
  assert.equal(st.profileRow.updateVisible, true, "drift → Update visible");
}

// ── Plain selection target ───────────────────────────────────────────────────
{
  const resolved = resolveSettings({ payload: null, profileStore: {}, scope: { level: "poem" } });
  const st = AshaarPanel.panelStateFor({
    resolved,
    pending: { set: {}, clear: [] },
    target: { kind: "selection", scope: { level: "poem" }, cellEnabled: false, gapEnabled: false },
  });
  assert.equal(st.header.title, "Selection");
  assert.equal(st.chips.length, 0, "no chips outside a block");
  assert.equal(st.profileRow.assignEnabled, false, "nowhere to store the link");
  assert.equal(st.footer.revertLabel, "Reset to defaults");
}

// ── Missing profile surfaces the restore affordance ──────────────────────────
{
  const resolved = resolveSettings({ payload: { profile: "Karbala", local: {}, profileCache: { strength: 9 } }, profileStore: {}, scope: { level: "poem" } });
  const st = AshaarPanel.panelStateFor({
    resolved,
    pending: { set: {}, clear: [] },
    target: { kind: "block", scope: { level: "poem" }, cellEnabled: false, gapEnabled: false },
  });
  assert.equal(st.profileRow.missing, true);
  assert.equal(st.header.title, "Poem — Karbala (not on this machine)");
}

// ── Scope narrows the control set ────────────────────────────────────────────
{
  const resolved = resolveSettings({ payload: { profile: "", local: {}, overrides: {} }, profileStore: {}, scope: { level: "cell", key: "A2:3" } });
  const st = AshaarPanel.panelStateFor({
    resolved,
    pending: { set: {}, clear: [] },
    target: { kind: "block", scope: { level: "cell", key: "A2:3" }, cellEnabled: true, gapEnabled: false, cellLabel: "A2:3" },
  });
  assert.equal(st.header.title, "Poem › Cell A2:3");
  assert.deepEqual(st.controls.map((c) => c.key), ["strength", "misraWidthPt", "capEm"]);
}

// ── mergePending / pendingToLocal ────────────────────────────────────────────
{
  let p = { set: {}, clear: [] };
  p = AshaarPanel.mergePending(p, "gap", 8);
  assert.deepEqual(p, { set: { gap: 8 }, clear: [] });
  p = AshaarPanel.mergePending(p, "gap", null); // ⟲ reset marks a clear
  assert.deepEqual(p, { set: {}, clear: ["gap"] });
  p = AshaarPanel.mergePending(p, "strength", 9);

  const local = { gap: 8, widthPct: 80 };
  const next = AshaarPanel.pendingToLocal(local, p, AshaarPanel.SCOPE_FIELDS.poem);
  assert.deepEqual(next, { widthPct: 80, strength: 9 }, "clear deletes, set adds, untouched survives");
  // Inputs not mutated.
  assert.deepEqual(local, { gap: 8, widthPct: 80 });
}

console.log("settings-panel tests passed");
