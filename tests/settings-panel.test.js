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

// ── pendingToLocal: clears are scope-gated like sets ─────────────────────────
{
  const local = { gap: 8, widthPct: 80 };
  const p = { set: {}, clear: ["gap"] };
  const next = AshaarPanel.pendingToLocal(local, p, AshaarPanel.SCOPE_FIELDS.cell);
  assert.deepEqual(next, { gap: 8, widthPct: 80 }, "cell scope cannot clear a poem-scope key");
  const nextPoem = AshaarPanel.pendingToLocal(local, p, AshaarPanel.SCOPE_FIELDS.poem);
  assert.deepEqual(nextPoem, { widthPct: 80 }, "owning scope clears it");
}

// ── panelStateFor: pending clear shows the inherited value, dirty ────────────
{
  const store = { K: profileFromSettings("K", Object.assign(defaultSettings(), { gap: 6 })) };
  const resolved = resolveSettings({ payload: { profile: "K", local: { gap: 8 } }, profileStore: store, scope: { level: "poem" } });
  const st = AshaarPanel.panelStateFor({
    resolved,
    pending: { set: {}, clear: ["gap"] },
    target: { kind: "block", scope: { level: "poem" }, cellEnabled: false, gapEnabled: false },
  });
  const gap = st.controls.find((c) => c.key === "gap");
  assert.equal(gap.value, 6, "clear falls back to profile layer");
  assert.equal(gap.dirty, true);
}

// ── §9 keys are poem-scope panel fields ──────────────────────────────────────
{
  assert.ok(AshaarPanel.SCOPE_FIELDS.poem.indexOf("lineHeightPt") !== -1);
  assert.ok(AshaarPanel.SCOPE_FIELDS.poem.indexOf("separatorPt") !== -1);
}

// ── §6 refresh-cost labels ───────────────────────────────────────────────────
{
  const resolved = resolveSettings({ payload: { profile: "", local: {} }, profileStore: {}, scope: { level: "poem" } });
  const t = { kind: "block", scope: { level: "poem" }, cellEnabled: false, gapEnabled: false };
  const RERENDER_SUFFIX = " · Re-render: always rebuilds poem tables (full re-layout)";
  const structural = AshaarPanel.panelStateFor({ resolved, pending: { set: { gap: 8 }, clear: [] }, target: t });
  assert.strictEqual(structural.footer.costLabel, "Apply — rebuilds poem tables" + RERENDER_SUFFIX);
  const light = AshaarPanel.panelStateFor({ resolved, pending: { set: { strength: 9 }, clear: [] }, target: t });
  assert.strictEqual(light.footer.costLabel, "Apply — re-justifies poem" + RERENDER_SUFFIX);
  const cellT = { kind: "block", scope: { level: "cell", key: "A2:3" }, cellEnabled: true, gapEnabled: false, cellLabel: "A2:3" };
  const cellResolved = resolveSettings({ payload: { profile: "", local: {} }, profileStore: {}, scope: { level: "cell", key: "A2:3" } });
  assert.strictEqual(AshaarPanel.panelStateFor({ resolved: cellResolved, pending: { set: {}, clear: [] }, target: cellT }).footer.costLabel,
    "Apply — re-justifies poem" + RERENDER_SUFFIX);
  // justifyMode none → unjustified suffix
  const noneResolved = resolveSettings({ payload: { profile: "", local: { justifyMode: "none" } }, profileStore: {}, scope: { level: "poem" } });
  assert.strictEqual(AshaarPanel.panelStateFor({ resolved: noneResolved, pending: { set: { gap: 8 }, clear: [] }, target: t }).footer.costLabel,
    "Apply — rebuilds poem tables (unjustified: Justification is None)" + RERENDER_SUFFIX);
  assert.deepStrictEqual(AshaarPanel.STRUCTURAL_KEYS, ["gap", "widthMode", "widthPct", "layoutMode", "colWidthMode", "separatorPt"]);

  // A plain (non-block) selection never shows Re-render — the button is
  // disabled for it, so the caption must not mention it either.
  const selT = { kind: "selection", scope: { level: "poem" }, cellEnabled: false, gapEnabled: false };
  const selSt = AshaarPanel.panelStateFor({ resolved, pending: { set: { gap: 8 }, clear: [] }, target: selT });
  assert.strictEqual(selSt.footer.costLabel, "Apply — rebuilds poem tables",
    "no Re-render suffix outside a block target");
}

// ── resettable: committed cell/bandh overrides keep the reset dot, not just
// local deltas or in-pane pending edits ──────────────────────────────────────
{
  assert.deepStrictEqual(AshaarPanel.RESETTABLE_SOURCES, ["local", "cell", "bandh"]);

  // "cell" source (committed override) → resettable even though not dirty.
  const cellResolved = resolveSettings({
    payload: { profile: "", local: {}, overrides: { "0:A1": { strength: 9 } } },
    profileStore: {},
    scope: { level: "cell", key: "0:A1" },
  });
  const cellT = { kind: "block", scope: { level: "cell", key: "0:A1" }, cellEnabled: true, gapEnabled: false, cellLabel: "A1" };
  const cellSt = AshaarPanel.panelStateFor({ resolved: cellResolved, pending: { set: {}, clear: [] }, target: cellT });
  const strengthC = cellSt.controls.find((c) => c.key === "strength");
  assert.equal(strengthC.source, "cell");
  assert.equal(strengthC.dirty, false, "committed, not a pending edit");
  assert.equal(strengthC.resettable, true, "committed cell override still offers a reset");

  // "bandh" source (committed width) → resettable.
  const bandhResolved = resolveSettings({
    payload: { profile: "", local: {}, widthPt: 300 },
    profileStore: {},
    scope: { level: "bandh" },
  });
  const bandhT = { kind: "block", scope: { level: "bandh" }, cellEnabled: false, gapEnabled: false };
  const bandhSt = AshaarPanel.panelStateFor({ resolved: bandhResolved, pending: { set: {}, clear: [] }, target: bandhT });
  const widthB = bandhSt.controls.find((c) => c.key === "misraWidthPt");
  assert.equal(widthB.source, "bandh");
  assert.equal(widthB.resettable, true, "committed bandh width still offers a reset");

  // "default" and "profile" sources → NOT resettable when not dirty.
  const store = { K: profileFromSettings("K", Object.assign(defaultSettings(), { strength: 9 })) };
  const profResolved = resolveSettings({ payload: { profile: "K", local: {} }, profileStore: store, scope: { level: "poem" } });
  const poemT = { kind: "block", scope: { level: "poem" }, cellEnabled: false, gapEnabled: false };
  const profSt = AshaarPanel.panelStateFor({ resolved: profResolved, pending: { set: {}, clear: [] }, target: poemT });
  const strengthP = profSt.controls.find((c) => c.key === "strength");
  assert.equal(strengthP.source, "profile");
  assert.equal(strengthP.resettable, false, "profile-sourced value has nothing local to reset");
  const gapP = profSt.controls.find((c) => c.key === "gap");
  assert.equal(gapP.source, "default");
  assert.equal(gapP.resettable, false, "default value has nothing to reset");

  // A dirty pending edit is resettable regardless of its underlying source.
  const dirtySt = AshaarPanel.panelStateFor({ resolved: profResolved, pending: { set: { gap: 8 }, clear: [] }, target: poemT });
  const gapDirty = dirtySt.controls.find((c) => c.key === "gap");
  assert.equal(gapDirty.dirty, true);
  assert.equal(gapDirty.resettable, true, "dirty pending edit is resettable even from a default source");
}

console.log("settings-panel tests passed");
