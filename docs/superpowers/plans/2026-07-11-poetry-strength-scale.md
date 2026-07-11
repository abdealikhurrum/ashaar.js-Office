# Poetry Kashida Strength (1–10 elongation:spacing ratio) — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Turn the poetry "Kashida strength" control into a 1–10 ratio of elongation→spacing: strength `s` sets `φ = (s−1)/9`, the share of each line's fill-gap closed by the font's elongation mechanism (Jameel swap / Mehr / generic tatweels); residual spacing fills the rest. Total fill still targets the column edge.

**Architecture:** New pure `strengthToElongationShare(s)→φ`. The three per-cell elongation branches in `justifySelection` target `natural + φ·gap` instead of the full column, then the existing residual spacing fills to the edge. Slider becomes 1–10; all other 0–24 strength consumers are rescaled to the 1–10 domain (no behavior break); stored profiles' 0–24 strength is remapped on load. Prose semantics and non-Jameel alternates are out of scope (ashaar-js#8 / #7).

**Tech Stack:** Vanilla JS (ES5/UMD, no build step), Office.js v1, Node `assert` tests.

**Spec:** `docs/superpowers/specs/2026-07-11-poetry-strength-scale-design.md`

## Global Constraints
- ES5/UMD only (`var`, `function`); never edit `src/vendor/`.
- Pure logic node-tested; Office.js glue in `taskpane.js` is manual-verify (Task 6).
- Reuse existing helpers (`capMicroSpaces`, `selectSwapRuns`, `justifyRuns`, per-cell `cellDesc`/`cellMech`, the hybrid residual code).
- `φ = (clamp(s,1,10) − 1) / 9`. Fill target = the column edge (`colPx`).
- All 0–24 strength consumers move to a **1–10** domain; stored profiles remap `round(1 + old/24·9)` clamped to [1,10].
- `npm test` green after every task.

---

### Task 1: `strengthToElongationShare` + slider rescale

**Files:** Modify `src/taskpane/word-html.js` (+ export); `tests/word-html.test.js`; `src/taskpane/taskpane.html`.

**Interfaces:** Produces `AshaarWord.strengthToElongationShare(strength) → number` in `[0,1]`.

- [ ] **Step 1: Failing test** (append to `tests/word-html.test.js`):
```js
// ── strengthToElongationShare: 1–10 → φ elongation share ─────────────────────
assert.strictEqual(AshaarWord.strengthToElongationShare(1), 0);
assert.strictEqual(AshaarWord.strengthToElongationShare(10), 1);
assert.ok(Math.abs(AshaarWord.strengthToElongationShare(5) - (4/9)) < 1e-9, "s5 → 4/9");
assert.strictEqual(AshaarWord.strengthToElongationShare(0), 0);   // clamps low
assert.strictEqual(AshaarWord.strengthToElongationShare(24), 1);  // clamps high
assert.strictEqual(AshaarWord.strengthToElongationShare(undefined), 0);
```
- [ ] **Step 2: Run to verify it fails** — `node tests/word-html.test.js`; Expected: FAIL (not a function).
- [ ] **Step 3: Implement** (add near `sliderToFill`, add to exports):
```js
// Poetry Kashida strength (1–10) → elongation share φ ∈ [0,1]: the fraction of a
// line's fill-gap closed by the font's elongation mechanism (the rest by
// spacing). s=1 → 0 (all spacing); s=10 → 1 (all elongation, minor spacing).
function strengthToElongationShare(strength) {
  var s = Number(strength);
  if (!isFinite(s)) s = 1;
  s = Math.max(1, Math.min(10, s));
  return (s - 1) / 9;
}
```
- [ ] **Step 4: Run to verify it passes** — `node tests/word-html.test.js`; Expected: PASS.
- [ ] **Step 5: Rescale the sliders** — in `src/taskpane/taskpane.html`, change both:
  - `#tatweel-count`: `min="1" max="10" value="7"` (was `min="0" max="24" value="6"`).
  - `#qaseeda-strength`: `min="1" max="10" value="7"`.
  Update the adjacent value labels' default text (`#tatweel-value`, `#qaseeda-strength-value`) to `7`.
- [ ] **Step 6: Full suite** — `npm test`; Expected: green.
- [ ] **Step 7: Commit**
```bash
git add src/taskpane/word-html.js tests/word-html.test.js src/taskpane/taskpane.html
git commit -m "feat(strength): strengthToElongationShare + 1-10 sliders"
```

---

### Task 2: Rescale the css/insert strength helpers to 1–10

**Files:** Modify `src/taskpane/word-html.js`; `tests/word-html.test.js`.

**Interfaces:** `strengthToKashidaLevel`, `kashidaExpansionFraction`, `sliderToFill` keep their signatures; only their input domain changes from 0–24 to 1–10.

- [ ] **Step 1: Update the existing tests** for the new domain (replace the old 0–24 assertions):
```js
// strengthToKashidaLevel — thirds of 1–10
assert.equal(AshaarWord.strengthToKashidaLevel(1),  "lowKashida");
assert.equal(AshaarWord.strengthToKashidaLevel(3),  "lowKashida");
assert.equal(AshaarWord.strengthToKashidaLevel(4),  "mediumKashida");
assert.equal(AshaarWord.strengthToKashidaLevel(6),  "mediumKashida");
assert.equal(AshaarWord.strengthToKashidaLevel(7),  "highKashida");
assert.equal(AshaarWord.strengthToKashidaLevel(10), "highKashida");
assert.equal(AshaarWord.strengthToKashidaLevel(undefined), "mediumKashida");
// kashidaExpansionFraction — 0 at 1, ~0.15 at 10
assert.equal(AshaarWord.kashidaExpansionFraction(1), 0);
assert.equal(AshaarWord.kashidaExpansionFraction(10), 0.15);
assert.equal(AshaarWord.kashidaExpansionFraction(999), 0.15); // clamp
```
Remove the superseded 0–24 assertions for these two functions.
- [ ] **Step 2: Run to verify it fails** — `node tests/word-html.test.js`; Expected: FAIL (old thirds/fractions).
- [ ] **Step 3: Implement** — replace the three bodies:
```js
function strengthToKashidaLevel(strength) {
  var s = Number(strength);
  if (!isFinite(s)) return "mediumKashida";
  if (s <= 3) return "lowKashida";
  if (s <= 6) return "mediumKashida";
  return "highKashida";
}
```
```js
function kashidaExpansionFraction(strength) {
  var s = Math.max(1, Math.min(10, Number(strength) || 1));
  return Math.round((0.15 * (s - 1) / 9) * 1000) / 1000;
}
```
```js
function sliderToFill(count) {
  var s = Math.max(1, Math.min(10, Number(count) || 1));
  return 0.90 + ((s - 1) / 9) * 0.10;
}
```
Update their `0–24` doc comments to `1–10`.
- [ ] **Step 4: Run to verify it passes** — `node tests/word-html.test.js`; Expected: PASS.
- [ ] **Step 5: Full suite** — `npm test`; Expected: green (poetry-corpus/word-html OOXML unaffected — these only shift numeric domains).
- [ ] **Step 6: Commit**
```bash
git add src/taskpane/word-html.js tests/word-html.test.js
git commit -m "feat(strength): rescale css/insert strength helpers to 1-10 domain"
```

---

### Task 3: Profile strength — rescale + remap stored 0–24 values

**Files:** Modify `src/taskpane/profiles.js`; `tests/profiles.test.js`.

**Interfaces:**
- `strengthToTargetFill(s)` domain → 1–10 (kept for any non-engine caller; the engine path uses φ, Task 4/5).
- Produces `AshaarProfiles.normalizeStrength(s) → int in [1,10]`: pass-through for 1–10; remap a legacy 0–24 value via `round(1 + s/24·9)`. Applied when a stored profile is loaded.

- [ ] **Step 1: Failing test** (append to `tests/profiles.test.js`):
```js
// strength domain is now 1–10
assert.ok(Math.abs(AshaarProfiles.strengthToTargetFill(1) - 0.90) < 1e-9);
assert.ok(Math.abs(AshaarProfiles.strengthToTargetFill(10) - 1.0) < 1e-9);
// legacy 0–24 values remap into 1–10
assert.strictEqual(AshaarProfiles.normalizeStrength(6), 6);   // already in-range → unchanged
assert.strictEqual(AshaarProfiles.normalizeStrength(0), 1);
assert.strictEqual(AshaarProfiles.normalizeStrength(24), 10);
assert.strictEqual(AshaarProfiles.normalizeStrength(12), 6);  // round(1+12/24*9)=round(5.5)=6
```
Wait — `normalizeStrength(6)` is ambiguous (valid in both ranges). Decision: values `≤10` are treated as already-1–10 (no remap); only `>10` are remapped from the legacy 0–24 scale. Encode that:
```js
assert.strictEqual(AshaarProfiles.normalizeStrength(6), 6);    // ≤10 → unchanged
assert.strictEqual(AshaarProfiles.normalizeStrength(24), 10);  // >10 → legacy remap
assert.strictEqual(AshaarProfiles.normalizeStrength(18), 8);   // round(1+18/24*9)=round(7.75)=8
```
- [ ] **Step 2: Run to verify it fails** — `node tests/profiles.test.js`; Expected: FAIL.
- [ ] **Step 3: Implement** (edit `strengthToTargetFill`; add `normalizeStrength` + export):
```js
function strengthToTargetFill(strength) {
  var s = Math.max(1, Math.min(10, Number(strength) || 1));
  return 0.90 + ((s - 1) / 9) * 0.10;
}
// Legacy 0–24 strengths (values >10) remap into the 1–10 scale; 1–10 pass through.
function normalizeStrength(strength) {
  var s = Number(strength) || 1;
  if (s > 10) s = Math.round(1 + (s / 24) * 9);
  return Math.max(1, Math.min(10, s));
}
```
- [ ] **Step 4: Apply the remap on profile load** — wherever a profile's `justify.strength` is read into the panel/opts (`profileToPanel` / `getProfile` consumers, `taskpane.js`), pass it through `AshaarProfiles.normalizeStrength(...)`. Add a one-line coercion at the read site so stored profiles surface a 1–10 value. (Manual-verify in Task 6.)
- [ ] **Step 5: Run to verify it passes** — `node tests/profiles.test.js`; Expected: PASS. Then `npm test`; Expected: green.
- [ ] **Step 6: Commit**
```bash
git add src/taskpane/profiles.js tests/profiles.test.js src/taskpane/taskpane.js
git commit -m "feat(strength): profile targetFill 1-10 domain + legacy strength remap"
```

---

### Task 4: φ-target in the elongation branches (Office.js — manual verify)

**Files:** Modify `src/taskpane/taskpane.js` (`justifySelection` phase-1 loop; the fixed targetFill).

**Interfaces:** Consumes `AshaarWord.strengthToElongationShare`. No node test (Office.js); verified in Task 6.

- [ ] **Step 1: Compute φ once per justify** — near the top of `justifySelection` (after `opts` is read), add:
```js
var elongShare = AshaarWord.strengthToElongationShare(opts.tatweelCount); // φ ∈ [0,1]
```
- [ ] **Step 2: Fix the fill target to the edge** — the engine poetry path fills to the column edge regardless of strength. Set `calibParams.targetFill = 1.0` for this path (replace the auto-calibrated/0.92 fill that strength used to drive), so `colPx` is the fill target and φ is the only strength lever. (Keep `fontQualityBoost` handling.)
- [ ] **Step 3: Jameel branch** — replace the target passed to `selectSwapRuns` (currently `colPx`) with the φ-scaled elongation target. Just before `taskpane.js:1661`:
```js
var jNatural = wb.reduce(function (a, b) { return a + b; }, 0);
var jTarget = jNatural + elongShare * Math.max(0, colPx - jNatural);
var sel = AshaarKashidaFontswap.selectSwapRuns(fss, wb, ww, jTarget);
```
Residual spacing below is unchanged — it still closes toward `colPx` (`colPx - sel.fill * jTarget`? no: keep residual against the edge). Set the residual’s achieved width from the swap result and target `colPx`: `capMicroSpaces(colPx - achievedJ, jGaps, jSpacePx, repSize*96/72)` where `achievedJ = sel.fill * jTarget`.
- [ ] **Step 4: Mehr branch** — same shape at `taskpane.js:1697`:
```js
var mNatural = mwb.reduce(function (a, b) { return a + b; }, 0);
var mTarget = mNatural + elongShare * Math.max(0, colPx - mNatural);
var msel = AshaarKashidaFontswap.selectSwapRuns(mtoks, mwb, mww, mTarget);
```
and the residual `capMicroSpaces(colPx - (msel.fill * mTarget), mGaps, mSpacePx, repSize*96/72)`.
- [ ] **Step 5: Generic branch** — at `taskpane.js:1763`, target `natural + φ·gap`:
```js
var gNatural = primRuns.reduce(function (a, r) { return a + r.measure(r.text); }, 0);
var gTarget = gNatural + elongShare * Math.max(0, colPx - gNatural);
var kout = AshaarJustify.justifyRuns(primRuns, gTarget, calibParams); // fill only φ of the gap
outTexts = kout.map(function (r) { return r.text; });
```
The existing generic residual-spacing step (if present) closes toward `colPx`; if the generic path has no residual step yet, leave spacing to the `else` branch — engine kashida at φ<1 relies on the mechanism only, so a short line is acceptable (documented accept-short). **Confirm in Task 6** whether generic needs an added residual pass; if so, mirror the Jameel/Mehr `capMicroSpaces` + `distributeMicroSpaces`.
- [ ] **Step 6: Syntax + suite** — `node --check src/taskpane/taskpane.js`; `npm test`; Expected: green (no node coverage of this glue; regression check only).
- [ ] **Step 7: Commit**
```bash
git add src/taskpane/taskpane.js
git commit -m "feat(strength): elongation targets natural+phi*gap per cell (poetry ratio)"
```

---

### Task 5: `applyProfileToQaseeda` adopts the φ model (Office.js — manual verify)

**Files:** Modify `src/taskpane/taskpane.js` (`applyProfileToQaseeda`).

- [ ] **Step 1:** Compute `elongShare = AshaarWord.strengthToElongationShare(profile.justify.strength)` and set the per-cell justify to target `natural + φ·gap` (mirror Task 4), instead of using `strengthToTargetFill` as the fill lever. Keep `targetFill = 1.0` (edge) for the residual. Route the profile's cells through the same φ-target logic as the free-form path (extract a shared helper if it reduces duplication).
- [ ] **Step 2:** Read `profile.justify.strength` through `AshaarProfiles.normalizeStrength(...)` so legacy profiles behave.
- [ ] **Step 3: Syntax + suite** — `node --check`; `npm test`; green.
- [ ] **Step 4: Commit**
```bash
git add src/taskpane/taskpane.js
git commit -m "feat(strength): qaseeda profile path uses phi elongation share"
```

---

### Task 6: Manual Word verification

**Files:** none.

- [ ] **Step 1:** `npm start` (or reload the pane). Insert an Arabic/Urdu poem.
- [ ] **Step 2 — ratio sweep:** With the dropdown font set so the cell has a real mechanism, set **strength 1** → Justify: the line fills by **spacing only** (letters not elongated; accept-short if wide). Set **strength 10** → Justify: fills by **elongation** (tatweels/Jameel swaps) with only minor spacing. Set **5** → a visible mix. Confirm the balance shifts smoothly.
- [ ] **Step 3 — per font:** repeat for **Jameel** (swaps scale with strength), **Mehr** (tatweel count scales), **generic/Fatemi** (tatweels scale). **Noto/Gulzar**: spacing-only at every strength (φ has no mechanism to drive).
- [ ] **Step 4 — profiles:** apply an existing (pre-migration) profile → confirm it still justifies sanely (strength remapped to 1–10). Save a new profile at strength 8, re-apply → consistent.
- [ ] **Step 5 — css mode:** "Word justify" at strengths 1/5/10 → Word kashida level low/medium/high respectively (thirds of 1–10).
- [ ] **Step 6 — idempotent:** re-justify twice at a fixed strength → no compounding.

## Self-Review notes
- Spec coverage: §1 φ helper → T1; §2 sliders → T1; §3 φ-target → T4; §4 fixed targetFill → T4 Step 2; §5 helper rescales → T2; §6 profiles → T3/T5; §7 wording → out of build (noted). Testing § → T1–T3 node, T6 manual.
- Type consistency: `strengthToElongationShare`, `normalizeStrength`, `sliderToFill`, `strengthToKashidaLevel`, `kashidaExpansionFraction` signatures unchanged across tasks; φ formula identical in T1/T4/T5.
- Open (plan-decided): slider default **7**; `sliderToFill` rescaled (not folded) to keep the insert path working — insert-path fill-vs-φ unification is a noted future cleanup, not in scope.
