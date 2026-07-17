/* ConversionPane — wires the "Convert" tab to AshaarConversion (word-conversion.js).
 * Runs preset-driven find-replace over the document (or selection) in either
 * direction (legacy double-press ⇄ modern), preserving run formatting by using
 * the Office.js search/replace API. Named presets are stored at the roaming
 * (per-user, cross-document) level.
 * See docs/superpowers/specs/2026-07-16-text-conversion-design.md.
 */
(function () {
  "use strict";

  var bound = false;
  var PRESET_KEY = "ashaar-conversion-presets"; // { name: [enabledId, ...] }

  function byId(id) { return document.getElementById(id); }

  function setStatus(msg, warn) {
    var el = byId("conv-status");
    if (!el) return;
    el.textContent = msg || "";
    el.classList.toggle("warn", !!warn);
  }

  var CATEGORY_LABEL = { letter: "Letters", mark: "Marks", symbol: "Symbols / honorifics" };

  // ── checklist ──────────────────────────────────────────────────────────────

  function renderChecklist() {
    var host = byId("conv-checklist");
    if (!host || typeof AshaarConversion === "undefined") return;
    host.innerHTML = "";
    AshaarConversion.groupsForUi().forEach(function (group) {
      var fs = document.createElement("fieldset");
      fs.className = "conv-group";
      var lg = document.createElement("legend");
      var allId = "conv-group-" + group.category;
      var allCb = document.createElement("input");
      allCb.type = "checkbox"; allCb.id = allId; allCb.checked = true;
      allCb.className = "conv-group-all";
      allCb.addEventListener("change", function () {
        group.rows.forEach(function (r) {
          var cb = byId("conv-row-" + r.id);
          if (cb) cb.checked = allCb.checked;
        });
      });
      var allLbl = document.createElement("label");
      allLbl.setAttribute("for", allId);
      allLbl.textContent = CATEGORY_LABEL[group.category] || group.category;
      lg.appendChild(allCb); lg.appendChild(allLbl);
      fs.appendChild(lg);

      group.rows.forEach(function (row) {
        var wrap = document.createElement("div");
        wrap.className = "conv-row";
        var cb = document.createElement("input");
        cb.type = "checkbox"; cb.id = "conv-row-" + row.id;
        cb.checked = true; // all on by default
        cb.setAttribute("data-conv-id", row.id);
        var lbl = document.createElement("label");
        lbl.setAttribute("for", cb.id);
        lbl.textContent = row.legacy + " ⇄ " + row.modern + "  (" + row.label + ")";
        wrap.appendChild(cb); wrap.appendChild(lbl);
        if (row.lossy) {
          var warn = document.createElement("span");
          warn.className = "conv-lossy";
          warn.title = "Lossy: the source character has an independent meaning; " +
            "uncheck this row for documents that use it genuinely.";
          warn.textContent = " ⚠";
          wrap.appendChild(warn);
        }
        fs.appendChild(wrap);
      });
      host.appendChild(fs);
    });
  }

  function checkedIds() {
    var ids = [];
    var boxes = document.querySelectorAll("#conv-checklist input[data-conv-id]");
    for (var i = 0; i < boxes.length; i++) {
      if (boxes[i].checked) ids.push(boxes[i].getAttribute("data-conv-id"));
    }
    return ids;
  }

  function setCheckedIds(ids) {
    var want = {};
    (ids || []).forEach(function (i) { want[i] = true; });
    var boxes = document.querySelectorAll("#conv-checklist input[data-conv-id]");
    for (var i = 0; i < boxes.length; i++) {
      boxes[i].checked = !!want[boxes[i].getAttribute("data-conv-id")];
    }
  }

  // ── presets (roaming: per-user, cross-document) ──────────────────────────────

  function loadPresets() {
    try {
      var raw = Office.context.roamingSettings.get(PRESET_KEY);
      return raw ? JSON.parse(raw) : {};
    } catch (e) { return {}; }
  }

  function savePresets(presets, done) {
    Office.context.roamingSettings.set(PRESET_KEY, JSON.stringify(presets || {}));
    Office.context.roamingSettings.saveAsync(function () { if (done) done(); });
  }

  function allIds() {
    return AshaarConversion.MAPPINGS.map(function (m) { return m.id; });
  }

  function populatePresetPicker() {
    var sel = byId("conv-preset");
    if (!sel) return;
    sel.innerHTML = "";
    var opt0 = document.createElement("option");
    opt0.value = "__all__"; opt0.textContent = "All conversions";
    sel.appendChild(opt0);
    var presets = loadPresets();
    Object.keys(presets).sort().forEach(function (name) {
      var opt = document.createElement("option");
      opt.value = name; opt.textContent = name;
      sel.appendChild(opt);
    });
  }

  function applyPreset(name) {
    if (name === "__all__" || !name) { setCheckedIds(allIds()); return; }
    var presets = loadPresets();
    if (presets[name]) setCheckedIds(presets[name]);
  }

  function saveCurrentAsPreset() {
    var name = (byId("conv-saveas-name").value || "").trim();
    if (!name) { setStatus("Enter a preset name.", true); return; }
    var presets = loadPresets();
    presets[name] = checkedIds();
    savePresets(presets, function () {
      populatePresetPicker();
      byId("conv-preset").value = name;
      byId("conv-saveas-row").hidden = true;
      byId("conv-saveas-name").value = "";
      setStatus("Saved preset “" + name + "”.");
    });
  }

  // ── the conversion run ───────────────────────────────────────────────────────

  function run(direction) {
    if (typeof Word === "undefined" || !Word.run) {
      setStatus("Word is not available.", true); return;
    }
    var enabled = checkedIds();
    if (!enabled.length) { setStatus("No conversions selected.", true); return; }
    var ops = AshaarConversion.buildOperations(direction, enabled);
    var scopeSel = (document.querySelector('input[name="conv-scope"]:checked') || {}).value || "document";
    var counts = { letter: 0, mark: 0, symbol: 0 };
    setStatus("Converting…");

    Word.run(function (ctx) {
      var scope = scopeSel === "selection" ? ctx.document.getSelection() : ctx.document.body;
      var chain = ctx.sync();
      ops.forEach(function (op) {
        chain = chain.then(function () {
          var res = scope.search(op.find, { matchCase: true, matchWholeWord: !!op.wholeWord });
          res.load("items");
          return ctx.sync().then(function () {
            for (var i = 0; i < res.items.length; i++) {
              res.items[i].insertText(op.replaceWith, Word.InsertLocation.replace);
            }
            counts[op.category] = (counts[op.category] || 0) + res.items.length;
            return ctx.sync();
          });
        });
      });
      return chain;
    }).then(function () {
      setStatus("Converted " + counts.letter + " letters, " + counts.mark +
        " marks, " + counts.symbol + " symbols.");
    }).catch(function (e) {
      setStatus("Convert failed: " + (e && e.message ? e.message : String(e)), true);
    });
  }

  // ── binding ──────────────────────────────────────────────────────────────────

  function bind() {
    if (bound) return;
    var toModern = byId("conv-to-modern");
    var toLegacy = byId("conv-to-legacy");
    if (!toModern || !toLegacy) return; // markup not present yet
    bound = true;

    renderChecklist();
    populatePresetPicker();

    toModern.addEventListener("click", function () { run(AshaarConversion.DIRECTIONS.TO_MODERN); });
    toLegacy.addEventListener("click", function () { run(AshaarConversion.DIRECTIONS.TO_LEGACY); });

    var presetSel = byId("conv-preset");
    if (presetSel) presetSel.addEventListener("change", function (e) { applyPreset(e.target.value); });

    var saveAsBtn = byId("conv-preset-saveas");
    if (saveAsBtn) saveAsBtn.addEventListener("click", function () {
      byId("conv-saveas-row").hidden = false;
      byId("conv-saveas-name").focus();
    });
    var saveOk = byId("conv-saveas-ok");
    if (saveOk) saveOk.addEventListener("click", saveCurrentAsPreset);
    var saveCancel = byId("conv-saveas-cancel");
    if (saveCancel) saveCancel.addEventListener("click", function () {
      byId("conv-saveas-row").hidden = true;
      byId("conv-saveas-name").value = "";
    });
  }

  function onTabShown() {
    // Re-render the checklist in case the mapping table changed since bind
    // (e.g. generated symbol rows loaded), and refresh the preset picker.
    if (!bound) { bind(); return; }
    renderChecklist();
    populatePresetPicker();
  }

  window.ConversionPane = { bind: bind, onTabShown: onTabShown };
}());
