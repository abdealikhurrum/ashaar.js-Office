(function () {
  var input = document.getElementById("poem-input");
  var preview = document.getElementById("preview");
  var message = document.getElementById("message");
  var hostStatus = document.getElementById("host-status");
  var modeTable = document.getElementById("mode-table");
  var modeConvert = document.getElementById("mode-convert");
  var tablePanel = document.getElementById("table-mode-panel");
  var convertPanel = document.getElementById("convert-mode-panel");
  var justifyMode = document.getElementById("justify-mode");
  var justifyFillMode = document.getElementById("justify-fill-mode");
  var layoutMode = document.getElementById("layout-mode");
  var widthMode = document.getElementById("width-mode");
  var bandhCount = document.getElementById("bandh-count");
  var misraCount = document.getElementById("misra-count");
  var layoutPreset = document.getElementById("layout-preset");
  var layoutSpec = document.getElementById("layout-spec");
  var fontMode = document.getElementById("font-mode");
  var tatweelCount = document.getElementById("tatweel-count");
  var tatweelValue = document.getElementById("tatweel-value");
  var gapWidth = document.getElementById("gap-width");
  var tableWidth = document.getElementById("table-width");
  var tableWidthValue = document.getElementById("table-width-value");
  var autoFitWidth = document.getElementById("auto-fit-width");

  // Table width as a fraction of the page text column (centred). 100% = full width.
  function tableWidthPct() {
    return Math.max(25, Math.min(100, Number((tableWidth && tableWidth.value) || 50)));
  }
  function scaledTextWidth(twips) {
    return Math.max(1, Math.round(twips * tableWidthPct() / 100));
  }

  // Smallest table width (twips) that fits the widest misra in each column at the
  // given font (ctx), leaving kashida headroom. Drives Auto-fit and the width nudge.
  function neededTableTwips(source, ctx, opts, pageTwips) {
    if (!ctx || typeof Ashaar === "undefined" || !Ashaar.parse) return pageTwips;
    var maxByPos = [];
    (Ashaar.parse(source) || []).forEach(function (p) {
      (p.stanzas || []).forEach(function (s) {
        (s.bayts || []).forEach(function (b) {
          var texts = (b.type === "row" && b.misras) ? b.misras.map(function (m) { return m.text; })
                    : (b.ajuz ? [b.sadr, b.ajuz] : [b.sadr]);
          texts.forEach(function (t, i) {
            var w = ctx.measureText(stripJustification(String(t || ""))).width;
            if (!(maxByPos[i] >= w)) maxByPos[i] = w;
          });
        });
      });
    });
    if (!maxByPos.length) return pageTwips;
    var kashidaOn = (opts.justifyMode === "kashida" || opts.justifyMode === "spacing") && Number(opts.tatweelCount || 0) > 0;
    var headroom = kashidaOn ? 0.9 : 0.98;        // leave room for kashida to fill
    var N = Math.max(maxByPos.length, 2);
    var gapCols = Math.max(1, Math.round(Number(opts.gapWidth || 1)));
    // 6 must mirror BASE_CPM in word-html.js (grid columns per misra) so the
    // gap-to-content ratio here matches the actual render; otherwise the needed
    // width (which now floors the table width) is overestimated and tables come
    // out too wide.
    var GRID = N * 6 + (N - 1) * gapCols, contentCols = N * 6;
    var sumContent = maxByPos.reduce(function (a, b) { return a + b; }, 0) / headroom;
    return Math.round(sumContent * GRID / contentCols * 1440 / 96);
  }
  var templateNameInput = document.getElementById("template-name");
  var templateList = document.getElementById("template-list");
  var importFileInput = document.getElementById("import-file");
  var sepMode = document.getElementById("sep-mode");
  var sepCustom = document.getElementById("sep-custom");
  var sepPair = document.getElementById("sep-pair");
  var layoutGridEl = document.getElementById("layout-grid");
  var layoutViewGridBtn = document.getElementById("layout-view-grid");
  var layoutViewNumbersBtn = document.getElementById("layout-view-numbers");
  var debugMode = document.getElementById("debug-mode");
  var debugOutput = document.getElementById("debug-output");
  var qaseedaName = document.getElementById("qaseeda-name");
  var qaseedaNames = document.getElementById("qaseeda-names");
  var qaseedaWidthMode = document.getElementById("qaseeda-width-mode");
  var qaseedaWidthPct = document.getElementById("qaseeda-width-pct");
  var qaseedaJustifyMode = document.getElementById("qaseeda-justify-mode");
  var qaseedaFillMode = document.getElementById("qaseeda-fill-mode");
  var qaseedaStrength = document.getElementById("qaseeda-strength");
  var qaseedaStrengthValue = document.getElementById("qaseeda-strength-value");
  var qaseedaCorrFont = document.getElementById("qaseeda-corr-font");
  var qaseedaCorrFactor = document.getElementById("qaseeda-corr-factor");
  var qaseedaDebugTatweel = document.getElementById("qaseeda-debug-tatweel");
  var qaseedaDebugTatweelOn = document.getElementById("qaseeda-debug-tatweel-on");
  var qaseedaDebugSpace = document.getElementById("qaseeda-debug-space");
  var qaseedaDebugSpaceOn = document.getElementById("qaseeda-debug-space-on");
  var qaseedaFontStatus = document.getElementById("qaseeda-font-status");

  // Format collected per-cell justification metrics into the Debug panel.
  function renderDebug(diags) {
    if (!debugOutput) return;
    if (!diags.length) { debugOutput.textContent = "(no kashida cells measured)"; return; }
    var head = "cell  font        res  col(in)   nat  target  final  fill  tw/cap  text";
    var rows = diags.map(function (d) {
      return [
        String(d.i).padEnd(4),
        d.font.padEnd(11),
        String(d.res || "?").padStart(3),
        (d.colPx + "(" + d.colIn + ")").padEnd(9),
        String(d.nat).padStart(4),
        String(d.target).padStart(6),
        String(d.fin).padStart(6),
        (d.fill + "%").padStart(5),
        (d.tw + "/" + d.cap).padStart(6),
        "  " + d.text
      ].join(" ");
    });
    debugOutput.textContent = head + "\n" + rows.join("\n");
  }

  // Whether the WebView can actually render `name` vs silently falling back.
  // document.fonts.check is unreliable for system fonts (true for unknown names),
  // so compare measured widths against generic families — if the font changes the
  // width over a generic baseline, it resolved; if it matches all three, it fell back.
  var _fontProbeCtx = null;
  function fontAvailable(name) {
    if (!name) return false;
    try {
      if (!_fontProbeCtx) _fontProbeCtx = document.createElement("canvas").getContext("2d");
      var ctx = _fontProbeCtx;
      if (!ctx) return true;
      var test = "mMgwiحيبٹكطولِ ظہور";
      var generics = ["monospace", "sans-serif", "serif"];
      for (var i = 0; i < generics.length; i++) {
        ctx.font = "72px " + generics[i];
        var base = ctx.measureText(test).width;
        ctx.font = "72px '" + name + "'," + generics[i];
        if (Math.abs(ctx.measureText(test).width - base) > 0.5) return true;
      }
      return false;
    } catch (e) { return true; }
  }

  var GRID_COLS = 12;
  var layoutView = "numbers";   // "grid" | "numbers"
  var gridMatrix = [];          // rows of 12 booleans, reading order (index 0 = visual right)

  var SEP_LABELS = {
    backslash: "\\", asterisk: "*", pipe: "|", dash: "dash",
    tab: "tab", spaces: "double space", custom: "custom", pairLines: "paired lines"
  };

  // Apply separator normalization to the current editor text using the import
  // options. Auto-detects by default; explicit choices and pair-lines override.
  // Re-renders the preview and notes what changed.
  function applyImportNormalization() {
    if (typeof AshaarSeparators === "undefined") { renderPreview(); return; }
    var res = AshaarSeparators.normalizeSeparators(input.value, {
      separator: sepMode ? sepMode.value : "auto",
      customPattern: sepCustom ? sepCustom.value : "",
      pairLines: sepPair ? sepPair.checked : false
    });
    if (res.changed) {
      input.value = res.text;
      setMessage("Converted separators (" + (SEP_LABELS[res.detected] || res.detected) + ") to standard \\ form.");
    }
    renderPreview();
  }

  // ── Visual layout grid (Grid mode) ──────────────────────────────────────────

  // Default row: a paired couplet — 5 on / 2 gap / 5 on (reading order).
  function defaultGridRow() {
    var r = [];
    for (var i = 0; i < 5; i++) r.push(true);
    for (i = 0; i < 2; i++) r.push(false);
    for (i = 0; i < 5; i++) r.push(true);
    return r;
  }

  function renderLayoutGrid() {
    if (!layoutGridEl || typeof AshaarLayoutGrid === "undefined") return;
    var html = "";
    gridMatrix.forEach(function (row, ri) {
      html += '<div class="lg-row">';
      html += '<button type="button" class="lg-rm" data-row="' + ri + '" title="Remove row" aria-label="Remove row">✕</button>';
      html += '<span class="lg-bubbles">';
      for (var c = 0; c < GRID_COLS; c++) {
        html += '<span class="lg-b' + (row[c] ? " on" : "") + '" data-row="' + ri + '" data-col="' + c + '"></span>';
      }
      html += "</span></div>";
    });
    html += '<button type="button" class="lg-add" id="lg-add-row">+ Add row</button>';

    var tpl = AshaarLayoutGrid.gridToTemplate(gridMatrix);
    html += '<div class="lg-prev"><span class="lg-prev-label">Preview</span>';
    tpl.rows.forEach(function (cells) {
      html += '<div class="lg-prev-row">';
      cells.forEach(function (cell) {
        if (cell.role === "gap") html += '<span class="lg-prev-gap" style="flex:' + cell.span + '"></span>';
        else html += '<span class="lg-prev-cell" style="flex:' + cell.span + '">·</span>';
      });
      html += "</div>";
    });
    html += "</div>";
    layoutGridEl.innerHTML = html;
  }

  function onLayoutGridClick(e) {
    var t = e.target;
    if (!t) return;
    if (t.id === "lg-add-row") { gridMatrix.push(defaultGridRow()); renderLayoutGrid(); return; }
    if (t.classList.contains("lg-rm")) {
      var ri = Number(t.getAttribute("data-row"));
      if (gridMatrix.length > 1) gridMatrix.splice(ri, 1); else gridMatrix = [defaultGridRow()];
      renderLayoutGrid();
      return;
    }
    if (t.classList.contains("lg-b")) {
      var r = Number(t.getAttribute("data-row"));
      var c = Number(t.getAttribute("data-col"));
      if (gridMatrix[r]) { gridMatrix[r][c] = !gridMatrix[r][c]; renderLayoutGrid(); }
    }
  }

  // Toggle between the visual Grid and the Numbers (text) view of the layout spec.
  // The two are kept in sync best-effort: Numbers→Grid parses the spec into bubbles;
  // Grid→Numbers serializes the bubbles back to the text spec.
  function setLayoutView(view) {
    if (typeof AshaarLayoutGrid === "undefined") return;
    layoutView = view === "grid" ? "grid" : "numbers";
    var grid = layoutView === "grid";
    layoutViewGridBtn.classList.toggle("is-active", grid);
    layoutViewNumbersBtn.classList.toggle("is-active", !grid);
    layoutViewGridBtn.setAttribute("aria-pressed", String(grid));
    layoutViewNumbersBtn.setAttribute("aria-pressed", String(!grid));
    layoutGridEl.hidden = !grid;
    layoutSpec.hidden = grid;
    if (grid) {
      var m = AshaarLayoutGrid.specToGrid(layoutSpec.value);
      gridMatrix = m.length ? m : [defaultGridRow()];
      renderLayoutGrid();
    } else {
      layoutSpec.value = AshaarLayoutGrid.gridToSpec(gridMatrix);
      renderPreview();
    }
  }

  function options() {
    return {
      justifyMode: justifyMode.value,
      justify: justifyMode.value === "none" ? false : justifyMode.value,
      fillMode: (justifyFillMode && justifyFillMode.value) || "natural-fit",
      layoutMode: layoutMode.value,
      layout: layoutMode.value,
      widthMode: widthMode.value,
      bandhCount: Number(bandhCount.value || 1),
      misraCount: Number(misraCount.value || 4),
      misraPattern: layoutPreset.value,
      layoutSpec: (!tablePanel.hidden) ? layoutSpec.value : "",
      fontMode: fontMode.value,
      tatweelCount: Number(tatweelCount.value || 0),
      gapWidth: Number(gapWidth.value || 4),
      tableWidthPct: tableWidthPct(),
      autoFitWidth: !!(autoFitWidth && autoFitWidth.checked),
      qaseeda: (qaseedaName && qaseedaName.value ? qaseedaName.value.trim() : "")
    };
  }

  function previewFontFamily(font) {
    var mode = font === "nastaliq" ? "noto" : font;
    var css = AshaarFonts.cssFamilyOf(mode);
    return css || "\"Times New Roman\", serif";
  }

  function updateFontNote() {
    var d = AshaarFonts.get(fontMode.value === "nastaliq" ? "noto" : fontMode.value);
    var note = document.getElementById("font-install-note");
    if (!note) return;
    if (d && d.readerNote) {
      note.hidden = false;
      note.textContent = (d.id === "jameel")
        ? "Readers need “Jameel Noori Nastaleeq” (Regular + Kasheeda) installed to see this correctly."
        : "Readers need “" + d.wordName + "” installed to see this correctly.";
    } else { note.hidden = true; }
  }

  function setMessage(text) {
    message.textContent = text || "";
  }

  // Office.js collapses many failures into the opaque string "GeneralException".
  // The useful detail (which API statement threw) lives in error.debugInfo. This
  // surfaces code + errorLocation to the pane and dumps the full error + debugInfo
  // to the WebView console (open devtools to see the stack / surrounding statements).
  function describeError(error) {
    try { console.error("[ashaar] error:", error, error && error.debugInfo); } catch (e) {}
    if (!error) return "Unknown error.";
    var di = error.debugInfo || {};
    var parts = [];
    var head = error.message || error.code || String(error);
    parts.push(head);
    if (di.code && di.code !== error.message) parts.push("code=" + di.code);
    if (di.errorLocation) parts.push("at " + di.errorLocation);
    // surroundingStatements pinpoints the API calls around the throw when
    // errorLocation is an internal resolver (e.g. _GetObjectByReferenceId).
    var ss = di.surroundingStatements;
    if (ss && ss.length) parts.push("near: " + ss.join(" | "));
    return parts.join(" — ");
  }

  function setMode(mode) {
    var isTable = mode === "table";
    modeTable.classList.toggle("is-active", isTable);
    modeConvert.classList.toggle("is-active", !isTable);
    modeTable.setAttribute("aria-selected", String(isTable));
    modeConvert.setAttribute("aria-selected", String(!isTable));
    tablePanel.classList.toggle("is-active", isTable);
    convertPanel.classList.toggle("is-active", !isTable);
    tablePanel.hidden = !isTable;
    convertPanel.hidden = isTable;
    setMessage(isTable ? "Table input mode: draw a blank grid, then type in Word." : "Ashaar.js conversion mode: paste source text, then insert a converted table.");
  }

  function renderPreview() {
    var opts = options();
    tatweelValue.textContent = String(opts.tatweelCount);
    if (tableWidthValue) tableWidthValue.textContent = String(opts.tableWidthPct);
    preview.className = "ashaar preview";
    // Mirror the chosen table width: a narrower, centred preview previews the insert.
    preview.style.maxWidth = opts.autoFitWidth ? "100%" : (opts.tableWidthPct + "%");
    preview.style.marginInline = "auto";
    preview.style.setProperty("--ashaar-font-family", previewFontFamily(opts.fontMode));
    preview.innerHTML = Ashaar.renderText(String(input.value || ""), { gapWidth: opts.gapWidth + "%" });
    Ashaar.applyRenderOptions(preview, { gapWidth: opts.gapWidth + "%" });
    if (opts.layout === "stacked") preview.classList.add("ashaar--stacked");
    if (opts.layout === "auto" || opts.layout === "compact") Ashaar.applyAutoLayout(preview, { layout: "auto" });
    if (opts.justify === "css") {
      preview.classList.add("ashaar--justify");
    } else if (opts.justify === "spacing") {
      Ashaar.justifyEl(preview, { method: "spacing", tatweel: false });
    } else if (opts.justify === "kashida") {
      // tatweelCount=0 suppresses tatweels; otherwise justifyEl fills to available width.
      // Fine-grained strength is applied on OOXML insertion (not controllable here without
      // modifying the vendor ashaar.js justifyMisra function).
      Ashaar.justifyEl(preview, opts.tatweelCount === 0 ? { tatweel: false } : {});
    }
  }

  function layoutSpecForPreset(preset, count) {
    count = Math.max(1, Number(count || 4));
    var rows = [];
    var i;
    if (preset === "centered-stack") {
      for (i = 1; i <= count; i++) rows.push("<" + i + ">");
      return rows.join("\n");
    }
    if (preset === "alternate-right") {
      for (i = 1; i <= count; i++) rows.push(i % 2 ? i + " >" : "< " + i);
      return rows.join("\n");
    }
    if (preset === "indented-stack") {
      for (i = 1; i <= count; i++) rows.push(new Array(i).join("  ") + i);
      return rows.join("\n");
    }
    if (preset === "karbala-refrain") {
      return "3 | 2 | 1\n<4>\n6 - 5";
    }
    for (i = 1; i <= count; i += 2) rows.push((i + 1 <= count ? i + 1 : "") + " - " + i);
    return rows.join("\n");
  }

  function applyLayoutPreset() {
    layoutSpec.value = layoutSpecForPreset(layoutPreset.value, misraCount.value);
  }

  async function withWord(callback) {
    if (typeof Word === "undefined") {
      setMessage("Open this task pane inside Word to update the document.");
      return;
    }
    try {
      await Word.run(callback);
      setMessage("Done.");
    } catch (error) {
      setMessage(describeError(error));
    }
  }

  // ── Qaseeda profiles — document store + block tagging (P2) ────────────────
  // Profiles live in Word document settings (one authoritative copy that travels
  // with the .docx), keyed by name. A block is linked to a qaseeda by the
  // `qaseeda` field in its content-control tag. Pure profile math is in
  // profiles.js (AshaarProfiles); this layer is the Office.js orchestration.
  var PROFILE_STORE_KEY = "ashaar:profiles";

  function loadProfileStore() {
    try {
      if (typeof Office === "undefined" || !Office.context || !Office.context.document) return {};
      var raw = Office.context.document.settings.get(PROFILE_STORE_KEY);
      var obj = raw ? (typeof raw === "string" ? JSON.parse(raw) : raw) : {};
      return (obj && typeof obj === "object") ? obj : {};
    } catch (e) { return {}; }
  }

  function saveProfileStore(store) {
    return new Promise(function (resolve) {
      try {
        if (typeof Office === "undefined" || !Office.context || !Office.context.document) { resolve(false); return; }
        Office.context.document.settings.set(PROFILE_STORE_KEY, JSON.stringify(store || {}));
        Office.context.document.settings.saveAsync(function (res) {
          resolve(!!(res && res.status === Office.AsyncResultStatus.Succeeded));
        });
      } catch (e) { resolve(false); }
    });
  }

  function getProfile(name) {
    var p = loadProfileStore()[name];
    return p ? AshaarProfiles.normalizeProfile(p) : AshaarProfiles.defaultProfile(name);
  }

  async function putProfile(profile) {
    var p = AshaarProfiles.normalizeProfile(profile);
    if (!p.name) return false;
    var store = loadProfileStore();
    store[p.name] = p;
    return await saveProfileStore(store);
  }

  function listProfileNames() {
    return Object.keys(loadProfileStore());
  }

  // Assign (or clear, with "") a qaseeda name onto the Ashaar Poem block at the cursor.
  async function setQaseedaOnSelection(name) {
    var done = false;
    await withWord(async function (context) {
      var cc = context.document.getSelection().parentContentControlOrNullObject;
      cc.load("title,tag");
      await context.sync();
      if (cc.isNullObject || cc.title !== "Ashaar Poem") {
        setMessage("Place the cursor inside an Ashaar Poem block first.");
        return;
      }
      cc.tag = AshaarWord.setTagQaseeda(cc.tag, name);
      await context.sync();
      done = true;
    });
    return done;
  }

  // Read the qaseeda name on the Ashaar Poem block at the cursor ("" if none).
  async function getQaseedaAtSelection() {
    var qaseeda = "";
    if (typeof Word === "undefined") return qaseeda;
    try {
      await Word.run(async function (context) {
        var cc = context.document.getSelection().parentContentControlOrNullObject;
        cc.load("title,tag");
        await context.sync();
        if (!cc.isNullObject && cc.title === "Ashaar Poem") {
          var payload = AshaarWord.parseContentControlTag(cc.tag);
          qaseeda = (payload && payload.qaseeda) || "";
        }
      });
    } catch (e) { /* leave qaseeda empty */ }
    return qaseeda;
  }

  // Read-only: show the bandh cell-map (labels + gaps) for the Ashaar Poem block
  // at the cursor. No document mutation. Labels can't be shown on the Word page
  // itself (no native per-cell text overlay), so the pane is their home.
  async function showCellMap() {
    var view = document.getElementById("cell-map-view");
    if (!view) return;
    if (typeof Word === "undefined") { setMessage("Open this task pane inside Word."); return; }
    var patterns = null;
    try {
      await Word.run(async function (context) {
        var cc = context.document.getSelection().parentContentControlOrNullObject;
        cc.load("title,tag");
        await context.sync();
        if (!cc.isNullObject && cc.title === "Ashaar Poem") {
          var payload = AshaarWord.parseContentControlTag(cc.tag);
          patterns = payload && payload.cells;
        }
      });
    } catch (e) { /* leave patterns null */ }

    if (!patterns || !patterns.length) {
      view.hidden = false;
      view.textContent = "No cell map on the block at the cursor (older or hand-drawn table).";
      return;
    }
    var html = "";
    patterns.forEach(function (pattern, bi) {
      var map = AshaarCellMap.buildBandhCellMap(pattern);
      html += "<div class=\"cell-map-bandh\"><b>Bandh " + (bi + 1) + "</b>";
      var lastRow = -1, rowHtml = "";
      function flush() { if (rowHtml) html += "<div class=\"cell-map-row\">" + rowHtml + "</div>"; rowHtml = ""; }
      map.forEach(function (e) {
        if (e.row !== lastRow) { flush(); lastRow = e.row; }
        rowHtml += e.kind === "content"
          ? "<span class=\"cell-map-cell\">" + e.label + "</span>"
          : "<span class=\"cell-map-gap\">(gap)</span>";
      });
      flush();
      html += "</div>";
    });
    view.hidden = false;
    view.innerHTML = html;
  }

  // ── Active-block / active-cell reflection (SP2) ───────────────────────────
  var _lastBlockTag = null;          // last-reflected block tag (resync only on change)
  var _reflectPending = false;       // debounce guard
  var _reflectBusy = false;          // suppress reflection while our own justify runs
  var _activeOvKey = null;           // override key of the content cell at the cursor (or null)
  var _activeDecorKey = null;        // slot-decor key of the spacing cell at the cursor (or null)
  var _activeSlot = null;            // slot-position (e.g. "A#1") of the spacing cell at the cursor

  // Populate the pane's block-level controls from a parsed tag payload.
  function syncBlockControls(payload) {
    if (!payload) return;
    if (payload.fontMode) fontMode.value = payload.fontMode;
    if (payload.justifyMode) justifyMode.value = payload.justifyMode;
    if (justifyFillMode && payload.fillMode) justifyFillMode.value = AshaarProfiles.normalizeFillMode(payload.fillMode);
    if (payload.tatweelCount != null) { tatweelCount.value = payload.tatweelCount; if (tatweelValue) tatweelValue.textContent = String(payload.tatweelCount); }
    if (payload.tableWidthPct != null && tableWidth) { tableWidth.value = payload.tableWidthPct; if (tableWidthValue) tableWidthValue.textContent = String(payload.tableWidthPct); }
    if (payload.qaseeda) { var st = loadProfileStore()[payload.qaseeda]; if (st) profileToPanel(st); }
  }

  // Reflect the Ashaar block (and the cell, Task 4) at the cursor in the pane.
  async function reflectActiveContext() {
    if (typeof Word === "undefined" || _reflectBusy) return;
    try {
      await Word.run(async function (context) {
        var sel = context.document.getSelection();
        var cc = sel.parentContentControlOrNullObject;
        cc.load("title,tag");
        await context.sync();
        var isBlock = !cc.isNullObject && cc.title === "Ashaar Poem";
        var payload = isBlock ? AshaarWord.parseContentControlTag(cc.tag) : null;
        // Resync block-level controls only when the active block changes, so
        // moving the cursor within a block never clobbers a mid-edit control.
        if (isBlock && cc.tag !== _lastBlockTag) { _lastBlockTag = cc.tag; syncBlockControls(payload); }
        if (!isBlock) { _lastBlockTag = null; }
        await reflectActiveCell(context, sel, cc, isBlock, payload);
      });
    } catch (e) { /* selection transient — ignore */ }
  }

  // Show/populate the per-cell override editor for the content cell at the
  // cursor. Resolves (tableIndex, label) via the SP1 cells map. Hides the editor
  // for gaps, non-content, or maps-absent blocks.
  async function reflectActiveCell(context, sel, cc, isBlock, payload) {
    var editor = document.getElementById("cell-override");
    _activeOvKey = null;
    if (!editor) return;
    if (!isBlock || !payload || !payload.cells) { editor.hidden = true; return; }

    var tcell = sel.parentTableCellOrNullObject;
    tcell.load("rowIndex,cellIndex,isNullObject");
    var tbls = cc.getRange().tables;
    tbls.load("items");
    await context.sync();
    if (tcell.isNullObject) { editor.hidden = true; return; }

    // §6a: which block table contains the selection? (No stable table id, so
    // match by range intersection.)
    var selRange = sel.getRange();
    var inters = tbls.items.map(function (tbl) {
      var r = tbl.getRange().intersectWithOrNullObject(selRange); r.load("isNullObject"); return r;
    });
    await context.sync();
    var tIdx = -1;
    for (var k = 0; k < inters.length; k++) { if (!inters[k].isNullObject) { tIdx = k; break; } }
    if (tIdx < 0 || !payload.cells[tIdx]) { editor.hidden = true; return; }

    var map = AshaarCellMap.buildBandhCellMap(payload.cells[tIdx]);
    var inRow = map.filter(function (e) { return e.row === tcell.rowIndex; });
    var entry = inRow[tcell.cellIndex];
    var decorEl = document.getElementById("slot-decor");
    if (!entry) { editor.hidden = true; if (decorEl) decorEl.hidden = true; _activeOvKey = null; _activeDecorKey = null; return; }
    if (entry.kind === "content") {
      if (decorEl) decorEl.hidden = true;
      _activeDecorKey = null;
      _activeOvKey = AshaarOverrides.overrideKey(tIdx, entry.label);
      populateCellEditor(entry.label, (payload.overrides || {})[_activeOvKey]);
      editor.hidden = false;
    } else { // spacing → decoration editor
      editor.hidden = true;
      _activeOvKey = null;
      _activeDecorKey = AshaarOverrides.overrideKey(tIdx, entry.slot);
      _activeSlot = entry.slot;
      populateDecorEditor(entry.slot, (payload.slotDecor || {})[_activeDecorKey]);
      if (decorEl) decorEl.hidden = false;
    }
  }

  function populateCellEditor(label, ov) {
    ov = ov || {};
    var lbl = document.getElementById("cell-override-label");
    if (lbl) lbl.textContent = label || "";
    document.getElementById("cell-ov-strength").value = (ov.strength != null) ? ov.strength : "";
    document.getElementById("cell-ov-width").value = (ov.widthPt != null) ? ov.widthPt : "";
    document.getElementById("cell-ov-cap").value = (ov.capEm != null) ? ov.capEm : "";
  }

  function readCellEditor() {
    function num(id) { var v = document.getElementById(id).value; return v === "" ? null : Number(v); }
    var ov = {};
    var s = num("cell-ov-strength"); if (s != null) ov.strength = Math.max(1, Math.min(10, s));
    var w = num("cell-ov-width"); if (w != null) ov.widthPt = w;
    var c = num("cell-ov-cap"); if (c != null) ov.capEm = c;
    return ov;
  }

  // Write the editor state to the active cell's override on the block tag, then
  // re-justify the whole block (via the existing path) for instant feedback.
  async function applyCellOverride(clear) {
    if (!_activeOvKey || typeof Word === "undefined") return;
    var ov = clear ? null : readCellEditor();
    _reflectBusy = true;
    try {
      await Word.run(async function (context) {
        var cc = context.document.getSelection().parentContentControlOrNullObject;
        cc.load("title,tag");
        await context.sync();
        if (cc.isNullObject || cc.title !== "Ashaar Poem") return;
        cc.tag = AshaarWord.setTagOverride(cc.tag, _activeOvKey, ov);
        await context.sync();
        _lastBlockTag = cc.tag;
      });
    } catch (e) { /* ignore */ } finally { _reflectBusy = false; }
    if (clear) {
      var lbl = document.getElementById("cell-override-label");
      populateCellEditor(lbl ? lbl.textContent : "", null);
    }
    await justifySelection(); // instant feedback via the existing path
  }

  function hexToWord(v) { return (v || "").replace(/^#/, ""); }
  function populateDecorEditor(slot, d) {
    d = d || {};
    var lbl = document.getElementById("slot-decor-label"); if (lbl) lbl.textContent = slot || "";
    document.getElementById("slot-decor-symbol").value = d.symbol || "";
    document.getElementById("slot-decor-fill-on").checked = !!d.fill;
    if (d.fill) document.getElementById("slot-decor-fill").value = "#" + d.fill;
    if (d.color) document.getElementById("slot-decor-color").value = "#" + d.color;
  }
  function readDecorEditor() {
    var d = {};
    var sym = document.getElementById("slot-decor-symbol").value;
    if (sym) d.symbol = sym;
    if (document.getElementById("slot-decor-fill-on").checked) d.fill = hexToWord(document.getElementById("slot-decor-fill").value);
    if (sym) d.color = hexToWord(document.getElementById("slot-decor-color").value);
    return d;
  }
  // Write the editor state to the active gap's per-slot decoration on the block
  // tag, then re-decorate via the qaseeda apply pass (needs a saved qaseeda).
  async function applySlotDecor(clear) {
    if (!_activeDecorKey || typeof Word === "undefined") return;
    var d = clear ? null : readDecorEditor();
    _reflectBusy = true;
    var qname = "";
    try {
      await Word.run(async function (context) {
        var cc = context.document.getSelection().parentContentControlOrNullObject;
        cc.load("title,tag");
        await context.sync();
        if (cc.isNullObject || cc.title !== "Ashaar Poem") return;
        cc.tag = AshaarWord.setTagSlotDecor(cc.tag, _activeDecorKey, d);
        await context.sync();
        _lastBlockTag = cc.tag;
        qname = (AshaarWord.parseContentControlTag(cc.tag) || {}).qaseeda || "";
      });
    } catch (e) { /* ignore */ } finally { _reflectBusy = false; }
    if (clear) populateDecorEditor(document.getElementById("slot-decor-label").textContent, null);
    if (qname && loadProfileStore()[qname]) await applyProfileToQaseeda(qname);
    else setMessage("Gap decoration saved — apply a qaseeda to this block to render it.");
  }

  // Save the current decor editor values as the qaseeda profile default for this
  // slot-position (e.g. every bandh's A#1), then re-apply across all blocks.
  async function saveSlotDecorToProfile() {
    if (!_activeSlot || typeof Word === "undefined") return;
    var qname = "";
    try {
      await Word.run(async function (context) {
        var cc = context.document.getSelection().parentContentControlOrNullObject;
        cc.load("title,tag");
        await context.sync();
        if (!cc.isNullObject && cc.title === "Ashaar Poem") {
          qname = (AshaarWord.parseContentControlTag(cc.tag) || {}).qaseeda || "";
        }
      });
    } catch (e) { /* ignore */ }
    if (!qname || !loadProfileStore()[qname]) { setMessage("Assign this block to a saved qaseeda first to set a profile-wide default."); return; }
    var profile = getProfile(qname);
    profile.spacingDecor = profile.spacingDecor || {};
    var d = readDecorEditor();
    if (d.symbol || d.fill || d.color) profile.spacingDecor[_activeSlot] = d;
    else delete profile.spacingDecor[_activeSlot];
    await putProfile(profile);
    await applyProfileToQaseeda(qname);
  }

  // Debounced entry point for the DocumentSelectionChanged event.
  function onSelectionChanged() {
    if (_reflectPending) return;
    _reflectPending = true;
    window.setTimeout(function () { _reflectPending = false; reflectActiveContext(); }, 150);
  }

  // ── Qaseeda apply/refresh engine (P3) ─────────────────────────────────────
  // Find every Ashaar Poem block linked to a qaseeda name.
  async function gatherQaseedaBlocks(context, name) {
    var ccs = context.document.contentControls;
    ccs.load("items/title,items/tag");
    await context.sync();
    return ccs.items.filter(function (cc) {
      if (cc.title !== "Ashaar Poem") return false;
      var p = AshaarWord.parseContentControlTag(cc.tag);
      return !!(name && p && p.qaseeda === name);
    });
  }

  // Per-qaseeda cache (this pane session) of the width signature last applied.
  // Rebuilding a block's OOXML is destructive (delete + re-insert its content
  // control) and heavy, so we only do it when the target WIDTH changes. Justify-
  // only changes (strength, fill mode, per-cell override) keep the same signature
  // → skip the rebuild and just re-justify the existing (already-sized) tables.
  var _appliedSizeSig = {};

  // Capture a qaseeda's blocks as measured tables ready for width sizing + justify.
  // Loads each block's tables/rows/cells (text + real font), captures plain values
  // (later edits invalidate proxies, so we never re-read body.*), reconstructs each
  // block's SOURCE, derives each content cell's grid span from that source (Word
  // can't report span-table column geometry), and measures natural (tatweel-free)
  // widths on a canvas. Geometry uses each block's OWN stored opts (from its tag)
  // so it matches how the block was/will be rendered. Returns everything both the
  // rebuild and justify passes need.
  async function captureQaseedaTables(context, blocks, profile) {
    var fallbackName = profile.font || "Times New Roman";

    var section = context.document.sections.getFirst();
    section.load("pageLayout/width,pageLayout/leftMargin,pageLayout/rightMargin");

    var blockTables = blocks.map(function (cc) { var t = cc.getRange().tables; t.load("items"); return t; });
    await context.sync();
    blockTables.forEach(function (t) { t.items.forEach(function (tbl) { tbl.rows.load("items"); }); });
    await context.sync();
    blockTables.forEach(function (t) { t.items.forEach(function (tbl) { tbl.rows.items.forEach(function (row) { row.cells.load("items"); }); }); });
    await context.sync();
    blockTables.forEach(function (t) { t.items.forEach(function (tbl) { tbl.rows.items.forEach(function (row) {
      row.cells.items.forEach(function (cell) {
        cell.body.load("text"); cell.body.font.load("name,size");
        // First-paragraph alignment + indent — the run-aware justify pass rebuilds
        // each cell via OOXML and must re-assert its visual side (sadr/ajuz/solo)
        // and any stacked-layout indent.
        cell.body.paragraphs.load("alignment,leftIndent");
      });
    }); }); });
    await context.sync();

    var pl = section.pageLayout;
    var pagePt = pl && pl.width ? (pl.width - (pl.leftMargin || 0) - (pl.rightMargin || 0)) : 468;
    var pageTwips = Math.round(pagePt * 20);

    var blockInfos = blocks.map(function (cc, bi) {
      var payload = AshaarWord.parseContentControlTag(cc.tag) || {};
      var pattern = payload.cells || null;
      var overrides = payload.overrides || {};
      var slotDecor = payload.slotDecor || {};
      var repFont = "", repSize = 0;
      var tableInfos = blockTables[bi].items.map(function (tbl, j) {
        var perRowCounts = tbl.rows.items.map(function (row) { return row.cells.items.length; });
        var tablePattern = pattern ? pattern[j] : null;
        var tblMap = AshaarCellMap.alignPatternToTable(perRowCounts, tablePattern)
          ? AshaarCellMap.buildBandhCellMap(tablePattern) : null;
        var seq = 0, cells = [], rowsText = [];
        tbl.rows.items.forEach(function (row, ri) {
          var cols = row.cells.items.length, rowText = [];
          row.cells.items.forEach(function (cell, ci) {
            var f = cell.body.font;
            var current = (cell.body.text || "").trim();
            var base = stripJustification(current);
            var mapped = tblMap ? tblMap[seq] : null;
            seq++;
            rowText.push(cell.body.text || "");
            if (f && f.name && !repFont) { repFont = f.name; if (f.size) repSize = f.size; }
            var p0 = cell.body.paragraphs.items && cell.body.paragraphs.items[0];
            var alv = p0 && p0.alignment;
            cells.push({
              cell: cell, current: current, base: base,
              measure: base.replace(/\s+/g, " ").trim(),
              matKey: mapped ? (mapped.label || mapped.slot) : AshaarMatrix.positionKey({ row: ri, col: ci, span: cols }),
              kind: mapped ? mapped.kind : null,
              slot: (mapped && mapped.kind === "spacing") ? mapped.slot : null,
              decorKey: (mapped && mapped.kind === "spacing" && mapped.slot) ? AshaarOverrides.overrideKey(j, mapped.slot) : null,
              ovKey: (mapped && mapped.kind === "content" && mapped.label) ? AshaarOverrides.overrideKey(j, mapped.label) : null,
              fontName: (f && f.name) || "", fontSize: (f && f.size) || 0,
              align: alv === "Right" ? "right" : alv === "Left" ? "left" : "center",
              indentTwips: (p0 && p0.leftIndent) ? Math.round(p0.leftIndent * 20) : 0
            });
          });
          rowsText.push(rowText);
        });
        return { tbl: tbl, cells: cells, rowsText: rowsText, overrides: overrides, slotDecor: slotDecor, blockIdx: j, grid: 0 };
      });
      var source = tableInfos.map(function (ti) {
        return AshaarTableAdopt.adoptTableToSource(ti.rowsText, { direction: "rtl" });
      }).filter(function (s) { return s.trim(); }).join("\n\n");
      return { cc: cc, oldTag: cc.tag, payload: payload, source: source, repFont: repFont, repSize: repSize, tableInfos: tableInfos };
    });

    // Representative font for the canvas baseline.
    var repName = fallbackName, repSize = 16;
    for (var b0 = 0; b0 < blockInfos.length && repName === fallbackName; b0++) {
      var tis = blockInfos[b0].tableInfos;
      for (var t0 = 0; t0 < tis.length && repName === fallbackName; t0++) {
        var cs0 = tis[t0].cells;
        for (var c0 = 0; c0 < cs0.length; c0++) {
          if (cs0[c0].fontName) { repName = cs0[c0].fontName; if (cs0[c0].fontSize) repSize = cs0[c0].fontSize; break; }
        }
      }
    }

    var canvasCtx = document.createElement("canvas").getContext("2d");
    if (canvasCtx) {
      canvasCtx.font = repSize + "pt \"" + repName + "\"";
      if (document.fonts && document.fonts.load) { try { await document.fonts.load(repSize + "pt \"" + repName + "\""); } catch (e) {} }
    }

    // Grid geometry per table from its OWN source + the block's stored opts (so it
    // matches the rendered table). Zips onto captured cells by emission order.
    blockInfos.forEach(function (blk) {
      var p = blk.payload;
      var geomOpts = { gapWidth: Number(p.gapWidth || 4), layoutMode: p.layoutMode || "balanced" };
      blk.tableInfos.forEach(function (info) {
        var flatGeo = [], grid = 0;
        try {
          var src = AshaarTableAdopt.adoptTableToSource(info.rowsText, { direction: "rtl" });
          AshaarWord.poemCellGeometry(src, geomOpts, Ashaar, pageTwips).forEach(function (st) {
            grid = Math.max(grid, st.GRID || 0);
            st.rows.forEach(function (row) { row.forEach(function (g) { flatGeo.push(g); }); });
          });
        } catch (e) { flatGeo = []; }
        if (flatGeo.length === info.cells.length) {
          info.cells.forEach(function (c, i) { c.gridCol = flatGeo[i].col; c.gridSpan = flatGeo[i].span; });
          info.grid = grid;
        } else {
          info.cells.forEach(function (c, i) { c.gridSpan = 1; c.gridCol = i; });
          info.grid = info.cells.length;
        }
      });
    });

    // Force-load EVERY distinct cell font (and its Kasheeda face, for font-swap)
    // before measuring — measureText silently substitutes an unloaded font and
    // returns wrong widths (mis-ranking swaps / mis-sizing the matrix). The rep
    // font alone isn't enough for multi-font poems or Jameel's wide face. See
    // memory font-measurement-model.
    if (canvasCtx && typeof document !== "undefined" && document.fonts && document.fonts.load) {
      var faceSet = {};
      blockInfos.forEach(function (blk) {
        blk.tableInfos.forEach(function (info) {
          info.cells.forEach(function (c) {
            var fnm = c.fontName || repName, fsz = c.fontSize || repSize;
            if (fnm) faceSet[fsz + "pt \"" + fnm + "\""] = true;
            var kn = (typeof AshaarFonts !== "undefined" && AshaarFonts.descriptorForFontName)
              ? AshaarFonts.descriptorForFontName(fnm).kasheedaName : null;
            if (kn) faceSet[fsz + "pt \"" + kn + "\""] = true;
          });
        });
      });
      var faceLoads = [];
      Object.keys(faceSet).forEach(function (s) { faceLoads.push(document.fonts.load(s).catch(function () {})); });
      try { await Promise.all(faceLoads); } catch (e) {}
    }

    // Measure natural widths + build the cross-block harmony matrix.
    var qMatrixCells = [];
    if (canvasCtx) {
      blockInfos.forEach(function (blk) {
        blk.tableInfos.forEach(function (info) {
          info.cells.forEach(function (c) {
            if (!AshaarMatrix.isContentCell(c.measure)) return;
            var fnm = c.fontName || repName, fsz = c.fontSize || repSize;
            // Measure in the BASE face: a re-applied Jameel cell reports the wider
            // Kasheeda face, which would inflate the harmony width and drift the
            // target on each apply. wordName maps Kasheeda→base; leaves others.
            var mfnm = (typeof AshaarFonts !== "undefined" && AshaarFonts.descriptorForFontName)
              ? (AshaarFonts.descriptorForFontName(fnm).wordName || fnm) : fnm;
            canvasCtx.font = fsz + "pt \"" + mfnm + "\"";
            c.natPx = AshaarProfiles.applyFontCorrection(canvasCtx.measureText(c.measure).width, mfnm, profile.fontCorrections);
            qMatrixCells.push({ key: c.matKey, natural: c.natPx });
          });
        });
      });
    }
    var qMatrix = AshaarMatrix.buildMatrix(qMatrixCells);

    return { blockInfos: blockInfos, pagePt: pagePt, pageTwips: pageTwips, repName: repName, repSize: repSize, canvasCtx: canvasCtx, qMatrix: qMatrix };
  }

  // Read each content cell's ORIGINAL per-word fonts (name/size/style/color)
  // BEFORE the SIZE rebuild flattens them to one representative font. Returns a
  // plain map keyed by block:table:cell index → { runs, align, indentTwips },
  // where runs are coalesced same-style segments. The rebuild regenerates cells
  // from font-less source text, so this is the only place the per-word fonts of
  // a mixed-font misra (e.g. Mehr + Amiri) still exist — pass 2 re-emits them.
  async function captureQaseedaCellRuns(context, cap) {
    var refs = [];
    cap.blockInfos.forEach(function (blk, b) {
      blk.tableInfos.forEach(function (info, t) {
        info.cells.forEach(function (c, i) {
          if (c.kind === "spacing" || !c.base) return;
          var wr = c.cell.body.getRange().getTextRanges([" "], true);
          wr.load("items");
          refs.push({ key: b + ":" + t + ":" + i, c: c, wr: wr });
        });
      });
    });
    if (!refs.length) return {};
    await context.sync();
    refs.forEach(function (r) {
      r.wr.items.forEach(function (w) { w.load("text"); w.font.load("name,size,bold,italic,color"); });
    });
    await context.sync();
    var out = {};
    refs.forEach(function (r) {
      var words = [];
      r.wr.items.forEach(function (w) {
        var txt = stripJustification(w.text || "");
        if (!txt) return;
        var f = w.font;
        var col = f && f.color;
        words.push({
          text: txt,
          name: (f && f.name) || r.c.fontName || "",
          size: (f && f.size) || r.c.fontSize || 0,
          bold: !!(f && f.bold), italic: !!(f && f.italic),
          color: (col && /^#?[0-9a-fA-F]{6}$/.test(col)) ? col : undefined
        });
      });
      if (!words.length) return;
      out[r.key] = { runs: AshaarWord.coalesceRuns(words), align: r.c.align, indentTwips: r.c.indentTwips || 0 };
    });
    return out;
  }

  // Apply a qaseeda's profile across ALL its blocks so they stay consistent. Two
  // passes: (1) SIZE — rebuild every block's table OOXML at one shared target
  // width (the only way to resize span tables; columns.setWidth garbles them —
  // see memory width-engine-rebuild-not-setwidth); same width for all bandhs →
  // same-GRID bandhs get an identical gridCol (harmony). (2) JUSTIFY — re-gather
  // the fresh bare tables and fill each cell to its box = span × (target/GRID).
  async function applyProfileToQaseeda(name) {
    if (typeof Word === "undefined") { setMessage("Open this task pane inside Word to apply a qaseeda."); return; }
    var profile = getProfile(name);
    var CELL_MARGIN_PT = 5.76;
    var MARGIN_PX = CELL_MARGIN_PT * 96 / 72;
    var strength = AshaarProfiles.normalizeStrength(profile.justify.strength);
    var fillMode = AshaarProfiles.normalizeFillMode(profile.justify.fillMode);
    var doKashida = profile.justify.mode === "kashida";
    // Both kashida and spacing fill the cell to its box; only "none"/"css" skip.
    var doFill = doKashida || profile.justify.mode === "spacing";
    var summary = "";
    var blockCount = 0, targetTwips = 0, sizeSig = "";
    // Per-word fonts of every content cell, captured before the rebuild flattens
    // them (block:table:cell → {runs, align, indentTwips}). Pass 2 re-emits these
    // so a mixed-font misra survives. Keyed by position; pass 1 and pass 2 gather
    // blocks in the same document order and rebuild from the same source, so the
    // indices line up (pass 2 also text-matches before trusting an entry).
    var origContent = {};

    try {
      // ── Pass 1: SIZE — rebuild each block at one shared target width ───────────
      await Word.run(async function (context) {
        var blocks = await gatherQaseedaBlocks(context, name);
        if (!blocks.length) { summary = "No blocks are tagged with qaseeda “" + name + "”."; return; }
        blockCount = blocks.length;
        var cap = await captureQaseedaTables(context, blocks, profile);
        // Snapshot per-word fonts NOW — the rebuild below discards them.
        origContent = await captureQaseedaCellRuns(context, cap);

        // Shared slot (px): auto-fit sizes it to hold every cell's natural text
        // (+kashida headroom, +cell margins → no wrap); fixed uses pct-of-page.
        var HEADROOM = doKashida ? 0.18 : 0.06;
        var bandhs = [];
        cap.blockInfos.forEach(function (blk) {
          blk.tableInfos.forEach(function (info) {
            var gcells = [];
            info.cells.forEach(function (c) { if (c.natPx != null) gcells.push({ natural: c.natPx, span: c.gridSpan || 1 }); });
            bandhs.push({ GRID: info.grid || 0, cells: gcells });
          });
        });
        var maxGRID = bandhs.reduce(function (m, b) { return Math.max(m, b.GRID || 0); }, 0);
        var pagePx = cap.pagePt * 96 / 72;
        var slotPx = AshaarMatrix.uniformSlotPx(bandhs, {
          mode: profile.width.mode === "fixed" ? "fixed" : "auto-fit",
          pct: profile.width.pct, pagePx: pagePx, headroom: HEADROOM, marginPx: MARGIN_PX
        });
        // One target width for all bandhs → same-GRID bandhs share an identical
        // cwt (harmony); a smaller-GRID bandh gets a wider cwt (still no wrap).
        // Capped at the page.
        targetTwips = Math.min(cap.pageTwips, Math.round(slotPx * maxGRID * 1440 / 96));
        if (targetTwips <= 0) targetTwips = cap.pageTwips;

        // Width signature: the target + each block's source. If unchanged since the
        // last apply this session, the tables are already sized correctly — skip
        // the destructive rebuild and let pass 2 just re-justify. Only a real width
        // change (mode/pct/text) triggers the rebuild.
        var srcSig = cap.blockInfos.map(function (b) {
          var h = 0, s = b.source; for (var i = 0; i < s.length; i++) h = ((h << 5) - h + s.charCodeAt(i)) | 0;
          return (h >>> 0).toString(16);
        }).join(",");
        sizeSig = targetTwips + "|" + srcSig;
        var needRebuild = _appliedSizeSig[name] !== sizeSig;
        if (!needRebuild) return; // pass 2 re-justifies the already-sized tables

        // Rebuild LAST block first so earlier blocks' ranges don't shift. Render
        // BARE (justifyMode none) with the block's own structural opts + pinned
        // font/size; the pattern is unchanged, so the old tag stays valid.
        for (var bi = cap.blockInfos.length - 1; bi >= 0; bi--) {
          var blk = cap.blockInfos[bi];
          if (!blk.source.trim()) continue;
          var p = blk.payload;
          var renderOpts = {
            layoutMode: p.layoutMode || "balanced",
            gapWidth: Number(p.gapWidth || 4),
            fontMode: p.fontMode || "document",
            misraPattern: p.misraPattern || "paired",
            misraCount: Number(p.misraCount || 4),
            tatweelCount: 0,
            justifyMode: "none"
          };
          if (blk.repFont) renderOpts.fontCsName = blk.repFont;
          if (blk.repSize) renderOpts.fontSizePt = blk.repSize;
          var ooxmlBody = AshaarWord.renderForWordOoxml(blk.source, renderOpts, Ashaar, targetTwips);
          if (!ooxmlBody) continue;
          // Embed the content control IN the OOXML (block-level w:sdt spanning all
          // tables) so insertOoxml creates a control over the WHOLE poem. Wrapping
          // the insertOoxml-returned range with insertContentControl() instead only
          // caught row 1 on Mac Word. insertOoxml("Replace") on a whole control
          // throws, so insert just after the old control, then delete the old.
          var ooxml = AshaarWord.wrapOoxmlControl(ooxmlBody, "Ashaar Poem", blk.oldTag);
          var afterRange = blk.cc.getRange("After");
          afterRange.insertOoxml(ooxml, Word.InsertLocation.start);
          blk.cc.delete(false);
          await context.sync();
        }
      });
      if (summary) { setMessage(summary); return; }

      // ── Pass 2: JUSTIFY — fill each cell of the fresh tables to its box ────────
      var changed = 0, coloured = 0;
      await Word.run(async function (context) {
        var blocks = await gatherQaseedaBlocks(context, name);
        if (!blocks.length) return;
        // Ensure the rebuilt (SDT-created) controls show the block outline.
        blocks.forEach(function (cc) { cc.appearance = "BoundingBox"; });
        var cap = await captureQaseedaTables(context, blocks, profile);
        var canvasCtx = cap.canvasCtx, repName = cap.repName, repSize = cap.repSize, qMatrix = cap.qMatrix;
        if (!canvasCtx) { summary = "Canvas unavailable; cannot measure."; return; }

        var MICRO_SPACE = " "; // hair space
        canvasCtx.font = repSize + "pt \"" + repName + "\"";
        if (canvasCtx.measureText(MICRO_SPACE).width <= 0) MICRO_SPACE = " "; // thin space

        // Force-load every ORIGINAL per-run font (+ Kasheeda face) referenced by
        // the captured cells. The rebuilt cells only report the flattened font, so
        // without this a mixed-font run would be measured with a substitute and
        // mis-elongated (see memory font-measurement-model).
        if (typeof document !== "undefined" && document.fonts && document.fonts.load) {
          var runFaces = {};
          Object.keys(origContent).forEach(function (k) {
            (origContent[k].runs || []).forEach(function (r) {
              var sz = r.size || repSize;
              if (r.name) runFaces[sz + "pt \"" + r.name + "\""] = true;
              var kn = AshaarFonts.descriptorForFontName(r.name).kasheedaName;
              if (kn) runFaces[sz + "pt \"" + kn + "\""] = true;
            });
          });
          var runLoads = [];
          Object.keys(runFaces).forEach(function (s) { runLoads.push(document.fonts.load(s).catch(function () {})); });
          try { await Promise.all(runLoads); } catch (e) {}
        }

        // Every content cell is (re)written as run-aware OOXML so each word keeps
        // its ORIGINAL font — the SIZE rebuild flattened the whole block to one
        // representative font, and a plain insertText would flatten it again.
        // Collected here, written after the spacing batch, one sync per cell.
        var cellPlans = [];

        // Build run-aware OOXML for one content cell: preserve each word's
        // original font/size/color (from origContent, captured before the rebuild)
        // and fill to the cell's target by the profile's fill mode + each run's
        // mechanism. Returns an OOXML paragraph string (never flat text, so cs
        // fonts are always explicit), or null to leave the bare rebuild.
        function buildContentCellOoxml(c, info, key, colPx) {
          var repFallback = c.fontName || repName;
          var sizeFallback = c.fontSize || repSize;
          // Original per-word runs; fall back to a single run when capture missed
          // or the reconstructed text no longer matches (rebuild changed words).
          var oc = origContent[key];
          var joined = (oc && oc.runs) ? oc.runs.map(function (r) { return r.text; }).join(" ").replace(/\s+/g, " ").trim() : "";
          var origRuns;
          if (oc && oc.runs && oc.runs.length && joined === c.base.replace(/\s+/g, " ").trim()) {
            origRuns = oc.runs.map(function (r) {
              return { text: r.text, name: r.name || repFallback, size: r.size || sizeFallback, color: r.color };
            });
          } else {
            origRuns = [{ text: c.base, name: repFallback, size: sizeFallback, color: undefined }];
          }
          var align = (oc && oc.align) || c.align || "center";
          var indentTwips = (oc && oc.indentTwips) || 0;
          var rep0Name = origRuns[0].name, rep0Size = origRuns[0].size;
          function measIn(text, nm, sz) { canvasCtx.font = sz + "pt \"" + nm + "\""; return canvasCtx.measureText(text).width; }

          // Regroup consecutive runs of the same FONT FAMILY into segments: both
          // Jameel faces map to one wordName (so a re-applied Jameel word, read
          // back with mixed base+Kasheeda faces, regroups into one segment and
          // re-derives its swap from clean text — idempotent); generic fonts key
          // on their own name so Amiri ≠ Fatemi. Each segment fills by its family's
          // mechanism. Natural widths are measured in the BASE face so a widened
          // Kasheeda read-back doesn't inflate the target on re-apply.
          function baseFaceOf(name) { return AshaarFonts.descriptorForFontName(name).wordName || name; }
          var segs = [];
          origRuns.forEach(function (r) {
            var d = AshaarFonts.descriptorForFontName(r.name);
            var fam = d.wordName || r.name;
            var prev = segs[segs.length - 1];
            if (prev && prev.fam === fam && prev.size === r.size && prev.color === r.color) {
              prev.text += " " + r.text;
            } else {
              segs.push({ fam: fam, mech: d.mechanism, desc: d, name: r.name, size: r.size, color: r.color, text: r.text });
            }
          });
          var interSpacePx = measIn(" ", baseFaceOf(segs[0].name), segs[0].size);

          // Flatten per-segment output runs, re-inserting the inter-segment space
          // (in the left segment's base face) that coalescing drops — so words
          // across a font boundary keep their gap instead of touching.
          function flattenSegs(segOut) {
            var out = [];
            segOut.forEach(function (arr, si) {
              if (si > 0) out.push({ text: " ", csName: baseFaceOf(segs[si - 1].name), sizePt: segs[si - 1].size, color: undefined });
              (arr || []).forEach(function (rr) { out.push(rr); });
            });
            return out;
          }

          // No fill (justify none/css) → re-emit each segment's text unchanged
          // (fonts + inter-segment spaces preserved), no elongation.
          if (!doFill || !(colPx > 0)) {
            var passOut = flattenSegs(segs.map(function (seg) {
              return [{ text: seg.text, csName: baseFaceOf(seg.name), sizePt: seg.size, color: seg.color }];
            }));
            return AshaarWord.misraRunsXml(passOut, align, rep0Size, { indentTwips: indentTwips });
          }

          // Natural width in BASE faces + the inter-segment spaces. Stable on
          // re-apply. Shared fill target (a per-cell width override wins): cell-fit
          // → toward the CELL EDGE; natural-fit → toward the position's HARMONY.
          var cNatural = segs.reduce(function (a, seg) { return a + measIn(seg.text, baseFaceOf(seg.name), seg.size); }, 0)
            + Math.max(0, segs.length - 1) * interSpacePx;
          var cOv = c.ovKey ? info.overrides[c.ovKey] : null;
          var cRes = AshaarOverrides.resolveCellOverride({ strength: strength, fillMode: fillMode }, cOv);
          var cPhi = AshaarWord.strengthToElongationShare(cRes.strength);
          var cCapEm = cRes.capEm != null ? cRes.capEm : undefined;
          var cMaxPos = AshaarWord.strengthToMaxPositions(cRes.strength);
          var cTarget;
          if (cRes.widthPt != null) cTarget = cRes.widthPt * 96 / 72;
          else if (fillMode === "cell-fit") cTarget = AshaarMatrix.cellFitBudget(cNatural, colPx, cPhi);
          else {
            var cReach = Math.max(cNatural, colPx - 0.28 * rep0Size * 96 / 72);
            var cWpos = qMatrix[c.matKey] || cNatural;
            cTarget = AshaarMatrix.naturalFitTarget(cWpos, cReach, cPhi);
          }
          cTarget = Math.min(cTarget, colPx); // no-wrap invariant

          // Elongate each segment by its OWN mechanism toward a proportional share
          // of the target; generic segments then absorb the slack the discrete
          // mechanisms (Jameel swap / Mehr tatweel) leave. Whitespace-shaping
          // segments never get tatweels. Under a spacing profile nothing elongates
          // and the capped hair-spaces do all the filling.
          var extra = Math.max(0, cTarget - cNatural);
          var segOut = new Array(segs.length);
          var genericIdx = [];
          var nonGenAchieved = Math.max(0, segs.length - 1) * interSpacePx; // inter-seg spaces
          segs.forEach(function (seg, si) {
            var bf = baseFaceOf(seg.name);
            var segNat = measIn(seg.text, bf, seg.size);
            var subTarget = cNatural > 0 ? segNat + extra * (segNat / cNatural) : segNat;
            if (doKashida && seg.mech === "font-swap") {
              var jw = seg.desc.kasheedaName || bf;
              var fss = AshaarKashidaFontswap.splitSpans(seg.text);
              var wb = [], ww = [];
              fss.forEach(function (s) { wb.push(measIn(s, bf, seg.size)); ww.push(measIn(s, jw, seg.size)); });
              var sel = AshaarKashidaFontswap.selectSwapRuns(fss, wb, ww, subTarget);
              segOut[si] = sel.runs.map(function (rr) { return { text: rr.text, csName: rr.swap ? jw : bf, sizePt: seg.size, color: seg.color }; });
              nonGenAchieved += fss.reduce(function (a, s, k) { return a + (sel.runs[k].swap ? ww[k] : wb[k]); }, 0);
            } else if (doKashida && seg.mech === "tatweel") {
              var mr = seg.desc.tatweelRules || {};
              var iso = {}, fin = {};
              (mr.isolatedInto || []).forEach(function (x) { iso[x] = true; });
              (mr.finalInto || []).forEach(function (x) { fin[x] = true; });
              var parts = seg.text.split(" "), toks = [];
              parts.forEach(function (wd, k) { if (k) toks.push(" "); toks.push(wd); });
              var elong = toks.map(function (t) { return t !== " " ? AshaarWord.mehrElongate(t, iso, fin) : t; });
              var mwb = [], mww = [];
              toks.forEach(function (t, k) { mwb.push(measIn(t, bf, seg.size)); mww.push(measIn(elong[k], bf, seg.size)); });
              var msel = AshaarKashidaFontswap.selectSwapRuns(toks, mwb, mww, subTarget);
              var mtext = msel.runs.map(function (rr, k) { return (rr.swap && mww[k] > mwb[k]) ? elong[k] : toks[k]; }).join("");
              segOut[si] = [{ text: mtext, csName: bf, sizePt: seg.size, color: seg.color }];
              nonGenAchieved += measIn(mtext, bf, seg.size);
            } else if (!doKashida || seg.mech === "whitespace") {
              segOut[si] = [{ text: seg.text, csName: bf, sizePt: seg.size, color: seg.color }];
              nonGenAchieved += segNat;
            } else {
              genericIdx.push(si); // generic — elongated jointly below
            }
          });
          if (genericIdx.length) {
            var primRuns = genericIdx.map(function (si) {
              var seg = segs[si], fstr = seg.size + "pt \"" + seg.name + "\"";
              return { text: seg.text, fontSize: seg.size, fontProfile: null,
                measure: function (s) { canvasCtx.font = fstr; return canvasCtx.measureText(s).width; } };
            });
            var genNat = primRuns.reduce(function (a, r) { return a + r.measure(r.text); }, 0);
            var genTarget = Math.max(genNat, cTarget - nonGenAchieved);
            var conc = AshaarJustify.justifyRunsConcentrated(primRuns, genTarget, { perPositionEm: 0.5, maxPositions: cMaxPos });
            genericIdx.forEach(function (si, k) {
              segOut[si] = [{ text: conc.runs[k].text, csName: segs[si].name, sizePt: segs[si].size, color: segs[si].color }];
            });
          }

          // Flatten (with inter-segment spaces) and close the residual with capped
          // hair-spaces across the word gaps.
          var outRuns = flattenSegs(segOut);
          var achievedTot = outRuns.reduce(function (a, rr) { return a + measIn(rr.text, rr.csName, rr.sizePt); }, 0);
          var totGaps = outRuns.reduce(function (a, rr) { return a + (rr.text.split(" ").length - 1); }, 0);
          var spacePx = measIn(MICRO_SPACE, rep0Name, rep0Size) || 1;
          var nSp = AshaarResidual.capMicroSpaces(cTarget - achievedTot, totGaps, spacePx, rep0Size * 96 / 72, cCapEm);
          var outTexts = AshaarWord.distributeMicroSpaces(outRuns.map(function (rr) { return rr.text; }), nSp, MICRO_SPACE);
          return AshaarWord.misraRunsXml(outRuns.map(function (rr, i) {
            return { text: outTexts[i], csName: rr.csName, sizePt: rr.sizePt, color: rr.color };
          }), align, rep0Size, { indentTwips: indentTwips });
        }

        cap.blockInfos.forEach(function (blk, bIdx) {
          blk.tableInfos.forEach(function (info, tIdx) {
            // Cell box comes from the width we just rebuilt to: cwt = target/GRID.
            var cwtPx = info.grid > 0 ? (targetTwips / info.grid) * 96 / 1440 : 0;
            info.cells.forEach(function (c, cIdx) {
              if (c.kind === "spacing") {
                // Decorate (not justify) a structural gap.
                var pDecor = c.slot ? (profile.spacingDecor || {})[c.slot] : null;
                var oDecor = c.decorKey ? info.slotDecor[c.decorKey] : null;
                var decor = AshaarOverrides.resolveSlotDecor(pDecor, oDecor);
                c.cell.body.clear();
                if (decor.symbol) {
                  c.cell.body.insertText(decor.symbol, Word.InsertLocation.replace);
                  c.cell.body.font.color = decor.color || "black";
                }
                // shadingColor rejects "" / "No color"; use "#FFFFFF" to clear.
                c.cell.shadingColor = decor.fill || "#FFFFFF";
                c.cell.body.paragraphs.getFirst().alignment = Word.Alignment.centered;
                changed++;
                return;
              }
              if (!c.base) return;
              // Fill box = span × cwt − cell margins (the text area we rebuilt to).
              var colPx = cwtPx > 0
                ? Math.max(1, (c.gridSpan || 1) * cwtPx - 2 * MARGIN_PX)
                : Math.max(1, (c.cell.columnWidth || 0) - 2 * CELL_MARGIN_PT) * 96 / 72;
              var ooxml = buildContentCellOoxml(c, info, bIdx + ":" + tIdx + ":" + cIdx, colPx);
              if (ooxml) cellPlans.push({ cell: c.cell, ooxml: ooxml });
            });
          });
        });
        await context.sync(); // commit the spacing-cell decorations

        // Write each content cell's run-aware OOXML: clear + insert. One sync per
        // cell so a single OOXML failure leaves that cell as its bare rebuild
        // instead of aborting the whole batch.
        for (var cpi = 0; cpi < cellPlans.length; cpi++) {
          var cp = cellPlans[cpi];
          try {
            cp.cell.body.clear();
            cp.cell.body.insertOoxml(AshaarWord.wrapOoxml(cp.ooxml), Word.InsertLocation.replace);
            await context.sync();
            changed++;
          } catch (e) { /* leave the cell as its bare rebuild */ }
        }

        // Debug colouring: tint inserted tatweels / micro-spaces so they're visible.
        var tatColor = (profile.debugColors && profile.debugColors.tatweel) || "";
        var spcColor = (profile.debugColors && profile.debugColors.space) || "";
        if (tatColor || spcColor) {
          var hits = [];
          cap.blockInfos.forEach(function (blk) {
            blk.tableInfos.forEach(function (info) {
              info.cells.forEach(function (c) {
                if (tatColor) { var st = c.cell.body.search("ـ"); st.load("items"); hits.push({ s: st, color: tatColor, hl: false }); }
                if (spcColor) {
                  var sh = c.cell.body.search(" "); sh.load("items"); hits.push({ s: sh, color: spcColor, hl: true });
                  var sn = c.cell.body.search(" "); sn.load("items"); hits.push({ s: sn, color: spcColor, hl: true });
                }
              });
            });
          });
          await context.sync();
          // Spaces have no ink, so tint them with the HIGHLIGHT (background); tatweels
          // are ink, so use the font color.
          hits.forEach(function (h) {
            h.s.items.forEach(function (r) {
              if (h.hl) r.font.highlightColor = h.color; else r.font.color = h.color;
              coloured++;
            });
          });
          await context.sync();
        }
      });

      // Remember the width we sized to, so a later justify-only apply (strength,
      // fill mode, per-cell override) skips the destructive rebuild.
      if (sizeSig) _appliedSizeSig[name] = sizeSig;
      summary = "Applied qaseeda “" + name + "” to " + blockCount + " block(s); justified " + changed + " cell(s)"
        + (coloured ? "; coloured " + coloured + " artifact(s)" : "") + ".";
    } catch (error) {
      summary = "Apply failed: " + describeError(error);
    }
    setMessage(summary);
  }

  // ── Qaseeda panel ↔ profile (P4 UI) ───────────────────────────────────────
  function panelToProfile() {
    var p = AshaarProfiles.defaultProfile((qaseedaName.value || "").trim());
    p.width.mode = qaseedaWidthMode.value;
    p.width.pct = Number(qaseedaWidthPct.value || 50);
    p.justify.mode = qaseedaJustifyMode.value;
    p.justify.fillMode = (qaseedaFillMode && qaseedaFillMode.value) || "natural-fit";
    p.justify.strength = Number(qaseedaStrength.value || 0);
    var corrFont = (qaseedaCorrFont.value || "").trim();
    p.fontCorrections = {};
    if (corrFont) p.fontCorrections[corrFont] = Number(qaseedaCorrFactor.value || 1);
    p.debugColors = {
      tatweel: (qaseedaDebugTatweelOn && qaseedaDebugTatweelOn.checked) ? (qaseedaDebugTatweel.value || "") : "",
      space: (qaseedaDebugSpaceOn && qaseedaDebugSpaceOn.checked) ? (qaseedaDebugSpace.value || "") : ""
    };
    return AshaarProfiles.normalizeProfile(p);
  }

  function profileToPanel(profile) {
    var p = AshaarProfiles.normalizeProfile(profile || {});
    qaseedaWidthMode.value = p.width.mode;
    qaseedaWidthPct.value = p.width.pct;
    qaseedaJustifyMode.value = p.justify.mode;
    if (qaseedaFillMode) qaseedaFillMode.value = AshaarProfiles.normalizeFillMode(p.justify.fillMode);
    var strength = AshaarProfiles.normalizeStrength(p.justify.strength);
    qaseedaStrength.value = strength;
    qaseedaStrengthValue.textContent = strength;
    var fonts = Object.keys(p.fontCorrections || {});
    qaseedaCorrFont.value = fonts[0] || "";
    qaseedaCorrFactor.value = fonts[0] ? p.fontCorrections[fonts[0]] : 1;
    var dc = p.debugColors || {};
    if (qaseedaDebugTatweelOn) qaseedaDebugTatweelOn.checked = !!dc.tatweel;
    if (dc.tatweel) qaseedaDebugTatweel.value = dc.tatweel;
    if (qaseedaDebugSpaceOn) qaseedaDebugSpaceOn.checked = !!dc.space;
    if (dc.space) qaseedaDebugSpace.value = dc.space;
  }

  function populateQaseedaNames() {
    if (!qaseedaNames) return;
    var names = listProfileNames();
    qaseedaNames.innerHTML = names.map(function (n) {
      return "<option value=\"" + String(n).replace(/"/g, "&quot;") + "\">";
    }).join("");
  }

  function loadQaseedaIntoPanel() {
    var name = (qaseedaName.value || "").trim();
    if (!name) return;
    var store = loadProfileStore();
    if (store[name]) profileToPanel(store[name]);
  }

  async function saveAndApplyQaseeda() {
    var p = panelToProfile();
    if (!p.name) { setMessage("Name the qaseeda first."); return; }
    await putProfile(p);
    populateQaseedaNames();
    await applyProfileToQaseeda(p.name);
  }

  async function assignBlockToQaseeda() {
    var name = (qaseedaName.value || "").trim();
    if (!name) { setMessage("Name the qaseeda first."); return; }
    if (!loadProfileStore()[name]) await putProfile(panelToProfile());
    await setQaseedaOnSelection(name);
    populateQaseedaNames();
  }

  function setQaseedaFontStatus(text, kind) {
    if (!qaseedaFontStatus) return;
    qaseedaFontStatus.textContent = text;
    qaseedaFontStatus.className = "qaseeda-font-status" + (kind === "ok" ? " is-ok" : kind === "warn" ? " is-warn" : "");
  }

  // Check whether the font of the block/selection at the cursor resolves in the
  // WebView; if not, justify metrics for it are only approximate.
  async function checkQaseedaFont() {
    if (typeof Word === "undefined") { setQaseedaFontStatus("Open in Word to check.", "warn"); return; }
    var fontName = "";
    try {
      await Word.run(async function (context) {
        var sel = context.document.getSelection();
        var cc = sel.parentContentControlOrNullObject;
        cc.load("title");
        await context.sync();
        var range = (!cc.isNullObject && cc.title === "Ashaar Poem") ? cc.getRange() : sel;
        range.font.load("name");
        await context.sync();
        fontName = range.font.name || "";
      });
    } catch (e) { /* ignore */ }
    if (!fontName) { setQaseedaFontStatus("No font found at the cursor.", "warn"); return; }
    if (document.fonts && document.fonts.load) { try { await document.fonts.load("16pt \"" + fontName + "\""); } catch (e) {} }
    if (fontAvailable(fontName)) setQaseedaFontStatus("“" + fontName + "” resolves — metrics are accurate.", "ok");
    else setQaseedaFontStatus("“" + fontName + "” is NOT resolvable here — metrics are approximate. Add its font file under Custom fonts.", "warn");
  }

  // ── Custom fonts (AshaarFontStore) ─────────────────────────────────────────
  // Let the user load a font from their machine so the justify canvas measures
  // the real outlines (see font-store.js). Registered under the exact name Word
  // reports for the text; persisted in IndexedDB and re-registered on startup.
  var fontUpload = document.getElementById("font-upload");
  var fontUploadName = document.getElementById("font-upload-name");
  var fontUploadStatus = document.getElementById("font-upload-status");
  var fontList = document.getElementById("font-list");

  function setFontUploadStatus(text, kind) {
    if (!fontUploadStatus) return;
    fontUploadStatus.textContent = text;
    fontUploadStatus.className = "qaseeda-font-status" + (kind === "ok" ? " is-ok" : kind === "warn" ? " is-warn" : "");
  }

  function renderFontList(fonts) {
    if (!fontList) return;
    fontList.innerHTML = "";
    (fonts || []).forEach(function (f) {
      var li = document.createElement("li");
      var span = document.createElement("span");
      span.textContent = f.family + (f.filename ? " (" + f.filename + ")" : "");
      var btn = document.createElement("button");
      btn.type = "button";
      btn.className = "button--secondary font-remove";
      btn.textContent = "×";
      btn.setAttribute("aria-label", "Remove " + f.family);
      btn.addEventListener("click", function () { removeFont(f.family); });
      li.appendChild(span);
      li.appendChild(btn);
      fontList.appendChild(li);
    });
  }

  function refreshFontList() {
    if (typeof AshaarFontStore === "undefined") return Promise.resolve();
    return AshaarFontStore.listFonts().then(renderFontList, function () {});
  }

  // Auto-detect the family from the picked file and prefill the name field.
  function onFontFilePicked() {
    setFontUploadStatus("", "");
    var file = fontUpload && fontUpload.files && fontUpload.files[0];
    if (!file || typeof AshaarFontStore === "undefined") return;
    file.arrayBuffer().then(function (buf) {
      var names = AshaarFontStore.parseNames(buf);
      if (names && names.family) {
        fontUploadName.value = names.family;
      } else if (!fontUploadName.value) {
        fontUploadName.value = (file.name || "").replace(/\.[^.]+$/, "");
        setFontUploadStatus("Couldn’t read the font’s name — confirm it matches what Word shows (Verify at cursor).", "warn");
      }
    }, function () {});
  }

  async function addFont() {
    if (typeof AshaarFontStore === "undefined") return;
    var file = fontUpload && fontUpload.files && fontUpload.files[0];
    var family = (fontUploadName.value || "").trim();
    if (!file) { setFontUploadStatus("Choose a font file first.", "warn"); return; }
    if (!family) { setFontUploadStatus("Enter the name Word uses for this font.", "warn"); return; }
    try {
      var res = await AshaarFontStore.addUserFont(family, file);
      setFontUploadStatus("Loaded “" + family + "”" +
        (res.persisted ? " — saved for future sessions." : " — this session only (storage unavailable)."), "ok");
      fontUpload.value = "";
      await refreshFontList();
    } catch (e) {
      setFontUploadStatus("Couldn’t load that file — is it a valid .ttf/.otf/.woff?", "warn");
    }
  }

  function removeFont(family) {
    if (typeof AshaarFontStore === "undefined") return;
    AshaarFontStore.deleteFont(family).then(function () {
      setFontUploadStatus("Removed “" + family + "”. Reload the add-in to fully unload it.", "");
      return refreshFontList();
    }, function () {});
  }

  // Compare the registered name against the font Word applies at the cursor.
  async function verifyFontAtCursor() {
    if (typeof Word === "undefined") { setFontUploadStatus("Open in Word to verify.", "warn"); return; }
    var want = (fontUploadName.value || "").trim();
    var fontName = "";
    try {
      await Word.run(async function (context) {
        var sel = context.document.getSelection();
        sel.font.load("name");
        await context.sync();
        fontName = sel.font.name || "";
      });
    } catch (e) { /* ignore */ }
    if (!fontName) { setFontUploadStatus("Put the cursor in the styled text, then Verify.", "warn"); return; }
    if (document.fonts && document.fonts.load) { try { await document.fonts.load("16pt \"" + fontName + "\""); } catch (e) {} }
    var resolves = fontAvailable(fontName);
    if (want !== fontName) {
      fontUploadName.value = fontName;
      setFontUploadStatus("Word uses “" + fontName + "” here — I set the name to match. Pick its file and click Add.", "warn");
    } else if (resolves) {
      setFontUploadStatus("✓ “" + fontName + "” matches and resolves — justify will be accurate.", "ok");
    } else {
      setFontUploadStatus("“" + fontName + "” isn’t resolvable yet — pick its font file above and click Add.", "warn");
    }
  }

  async function insertPoem(replaceSelection, optsOverride) {
    await withWord(async function (context) {
      var opts = options();
      // Re-render passes overrides (e.g. justifyMode:"none" for a bare rebuild,
      // fontCsName to pin the poem's existing font). Merge over the pane opts.
      if (optsOverride) opts = Object.assign({}, opts, optsOverride);
      var source = String(input.value || "");

      var sectionP = context.document.sections.getFirst();
      sectionP.load("pageLayout/width,pageLayout/leftMargin,pageLayout/rightMargin");
      var selFontP = context.document.getSelection();
      selFontP.load("font/size,font/name");
      await context.sync();

      // pageLayout requires WordApi 1.5; fall back to US-Letter 6.5" on older builds
      var plP = sectionP.pageLayout;
      var pageTwips = plP && plP.width
        ? Math.round((plP.width - (plP.leftMargin || 0) - (plP.rightMargin || 0)) * 20)
        : 9360;

      // Measurement canvas at the selection's font — used for auto-fit, the nudge, and kashida.
      var fontSizeP = selFontP.font.size || 12;
      // Size-preserving re-render: when REPLACING an existing poem (justify /
      // adjust re-render), carry the selection's real font size into the rebuild
      // so the replaced table keeps its size instead of reverting to Word's
      // default. Only when the selection reports a single size (mixed → null →
      // leave the default). Fresh insert (replaceSelection false) is unchanged.
      if (replaceSelection && selFontP.font.size) opts.fontSizePt = selFontP.font.size;
      var modeP = opts.fontMode === "nastaliq" ? "noto" : opts.fontMode;
      var fontNameP = AshaarFonts.wordNameOf(modeP)
                    || selFontP.font.name || "Times New Roman";
      var ctxP = document.createElement("canvas").getContext("2d");
      if (ctxP) ctxP.font = fontSizeP + "pt \"" + fontNameP + "\"";

      var neededTwips = ctxP ? neededTableTwips(source, ctxP, opts, pageTwips) : pageTwips;
      var textWidthTwips;
      if (opts.autoFitWidth) {
        textWidthTwips = Math.min(pageTwips, neededTwips);
      } else {
        textWidthTwips = scaledTextWidth(pageTwips);
        // Never render narrower than needed to avoid misra word-wrap — floor at
        // the needed width (still capped at the page width), even when the user
        // hasn't enabled Auto-fit. A wider user preference is still honored.
        if (ctxP && neededTwips > textWidthTwips) {
          textWidthTwips = Math.min(pageTwips, neededTwips);
        }
      }

      // "Let Word fill it" (§4): expand the render width by a strength-driven
      // fraction (0..15%) so higher strength visibly fills a wider table. The
      // base width above still comes from the Table-width slider / auto-fit;
      // this only scales it up, capped at the page width.
      if (opts.justifyMode === "css") {
        var expandFrac = AshaarWord.kashidaExpansionFraction(opts.tatweelCount);
        textWidthTwips = Math.min(pageTwips, Math.round(textWidthTwips * (1 + expandFrac)));
      }

      if ((opts.justifyMode === "kashida" || opts.justifyMode === "spacing") && ctxP) {
        opts._textWidthPx = textWidthTwips * 96 / 1440;
        opts._justifyCtx = ctxP;
      }

      var ooxmlBody;
      try {
        ooxmlBody = AshaarWord.renderForWordOoxml(source, opts, Ashaar, textWidthTwips);
      } catch (err) {
        setMessage("Render error: " + (err.message || String(err)));
        return;
      }
      if (!ooxmlBody) { setMessage("No content generated."); return; }

      var ooxml = AshaarWord.wrapOoxml(ooxmlBody);
      var newTag = AshaarWord.contentControlTag(source, opts, AshaarWord.poemCellPatterns(source, opts, Ashaar));
      var selection = context.document.getSelection();

      // In-place re-render (Re-render / word-fill re-render). insertOoxml("Replace")
      // on a Range that spans an ENTIRE "Ashaar Poem" content control — its boundary
      // markers plus the poem tables — throws GeneralException. Replace the control's
      // CONTENT instead (ContentControl.insertOoxml keeps the wrapper), so nothing
      // crosses the boundary and we don't nest a second control inside the old one.
      // Fresh inserts and manual "Replace Selection" over plain text fall through
      // to the original selection-scope insert unchanged.
      if (replaceSelection) {
        // Robustly find the enclosing "Ashaar Poem" control. We can't use
        // parentContentControlOrNullObject: Re-render selects the control's WHOLE
        // range (boundary markers included), so the selection has no strict parent
        // control and that lookup returns null (which is why the boundary-crossing
        // Range.insertOoxml still fired). Intersect each control's range with the
        // selection instead — the SP2 detection pattern — which matches even when
        // the selection spans the whole control.
        var poemCCs = context.document.contentControls;
        poemCCs.load("items/title");
        await context.sync();
        var xs = poemCCs.items.map(function (c) {
          return { cc: c, hit: c.getRange().intersectWithOrNullObject(selection) };
        });
        xs.forEach(function (x) { x.hit.load("isNullObject"); });
        await context.sync();
        var poemCC = null;
        for (var xi = 0; xi < xs.length; xi++) {
          if (xs[xi].cc.title === "Ashaar Poem" && !xs[xi].hit.isNullObject) { poemCC = xs[xi].cc; break; }
        }
        if (poemCC) {
          // This build also rejects ContentControl.insertOoxml("Replace") with
          // tables (GeneralException) — replacing INTO a control with block tables
          // fails at every scope. Only body-scope table insert works (proven by the
          // End path). So insert the rebuilt poem into the BODY just after the old
          // control, wrap it in a fresh "Ashaar Poem" control, then delete the old
          // control and its content. Net effect: the poem is replaced in place.
          var afterRange = poemCC.getRange("After");
          var insertedRange = afterRange.insertOoxml(ooxml, Word.InsertLocation.start);
          var newControl = insertedRange.insertContentControl();
          newControl.title = "Ashaar Poem";
          newControl.tag = newTag;
          newControl.appearance = "BoundingBox";
          poemCC.delete(false); // remove old control + its content
          await context.sync();
          return;
        }
      }

      var inserted = selection.insertOoxml(ooxml,
        replaceSelection ? Word.InsertLocation.replace : Word.InsertLocation.end);
      var control = inserted.insertContentControl();
      control.title = "Ashaar Poem";
      control.tag = newTag;
      control.appearance = "BoundingBox";
      await context.sync();
    });
  }

  async function insertStructure() {
    await withWord(async function (context) {
      var opts = options();

      // Grid mode: build the 12-column span template from the bubble grid and
      // insert it (repeated per bandh) via the existing template-OOXML path.
      if (layoutView === "grid" && typeof AshaarLayoutGrid !== "undefined") {
        var sectionG = context.document.sections.getFirst();
        sectionG.load("pageLayout/width,pageLayout/leftMargin,pageLayout/rightMargin");
        await context.sync();
        var plG = sectionG.pageLayout;
        var twG = plG && plG.width
          ? Math.round((plG.width - (plG.leftMargin || 0) - (plG.rightMargin || 0)) * 20)
          : 9360;
        var tmplG = AshaarLayoutGrid.gridToTemplate(gridMatrix);
        if (!tmplG.rows.length) { setMessage("Draw at least one row of bubbles in the grid."); return; }
        var countG = Math.max(1, Math.min(20, Number(opts.bandhCount || 1)));
        var bodyG = [];
        var twGs = scaledTextWidth(twG);
        for (var bi = 0; bi < countG; bi++) bodyG.push(AshaarWord.templateToOoxml(tmplG, twGs, opts));
        var selG = context.document.getSelection();
        var insG = selG.insertOoxml(AshaarWord.wrapOoxml(bodyG.join("<w:p/>")), Word.InsertLocation.end);
        var ccG = insG.insertContentControl();
        ccG.title = "Ashaar Poem";
        ccG.tag = AshaarWord.contentControlTag("grid", opts);
        ccG.appearance = "BoundingBox";
        await context.sync();
        return;
      }

      var tables = AshaarWord.layoutTablesForTemplate(opts);

      // Both span-based and plain (Numbers-view) layouts go through OOXML so
      // every table carries <w:bidiVisual/> and is a genuine RTL table. (The
      // native Word.insertTable API has no per-table RTL flag, so it always
      // yields an LTR table whose cell/tab order runs left-to-right even when
      // the content looks right — hence OOXML for every layout.)
      if (tables.length) {
        var section = context.document.sections.getFirst();
        section.load("pageLayout/width,pageLayout/leftMargin,pageLayout/rightMargin");
        await context.sync();
        var pl = section.pageLayout;
        var textWidthTwips = pl && pl.width
          ? Math.round((pl.width - (pl.leftMargin || 0) - (pl.rightMargin || 0)) * 20)
          : 9360;
        var scaled = scaledTextWidth(textWidthTwips);
        var ooxmlBody = tables.map(function (t) {
          return t.spanBased
            ? AshaarWord.templateToOoxml(t, scaled, opts)
            : AshaarWord.layoutTableToOoxml(t, scaled, opts);
        }).join("<w:p/>");
        var selection = context.document.getSelection();
        var inserted = selection.insertOoxml(AshaarWord.wrapOoxml(ooxmlBody), Word.InsertLocation.end);
        var control = inserted.insertContentControl();
        control.title = "Ashaar Poem";
        control.tag = AshaarWord.contentControlTag("template", opts);
        control.appearance = "BoundingBox";
        await context.sync();
        return;
      }

      var html = AshaarWord.renderTemplateForWord(opts);
      var selection = context.document.getSelection();
      var inserted = selection.insertHtml(html, Word.InsertLocation.end);
      var control = inserted.insertContentControl();
      control.title = "Ashaar Poem";
      control.tag = AshaarWord.contentControlTag("template", opts);
      control.appearance = "BoundingBox";
      await context.sync();
    });
  }

  async function insertTabStopPoem() {
    await withWord(async function (context) {
      // Read the actual text-area width so tab stops fit the document's page size and margins.
      // pageLayout properties are in points; multiply by 20 to convert to twips.
      var section = context.document.sections.getFirst();
      section.load("pageLayout/width,pageLayout/leftMargin,pageLayout/rightMargin");
      var selFont = context.document.getSelection();
      selFont.load("font/size,font/name");
      await context.sync();
      var pl = section.pageLayout;
      // pageLayout requires WordApi 1.5; fall back to US-Letter 6.5" text width on older builds
      var textWidthTwips = pl && pl.width
        ? Math.round((pl.width - (pl.leftMargin || 0) - (pl.rightMargin || 0)) * 20)
        : 9360;

      var opts = options();
      if (opts.justifyMode === "kashida" || opts.justifyMode === "spacing") {
        var fontSize = selFont.font.size || 12;
        var mode = opts.fontMode === "nastaliq" ? "noto" : opts.fontMode;
        var fontName = AshaarFonts.wordNameOf(mode)
                     || selFont.font.name || "Times New Roman";
        var canvas = document.createElement("canvas");
        var ctx = canvas.getContext("2d");
        if (ctx) {
          ctx.font = fontSize + "pt \"" + fontName + "\"";
          opts._justifyCtx = ctx;
        }
      }
      var source = String(input.value || "");
      var content;
      try {
        content = AshaarTabStop.poemToOoxml(source, opts, Ashaar, textWidthTwips);
      } catch (e) {
        setMessage("Paragraph engine error: " + (e.message || String(e)));
        return;
      }
      if (!content) {
        setMessage("No content generated.");
        return;
      }
      var ooxml = AshaarTabStop.wrapOoxml(content);
      var selection = context.document.getSelection();
      var inserted = selection.insertOoxml(ooxml, Word.InsertLocation.end);
      var control = inserted.insertContentControl();
      control.title = "Ashaar Poem";
      control.tag = AshaarWord.contentControlTag(source, opts);
      control.appearance = "BoundingBox";
      await context.sync();
    });
  }

  async function loadSelection() {
    var picked = null;
    await withWord(async function (context) {
      var selection = context.document.getSelection();
      selection.load("text");
      await context.sync();
      picked = selection.text || "";
    });
    if (picked) { setMode("convert"); input.value = picked; }
    // Run after withWord (which sets "Done.") so the conversion note is visible.
    applyImportNormalization();
  }

  // Kashida (U+0640) and the micro-spaces inserted by spacing justification
  // (hair U+200A, thin U+2009) are elongation artifacts, not part of the source
  // line. Strip them so every justification pass re-derives from the bare text —
  // making justification idempotent and reducible: re-justify any number of times
  // and the kashida count follows the current width / font / fill, never compounds.
  function stripJustification(s) {
    return String(s || "").replace(/[ـ  ]/g, "");
  }

  // "Let Word fill it" mode: no probe/calibrate/tatweel — Word's own kashida
  // renderer does the work via the paragraph's jc (justification) value.
  //
  // A per-cell insertOoxml replace (the original approach) is WRONG: verified
  // in Word that it drops the kashida jc — Word resets a cell paragraph's
  // alignment to the cell's own positional value on a body-scope replace.
  // insertOoxml at SELECTION scope, by contrast, preserves jc (that's exactly
  // what insertPoem/adoptTable already rely on). So instead of rebuilding
  // OOXML per cell, we reconstruct the poem's plain-text SOURCE from the
  // table(s)' cells (mirrors adoptTable's AshaarTableAdopt.adoptTableToSource
  // reconstruction) and re-run the normal insertPoem(true) pipeline — which
  // renders via renderForWordOoxml → wrapOoxml → selection.insertOoxml(...,
  // replace) and re-wraps the "Ashaar Poem" content control. Because the
  // §6 "Reset to unstretched" (core): strip justification artifacts — tatweels
  // (U+0640) and hairline/thin micro-spaces (U+200A / U+2009) — from the poem or
  // table at the cursor, IN PLACE. Never rebuilds the table, so font size and
  // run formatting survive. Reuses stripJustification (the same strip re-justify
  // already runs before each pass, so this leaves the exact bare text). The
  // fuller §6 reset (uniform font-scale, column widening, word-fill jc/break) is
  // a separate follow-up; this covers the "clear kashida & spaces" case.
  async function resetJustification() {
    if (typeof Word === "undefined") { setMessage("Open this task pane inside Word to reset."); return; }
    await withWord(async function (context) {
      var selection = context.document.getSelection();
      var cc = selection.parentContentControlOrNullObject;
      cc.load("title");
      await context.sync();
      var workRange = (!cc.isNullObject && cc.title === "Ashaar Poem") ? cc.getRange() : selection;
      var tables = workRange.tables;
      tables.load("items");
      await context.sync();
      if (!tables.items.length) { setMessage("Click inside an Ashaar table to reset."); return; }

      tables.items.forEach(function (tbl) { tbl.rows.load("items"); });
      await context.sync();
      tables.items.forEach(function (tbl) { tbl.rows.items.forEach(function (row) { row.cells.load("items"); }); });
      await context.sync();
      var cells = [];
      tables.items.forEach(function (tbl) {
        tbl.rows.items.forEach(function (row) {
          row.cells.items.forEach(function (cell) { cells.push(cell); cell.body.load("text"); });
        });
      });
      await context.sync();

      // One sync per changed cell so a single-cell failure falls out of the
      // batch without aborting the rest (mirrors justifySelection's write phase).
      var cleared = 0;
      for (var i = 0; i < cells.length; i++) {
        var cur = cells[i].body.text || "";
        var stripped = stripJustification(cur);
        if (stripped === cur) continue; // already clean — no tatweels/micro-spaces
        try {
          cells[i].body.paragraphs.getFirst().insertText(stripped, Word.InsertLocation.replace);
          await context.sync();
          cleared++;
        } catch (e) { /* skip a cell that fails to rewrite */ }
      }
      setMessage("Reset " + cleared + " cell(s) — kashida & micro-spaces cleared.");
    });
  }

  // Single-button Re-render for a managed poem: reconstruct it from the current
  // cells and rebuild through the normal render path applying the pane's gap /
  // table-width, but BARE (justifyMode "none") — renderForWordOoxml is
  // mechanism-unaware, so baking kashida here would inject generic tatweels into
  // a Jameel/Mehr poem. Font + size are preserved: size rides insertPoem's
  // replace-capture; font is the dropdown font when one is explicitly chosen
  // (so a mode that needs a specific font adopts it), else the poem's existing
  // font (pinned via opts.fontCsName). After the bare rebuild, fill in place
  // with the correct per-cell mechanism by delegating to justifySelection.
  async function reRender() {
    if (typeof Word === "undefined") { setMessage("Open this task pane inside Word to re-render."); return; }
    var opts = options();
    var source = "";
    var existingFont = "";

    await withWord(async function (context) {
      var selection = context.document.getSelection();
      var cc = selection.parentContentControlOrNullObject;
      cc.load("title");
      await context.sync();
      var workRange = (!cc.isNullObject && cc.title === "Ashaar Poem") ? cc.getRange() : selection;
      var tables = workRange.tables;
      tables.load("items");
      await context.sync();
      if (!tables.items.length) { setMessage("Click inside an Ashaar table to re-render."); return; }

      tables.items.forEach(function (tbl) { tbl.rows.load("items"); });
      await context.sync();
      tables.items.forEach(function (tbl) { tbl.rows.items.forEach(function (row) { row.cells.load("items"); }); });
      await context.sync();
      tables.items.forEach(function (tbl) {
        tbl.rows.items.forEach(function (row) {
          row.cells.items.forEach(function (cell) { cell.body.load("text"); cell.body.font.load("name"); });
        });
      });
      await context.sync();

      // Reconstruct source (same reconstruction Adopt / word-fill use).
      source = tables.items.map(function (tbl) {
        var rows = tbl.rows.items.map(function (row) {
          return row.cells.items.map(function (cell) { return cell.body.text || ""; });
        });
        return AshaarTableAdopt.adoptTableToSource(rows, { direction: "rtl" });
      }).filter(function (s) { return s.trim(); }).join("\n\n");

      // Representative existing font (first cell reporting one) to preserve.
      tables.items.forEach(function (tbl) {
        tbl.rows.items.forEach(function (row) {
          row.cells.items.forEach(function (cell) {
            if (!existingFont && cell.body.font && cell.body.font.name) existingFont = cell.body.font.name;
          });
        });
      });

      if (!source.trim()) { setMessage("That table didn't contain any text to re-render."); return; }
      workRange.select();
      await context.sync();
    });

    if (!source.trim()) return; // a friendly message was already shown

    // Font pin: an explicitly-chosen dropdown font wins (so a mode that needs a
    // specific font adopts it); otherwise preserve the poem's existing font.
    var dropdownFont = AshaarFonts.wordNameOf(opts.fontMode === "nastaliq" ? "noto" : opts.fontMode);
    var fontCsName = dropdownFont || existingFont || null;

    input.value = source;
    // Step 1: bare rebuild — gap/width from the pane, font pinned, size preserved.
    await insertPoem(true, { justifyMode: "none", fontCsName: fontCsName });
    // Step 2: fill in place with the correct per-cell mechanism for the chosen
    // mode (skipped when the pane mode is "none").
    if (opts.justifyMode && opts.justifyMode !== "none") await justifySelection();
    setMessage("Re-rendered (font & size preserved).");
  }

  // justify-mode dropdown is "Word justify" (opts.justifyMode === "css") on
  // this path, that re-render emits the word-fill kashida jc + shrunk break
  // (misraParaXml), with width %/gap/strength/fontMode flowing through
  // exactly as they do for a fresh insert.
  async function justifySelectionWordFill(opts) {
    setMessage("Justifying…");
    var source = "";

    await withWord(async function (context) {
      var selection = context.document.getSelection();

      // Find enclosing Ashaar Poem content control (mirrors the kashida/spacing path).
      var cc = selection.parentContentControlOrNullObject;
      cc.load("title");
      await context.sync();

      var workRange = (!cc.isNullObject && cc.title === "Ashaar Poem")
        ? cc.getRange() : selection;

      var tables = workRange.tables;
      tables.load("items");
      await context.sync();

      if (!tables.items.length) {
        setMessage("Select an Ashaar table to fill.");
        return;
      }

      tables.items.forEach(function (tbl) { tbl.rows.load("items"); });
      await context.sync();
      tables.items.forEach(function (tbl) {
        tbl.rows.items.forEach(function (row) { row.cells.load("items"); });
      });
      await context.sync();
      tables.items.forEach(function (tbl) {
        tbl.rows.items.forEach(function (row) {
          row.cells.items.forEach(function (cell) { cell.body.load("text"); });
        });
      });
      await context.sync();

      // Each table → one stanza; multiple tables in scope → stanza-separated.
      // Same reconstruction adoptTable uses, so the round-trip parsing rules
      // (misra/bayt/refrain detection from row layout) stay in one place.
      source = tables.items.map(function (tbl) {
        var rows = tbl.rows.items.map(function (row) {
          return row.cells.items.map(function (cell) { return cell.body.text || ""; });
        });
        return AshaarTableAdopt.adoptTableToSource(rows, { direction: "rtl" });
      }).filter(function (s) { return s.trim(); }).join("\n\n");

      if (!source.trim()) {
        setMessage("That table didn't contain any text to fill.");
        return;
      }

      // Put the selection on the content insertPoem(true) will replace.
      workRange.select();
      await context.sync();
    });

    if (!source.trim()) return; // a friendly message was already shown

    input.value = source;
    await insertPoem(true);
  }

  // Wrapper: suppress active-context reflection while our own justify mutates
  // the document/selection, and guarantee the flag resets on every path.
  async function justifySelection() {
    _reflectBusy = true;
    try { return await justifySelectionInner(); }
    finally { _reflectBusy = false; }
  }

  async function justifySelectionInner() {
    var opts = options();
    var fontId = opts.fontMode === "nastaliq" ? "noto" : opts.fontMode;
    var mechanism = AshaarFonts.mechanismOf(fontId);

    // Hybrid qaseeda trigger: if the cursor's block belongs to a qaseeda that has
    // a stored profile, justify by applying that profile across ALL its blocks so
    // they stay consistent — instead of the free-form local justify below. Only
    // fires for tagged blocks; untagged blocks justify exactly as before.
    try {
      var qname = await getQaseedaAtSelection();
      if (qname && loadProfileStore()[qname]) { await applyProfileToQaseeda(qname); return; }
    } catch (e) { /* fall through to normal justify */ }

    // "Let Word fill it": native Word kashida (jc) instead of manual tatweel
    // insertion / spacing math. Entirely different code path — skip the
    // probe/calibrate/tatweel machinery below and delegate.
    if (opts.justifyMode === "css") { await justifySelectionWordFill(opts); return; }

    // Fallback font from the pane — used only when a cell reports no explicit font.
    var fbMode = opts.fontMode === "nastaliq" ? "noto" : opts.fontMode;
    var fallbackName = AshaarFonts.wordNameOf(fbMode) || "Times New Roman";
    var doKashida = opts.justifyMode === "kashida" || opts.justifyMode === "spacing";
    var CELL_MARGIN_PT = 5.76; // Word default cell side margin (0.08") reserved for text
    var debug = !!(debugMode && debugMode.checked);
    var diags = [];

    setMessage("Justifying…");

    // Kashida mechanism is resolved PER CELL/RUN from each run's REAL font
    // (see the generic run-aware path below), NOT from the pane dropdown.
    // Mehr (tatweel) and Jameel (font-swap) still get their explicit
    // dropdown-driven branches; every other pane selection ("Document default",
    // Noto, Gulzar, Arabic serif) is "whitespace" as an id but falls through to
    // the generic path, which decides kashida-vs-spacing from the actual font
    // of each run: arbitrary Arabic fonts (e.g. Fatemi Maqala) run the generic
    // tatweel engine; true whitespace-shaping fonts (Noto/Gulzar/Scheherazade)
    // fall back to spacing. There is deliberately no blanket dropdown-based
    // downgrade here — it forced "document" and every unrecognised font to
    // spacing, which is the regression this restores.

    await withWord(async function (context) {
      var selection = context.document.getSelection();

      // Find enclosing Ashaar Poem content control (the poem is the calibration unit)
      var cc = selection.parentContentControlOrNullObject;
      cc.load("title,tag");
      await context.sync();

      var workRange = (!cc.isNullObject && cc.title === "Ashaar Poem")
        ? cc.getRange() : selection;

      // Persisted bandh cell-map (content/spacing tag + labels) for this block,
      // when present — one pattern per stanza table, in document order.
      var ccCells = null, ccOverrides = {};
      if (!cc.isNullObject && cc.title === "Ashaar Poem") {
        var ccPayload = AshaarWord.parseContentControlTag(cc.tag);
        ccCells = ccPayload && ccPayload.cells;
        ccOverrides = (ccPayload && ccPayload.overrides) || {};
      }

      var tables = workRange.tables;
      tables.load("items");
      await context.sync();

      if (!tables.items.length) {
        // No tables — justify plain selection text, measuring with the selection's own font.
        selection.load("text");
        selection.font.load("name,size");
        await context.sync();
        // Resolve the mechanism from the selection's REAL font: true
        // whitespace-shaping fonts (Noto/Gulzar/Scheherazade) shatter under
        // injected tatweels, so downgrade kashida→spacing for them; generic /
        // arbitrary Arabic fonts (Fatemi Maqala, …) keep kashida.
        var plainOpts = opts;
        if (opts.justifyMode === "kashida" &&
            AshaarFonts.mechanismForFontName(selection.font.name) === "whitespace") {
          plainOpts = Object.assign({}, opts, { justifyMode: "spacing" });
          setMessage("“" + (selection.font.name || "This font") +
            "” can’t stretch letters in Word — filling by spacing instead.");
        }
        if (doKashida) {
          var pc = document.createElement("canvas").getContext("2d");
          if (pc) {
            pc.font = (selection.font.size || 16) + "pt \"" + (selection.font.name || fallbackName) + "\"";
            plainOpts._justifyCtx = pc;
          }
        }
        var justifiedText = AshaarWord.justifyPlainTextBlock(stripJustification(selection.text), plainOpts);
        selection.insertText(justifiedText, Word.InsertLocation.replace);
        await context.sync();
        return;
      }

      // Load rows → cells, including each cell's REAL font name/size (not a guess).
      tables.items.forEach(function (tbl) { tbl.rows.load("items"); });
      await context.sync();
      tables.items.forEach(function (tbl) {
        tbl.rows.items.forEach(function (row) { row.cells.load("items/columnWidth"); });
      });
      await context.sync();

      var allCells = [];
      tables.items.forEach(function (tbl, ti) {
        // Prefer the persisted bandh map for this table; fall back to geometry
        // (adopted/hand-drawn tables, older v1 tags → no map).
        var pattern = ccCells && ccCells[ti];
        var perRowCounts = tbl.rows.items.map(function (row) { return row.cells.items.length; });
        var tblMap = AshaarCellMap.alignPatternToTable(perRowCounts, pattern)
          ? AshaarCellMap.buildBandhCellMap(pattern) : null;
        var cellSeq = 0;
        tbl.rows.items.forEach(function (row, ri) {
          var cols = row.cells.items.length;
          row.cells.items.forEach(function (cell, ci) {
            allCells.push(cell);
            // Harmony key + content/spacing: the label from the persisted map
            // (A1 matches A1 across bandhs) when available, else the geometric
            // signature. `__kind` lets an empty content cell stay content and a
            // tagged gap be skipped regardless of its text.
            var mapped = tblMap ? tblMap[cellSeq] : null;
            cellSeq++;
            if (mapped) {
              cell.__kind = mapped.kind;
              cell.__matKey = mapped.label || mapped.slot;
              cell.__ovKey = (mapped.kind === "content" && mapped.label)
                ? AshaarOverrides.overrideKey(ti, mapped.label) : null;
            } else {
              cell.__kind = null;
              cell.__matKey = AshaarMatrix.positionKey({ row: ri, col: ci, span: cols });
              cell.__ovKey = null;
            }
            cell.body.load("text");
            cell.body.font.load("name,size");
            // Alignment of the cell's own first paragraph — used by the
            // font-swap (Jameel) path to preserve the misra's visual side
            // (sadr/ajuz/solo) when it rebuilds the cell via OOXML.
            cell.body.paragraphs.load("alignment");
          });
        });
      });
      await context.sync();

      // Split each cell into word-ranges so justify can read a font per word and
      // rebuild the cell as an ordered list of runs (run-aware justification).
      allCells.forEach(function (cell) {
        cell.__wordRanges = cell.body.getRange().getTextRanges([" "], true);
        cell.__wordRanges.load("items");
      });
      await context.sync();
      allCells.forEach(function (cell) {
        cell.__wordRanges.items.forEach(function (wr) {
          wr.load("text");
          wr.font.load("name,size,bold,italic");
        });
      });
      await context.sync();

      // Representative font taken from the cells themselves (fall back to the pane).
      var repName = fallbackName, repSize = 16;
      for (var ci = 0; ci < allCells.length; ci++) {
        var rf = allCells[ci].body.font;
        if (rf && rf.name) { repName = rf.name; if (rf.size) repSize = rf.size; break; }
      }

      // Content width = cell width minus the side margins Word reserves for text.
      function contentPx(cell) {
        return Math.max(1, (cell.columnWidth || 0) - 2 * CELL_MARGIN_PT) * 96 / 72;
      }

      // The cell's own paragraph alignment ("Right"/"Left"/"Centered"/…) maps
      // to the "right"/"left"/"center" jc the font-swap path needs when it
      // rebuilds the cell as fresh OOXML — so the misra keeps its visual side.
      function cellAlignOf(cell) {
        var p0 = cell.body.paragraphs.items && cell.body.paragraphs.items[0];
        var al = p0 && p0.alignment;
        if (al === "Right") return "right";
        if (al === "Left") return "left";
        return "center";
      }

      // Build the measurement canvas with the REAL font + size.
      var canvasCtx = null;
      if (doKashida) {
        var c = document.createElement("canvas").getContext("2d");
        if (c) { c.font = repSize + "pt \"" + repName + "\""; canvasCtx = c; opts._justifyCtx = c; }
      }

      // Ensure a bundled @font-face (e.g. FatemiMaqala) finishes loading before we
      // measure, so the canvas measures the same outlines Word renders. @font-face
      // fonts load lazily on first use; this forces the load and awaits it.
      if (canvasCtx && typeof document !== "undefined" && document.fonts && document.fonts.load) {
        try { await document.fonts.load(repSize + "pt \"" + repName + "\""); } catch (e) {}
      }

      // Jameel font-swap also measures fasls in the Kasheeda (wide) face on
      // this same canvas — force that @font-face to finish loading too, or
      // measureText silently falls back to a substitute font and corrupts
      // selectSwapRuns' gain ranking (wrong fasls get swapped).
      // Both font-swap (Jameel) and tatweel (Mehr) measure a specific Arabic
      // w:cs face on this canvas, not repName (the cell's reported Latin/hAnsi
      // font). If that face isn't force-loaded, measureText silently falls back
      // to a substitute and the elongation measures ~zero width — so Mehr's
      // trailing tatweel never registers as wider (no final tatweels selected)
      // and Jameel's gain ranking picks the wrong fasls.
      if (canvasCtx && (mechanism === "font-swap" || mechanism === "tatweel") &&
          typeof document !== "undefined" && document.fonts && document.fonts.load) {
        var bName = AshaarFonts.wordNameOf(fontId);
        if (bName) { try { await document.fonts.load(repSize + "pt \"" + bName + "\""); } catch (e) {} }
        if (mechanism === "font-swap") {
          var kName = AshaarFonts.kasheedaNameOf(fontId);
          if (kName) { try { await document.fonts.load(repSize + "pt \"" + kName + "\""); } catch (e) {} }
        }
      }

      // Auto-fit (in place): widen each table's columns so the widest misra has
      // kashida headroom, then justify into the new widths. Uses the desktop-only
      // TableColumn API (WordApiDesktop 1.3); on hosts without it, justify proceeds
      // at the current widths (no resize).
      var canResize = (typeof Office !== "undefined" && Office.context && Office.context.requirements
        && Office.context.requirements.isSetSupported
        && Office.context.requirements.isSetSupported("WordApiDesktop", "1.3"));
      if (opts.autoFitWidth && canvasCtx && canResize) {
        var sectionA = context.document.sections.getFirst();
        sectionA.load("pageLayout/width,pageLayout/leftMargin,pageLayout/rightMargin");
        await context.sync();
        var plA = sectionA.pageLayout;
        var pagePt = plA && plA.width ? (plA.width - (plA.leftMargin || 0) - (plA.rightMargin || 0)) : 468;
        var kOn = (opts.justifyMode === "kashida" || opts.justifyMode === "spacing") && Number(opts.tatweelCount || 0) > 0;
        var headroom = kOn ? 0.9 : 0.98;

        // Per table: the scale needed so the tightest cell gains headroom, capped at page width.
        var scaleByTable = tables.items.map(function (tbl) {
          var maxScale = 1, tableWpt = 0;
          tbl.rows.items.forEach(function (row, ri) {
            row.cells.items.forEach(function (cell) {
              if (ri === 0) tableWpt += (cell.columnWidth || 0);
              var t = stripJustification(cell.body.text || "").replace(/\s+/g, " ").trim();
              if (!t) return;
              var cf = cell.body.font;
              canvasCtx.font = ((cf && cf.size) || repSize) + "pt \"" + ((cf && cf.name) || repName) + "\"";
              var colWpx = (cell.columnWidth || 0) * 96 / 72;
              if (colWpx > 0) maxScale = Math.max(maxScale, canvasCtx.measureText(t).width / (headroom * colWpx));
            });
          });
          if (tableWpt > 0 && tableWpt * maxScale > pagePt) maxScale = pagePt / tableWpt;
          return maxScale;
        });

        if (scaleByTable.some(function (s) { return s > 1.01; })) {
          var colSets = tables.items.map(function (tbl, i) {
            if (scaleByTable[i] <= 1.01) return null;
            var cols = tbl.columns; cols.load("items/width"); return cols;
          });
          await context.sync();
          colSets.forEach(function (cols, i) {
            if (!cols) return;
            cols.items.forEach(function (col) { col.width = Math.round(col.width * scaleByTable[i] * 100) / 100; });
          });
          await context.sync();
          // Re-read cell widths (now changed) so justify targets the resized columns.
          tables.items.forEach(function (tbl) {
            tbl.rows.items.forEach(function (row) { row.cells.load("items/columnWidth"); });
          });
          await context.sync();
        }
      } else if (!opts.autoFitWidth && canvasCtx && canResize && opts.tableWidthPct) {
        // Table-width % applied IN PLACE (no rebuild): scale each table's columns
        // so its total width = tableWidthPct% of the page content width. Mirrors
        // how insert treats the slider when auto-fit is off, but on the EXISTING
        // table — so a width change no longer requires copy-and-replace. Uniform
        // scale preserves the gap:content proportions (layout shape intact).
        var sectionW = context.document.sections.getFirst();
        sectionW.load("pageLayout/width,pageLayout/leftMargin,pageLayout/rightMargin");
        await context.sync();
        var plW = sectionW.pageLayout;
        var pageW = plW && plW.width ? (plW.width - (plW.leftMargin || 0) - (plW.rightMargin || 0)) : 468;
        var targetW = Math.max(1, (Number(opts.tableWidthPct) / 100) * pageW);
        var colSetsW = tables.items.map(function (tbl) { var c = tbl.columns; c.load("items/width"); return c; });
        await context.sync();
        var didResize = false;
        colSetsW.forEach(function (cols) {
          var cur = 0; cols.items.forEach(function (col) { cur += (col.width || 0); });
          if (cur <= 0) return;
          var scale = targetW / cur;
          if (Math.abs(scale - 1) < 0.005) return;
          cols.items.forEach(function (col) { col.width = Math.round(col.width * scale * 100) / 100; });
          didResize = true;
        });
        if (didResize) {
          await context.sync();
          tables.items.forEach(function (tbl) {
            tbl.rows.items.forEach(function (row) { row.cells.load("items/columnWidth"); });
          });
          await context.sync();
        }
      }

      // Probe + calibrate using the real font/size and content widths.
      var fontProfile = null;
      if (canvasCtx && typeof AshaarTune !== "undefined") {
        try { fontProfile = await AshaarTune.probeFont({ fontFamily: repName, fontSize: 64 }); }
        catch (e) { /* degrade gracefully */ }
      }
      if (fontProfile) opts._fontProfile = fontProfile;

      var calibParams = { targetFill: 0.92 };
      if (fontProfile) calibParams.fontQualityBoost = 1.8;
      if (canvasCtx && typeof AshaarTune !== "undefined") {
        var lineTexts = [];
        var totalPx = 0, n = 0;
        allCells.forEach(function (cell) {
          var t = stripJustification(cell.body.text || "").replace(/[\r\n]+/g, " ").trim();
          if (t) lineTexts.push(t);
          if (cell.columnWidth > 0) { totalPx += contentPx(cell); n++; }
        });
        var avgPx = n ? totalPx / n : 300;
        if (lineTexts.length) {
          try {
            var session = await AshaarTune.calibrate({
              texts: lineTexts, fontFamily: repName, fontSize: repSize,
              containerWidth: avgPx, mode: "poetry", fontProfile: fontProfile, iterations: 50
            });
            calibParams = Object.assign({}, session.params);
            if (fontProfile) calibParams.fontQualityBoost = calibParams.fontQualityBoost || 1.8;
          } catch (e) { /* keep defaults */ }
        }
      }

      // Canvas font shorthand for one run: "[italic] [bold] Npt \"Family\"".
      function runFontStr(name, size, bold, italic) {
        return (italic ? "italic " : "") + (bold ? "bold " : "") +
          ((size || repSize)) + "pt \"" + (name || repName) + "\"";
      }
      // Micro-space glyph used to realize word-spacing in Word (text-mutating).
      var MICRO_SPACE = " "; // hair space
      if (canvasCtx) {
        canvasCtx.font = runFontStr(repName, repSize, false, false);
        if (canvasCtx.measureText(MICRO_SPACE).width <= 0) MICRO_SPACE = " "; // thin space
      }

      // Force-load EVERY distinct run font across all cells before measuring —
      // not just repName. A font the WebView CAN see (system-exposed, or a
      // bundled/uploaded @font-face) loads lazily on first use; without this,
      // measureText in a mixed-font cell silently falls back to a substitute
      // for the runs whose face isn't loaded yet, so those runs' metrics are
      // wrong. This is what lets dual-accessible fonts auto-measure correctly;
      // fonts the sandbox can't reach are supplied via the Custom-fonts
      // uploader, which registers an @font-face loaded the exact same way.
      if (canvasCtx && typeof document !== "undefined" && document.fonts && document.fonts.load) {
        var faceStrs = {};
        allCells.forEach(function (cell) {
          (cell.__wordRanges.items || []).forEach(function (wr) {
            var f = wr.font;
            var nm = (f && f.name) || repName, sz = (f && f.size) || repSize;
            faceStrs[runFontStr(nm, sz, !!(f && f.bold), !!(f && f.italic))] = true;
            // A font-swap font (Jameel) also measures fasls in its wider Kasheeda
            // face — load that too, or measureText falls back to a substitute and
            // mis-ranks the swaps. Needed under any dropdown (incl. Document
            // default), since per-cell dispatch may pick font-swap here.
            var kn = AshaarFonts.descriptorForFontName(nm).kasheedaName;
            if (kn) faceStrs[runFontStr(kn, sz, false, false)] = true;
          });
        });
        var faceLoads = [];
        Object.keys(faceStrs).forEach(function (s) { faceLoads.push(document.fonts.load(s).catch(function () {})); });
        try { await Promise.all(faceLoads); } catch (e) {}
      }

      // Natural-width matrix (harmony): the longest tatweel-free width per grid
      // position across every content cell in the work range. Natural-fit fills
      // each cell up to its position's Wpos (φ=1 pushes further, to the edge).
      var fillMode = opts.fillMode === "cell-fit" ? "cell-fit" : "natural-fit";
      var matrixCells = [];
      allCells.forEach(function (cell) {
        var base = stripJustification(cell.body.text || "").replace(/\s+/g, " ").trim();
        var isContent = cell.__kind === "content" || (cell.__kind == null && AshaarMatrix.isContentCell(base));
        if (!isContent) return; // tagged spacing (even with stray text) excluded from the matrix
        var mf = cell.body.font;
        var mnm = (mf && mf.name) || repName, msz = (mf && mf.size) || repSize;
        var natPx = 0;
        if (canvasCtx) { canvasCtx.font = runFontStr(mnm, msz, false, false); natPx = canvasCtx.measureText(base).width; }
        cell.__natPx = natPx;
        matrixCells.push({ key: cell.__matKey, natural: natPx });
      });
      var widthMatrix = AshaarMatrix.buildMatrix(matrixCells);

      // Phase 1 (pure, no sync): rebuild each cell as an ordered list of style
      // runs and justify measuring each run in its OWN font. Produces per-cell
      // write plans consumed in phase 2.
      var plans = [];
      allCells.forEach(function (cell) {
        var current = (cell.body.text || "").trim();
        if (cell.__kind === "spacing") return; // structural gap — never justified
        if (!stripJustification(current)) return;
        var colPx = contentPx(cell);

        // Per-cell override (SP2): strength / target width / cap-lift deviations
        // for this one cell, merged onto the block's justify defaults.
        var cellOv = cell.__ovKey ? ccOverrides[cell.__ovKey] : null;
        var resolved = AshaarOverrides.resolveCellOverride({ strength: opts.tatweelCount, fillMode: fillMode }, cellOv);
        var cellPhi = AshaarWord.strengthToElongationShare(resolved.strength);
        var cellMaxPos = AshaarWord.strengthToMaxPositions(resolved.strength);
        var cellCapEm = resolved.capEm != null ? resolved.capEm : undefined;

        // Resolve THIS cell's mechanism from its OWN real font (per-cell
        // dispatch), not the pane dropdown — so any dropdown (incl. Document
        // default) routes each cell to its font's correct mechanism instead of
        // shattering Jameel/Mehr with generic tatweels. The dropdown font is
        // only the fallback when a cell reports no resolvable font.
        var cellFontName = (cell.body.font && cell.body.font.name) || repName;
        var cellDesc = AshaarFonts.descriptorForFontName(cellFontName);
        var cellMech = cellDesc.mechanism;

        // Jameel font-swap: measure each fasl (connected segment) in the base
        // vs Kasheeda face, greedily swap the highest-gain fasls to the wider
        // face until the misra fills the column, and rebuild the cell as
        // OOXML with a per-run w:cs. Its own path — no word-range/tatweel/
        // spacing handling applies, so it returns before that machinery.
        if (cellMech === "font-swap") {
          if (!canvasCtx || colPx <= 0) return; // no measurement context — leave the cell as-is
          var cellAlign = cellAlignOf(cell);
          var wideCss = "\"" + (cellDesc.kasheedaName || repName) + "\"";
          var baseCss = "\"" + (cellDesc.wordName || repName) + "\"";
          var fss = AshaarKashidaFontswap.splitSpans(stripJustification(current));
          var wb = [], ww = [];
          fss.forEach(function (s) {
            canvasCtx.font = repSize + "pt " + baseCss; wb.push(canvasCtx.measureText(s).width);
            canvasCtx.font = repSize + "pt " + wideCss; ww.push(canvasCtx.measureText(s).width);
          });
          var jNatural = wb.reduce(function (a, b) { return a + b; }, 0);
          if (fillMode === "cell-fit") {
            // Cell-fit: swap fasls up to the φ elongation budget (no buffer),
            // then let Word distribute the residual to the true edge.
            var jBudget = AshaarMatrix.cellFitBudget(jNatural, colPx, cellPhi);
            var jSelC = AshaarKashidaFontswap.selectSwapRuns(fss, wb, ww, jBudget);
            var jRunsC = jSelC.runs.map(function (r) {
              return { text: r.text, csName: r.swap ? (cellDesc.kasheedaName || repName) : (cellDesc.wordName || repName), sizePt: repSize };
            });
            plans.push({ cell: cell, ooxml: AshaarWord.misraDistributeXml(jRunsC, repSize) });
            return;
          }
          // Natural-fit: fill to the position's matrix width (φ pushes toward the
          // buffered edge); capped hair-spaces backfill what the swaps miss.
          var jReach = colPx - 0.28 * repSize * 96 / 72;
          var jWpos = widthMatrix[cell.__matKey] || jNatural;
          var jTarget = (resolved.widthPt != null) ? resolved.widthPt * 96 / 72
            : AshaarMatrix.naturalFitTarget(jWpos, jReach, cellPhi);
          var sel = AshaarKashidaFontswap.selectSwapRuns(fss, wb, ww, jTarget);
          // Hybrid fill: font-swap elongation undershoots (only fasls with a
          // Kasheeda variant widen) — close the residual with capped hair-spaces
          // in the inter-word gap runs. Accept-short if the cap binds.
          var jGaps = 0;
          for (var jgi = 0; jgi < sel.runs.length; jgi++) { if (sel.runs[jgi].text === " ") jGaps++; }
          canvasCtx.font = repSize + "pt " + baseCss;
          var jSpacePx = canvasCtx.measureText(MICRO_SPACE).width || 1;
          var jn = AshaarResidual.capMicroSpaces(jTarget - sel.fill * jTarget, jGaps, jSpacePx, repSize * 96 / 72, cellCapEm);
          var jRuns = AshaarResidual.injectSpaceRuns(sel.runs, jn, MICRO_SPACE);
          var swapXml = AshaarWord.runsToMisraXml(jRuns, cellAlign, opts, repSize);
          plans.push({ cell: cell, ooxml: swapXml });
          return; // handled — skip the tatweel/spacing paths for this cell
        }

        // Mehr tatweel: DISCRETE trailing elongation. Mehr renders a clean
        // kashida only from ONE trailing tatweel after a word ending in a
        // whitelisted final letter (medial U+0640 is zero-width on the canvas
        // we measure with; Word-native highKashida does nothing for Mehr). So
        // Mehr fits by the SAME discrete subset-selection as Jameel — choose
        // which eligible words get a trailing tatweel. Single-font text output.
        if (cellMech === "tatweel" && opts.justifyMode === "kashida") {
          if (!canvasCtx || colPx <= 0) return;
          var mehrFont = repSize + "pt \"" + (cellDesc.wordName || repName) + "\"";
          var mRules = cellDesc.tatweelRules || {};
          var isoSet = {}, finSet = {};
          (mRules.isolatedInto || []).forEach(function (c) { isoSet[c] = true; });
          (mRules.finalInto || []).forEach(function (c) { finSet[c] = true; });
          var mline = stripJustification(current);
          var mparts = mline.split(" "), mtoks = [];
          mparts.forEach(function (wd, i) { if (i) mtoks.push(" "); mtoks.push(wd); });
          // Form-aware: trailing tatweel only on allowed isolated/final letters.
          var melong = mtoks.map(function (t) { return t !== " " ? AshaarWord.mehrElongate(t, isoSet, finSet) : t; });
          var mwb = [], mww = [];
          canvasCtx.font = mehrFont;
          for (var mi = 0; mi < mtoks.length; mi++) { mwb.push(canvasCtx.measureText(mtoks[mi]).width); mww.push(canvasCtx.measureText(melong[mi]).width); }
          var mNatural = mwb.reduce(function (a, b) { return a + b; }, 0);
          if (fillMode === "cell-fit") {
            var mBudget = AshaarMatrix.cellFitBudget(mNatural, colPx, cellPhi);
            var mselC = AshaarKashidaFontswap.selectSwapRuns(mtoks, mwb, mww, mBudget);
            var moutC = mselC.runs.map(function (r, i) { return (r.swap && mww[i] > mwb[i]) ? melong[i] : mtoks[i]; }).join("");
            plans.push({ cell: cell, ooxml: AshaarWord.misraDistributeXml([{ text: moutC, csName: cellDesc.wordName || repName, sizePt: repSize }], repSize) });
            return;
          }
          var mReach = colPx - 0.28 * repSize * 96 / 72;
          var mWpos = widthMatrix[cell.__matKey] || mNatural;
          var mTarget = (resolved.widthPt != null) ? resolved.widthPt * 96 / 72
            : AshaarMatrix.naturalFitTarget(mWpos, mReach, cellPhi);
          var msel = AshaarKashidaFontswap.selectSwapRuns(mtoks, mwb, mww, mTarget);
          var mout = msel.runs.map(function (r, i) { return (r.swap && mww[i] > mwb[i]) ? melong[i] : mtoks[i]; }).join("");
          // Hybrid fill: Mehr elongates only at whitelisted word-endings, so it
          // undershoots — close the residual with capped hair-spaces at the word
          // gaps (reusing distributeMicroSpaces). Accept-short if the cap binds.
          var mGaps = mout.split(" ").length - 1;
          canvasCtx.font = mehrFont;
          var mSpacePx = canvasCtx.measureText(MICRO_SPACE).width || 1;
          var mn = AshaarResidual.capMicroSpaces(mTarget - msel.fill * mTarget, mGaps, mSpacePx, repSize * 96 / 72, cellCapEm);
          var mfinal = AshaarWord.distributeMicroSpaces([mout], mn, MICRO_SPACE)[0];
          if (mfinal !== current) plans.push({ cell: cell, flat: mfinal, align: cellAlignOf(cell) });
          return;
        }

        // Per-word style tuples from the word ranges, then coalesce to runs.
        var words = [];
        (cell.__wordRanges.items || []).forEach(function (wr) {
          var t = stripJustification(wr.text || "");
          if (!t) return;
          var f = wr.font;
          words.push({
            text: t,
            name: (f && f.name) || repName,
            size: (f && f.size) || repSize,
            bold: !!(f && f.bold),
            italic: !!(f && f.italic),
            range: wr
          });
        });
        if (!words.length) return;
        var runs = AshaarWord.coalesceRuns(words);

        // Fallback: without a measurement canvas we cannot do run-aware work —
        // justify the flattened line as before (single-font behavior).
        if (!canvasCtx || colPx <= 0) {
          var flat = AshaarWord.justifyPlainTextBlock(stripJustification(current), opts, colPx);
          if (flat !== current) plans.push({ cell: cell, flat: flat });
          return;
        }

        // Primitive runs: each carries a measure() bound to its own font.
        var primRuns = runs.map(function (r) {
          var fstr = runFontStr(r.name, r.size, r.bold, r.italic);
          return {
            text: r.text,
            fontSize: r.size,
            fontProfile: fontProfile || null,
            measure: function (s) { canvasCtx.font = fstr; return canvasCtx.measureText(s).width; }
          };
        });

        var outTexts; // per-run text to write back (null when spacing writes properties only)
        var sp = null;

        // Per-cell mechanism from the runs' REAL fonts. When kashida is chosen
        // and every run is a font the tatweel engine can elongate (generic /
        // arbitrary Arabic fonts like Fatemi Maqala), kashida-fill via
        // justifyRuns — each run measured in its OWN font, so mixed-font misras
        // stretch correctly (Task A3 Step 3). If ANY run is a whitespace-shaping
        // font (Noto/Gulzar/Scheherazade), where injected tatweels shatter the
        // shaping, or the user chose spacing, fall back to run-aware spacing.
        // (Mehr/Jameel cells never reach here — handled by the branches above.)
        var anyWhitespaceRun = runs.some(function (r) {
          return AshaarFonts.mechanismForFontName(r.name) === "whitespace";
        });
        if (opts.justifyMode === "kashida" && !anyWhitespaceRun) {
          var gNatural = primRuns.reduce(function (a, r) { return a + r.measure(r.text); }, 0);
          var gMax = { perPositionEm: 0.5, maxPositions: cellMaxPos };
          if (fillMode === "cell-fit") {
            // Cell-fit: concentrate tatweels to the φ budget (no buffer); Word's
            // distribute jc stretches the inter-word gaps to the true edge.
            var gBudgetC = AshaarMatrix.cellFitBudget(gNatural, colPx, cellPhi);
            var concC = AshaarJustify.justifyRunsConcentrated(primRuns, gBudgetC, Object.assign({}, calibParams, gMax));
            var cfRuns = concC.runs.map(function (r, i) { return { text: r.text, csName: runs[i].name, sizePt: runs[i].size }; });
            plans.push({ cell: cell, ooxml: AshaarWord.misraDistributeXml(cfRuns, repSize) });
            return;
          }
          // Natural-fit: fill to the position's matrix width; capped micro-spaces
          // backfill whatever the concentrated tatweels didn't cover — so low
          // strength = spacing-dominant, harmony baseline at φ=0.
          var gReach = colPx - 0.28 * repSize * 96 / 72;
          var gWpos = widthMatrix[cell.__matKey] || gNatural;
          var gTarget = (resolved.widthPt != null) ? resolved.widthPt * 96 / 72
            : AshaarMatrix.naturalFitTarget(gWpos, gReach, cellPhi);
          var conc = AshaarJustify.justifyRunsConcentrated(primRuns, gTarget, Object.assign({}, calibParams, gMax));
          outTexts = conc.runs.map(function (r) { return r.text; });
          var gGaps = primRuns.reduce(function (a, r) { return a + (r.text.split(" ").length - 1); }, 0);
          canvasCtx.font = runFontStr(repName, repSize, false, false);
          var gSpacePx = canvasCtx.measureText(MICRO_SPACE).width || 1;
          var gN = AshaarResidual.capMicroSpaces(gTarget - conc.achievedPx, gGaps, gSpacePx, repSize * 96 / 72, cellCapEm);
          outTexts = AshaarWord.distributeMicroSpaces(outTexts, gN, MICRO_SPACE);
        } else {
          // spacing/scale: single wordSpacing + uniform fontScale from run-aware widths.
          sp = AshaarJustify.computeRunSpacing(primRuns, colPx, calibParams);
          var gaps = runs.reduce(function (a, r) { return a + (r.text.split(" ").length - 1); }, 0);
          canvasCtx.font = runFontStr(repName, repSize, false, false);
          var spaceGlyphPx = canvasCtx.measureText(MICRO_SPACE).width || 1;
          var n = Math.max(0, Math.round(sp.wordSpacing * gaps / spaceGlyphPx));
          outTexts = AshaarWord.distributeMicroSpaces(runs.map(function (r) { return r.text; }), n, MICRO_SPACE);
        }

        if (debug) {
          var natSum = 0, finSum = 0, twCount = 0;
          primRuns.forEach(function (pr, i) { natSum += pr.measure(runs[i].text); finSum += pr.measure(outTexts[i]); });
          outTexts.forEach(function (t) { twCount += (t.match(/ـ/g) || []).length; });
          diags.push({
            i: diags.length,
            font: runs.length + " run(s), " + repSize + "pt " + repName,
            res: fontAvailable(runs[0].name) ? "yes" : "NO",
            colPx: Math.round(colPx),
            colIn: (colPx / 96).toFixed(2),
            nat: Math.round(natSum),
            target: Math.round(colPx * (calibParams.targetFill || 1)),
            fin: Math.round(finSum),
            fill: colPx ? Math.round(finSum / colPx * 100) : 0,
            tw: twCount + (sp ? " ws" + sp.wordSpacing + " x" + sp.fontScale : ""),
            cap: runs.reduce(function (a, r) { return a + r.text.replace(/\s/g, "").length; }, 0),
            text: runs.map(function (r) { return r.text; }).join(" ").slice(0, 14)
          });
        }

        // Each run's justified text must split 1:1 back onto its source word
        // ranges (tatweels/micro-spaces never add ASCII spaces). If that ever
        // fails, route the cell to the flattened path instead of a partial write.
        var alignedOk = runs.every(function (r, i) {
          return outTexts[i].split(" ").length === r.refs.length;
        });
        if (!alignedOk) { plans.push({ cell: cell, flat: outTexts.join(" "), align: cellAlignOf(cell) }); return; }

        plans.push({ cell: cell, runs: runs, outTexts: outTexts, sp: sp, align: cellAlignOf(cell) });
      });

      // Phase 2 (write): one context.sync() per cell so a range failure on one
      // cell falls back to a flattened whole-cell replace without aborting the
      // batch (the run-aware write can only error at sync, not synchronously).
      // Map a cell's own alignment → an Office enum. Applied on flat/run-aware
      // writes so re-justifying a cell that was previously Cell-fit (paragraph
      // jc=distribute) clears the distribute — Office.js has no "distribute"
      // Alignment, so we re-assert the cell's intended side.
      function officeAlign(a) {
        if (a === "right") return Word.Alignment.right;
        if (a === "left") return Word.Alignment.left;
        return Word.Alignment.centered;
      }

      var changed = 0;
      for (var pi = 0; pi < plans.length; pi++) {
        var p = plans[pi];
        if (p.ooxml) {
          try {
            p.cell.body.clear();
            p.cell.body.insertOoxml(AshaarWord.wrapOoxml(p.ooxml), Word.InsertLocation.replace);
            await context.sync();
            changed++;
          } catch (e) {
            if (debug) diags.push({ i: diags.length, font: "OOXML-FAIL", text: (e && e.message || "").slice(0, 14) });
          }
          continue;
        }
        if (p.flat != null) {
          var flatPara = p.cell.body.paragraphs.getFirst();
          flatPara.insertText(p.flat, Word.InsertLocation.replace);
          if (p.align) flatPara.alignment = officeAlign(p.align);
          await context.sync();
          changed++;
          continue;
        }
        try {
          var cellChanged = false;
          p.runs.forEach(function (r, i) {
            // outTexts[i] splits 1:1 onto the run's original word ranges
            // (validated in phase 1) — write each word range independently
            // (disjoint; no union/expand needed).
            var pieces = p.outTexts[i].split(" ");
            r.refs.forEach(function (w, j) {
              if (p.sp && p.sp.fontScale !== 1) { w.range.font.size = r.size * p.sp.fontScale; cellChanged = true; }
              if (pieces[j] !== w.text) { w.range.insertText(pieces[j], Word.InsertLocation.replace); cellChanged = true; }
            });
          });
          if (p.align) { p.cell.body.paragraphs.getFirst().alignment = officeAlign(p.align); cellChanged = true; }
          if (cellChanged) { await context.sync(); changed++; }
        } catch (e) {
          // Queued range write failed at sync (or count mismatch) — flatten.
          p.cell.body.paragraphs.getFirst().insertText(p.outTexts.join(" "), Word.InsertLocation.replace);
          await context.sync();
          changed++;
          if (debug) diags.push({ i: diags.length, font: "RANGE-FALLBACK", text: (e && e.message || "").slice(0, 14) });
        }
      }

      setMessage("Justified " + changed + " cell(s) across " + tables.items.length + " table(s).");
      if (debug) renderDebug(diags);
    });
  }

  // Adopt an existing Word table of poetry: read its cells, reconstruct the
  // canonical Ashaar source, and (by default) replace the table in place with a
  // managed, content-controlled Ashaar block. Uses Word.run directly so we keep
  // control of messaging (withWord forces a "Done." message).
  async function adoptTable() {
    if (typeof Word === "undefined") {
      setMessage("Open this task pane inside Word to adopt a table.");
      return;
    }
    var reviewOnly = document.getElementById("adopt-review").checked;
    var dirChoice = document.getElementById("adopt-direction").value;     // auto | rtl | ltr
    var direction = dirChoice === "ltr" ? "ltr" : "rtl";                   // auto → rtl
    var scope = document.getElementById("adopt-scope").value;             // cursor | selection
    var source = "";

    try {
      await Word.run(async function (context) {
        var selection = context.document.getSelection();
        var targetTables;

        if (scope === "selection") {
          var tbls = selection.tables;
          tbls.load("items");
          await context.sync();
          targetTables = tbls.items;
          if (!targetTables.length) { setMessage("Select one or more tables to adopt."); return; }
        } else {
          var t = selection.parentTableOrNullObject;
          t.load("rows");
          await context.sync();
          if (t.isNullObject) { setMessage("Place the cursor inside a table to adopt it."); return; }
          targetTables = [t];
        }

        targetTables.forEach(function (tbl) { tbl.rows.load("items"); });
        await context.sync();
        targetTables.forEach(function (tbl) {
          tbl.rows.items.forEach(function (row) { row.cells.load("items"); });
        });
        await context.sync();
        targetTables.forEach(function (tbl) {
          tbl.rows.items.forEach(function (row) {
            row.cells.items.forEach(function (cell) { cell.body.load("text"); });
          });
        });
        await context.sync();

        // Each table → one stanza; multiple selected tables → stanza-separated.
        source = targetTables.map(function (tbl) {
          var rows = tbl.rows.items.map(function (row) {
            return row.cells.items.map(function (cell) { return cell.body.text || ""; });
          });
          return AshaarTableAdopt.adoptTableToSource(rows, { direction: direction });
        }).filter(function (s) { return s.trim(); }).join("\n\n");

        if (!source.trim()) { setMessage("That table didn't contain any text to adopt."); return; }

        // Put the selection on the content we'll replace, so insertPoem(true) targets it.
        var range = (scope === "cursor") ? targetTables[0].getRange() : selection.getRange();
        range.select();
        await context.sync();
      });
    } catch (e) {
      setMessage("Adopt failed: " + (e && e.message ? e.message : String(e)));
      return;
    }

    if (!source.trim()) return; // a friendly message was already shown

    // Show the recovered source + preview (transparent, editable).
    setMode("convert");
    input.value = source;
    renderPreview();

    if (reviewOnly) {
      setMessage("Adopted the table into the editor. Review the text, then click Replace Selection.");
      return;
    }

    // One-click: replace the selected table with the regenerated Ashaar block.
    await insertPoem(true);
    setMessage("Table adopted and replaced with a formatted Ashaar block.");
  }

  // ── Template persistence helpers ───────────────────────────────────────────

  function loadTemplates() {
    try { return JSON.parse(localStorage.getItem("ashaar-templates") || "[]"); }
    catch (e) { return []; }
  }

  function saveTemplates(templates) {
    localStorage.setItem("ashaar-templates", JSON.stringify(templates));
  }

  function renderTemplateList() {
    var templates = loadTemplates();
    templateList.innerHTML = "";
    if (!templates.length) {
      var opt = document.createElement("option");
      opt.value = "";
      opt.textContent = "— no saved templates —";
      opt.disabled = true;
      opt.selected = true;
      templateList.appendChild(opt);
      return;
    }
    templates.forEach(function (t) {
      var opt = document.createElement("option");
      opt.value = t.id;
      opt.textContent = t.name;
      templateList.appendChild(opt);
    });
  }

  // ── Drop bare 12-column grid ───────────────────────────────────────────────

  async function insertBareGrid() {
    await withWord(async function (context) {
      var section = context.document.sections.getFirst();
      section.load("pageLayout/width,pageLayout/leftMargin,pageLayout/rightMargin");
      await context.sync();
      var pl = section.pageLayout;
      var textWidthTwips = pl && pl.width
        ? Math.round((pl.width - (pl.leftMargin || 0) - (pl.rightMargin || 0)) * 20)
        : 9360;
      var ooxml = AshaarWord.wrapOoxml(AshaarWord.generateBareGrid12Ooxml(scaledTextWidth(textWidthTwips)));
      var selection = context.document.getSelection();
      var inserted = selection.insertOoxml(ooxml, Word.InsertLocation.end);
      var control = inserted.insertContentControl();
      control.title = "Ashaar Poem";
      control.tag = AshaarWord.contentControlTag("grid12", options());
      control.appearance = "BoundingBox";
      await context.sync();
      setMessage("12-column grid inserted. Merge cells in Word, then Capture as a template.");
    });
  }

  // ── Capture selected table layout ─────────────────────────────────────────

  async function captureSelectedTableLayout() {
    var name = (templateNameInput.value || "").trim();
    if (!name) { setMessage("Enter a template name first."); return; }

    await withWord(async function (context) {
      var selection = context.document.getSelection();
      var table = selection.parentTableOrNullObject;
      table.load("rows");
      await context.sync();
      if (table.isNullObject) { setMessage("Click inside a table first, then capture."); return; }

      table.rows.load("items");
      await context.sync();
      table.rows.items.forEach(function (row) { row.cells.load("items/columnWidth"); });
      await context.sync();

      // Infer total table width (sum of first row cell widths in points)
      var firstRow = table.rows.items[0];
      var totalWidthPt = 0;
      firstRow.cells.items.forEach(function (cell) { totalWidthPt += (cell.columnWidth || 0); });
      if (totalWidthPt <= 0) { setMessage("Could not read table cell widths."); return; }

      var GRID = 12;
      var baseColPt = totalWidthPt / GRID;
      var rows = table.rows.items.map(function (row) {
        return row.cells.items.map(function (cell) {
          var span = Math.max(1, Math.min(GRID, Math.round((cell.columnWidth || baseColPt) / baseColPt)));
          return { span: span };
        });
      });

      var id = String(Date.now());
      var template = {
        id: id,
        name: name,
        columnCount: GRID,
        rows: rows,
        fontMode: fontMode.value,
        justifyMode: justifyMode.value,
        tatweelCount: Number(tatweelCount.value || 0),
        gapWidth: Number(gapWidth.value || 4)
      };

      var templates = loadTemplates();
      templates.push(template);
      saveTemplates(templates);
      renderTemplateList();
      // Select the newly saved template
      for (var i = 0; i < templateList.options.length; i++) {
        if (templateList.options[i].value === id) { templateList.selectedIndex = i; break; }
      }
      templateNameInput.value = "";
      setMessage("Template \"" + name + "\" saved.");
    });
  }

  // ── Apply saved template ───────────────────────────────────────────────────

  async function applyTemplate() {
    var id = templateList.value;
    if (!id) { setMessage("Select a template first."); return; }
    var templates = loadTemplates();
    var tmpl = null;
    for (var i = 0; i < templates.length; i++) {
      if (templates[i].id === id) { tmpl = templates[i]; break; }
    }
    if (!tmpl) { setMessage("Template not found."); return; }

    await withWord(async function (context) {
      var section = context.document.sections.getFirst();
      section.load("pageLayout/width,pageLayout/leftMargin,pageLayout/rightMargin");
      await context.sync();
      var pl = section.pageLayout;
      var textWidthTwips = pl && pl.width
        ? Math.round((pl.width - (pl.leftMargin || 0) - (pl.rightMargin || 0)) * 20)
        : 9360;

      var opts = options();
      var ooxml = AshaarWord.wrapOoxml(AshaarWord.templateToOoxml(tmpl, scaledTextWidth(textWidthTwips), opts));
      var selection = context.document.getSelection();
      var inserted = selection.insertOoxml(ooxml, Word.InsertLocation.end);
      var control = inserted.insertContentControl();
      control.title = "Ashaar Poem";
      control.tag = AshaarWord.contentControlTag("template:" + tmpl.name, opts);
      control.appearance = "BoundingBox";
      await context.sync();
      setMessage("Template \"" + tmpl.name + "\" inserted.");
    });
  }

  // ── Delete template ────────────────────────────────────────────────────────

  function deleteTemplate() {
    var id = templateList.value;
    if (!id) return;
    var templates = loadTemplates().filter(function (t) { return t.id !== id; });
    saveTemplates(templates);
    renderTemplateList();
    setMessage("Template deleted.");
  }

  // ── Export / Import ────────────────────────────────────────────────────────

  function exportTemplates() {
    var templates = loadTemplates();
    if (!templates.length) { setMessage("No templates to export."); return; }
    var json = JSON.stringify(templates, null, 2);
    var blob = new Blob([json], { type: "application/json" });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = "ashaar-templates.json";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  function importTemplates() {
    importFileInput.click();
  }

  function onImportFile(event) {
    var file = event.target.files && event.target.files[0];
    if (!file) return;
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var incoming = JSON.parse(e.target.result);
        if (!Array.isArray(incoming)) throw new Error("Expected array");
        var existing = loadTemplates();
        var existingIds = {};
        existing.forEach(function (t) { existingIds[t.id] = true; });
        var added = 0;
        incoming.forEach(function (t) {
          if (t && t.id && t.name && t.rows && !existingIds[t.id]) {
            existing.push(t);
            added++;
          }
        });
        saveTemplates(existing);
        renderTemplateList();
        setMessage("Imported " + added + " template(s).");
      } catch (err) {
        setMessage("Import failed: " + (err.message || String(err)));
      }
      importFileInput.value = "";
    };
    reader.readAsText(file);
  }

  var isBound = false;

  function bind() {
    if (isBound) return;
    isBound = true;
    (function populateFontModes() {
      var order = ["document", "arabic-serif", "noto", "mehr", "jameel", "gulzar"];
      fontMode.innerHTML = "";
      order.forEach(function (id) {
        var d = AshaarFonts.get(id);
        if (!d) return;
        var o = document.createElement("option");
        o.value = id; o.textContent = d.label;
        if (id === "document") o.selected = true;
        fontMode.appendChild(o);
      });
    })();
    autoFitWidth.addEventListener("change", function () {
      tableWidth.disabled = autoFitWidth.checked;
      renderPreview();
    });
    [input, justifyMode, layoutMode, widthMode, bandhCount, misraCount, layoutPreset, layoutSpec, fontMode, tatweelCount, gapWidth, tableWidth, autoFitWidth].forEach(function (el) {
      el.addEventListener("input", renderPreview);
      el.addEventListener("change", renderPreview);
    });
    fontMode.addEventListener("change", updateFontNote);
    layoutPreset.addEventListener("change", applyLayoutPreset);
    misraCount.addEventListener("change", applyLayoutPreset);
    modeTable.addEventListener("click", function () { setMode("table"); });
    modeConvert.addEventListener("click", function () { setMode("convert"); });
    document.getElementById("insert-structure").addEventListener("click", insertStructure);
    document.getElementById("insert-poem").addEventListener("click", function () { insertPoem(false); });
    document.getElementById("insert-tabstop").addEventListener("click", insertTabStopPoem);
    document.getElementById("replace-selection").addEventListener("click", function () { insertPoem(true); });
    document.getElementById("justify-selection").addEventListener("click", justifySelection);
    var showMapBtn = document.getElementById("show-cell-map");
    if (showMapBtn) showMapBtn.addEventListener("click", showCellMap);
    if (typeof Office !== "undefined" && Office.context && Office.context.document &&
        Office.context.document.addHandlerAsync && typeof Word !== "undefined") {
      Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChanged);
    }
    ["cell-ov-strength", "cell-ov-width", "cell-ov-cap"].forEach(function (id) {
      var el = document.getElementById(id);
      if (el) el.addEventListener("change", function () { applyCellOverride(false); });
    });
    var ovClear = document.getElementById("cell-ov-clear");
    if (ovClear) ovClear.addEventListener("click", function () { applyCellOverride(true); });
    ["slot-decor-symbol", "slot-decor-fill", "slot-decor-color", "slot-decor-fill-on"].forEach(function (id) {
      var el = document.getElementById(id);
      if (el) el.addEventListener("change", function () { applySlotDecor(false); });
    });
    var decorClear = document.getElementById("slot-decor-clear");
    if (decorClear) decorClear.addEventListener("click", function () { applySlotDecor(true); });
    var decorSaveProfile = document.getElementById("slot-decor-save-profile");
    if (decorSaveProfile) decorSaveProfile.addEventListener("click", saveSlotDecorToProfile);
    document.getElementById("re-render").addEventListener("click", reRender);
    document.getElementById("reset-justification").addEventListener("click", resetJustification);
    document.getElementById("load-selection").addEventListener("click", loadSelection);
    // Import-options (separator flexibility): auto-normalize on paste; manual overrides.
    input.addEventListener("paste", function () { setTimeout(applyImportNormalization, 0); });
    sepMode.addEventListener("change", function () {
      sepCustom.hidden = sepMode.value !== "custom";
      applyImportNormalization();
    });
    sepCustom.addEventListener("change", applyImportNormalization);
    sepPair.addEventListener("change", applyImportNormalization);
    document.getElementById("sep-apply").addEventListener("click", applyImportNormalization);
    document.getElementById("drop-grid").addEventListener("click", insertBareGrid);
    layoutViewGridBtn.addEventListener("click", function () { setLayoutView("grid"); });
    layoutViewNumbersBtn.addEventListener("click", function () { setLayoutView("numbers"); });
    layoutGridEl.addEventListener("click", onLayoutGridClick);
    document.getElementById("adopt-table").addEventListener("click", adoptTable);
    document.getElementById("capture-template").addEventListener("click", captureSelectedTableLayout);
    document.getElementById("apply-template").addEventListener("click", applyTemplate);
    document.getElementById("delete-template").addEventListener("click", deleteTemplate);
    document.getElementById("export-templates").addEventListener("click", exportTemplates);
    document.getElementById("import-templates").addEventListener("click", importTemplates);
    importFileInput.addEventListener("change", onImportFile);
    // Qaseeda profile panel
    qaseedaStrength.addEventListener("input", function () { qaseedaStrengthValue.textContent = qaseedaStrength.value; });
    qaseedaName.addEventListener("change", loadQaseedaIntoPanel);
    document.getElementById("qaseeda-assign").addEventListener("click", assignBlockToQaseeda);
    document.getElementById("qaseeda-apply").addEventListener("click", saveAndApplyQaseeda);
    document.getElementById("qaseeda-font-check").addEventListener("click", checkQaseedaFont);

    // Custom fonts: register any stored fonts before measurement, wire the UI.
    if (typeof AshaarFontStore !== "undefined") {
      AshaarFontStore.registerAll().then(refreshFontList, function () {});
    }
    if (fontUpload) fontUpload.addEventListener("change", onFontFilePicked);
    var fontAddBtn = document.getElementById("font-upload-add");
    if (fontAddBtn) fontAddBtn.addEventListener("click", addFont);
    var fontVerifyBtn = document.getElementById("font-upload-verify");
    if (fontVerifyBtn) fontVerifyBtn.addEventListener("click", verifyFontAtCursor);

    applyLayoutPreset();
    renderPreview();
    updateFontNote();
    setMode("table");
    renderTemplateList();
    populateQaseedaNames();
  }

  if (window.Office && Office.onReady) {
    Office.onReady(function (info) {
      hostStatus.textContent = info.host === Office.HostType.Word ? "Connected to Word" : "Preview mode";
      bind();
    });
    window.setTimeout(function () {
      if (!isBound) {
        hostStatus.textContent = "Browser preview mode";
        bind();
      }
    }, 1200);
  } else {
    hostStatus.textContent = "Browser preview mode";
    bind();
  }
}());
