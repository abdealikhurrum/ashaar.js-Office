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
  var templateNameInput = document.getElementById("template-name");
  var templateList = document.getElementById("template-list");
  var importFileInput = document.getElementById("import-file");

  function options() {
    return {
      justifyMode: justifyMode.value,
      justify: justifyMode.value === "none" ? false : justifyMode.value,
      layoutMode: layoutMode.value,
      layout: layoutMode.value,
      widthMode: widthMode.value,
      bandhCount: Number(bandhCount.value || 1),
      misraCount: Number(misraCount.value || 4),
      misraPattern: layoutPreset.value,
      layoutSpec: (!tablePanel.hidden) ? layoutSpec.value : "",
      fontMode: fontMode.value,
      tatweelCount: Number(tatweelCount.value || 0),
      gapWidth: Number(gapWidth.value || 4)
    };
  }

  function previewFontFamily(font) {
    if (font === "nastaliq") return "\"Noto Nastaliq Urdu\", \"Jameel Noori Nastaleeq\", serif";
    if (font === "arabic-serif") return "\"Scheherazade New\", \"Amiri\", \"Times New Roman\", serif";
    return "\"Times New Roman\", serif";
  }

  function setMessage(text) {
    message.textContent = text || "";
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
    preview.className = "ashaar preview";
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
      setMessage(error && error.message ? error.message : String(error));
    }
  }

  function wordAlignment(align) {
    if (align === "left") return "Left";
    if (align === "right") return "Right";
    return "Centered";
  }

  function formatNativeLayoutTable(table, layoutTable, source, opts) {
    table.alignment = "Centered";
    table.horizontalAlignment = "Centered";
    table.styleFirstColumn = false;
    table.styleLastColumn = false;
    table.styleBandedColumns = false;
    table.styleBandedRows = false;

    layoutTable.widths.forEach(function (width, columnIndex) {
      var cell = table.getCell(0, columnIndex);
      cell.columnWidth = Math.max(8, 468 * width / 100);
    });

    layoutTable.rows.forEach(function (row, rowIndex) {
      row.forEach(function (layoutCell, columnIndex) {
        var cell = table.getCell(rowIndex, columnIndex);
        cell.horizontalAlignment = wordAlignment(layoutCell.align);
      });
    });

    var control = table.insertContentControl();
    control.title = "Ashaar Poem";
    control.tag = AshaarWord.contentControlTag(source, opts);
    control.appearance = "BoundingBox";
  }

  async function insertNativeLayoutTables(context, tables, opts, source, replaceSelection) {
    if (!tables.length) return false;

    var selection = context.document.getSelection();
    if (replaceSelection) {
      selection.insertText("", Word.InsertLocation.replace);
      await context.sync();
      selection = context.document.getSelection();
    }

    var anchor = selection;
    tables.forEach(function (layoutTable, index) {
      var values = layoutTable.rows.map(function (row) {
        return row.map(function (cell, colIdx) {
          var colWidthPx = opts._textWidthPx && layoutTable.widths && layoutTable.widths[colIdx]
            ? layoutTable.widths[colIdx] / 100 * opts._textWidthPx : 0;
          return AshaarWord.justifyPlainTextBlock(cell.text || "", opts, colWidthPx);
        });
      });
      var insertLocation = index === 0 ? Word.InsertLocation.after : Word.InsertLocation.after;
      var table = anchor.insertTable(layoutTable.rows.length, layoutTable.columnCount, insertLocation, values);
      formatNativeLayoutTable(table, layoutTable, source, opts);
      anchor = table.insertParagraph("", Word.InsertLocation.after);
    });
    await context.sync();
    return true;
  }

  async function insertPoem(replaceSelection) {
    await withWord(async function (context) {
      var opts = options();
      var source = String(input.value || "");

      var sectionP = context.document.sections.getFirst();
      sectionP.load("pageLayout/width,pageLayout/leftMargin,pageLayout/rightMargin");
      var selFontP = context.document.getSelection();
      selFontP.load("font/size,font/name");
      await context.sync();

      // pageLayout requires WordApi 1.5; fall back to US-Letter 6.5" on older builds
      var plP = sectionP.pageLayout;
      var textWidthTwips = plP && plP.width
        ? Math.round((plP.width - (plP.leftMargin || 0) - (plP.rightMargin || 0)) * 20)
        : 9360;

      if (opts.justifyMode === "kashida" || opts.justifyMode === "spacing") {
        opts._textWidthPx = textWidthTwips * 96 / 1440;
        var fontSizeP = selFontP.font.size || 12;
        var fontNameP = opts.fontMode === "nastaliq" ? "Noto Nastaliq Urdu"
                      : opts.fontMode === "arabic-serif" ? "Scheherazade New"
                      : (selFontP.font.name || "Times New Roman");
        var canvasP = document.createElement("canvas");
        var ctxP = canvasP.getContext("2d");
        if (ctxP) {
          ctxP.font = fontSizeP + "pt \"" + fontNameP + "\"";
          opts._justifyCtx = ctxP;
        }
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
      var selection = context.document.getSelection();
      var inserted = selection.insertOoxml(ooxml,
        replaceSelection ? Word.InsertLocation.replace : Word.InsertLocation.end);
      var control = inserted.insertContentControl();
      control.title = "Ashaar Poem";
      control.tag = AshaarWord.contentControlTag(source, opts);
      control.appearance = "BoundingBox";
      await context.sync();
    });
  }

  async function insertStructure() {
    await withWord(async function (context) {
      var opts = options();
      var tables = AshaarWord.layoutTablesForTemplate(opts);

      if (tables.length && tables.some(function (t) { return t.spanBased; })) {
        var section = context.document.sections.getFirst();
        section.load("pageLayout/width,pageLayout/leftMargin,pageLayout/rightMargin");
        await context.sync();
        var pl = section.pageLayout;
        var textWidthTwips = pl && pl.width
          ? Math.round((pl.width - (pl.leftMargin || 0) - (pl.rightMargin || 0)) * 20)
          : 9360;
        var ooxmlBody = tables.map(function (t) {
          return AshaarWord.templateToOoxml(t, textWidthTwips, opts);
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

      if (tables.length && await insertNativeLayoutTables(context, tables, opts, "template", false)) return;
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
        var fontName = opts.fontMode === "nastaliq" ? "Noto Nastaliq Urdu"
                     : opts.fontMode === "arabic-serif" ? "Scheherazade New"
                     : (selFont.font.name || "Times New Roman");
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
    await withWord(async function (context) {
      var selection = context.document.getSelection();
      selection.load("text");
      await context.sync();
      input.value = selection.text || input.value;
      renderPreview();
    });
  }

  // Kashida (U+0640) and the micro-spaces inserted by spacing justification
  // (hair U+200A, thin U+2009) are elongation artifacts, not part of the source
  // line. Strip them so every justification pass re-derives from the bare text —
  // making justification idempotent and reducible: re-justify any number of times
  // and the kashida count follows the current width / font / fill, never compounds.
  function stripJustification(s) {
    return String(s || "").replace(/[ـ  ]/g, "");
  }

  async function justifySelection() {
    var opts = options();

    // Build canvas context for kashida measurement
    var canvasCtx = null;
    var fontName = opts.fontMode === "nastaliq" ? "Noto Nastaliq Urdu"
                 : opts.fontMode === "arabic-serif" ? "Scheherazade New"
                 : "Times New Roman";
    if (opts.justifyMode === "kashida" || opts.justifyMode === "spacing") {
      var c = document.createElement("canvas").getContext("2d");
      if (c) {
        c.font = "16pt \"" + fontName + "\"";
        canvasCtx = c;
        opts._justifyCtx = c;
      }
    }

    // Probe font once for quality-aware kashida slot selection
    var fontProfile = null;
    if (canvasCtx && typeof AshaarTune !== "undefined") {
      try { fontProfile = await AshaarTune.probeFont({ fontFamily: fontName, fontSize: 64 }); }
      catch (e) { /* degrade gracefully */ }
    }
    if (fontProfile) opts._fontProfile = fontProfile;

    setMessage("Justifying…");

    await withWord(async function (context) {
      var selection = context.document.getSelection();

      // Find enclosing Ashaar Poem content control (the poem is the calibration unit)
      var cc = selection.parentContentControlOrNullObject;
      cc.load("title");
      await context.sync();

      // Determine the range to work on: content control if we're inside one, else selection
      var workRange = (!cc.isNullObject && cc.title === "Ashaar Poem")
        ? cc.getRange() : selection;

      // Gather tables in the work range
      var tables = workRange.tables;
      tables.load("items");
      await context.sync();

      if (!tables.items.length) {
        // No tables — justify plain selection text
        selection.load("text");
        await context.sync();
        var justifiedText = AshaarWord.justifyPlainTextBlock(stripJustification(selection.text), opts);
        selection.insertText(justifiedText, Word.InsertLocation.replace);
        await context.sync();
        return;
      }

      // Load rows → cells (columnWidth + body text) across all tables
      tables.items.forEach(function (tbl) { tbl.rows.load("items"); });
      await context.sync();
      tables.items.forEach(function (tbl) {
        tbl.rows.items.forEach(function (row) { row.cells.load("items/columnWidth"); });
      });
      await context.sync();

      var allCells = [];
      tables.items.forEach(function (tbl) {
        tbl.rows.items.forEach(function (row) {
          row.cells.items.forEach(function (cell) {
            allCells.push(cell);
            cell.body.load("text");
          });
        });
      });
      await context.sync();

      // Calibrate params treating the whole poem as one unit (using AshaarTune)
      var calibParams = { targetFill: 0.92 };
      if (fontProfile) calibParams.fontQualityBoost = 1.8;

      if (canvasCtx && typeof AshaarTune !== "undefined") {
        var lineTexts = [];
        var totalColPx = 0, colCount = 0;
        allCells.forEach(function (cell) {
          var t = stripJustification(cell.body.text || "").replace(/[\r\n]+/g, " ").trim();
          if (t) lineTexts.push(t);
          if (cell.columnWidth > 0) { totalColPx += cell.columnWidth * 96 / 72; colCount++; }
        });
        var avgColPx = colCount ? totalColPx / colCount : 300;

        if (lineTexts.length) {
          try {
            var session = await AshaarTune.calibrate({
              texts: lineTexts,
              fontFamily: fontName,
              fontSize: 16,
              containerWidth: avgColPx,
              mode: "poetry",
              fontProfile: fontProfile,
              iterations: 50
            });
            calibParams = Object.assign({}, session.params);
            if (fontProfile) calibParams.fontQualityBoost = calibParams.fontQualityBoost || 1.8;
          } catch (e) { /* keep defaults */ }
        }
      }

      // Apply justified text back to each cell
      var changed = 0;
      allCells.forEach(function (cell) {
        var current = (cell.body.text || "").trim();
        var base = stripJustification(current); // re-justify from the bare line, not prior kashidas
        if (!base) return;
        var colPx = (cell.columnWidth || 0) * 96 / 72;
        var justified;
        if (canvasCtx && colPx > 0 && opts.justifyMode === "kashida") {
          justified = AshaarJustify.justifyLine(base, colPx, canvasCtx, calibParams, fontProfile || null);
        } else {
          // spacing mode: justifyText dispatches to justifyWordSpacing via justifyPlainTextBlock
          justified = AshaarWord.justifyPlainTextBlock(base, opts, colPx);
        }
        // Compare against the CURRENT cell text so a reduction (fewer or zero
        // kashidas) is written back even when the result equals the bare base.
        if (justified !== current) {
          // Use paragraph.insertText rather than body.insertText so paragraph-level
          // properties (jc, spacing, indents — including jc="both" for spacing mode)
          // are preserved. body.insertText replaces the entire cell content including
          // those properties.
          cell.body.paragraphs.getFirst().insertText(justified, Word.InsertLocation.replace);
          changed++;
        }
      });

      await context.sync();
      setMessage("Justified " + changed + " cell(s) across " + tables.items.length + " table(s).");
    });
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
      var ooxml = AshaarWord.wrapOoxml(AshaarWord.generateBareGrid12Ooxml(textWidthTwips));
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
      var ooxml = AshaarWord.wrapOoxml(AshaarWord.templateToOoxml(tmpl, textWidthTwips, opts));
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
    [input, justifyMode, layoutMode, widthMode, bandhCount, misraCount, layoutPreset, layoutSpec, fontMode, tatweelCount, gapWidth].forEach(function (el) {
      el.addEventListener("input", renderPreview);
      el.addEventListener("change", renderPreview);
    });
    layoutPreset.addEventListener("change", applyLayoutPreset);
    misraCount.addEventListener("change", applyLayoutPreset);
    modeTable.addEventListener("click", function () { setMode("table"); });
    modeConvert.addEventListener("click", function () { setMode("convert"); });
    document.getElementById("insert-structure").addEventListener("click", insertStructure);
    document.getElementById("insert-poem").addEventListener("click", function () { insertPoem(false); });
    document.getElementById("insert-tabstop").addEventListener("click", insertTabStopPoem);
    document.getElementById("replace-selection").addEventListener("click", function () { insertPoem(true); });
    document.getElementById("justify-selection").addEventListener("click", justifySelection);
    document.getElementById("load-selection").addEventListener("click", loadSelection);
    document.getElementById("drop-grid").addEventListener("click", insertBareGrid);
    document.getElementById("capture-template").addEventListener("click", captureSelectedTableLayout);
    document.getElementById("apply-template").addEventListener("click", applyTemplate);
    document.getElementById("delete-template").addEventListener("click", deleteTemplate);
    document.getElementById("export-templates").addEventListener("click", exportTemplates);
    document.getElementById("import-templates").addEventListener("click", importTemplates);
    importFileInput.addEventListener("change", onImportFile);
    applyLayoutPreset();
    renderPreview();
    setMode("table");
    renderTemplateList();
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
