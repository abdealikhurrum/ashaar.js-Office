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
      Ashaar.justifyEl(preview, {});
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
      for (i = 1; i <= count; i++) rows.push(new Array(Math.max(0, count - i) + 1).join("  ") + i);
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
        return row.map(function (cell) {
          return AshaarWord.justifyPlainTextBlock(cell.text || "", opts);
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
      var html;
      var diagLines = [];

      function diag(msg) {
        console.log("[Ashaar]", msg);
        diagLines.push(msg);
      }

      diag("layoutTablesForPoem? " + (typeof AshaarWord.layoutTablesForPoem));
      diag("source chars: " + source.length);

      if (opts.layoutSpec && opts.layoutSpec.trim()) {
        diag("path: layoutSpec");
        var tables = AshaarWord.layoutTablesForText(source, opts);
        if (await insertNativeLayoutTables(context, tables, opts, source, replaceSelection)) return;
        html = "";
      } else {
        var poemTables = null;
        try {
          poemTables = AshaarWord.layoutTablesForPoem(source, opts, Ashaar);
        } catch (e) {
          diag("layoutTablesForPoem threw: " + e.message);
        }

        if (poemTables) {
          diag("path: native (" + poemTables.length + " tables, cols: " + poemTables.map(function(t){return t.columnCount;}).join(",") + ")");
          if (await insertNativeLayoutTables(context, poemTables, opts, source, replaceSelection)) return;
          diag("insertNativeLayoutTables returned false");
        } else {
          diag("path: html (poemTables was null)");
        }

        try {
          html = AshaarWord.renderForWord(source, opts, Ashaar);
          diag("renderForWord ok, html length: " + html.length);
        } catch (error) {
          diag("renderForWord threw: " + error.message);
          html = "";
        }
      }

      setMessage(diagLines.join(" | "));

      if (!html) {
        setMessage(diagLines.join(" | ") + " | no html");
        return;
      }
      var selection = context.document.getSelection();
      var inserted = selection.insertHtml(html, replaceSelection ? Word.InsertLocation.replace : Word.InsertLocation.end);
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
      if (opts.layoutSpec && opts.layoutSpec.trim()) {
        var tables = AshaarWord.layoutTablesForTemplate(opts);
        if (await insertNativeLayoutTables(context, tables, opts, "template", false)) return;
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
      await context.sync();
      var pl = section.pageLayout;
      var textWidthTwips = Math.round((pl.width - pl.leftMargin - pl.rightMargin) * 20);

      var opts = options();
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

  async function justifySelection() {
    await withWord(async function (context) {
      var selection = context.document.getSelection();
      selection.load("text");
      await context.sync();
      var justified = AshaarWord.justifyPlainTextBlock(selection.text, options());
      selection.insertText(justified, Word.InsertLocation.replace);
      await context.sync();
    });
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
    applyLayoutPreset();
    renderPreview();
    setMode("table");
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
