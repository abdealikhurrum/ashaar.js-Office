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
  var misraPattern = document.getElementById("misra-pattern");
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
      misraPattern: misraPattern.value,
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
    preview.innerHTML = Ashaar.renderText(input.value, { gapWidth: opts.gapWidth + "%" });
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

  async function insertPoem(replaceSelection) {
    await withWord(async function (context) {
      var html = AshaarWord.renderForWord(input.value, options(), Ashaar);
      if (!html) {
        setMessage("Enter poetry text first.");
        return;
      }
      var selection = context.document.getSelection();
      var inserted = selection.insertHtml(html, replaceSelection ? Word.InsertLocation.replace : Word.InsertLocation.end);
      var control = inserted.insertContentControl();
      control.title = "Ashaar Poem";
      control.tag = AshaarWord.contentControlTag(input.value, options());
      control.appearance = "BoundingBox";
      await context.sync();
    });
  }

  async function insertStructure() {
    await withWord(async function (context) {
      var opts = options();
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
    [input, justifyMode, layoutMode, widthMode, bandhCount, misraCount, misraPattern, fontMode, tatweelCount, gapWidth].forEach(function (el) {
      el.addEventListener("input", renderPreview);
      el.addEventListener("change", renderPreview);
    });
    modeTable.addEventListener("click", function () { setMode("table"); });
    modeConvert.addEventListener("click", function () { setMode("convert"); });
    document.getElementById("insert-structure").addEventListener("click", insertStructure);
    document.getElementById("insert-poem").addEventListener("click", function () { insertPoem(false); });
    document.getElementById("replace-selection").addEventListener("click", function () { insertPoem(true); });
    document.getElementById("justify-selection").addEventListener("click", justifySelection);
    document.getElementById("load-selection").addEventListener("click", loadSelection);
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
