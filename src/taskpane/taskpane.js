(function () {
  var input = document.getElementById("poem-input");
  var preview = document.getElementById("preview");
  var message = document.getElementById("message");
  var hostStatus = document.getElementById("host-status");
  var justifyMode = document.getElementById("justify-mode");
  var layoutMode = document.getElementById("layout-mode");
  var tatweelCount = document.getElementById("tatweel-count");
  var tatweelValue = document.getElementById("tatweel-value");
  var gapWidth = document.getElementById("gap-width");

  function options() {
    return {
      justifyMode: justifyMode.value,
      justify: justifyMode.value === "none" ? false : justifyMode.value,
      layoutMode: layoutMode.value,
      layout: layoutMode.value,
      tatweelCount: Number(tatweelCount.value || 0),
      gapWidth: Number(gapWidth.value || 4)
    };
  }

  function setMessage(text) {
    message.textContent = text || "";
  }

  function renderPreview() {
    var opts = options();
    tatweelValue.textContent = String(opts.tatweelCount);
    preview.className = "ashaar preview";
    preview.innerHTML = Ashaar.renderText(input.value, { gapWidth: opts.gapWidth + "%" });
    Ashaar.applyRenderOptions(preview, { gapWidth: opts.gapWidth + "%" });
    if (opts.layout === "stacked") preview.classList.add("ashaar--stacked");
    if (opts.layout === "auto") Ashaar.applyAutoLayout(preview, { layout: "auto" });
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
      selection.insertHtml(html, replaceSelection ? Word.InsertLocation.replace : Word.InsertLocation.end);
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

  function bind() {
    [input, justifyMode, layoutMode, tatweelCount, gapWidth].forEach(function (el) {
      el.addEventListener("input", renderPreview);
      el.addEventListener("change", renderPreview);
    });
    document.getElementById("insert-poem").addEventListener("click", function () { insertPoem(false); });
    document.getElementById("replace-selection").addEventListener("click", function () { insertPoem(true); });
    document.getElementById("justify-selection").addEventListener("click", justifySelection);
    document.getElementById("load-selection").addEventListener("click", loadSelection);
    renderPreview();
  }

  if (window.Office && Office.onReady) {
    Office.onReady(function (info) {
      hostStatus.textContent = info.host === Office.HostType.Word ? "Connected to Word" : "Preview mode";
      bind();
    });
  } else {
    hostStatus.textContent = "Browser preview mode";
    bind();
  }
}());
