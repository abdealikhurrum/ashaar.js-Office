/* BookletPane — wires the "Booklet" tab to Impose/ImposePdf (impose-core.js,
 * impose-pdf.js). Pulls the open Word document as a PDF via Office.js
 * (getFileAsync) instead of requiring a manual export/upload round-trip —
 * that's the whole point of living inside the add-in rather than just using
 * the standalone Awraq web tool.
 */
(function () {
  "use strict";

  var els = {};
  var docBytes = null;
  var pageCount = 0;
  var bound = false;
  // Stashed from the last "Make booklet PDF" run in manual-duplex mode, so
  // the separate "Download backs" click (a fresh user gesture — required
  // because Word's embedded webview blocks a second auto-triggered download
  // from the same click) reimposes the SAME document state, not a possibly
  // newer one.
  var pendingBacks = null;

  function setStatus(msg, warn) {
    els.status.textContent = msg;
    els.status.classList.toggle("warn", !!warn);
  }
  function setDocStatus(msg, warn) {
    els.docStatus.textContent = msg;
    els.docStatus.classList.toggle("warn", !!warn);
  }

  function opts() {
    return {
      scheme: els.scheme.value,
      direction: els.direction.value,
      flip: els.flip.value,
      duplex: els.duplex.value,
      paperSize: els.paperSize.value,
      reverse: els.reverse.checked,
    };
  }

  // Single-sided schemes (zinefold8) have no duplex/flip to configure.
  function syncSchemeUi() {
    var withBack = Impose.hasBack(els.scheme.value);
    els.duplexField.hidden = !withBack;
    els.flipField.hidden = !withBack;
    els.zineNote.hidden = withBack;
    if (!withBack) els.duplex.value = "auto";
    syncDuplexUi();
  }

  function syncDuplexUi() {
    els.reverseLine.hidden = els.duplex.value !== "manual";
  }

  function download(bytes, name) {
    var blob = new Blob([bytes], { type: "application/pdf" });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = name;
    a.rel = "noopener";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 30000);
  }

  // Reads the open document as PDF bytes via Office.js. getFileAsync hands
  // back the file in slices; concatenate them in order, then close the file
  // handle (required by the API regardless of success/failure).
  function getDocumentAsPdfBytes() {
    return new Promise(function (resolve, reject) {
      if (typeof Office === "undefined" || !Office.context || !Office.context.document) {
        reject(new Error("Word document is not available."));
        return;
      }
      Office.context.document.getFileAsync(Office.FileType.Pdf, { sliceSize: 4194304 }, function (result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error((result.error && result.error.message) || "Could not export the document as PDF."));
          return;
        }
        var file = result.value;
        var sliceCount = file.sliceCount;
        if (!sliceCount) {
          file.closeAsync(function () {});
          resolve(new Uint8Array(0));
          return;
        }
        var slices = new Array(sliceCount);
        var received = 0;
        function getSlice(i) {
          file.getSliceAsync(i, function (sliceResult) {
            if (sliceResult.status !== Office.AsyncResultStatus.Succeeded) {
              file.closeAsync(function () {});
              reject(new Error((sliceResult.error && sliceResult.error.message) || "Could not read the document."));
              return;
            }
            slices[i] = sliceResult.value.data;
            received++;
            if (received === sliceCount) {
              file.closeAsync(function () {});
              var total = 0;
              for (var k = 0; k < sliceCount; k++) total += slices[k].length;
              var bytes = new Uint8Array(total);
              var offset = 0;
              for (var j = 0; j < sliceCount; j++) {
                bytes.set(slices[j], offset);
                offset += slices[j].length;
              }
              resolve(bytes);
            } else {
              getSlice(i + 1);
            }
          });
        }
        getSlice(0);
      });
    });
  }

  function renderPlan() {
    if (!pageCount) { setStatus(""); return; }
    var o = opts();
    var p = Impose.plan({ pageCount: pageCount, scheme: o.scheme, direction: o.direction });
    var blankNote = p.blanks.length ? " (+" + p.blanks.length + " blank to fill the fold)" : "";
    setStatus(pageCount + " pages → " + p.sheets.length + " sheet" +
      (p.sheets.length === 1 ? "" : "s") + ", " + p.paddedCount + " sides" + blankNote + ".");
  }

  // Reads the document fresh and updates the cached docBytes/pageCount.
  // Called by BOTH "Get current document" (a preview) and "Make booklet PDF"
  // (the real thing) — Make never trusts a possibly-stale previous read, so
  // editing the document and clicking Make again always reflects the edit,
  // with no separate "refresh" step to remember.
  async function refreshDocument() {
    var bytes = await getDocumentAsPdfBytes();
    if (!bytes.length) throw new Error("The document appears to be empty.");
    var doc = await PDFLib.PDFDocument.load(bytes);
    docBytes = bytes;
    pageCount = doc.getPageCount();
    setDocStatus(pageCount + " page" + (pageCount === 1 ? "" : "s") + " loaded from the open document.");
    renderPlan();
  }

  async function getCurrentDocument() {
    els.getDoc.disabled = true;
    setDocStatus("Reading the open document as PDF…");
    try {
      await refreshDocument();
    } catch (e) {
      console.error(e);
      docBytes = null;
      pageCount = 0;
      setDocStatus("Couldn’t read the document: " + (e.message || e), true);
    } finally {
      els.getDoc.disabled = false;
    }
  }

  async function makeBooklet() {
    els.make.disabled = true;
    els.downloadBacks.hidden = true;
    setStatus("Reading the open document…");
    try {
      await refreshDocument();
      setStatus(els.status.textContent + " Imposing…");
      var o = opts();
      o.pdfLib = PDFLib;
      if (o.duplex === "manual") {
        var fronts = await ImposePdf.imposePdf(docBytes, Object.assign({}, o, { faces: "fronts", reverse: false }));
        download(fronts, "booklet-1-fronts.pdf");
        pendingBacks = { docBytes: docBytes, opts: o };
        els.downloadBacks.hidden = false;
        els.downloadBacks.disabled = false;
        setStatus(els.status.textContent + ' Fronts downloaded — print them, reinsert the stack, then click "Download backs".');
      } else {
        var out = await ImposePdf.imposePdf(docBytes, o);
        download(out, "booklet.pdf");
        setStatus(els.status.textContent + " Booklet PDF downloaded.");
      }
    } catch (e) {
      console.error(e);
      setStatus("Couldn’t make the booklet: " + (e.message || e), true);
    } finally {
      els.make.disabled = false;
    }
  }

  // A separate click (not auto-triggered after fronts) because Word's
  // embedded webview blocks a second programmatic download fired from the
  // same click handler — each file needs its own genuine user gesture.
  async function downloadBackFaces() {
    if (!pendingBacks) return;
    els.downloadBacks.disabled = true;
    try {
      var backs = await ImposePdf.imposePdf(pendingBacks.docBytes,
        Object.assign({}, pendingBacks.opts, { faces: "backs" }));
      download(backs, "booklet-2-backs.pdf");
      setStatus(els.status.textContent + " Backs downloaded.");
    } catch (e) {
      console.error(e);
      setStatus("Couldn’t make the backs file: " + (e.message || e), true);
    } finally {
      els.downloadBacks.disabled = false;
    }
  }

  function bind() {
    els.scheme = document.getElementById("booklet-scheme");
    els.direction = document.getElementById("booklet-direction");
    els.paperSize = document.getElementById("booklet-papersize");
    els.duplex = document.getElementById("booklet-duplex");
    els.flip = document.getElementById("booklet-flip");
    els.duplexField = document.getElementById("booklet-duplex-field");
    els.flipField = document.getElementById("booklet-flip-field");
    els.zineNote = document.getElementById("booklet-zine-note");
    els.status = document.getElementById("booklet-status");
    els.docStatus = document.getElementById("booklet-doc-status");
    els.getDoc = document.getElementById("booklet-get-doc");
    els.make = document.getElementById("booklet-make");
    els.downloadBacks = document.getElementById("booklet-download-backs");
    els.reverse = document.getElementById("booklet-reverse");
    els.reverseLine = document.getElementById("booklet-reverse-line");

    els.getDoc.addEventListener("click", getCurrentDocument);
    els.make.addEventListener("click", makeBooklet);
    els.downloadBacks.addEventListener("click", downloadBackFaces);
    ["scheme", "direction", "flip", "duplex"].forEach(function (key) {
      els[key].addEventListener("change", function () {
        syncSchemeUi();
        renderPlan();
      });
    });
    syncSchemeUi();
  }

  function onShow() {
    if (!bound) { bind(); bound = true; }
  }

  window.BookletPane = { onShow: onShow };
})();
