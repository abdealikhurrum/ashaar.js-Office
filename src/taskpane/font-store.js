/**
 * AshaarFontStore — load user-supplied font files so the justify canvas can
 * measure the same outlines Word renders, without bundling/redistributing fonts.
 *
 * A FontFace built from bytes is an explicitly-provided web font, so it is
 * measurable even in WebKit (Mac Word), which hides arbitrary *installed* fonts
 * from CSS/canvas for anti-fingerprinting. Fonts are persisted in IndexedDB and
 * re-registered on startup, so loading is a one-time action per machine.
 *
 * parseNames() is pure (sfnt name-table parser, unit-tested in Node); the
 * FontFace/IndexedDB helpers are browser-only and guarded accordingly.
 */
(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) {
    module.exports = factory();
  } else {
    root.AshaarFontStore = factory();
  }
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  // ── Pure: sfnt (TTF/OTF) name-table family parser ──────────────────────────

  function toDataView(input) {
    if (input instanceof DataView) return input;
    if (input instanceof ArrayBuffer) return new DataView(input);
    if (input && input.buffer instanceof ArrayBuffer && typeof input.byteOffset === "number") {
      return new DataView(input.buffer, input.byteOffset, input.byteLength);
    }
    return null;
  }

  function decodeName(dv, offset, length, platformID) {
    var s = "";
    if (platformID === 3 || platformID === 0) {
      // Windows / Unicode → UTF-16BE
      for (var i = 0; i + 1 < length; i += 2) s += String.fromCharCode(dv.getUint16(offset + i));
    } else {
      // Mac (platform 1) / other → treat as Latin-1 (good enough for ASCII family names)
      for (var j = 0; j < length; j++) s += String.fromCharCode(dv.getUint8(offset + j));
    }
    return s.replace(/\0+$/, "").trim();
  }

  // Returns { family, fullName, postScript } or null for non-sfnt / unparseable input.
  function parseNames(input) {
    var dv = toDataView(input);
    if (!dv || dv.byteLength < 12) return null;

    var sfnt = dv.getUint32(0);
    // 0x00010000 TrueType, 'OTTO', 'true', 'typ1'. Reject 'ttcf' (collection), WOFF/WOFF2.
    if (sfnt !== 0x00010000 && sfnt !== 0x4F54544F && sfnt !== 0x74727565 && sfnt !== 0x74797031) {
      return null;
    }

    var numTables = dv.getUint16(4);
    var nameOffset = 0;
    for (var t = 0, p = 12; t < numTables; t++, p += 16) {
      if (p + 16 > dv.byteLength) return null;
      if (dv.getUint32(p) === 0x6E616D65 /* 'name' */) { nameOffset = dv.getUint32(p + 8); break; }
    }
    if (!nameOffset || nameOffset + 6 > dv.byteLength) return null;

    var count = dv.getUint16(nameOffset + 2);
    var storage = nameOffset + dv.getUint16(nameOffset + 4);
    var best = {}; // nameID -> { score, value }

    for (var r = 0, rec = nameOffset + 6; r < count; r++, rec += 12) {
      if (rec + 12 > dv.byteLength) break;
      var platformID = dv.getUint16(rec);
      var nameID = dv.getUint16(rec + 6);
      if (nameID !== 1 && nameID !== 4 && nameID !== 6 && nameID !== 16) continue;
      var len = dv.getUint16(rec + 8);
      var off = storage + dv.getUint16(rec + 10);
      if (off + len > dv.byteLength) continue;
      var score = platformID === 3 ? 2 : 1; // prefer Windows records
      if (!best[nameID] || score > best[nameID].score) {
        var val = decodeName(dv, off, len, platformID);
        if (val) best[nameID] = { score: score, value: val };
      }
    }

    var family = (best[16] && best[16].value) || (best[1] && best[1].value) || "";
    if (!family) return null;
    return {
      family: family,
      fullName: (best[4] && best[4].value) || "",
      postScript: (best[6] && best[6].value) || ""
    };
  }

  // ── Browser-only: FontFace registration + IndexedDB persistence ────────────

  var DB_NAME = "ashaar-fonts", STORE = "fonts", DB_VERSION = 1;

  function openDB() {
    return new Promise(function (resolve, reject) {
      if (typeof indexedDB === "undefined") { reject(new Error("IndexedDB unavailable")); return; }
      var req = indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = function () {
        var db = req.result;
        if (!db.objectStoreNames.contains(STORE)) db.createObjectStore(STORE, { keyPath: "family" });
      };
      req.onsuccess = function () { resolve(req.result); };
      req.onerror = function () { reject(req.error); };
    });
  }

  function withStore(mode, fn) {
    return openDB().then(function (db) {
      return new Promise(function (resolve, reject) {
        var req = fn(db.transaction(STORE, mode).objectStore(STORE));
        req.onsuccess = function () { resolve(req.result); };
        req.onerror = function () { reject(req.error); };
      });
    });
  }

  function saveFont(family, filename, buffer) {
    return withStore("readwrite", function (s) {
      return s.put({ family: family, filename: filename || "", bytes: buffer });
    });
  }
  function listFonts() {
    return withStore("readonly", function (s) { return s.getAll(); }).then(function (r) { return r || []; });
  }
  function deleteFont(family) {
    return withStore("readwrite", function (s) { return s.delete(family); });
  }

  // Register a font's bytes with the document so canvas/CSS can measure it.
  function loadFont(family, buffer) {
    if (typeof FontFace === "undefined" || typeof document === "undefined" || !document.fonts) {
      return Promise.reject(new Error("FontFace API unavailable"));
    }
    return new FontFace(family, buffer).load().then(function (ff) {
      document.fonts.add(ff);
      return ff;
    });
  }

  // Load + persist a user-picked file. Persists best-effort (session-only if no IndexedDB).
  function addUserFont(family, file) {
    return file.arrayBuffer().then(function (buf) {
      return loadFont(family, buf).then(function () {
        return saveFont(family, file.name, buf).then(
          function () { return { family: family, persisted: true }; },
          function () { return { family: family, persisted: false }; }
        );
      });
    });
  }

  // Re-register every stored font. Call once on startup, before any measurement.
  function registerAll() {
    return listFonts().then(function (fonts) {
      return Promise.all(fonts.map(function (f) {
        return loadFont(f.family, f.bytes).then(
          function () { return { family: f.family, filename: f.filename }; },
          function () { return null; }
        );
      }));
    }).then(function (arr) { return arr.filter(Boolean); }, function () { return []; });
  }

  return {
    parseNames: parseNames,
    loadFont: loadFont,
    addUserFont: addUserFont,
    saveFont: saveFont,
    listFonts: listFonts,
    deleteFont: deleteFont,
    registerAll: registerAll
  };
}));
