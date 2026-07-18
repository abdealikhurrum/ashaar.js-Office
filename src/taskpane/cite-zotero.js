(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) {
    module.exports = factory();
  } else {
    root.CiteZotero = factory();
  }
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  // In-memory citekey -> CSL-JSON item cache. Reset via clearCache() (tests;
  // harmless to call in prod too).
  var cache = {};

  // In-memory citekey -> [tagStrings] cache (parallel to `cache`).
  var tagCache = {};

  function clearCache() {
    cache = {};
    tagCache = {};
  }

  // --- pure: request/response shaping (no I/O) ---

  function buildExportRequest(citekeys) {
    return {
      jsonrpc: "2.0",
      method: "item.export",
      params: [citekeys, "Better CSL JSON"],
      id: 1
    };
  }

  function parseExportResult(rpcResponse, citekeys) {
    if (rpcResponse && rpcResponse.error) {
      var err = rpcResponse.error;
      var message = (err && err.message) ? err.message
        : (typeof err === "string" ? err : JSON.stringify(err));
      throw new Error(message);
    }
    var items = JSON.parse(rpcResponse.result) || [];
    var map = {};
    items.forEach(function (item) {
      if (item && item.id !== undefined && item.id !== null) {
        map[item.id] = item;
      }
    });
    // Defensive fallback: if a requested citekey is absent from the map but
    // there is a matching-index entry in the array, key that entry by the
    // requested citekey too (guards against a translator that doesn't set
    // id = citekey; "Better CSL JSON" already matches, this is belt & braces).
    // Only safe when the parsed array's length exactly matches the requested
    // citekeys' length: only then does index i in items correspond to index i
    // in citekeys. If lengths differ (e.g. Zotero silently dropped or
    // reordered items), positional keying could alias a requested key onto
    // the wrong item, which is worse than leaving it genuinely absent.
    var keys = citekeys || [];
    if (items.length === keys.length) {
      keys.forEach(function (key, i) {
        if (!Object.prototype.hasOwnProperty.call(map, key) && items[i]) {
          map[key] = items[i];
        }
      });
    }
    return map;
  }

  function buildTagsRequest(citekeys) {
    return {
      jsonrpc: "2.0",
      method: "item.export",
      params: [citekeys, "BetterBibTeX JSON"],
      id: 1
    };
  }

  function normalizeTag(t) {
    if (t && typeof t === "object") { return t.tag; }
    return t;
  }

  function parseTagsResult(rpcResponse, citekeys) {
    if (rpcResponse && rpcResponse.error) {
      var err = rpcResponse.error;
      var message = (err && err.message) ? err.message
        : (typeof err === "string" ? err : JSON.stringify(err));
      throw new Error(message);
    }
    var payload = JSON.parse(rpcResponse.result) || {};
    var items = payload.items || [];
    var map = {};
    items.forEach(function (item) {
      if (!item) { return; }
      var key = item.citationKey || item.citekey || item.id;
      if (key === undefined || key === null) { return; }
      var tags = (item.tags || []).map(normalizeTag).filter(function (t) {
        return typeof t === "string" && t.length > 0;
      });
      map[key] = tags;
    });
    return map;
  }

  var LOC_LABELS = [
    { re: /^(pp?\.?|pages?)$/i, label: "page" },
    { re: /^(chap\.?|chapters?)$/i, label: "chapter" },
    { re: /^(sec\.?|section|§)$/i, label: "section" },
    { re: /^(vv?\.?|verses?)$/i, label: "verse" }
  ];

  // Parse one pandoc item body ("@key" or "@key, p. 42") → {citekey, locator?, label?}.
  function parseCaywItem(body) {
    var trimmed = body.trim();
    if (trimmed.charAt(0) === "@") { trimmed = trimmed.slice(1); }
    var comma = trimmed.indexOf(",");
    if (comma === -1) {
      var bareKey = trimmed.trim();
      return bareKey ? { citekey: bareKey } : null;
    }
    var citekey = trimmed.slice(0, comma).trim();
    if (!citekey) { return null; }
    var suffix = trimmed.slice(comma + 1).trim();
    // suffix forms: "p. 42" | "pp. 42-45" | "chap. 3" | "42"
    var mLabel = /^([A-Za-z.§]+)\s*(.+)$/.exec(suffix);
    if (mLabel) {
      for (var i = 0; i < LOC_LABELS.length; i++) {
        if (LOC_LABELS[i].re.test(mLabel[1])) {
          return { citekey: citekey, locator: mLabel[2].trim(), label: LOC_LABELS[i].label };
        }
      }
      return { citekey: citekey }; // unrecognized label ⇒ no locator
    }
    if (/^[0-9]/.test(suffix)) { return { citekey: citekey, locator: suffix, label: "page" }; }
    return { citekey: citekey };
  }

  function parseCaywResult(text) {
    if (text === null || text === undefined) { return []; }
    var trimmed = String(text).trim();
    if (trimmed === "") { return []; }
    // Strip a single pair of surrounding {} or [] if present.
    if ((trimmed[0] === "{" && trimmed[trimmed.length - 1] === "}") ||
        (trimmed[0] === "[" && trimmed[trimmed.length - 1] === "]")) {
      trimmed = trimmed.slice(1, -1).trim();
    }
    if (trimmed === "") { return []; }
    return trimmed.split(";")
      .map(function (part) { return parseCaywItem(part); })
      .filter(function (it) { return it && it.citekey; });
  }

  // --- I/O: talk only to same-origin /zotero/* routes ---

  function ping(fetchImpl) {
    var f = fetchImpl || (typeof fetch !== "undefined" ? fetch : undefined);
    return Promise.resolve()
      .then(function () { return f("/zotero/ping"); })
      .then(function (res) { return !!(res && res.ok); })
      .catch(function () { return false; });
  }

  function caywPick(fetchImpl) {
    var f = fetchImpl || (typeof fetch !== "undefined" ? fetch : undefined);
    // BBT CAYW has no "citekeys" format; the pandoc format returns the picked
    // items as `[@citekey1; @citekey2]` (each optionally followed by a
    // `, <label> <locator>` suffix), which parseCaywResult parses into
    // `{citekey, locator?, label?}` objects.
    return Promise.resolve()
      .then(function () { return f("/zotero/cayw?format=pandoc"); })
      .then(function (res) {
        if (!res.ok) { throw new Error("cayw HTTP " + res.status); }
        return res.text();
      })
      .then(function (text) { return parseCaywResult(text); });
  }

  function fetchCslJson(citekeys, fetchImpl) {
    var f = fetchImpl || (typeof fetch !== "undefined" ? fetch : undefined);
    var missing = (citekeys || []).filter(function (key) {
      return !Object.prototype.hasOwnProperty.call(cache, key);
    });

    var fetchStep = missing.length === 0
      ? Promise.resolve(null)
      : Promise.resolve()
        .then(function () {
          return f("/zotero/json-rpc", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(buildExportRequest(missing))
          });
        })
        .then(function (res) {
          if (!res.ok) { throw new Error("json-rpc HTTP " + res.status); }
          return res.json();
        })
        .then(function (rpcResponse) {
          var parsed = parseExportResult(rpcResponse, missing);
          Object.keys(parsed).forEach(function (key) {
            cache[key] = parsed[key];
          });
        });

    return fetchStep.then(function () {
      var result = {};
      (citekeys || []).forEach(function (key) {
        if (Object.prototype.hasOwnProperty.call(cache, key)) {
          result[key] = cache[key];
        }
      });
      return result;
    });
  }

  function fetchTags(citekeys, fetchImpl) {
    var f = fetchImpl || (typeof fetch !== "undefined" ? fetch : undefined);
    var missing = (citekeys || []).filter(function (key) {
      return !Object.prototype.hasOwnProperty.call(tagCache, key);
    });
    var fetchStep = missing.length === 0
      ? Promise.resolve(null)
      : Promise.resolve()
        .then(function () {
          return f("/zotero/json-rpc", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(buildTagsRequest(missing))
          });
        })
        .then(function (res) {
          if (!res.ok) { throw new Error("json-rpc HTTP " + res.status); }
          return res.json();
        })
        .then(function (rpcResponse) {
          var parsed = parseTagsResult(rpcResponse, missing);
          // Cache every REQUESTED key so an item with no tags (absent from the
          // export) is remembered as [] rather than re-fetched every time.
          missing.forEach(function (key) {
            tagCache[key] = Object.prototype.hasOwnProperty.call(parsed, key) ? parsed[key] : [];
          });
        });
    return fetchStep.then(function () {
      var result = {};
      (citekeys || []).forEach(function (key) {
        result[key] = Object.prototype.hasOwnProperty.call(tagCache, key) ? tagCache[key] : [];
      });
      return result;
    });
  }

  return {
    buildExportRequest: buildExportRequest,
    parseExportResult: parseExportResult,
    parseCaywResult: parseCaywResult,
    ping: ping,
    caywPick: caywPick,
    fetchCslJson: fetchCslJson,
    buildTagsRequest: buildTagsRequest,
    parseTagsResult: parseTagsResult,
    fetchTags: fetchTags,
    clearCache: clearCache
  };
}));
