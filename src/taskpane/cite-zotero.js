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

  function clearCache() {
    cache = {};
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
    return trimmed
      .split(/[\s,;]+/)
      .map(function (tok) { return tok.trim(); })
      .filter(function (tok) { return tok.length > 0; })
      .map(function (tok) { return tok[0] === "@" ? tok.slice(1) : tok; })
      .filter(function (tok) { return tok.length > 0; });
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
    return Promise.resolve()
      .then(function () { return f("/zotero/cayw?format=citekeys"); })
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

  return {
    buildExportRequest: buildExportRequest,
    parseExportResult: parseExportResult,
    parseCaywResult: parseCaywResult,
    ping: ping,
    caywPick: caywPick,
    fetchCslJson: fetchCslJson,
    clearCache: clearCache
  };
}));
