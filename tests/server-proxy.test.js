"use strict";
const assert = require("assert");
const zoteroProxy = require("../zotero-proxy.js");

assert.strictEqual(zoteroProxy.ZOTERO_BASE, "http://localhost:23119");

assert.strictEqual(zoteroProxy.zoteroProxyTarget("/zotero/ping"), "/connector/ping");
assert.strictEqual(zoteroProxy.zoteroProxyTarget("/zotero/cayw"), "/better-bibtex/cayw");
assert.strictEqual(zoteroProxy.zoteroProxyTarget("/zotero/json-rpc"), "/better-bibtex/json-rpc");

assert.strictEqual(zoteroProxy.zoteroProxyTarget("/zotero/"), null);
assert.strictEqual(zoteroProxy.zoteroProxyTarget("/zotero/unknown"), null);
assert.strictEqual(zoteroProxy.zoteroProxyTarget("/src/taskpane/taskpane.html"), null);
assert.strictEqual(zoteroProxy.zoteroProxyTarget("/"), null);

console.log("server-proxy.test.js: all assertions passed");
