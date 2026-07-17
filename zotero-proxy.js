"use strict";

// Pure route-decision module for the /zotero/* reverse proxy in server.mjs.
// No sockets, no node:http — kept requirable standalone by tests.

const ZOTERO_BASE = "http://localhost:23119";

const ROUTES = {
  "/zotero/ping": "/connector/ping",
  "/zotero/cayw": "/better-bibtex/cayw",
  "/zotero/json-rpc": "/better-bibtex/json-rpc"
};

function zoteroProxyTarget(pathname) {
  return Object.prototype.hasOwnProperty.call(ROUTES, pathname) ? ROUTES[pathname] : null;
}

module.exports = { ZOTERO_BASE, zoteroProxyTarget };
