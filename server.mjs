import { createServer } from "node:https";
import { request as httpRequest } from "node:http";
import { readFile } from "node:fs/promises";
import { createReadStream } from "node:fs";
import { extname, join, normalize } from "node:path";
import zoteroProxy from "./zotero-proxy.js";

const root = process.cwd();
const port = Number(process.env.PORT || 3000);

const mimeTypes = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml; charset=utf-8",
  ".xml": "application/xml; charset=utf-8"
};

function filePathFor(url) {
  const path = decodeURIComponent(new URL(url, `https://localhost:${port}`).pathname);
  const normalized = normalize(path).replace(/^(\.\.[/\\])+/, "");
  const filePath = join(root, normalized === "/" ? "src/taskpane/taskpane.html" : normalized);
  return filePath.startsWith(root) ? filePath : join(root, "src/taskpane/taskpane.html");
}

// Reverse-proxies /zotero/* to the local Zotero Better BibTeX HTTP server so the
// (HTTPS-served) pane can reach it same-origin without a mixed-content error.
// CAYW long-polls until the user finishes/cancels the picker, so no timeout is set.
function proxyToZotero(req, res, target, search) {
  const upstreamUrl = zoteroProxy.ZOTERO_BASE + target + (search || "");
  const headers = {};
  if (req.headers["content-type"]) headers["content-type"] = req.headers["content-type"];
  // Zotero's connector rejects chunked/length-less POSTs ("Content-length not
  // provided"); forward the incoming Content-Length since we pipe the body verbatim.
  if (req.headers["content-length"]) headers["content-length"] = req.headers["content-length"];
  const upstreamReq = httpRequest(upstreamUrl, { method: req.method, headers, timeout: 0 }, (upstreamRes) => {
    res.writeHead(upstreamRes.statusCode || 502, {
      "Content-Type": upstreamRes.headers["content-type"] || "application/octet-stream"
    });
    upstreamRes.on("error", () => {
      if (!res.writableEnded) res.destroy();
    });
    upstreamRes.pipe(res);
  });
  upstreamReq.on("error", (err) => {
    if (res.headersSent || res.writableEnded) {
      res.end();
      return;
    }
    res.writeHead(502, { "Content-Type": "application/json" });
    res.end(JSON.stringify({ error: "zotero-unreachable", detail: err.message }));
  });
  req.on("error", () => upstreamReq.destroy());
  req.pipe(upstreamReq);
}

const server = createServer({
  cert: await readFile(join(root, "localhost.pem")),
  key: await readFile(join(root, "localhost-key.pem"))
}, async (req, res) => {
  const requestUrl = new URL(req.url || "/", `https://localhost:${port}`);
  const target = zoteroProxy.zoteroProxyTarget(requestUrl.pathname);
  if (target !== null) {
    proxyToZotero(req, res, target, requestUrl.search);
    return;
  }

  const filePath = filePathFor(req.url || "/");
  const stream = createReadStream(filePath);
  stream.on("open", () => {
    res.writeHead(200, {
      "Cache-Control": "no-store",
      "Content-Type": mimeTypes[extname(filePath)] || "application/octet-stream"
    });
    stream.pipe(res);
  });
  stream.on("error", () => {
    res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
    res.end("Not found");
  });
});

server.listen(port, () => {
  console.log(`Ashaar Office add-in served at https://localhost:${port}/src/taskpane/taskpane.html`);
});
