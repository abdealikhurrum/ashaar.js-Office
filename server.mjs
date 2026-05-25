import { createServer } from "node:https";
import { readFile } from "node:fs/promises";
import { createReadStream } from "node:fs";
import { extname, join, normalize } from "node:path";

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

const server = createServer({
  cert: await readFile(join(root, "localhost.pem")),
  key: await readFile(join(root, "localhost-key.pem"))
}, async (req, res) => {
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
