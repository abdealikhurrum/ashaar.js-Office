// Task 1 Step 2 verification: `node -e "const C=require('./vendor/citeproc-js/citeproc.js');
// console.log(typeof C, typeof (C&&C.Engine));"` printed `object undefined` — citeproc.js
// does not attach CSL to module.exports. `citeproc_commonjs.js` is byte-identical except
// for a trailing `module.exports = CSL` line, and printed `object function` as expected, so
// it is the file vendored here (as src/vendor/citeproc.js) for both Node (require) and
// browser (global `var CSL = {...}` declared at top level of the same file) use.
//
// That trailing `module.exports = CSL` is bare/unguarded, so loading the vendored file as a
// plain browser <script> throws `ReferenceError: module is not defined` (no `module` global
// there). `var CSL = {` at line 60 of the bundle already makes CSL a browser global in a
// classic script, so the only fix needed is to guard the export line so it's a no-op outside
// Node. This post-processing step rewrites that trailing statement after copying.
import { execFileSync } from "node:child_process";
import { copyFile, mkdir, readFile, writeFile } from "node:fs/promises";
import { join } from "node:path";

// Guard so `module.exports = CSL` only runs under Node (CommonJS); a no-op under a browser
// <script> tag, where `module` is undefined and CSL is already the global set at bundle top.
const GUARDED_EXPORT = 'if (typeof module !== "undefined" && module.exports) { module.exports = CSL; }';
const BARE_EXPORT_RE = /module\.exports\s*=\s*CSL;?\s*$/;

function guardModuleExports(source) {
  if (source.includes(GUARDED_EXPORT)) {
    return source; // already guarded (idempotent re-sync)
  }
  if (!BARE_EXPORT_RE.test(source)) {
    throw new Error("citeproc bundle: expected trailing `module.exports = CSL` not found — upstream format changed, update guardModuleExports()");
  }
  return source.replace(BARE_EXPORT_RE, GUARDED_EXPORT);
}

const root = process.cwd();
const engineRepo = join(root, "vendor", "citeproc-js");
const localesRepo = join(root, "vendor", "csl-locales");
const stylesRepo = join(root, "vendor", "csl-styles");
const dest = join(root, "src", "vendor");
const localesDest = join(dest, "csl-locales");
const stylesDest = join(dest, "csl-styles");

// The engine bundle that exposes CSL for both browser + Node (verified in Task 1 Step 2).
const engineFile = ["citeproc_commonjs.js", "citeproc.js"];
// citeproc needs a "us" fallback locale; en-US serves that role.
const locales = ["locales-ar.xml", "locales-en-US.xml"];
// Curated style subset ONLY (the styles repo is thousands of files).
const stockStyles = ["chicago-notes-bibliography.csl", "apa.csl"];

function git(dir, args) {
  return execFileSync("git", ["-C", dir, ...args], { encoding: "utf8" }).trim();
}

await mkdir(localesDest, { recursive: true });
await mkdir(stylesDest, { recursive: true });

const engineSource = await readFile(join(engineRepo, engineFile[0]), "utf8");
await writeFile(join(dest, engineFile[1]), guardModuleExports(engineSource));
for (const name of locales) {
  await copyFile(join(localesRepo, name), join(localesDest, name));
}
for (const name of stockStyles) {
  await copyFile(join(stylesRepo, name), join(stylesDest, name));
}

const stamp = [engineRepo, localesRepo, stylesRepo].map((dir) => {
  const remote = git(dir, ["config", "--get", "remote.origin.url"]);
  const commit = git(dir, ["rev-parse", "HEAD"]);
  const date = git(dir, ["log", "-1", "--format=%cI"]);
  const subject = git(dir, ["log", "-1", "--format=%s"]);
  return [`repo=${remote}`, `commit=${commit}`, `date=${date}`, `subject=${subject}`, ""].join("\n");
}).join("\n");

await writeFile(join(dest, "CITEPROC_UPSTREAM_VERSION"), stamp);
console.log("citeproc vendor sync complete");
