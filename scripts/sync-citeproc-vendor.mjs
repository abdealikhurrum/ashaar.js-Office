// Task 1 Step 2 verification: `node -e "const C=require('./vendor/citeproc-js/citeproc.js');
// console.log(typeof C, typeof (C&&C.Engine));"` printed `object undefined` — citeproc.js
// does not attach CSL to module.exports. `citeproc_commonjs.js` is byte-identical except
// for a trailing `module.exports = CSL` line, and printed `object function` as expected, so
// it is the file vendored here (as src/vendor/citeproc.js) for both Node (require) and
// browser (global `var CSL = {...}` declared at top level of the same file) use.
import { execFileSync } from "node:child_process";
import { copyFile, mkdir, writeFile } from "node:fs/promises";
import { join } from "node:path";

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

await copyFile(join(engineRepo, engineFile[0]), join(dest, engineFile[1]));
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
