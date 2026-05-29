import { execFileSync } from "node:child_process";
import { copyFile, mkdir, readFile, writeFile } from "node:fs/promises";
import { join } from "node:path";

const root = process.cwd();
const upstream = join(root, "vendor", "ashaar-js");
const dest = join(root, "src", "vendor");

const files = [
  ["ashaar.js", "ashaar.js"],
  ["ashaar-justify.js", "ashaar-justify.js"],
  ["ashaar-tune.js", "ashaar-autotune.js"],
  ["stylesheet.css", "ashaar.css"]
];

function git(args) {
  return execFileSync("git", ["-C", upstream, ...args], { encoding: "utf8" }).trim();
}

await mkdir(dest, { recursive: true });

for (const [sourceName, destName] of files) {
  await copyFile(join(upstream, sourceName), join(dest, destName));
}

const commit = git(["rev-parse", "HEAD"]);
const date = git(["log", "-1", "--format=%cI"]);
const subject = git(["log", "-1", "--format=%s"]);
const remote = git(["config", "--get", "remote.origin.url"]);
const packageJson = JSON.parse(await readFile(join(upstream, "package.json"), "utf8"));

await writeFile(join(dest, "ASHAAR_UPSTREAM_VERSION"), [
  `repo=${remote}`,
  `branch=master`,
  `commit=${commit}`,
  `date=${date}`,
  `package=${packageJson.name || "ashaar-js"}`,
  `version=${packageJson.version || "unversioned"}`,
  `subject=${subject}`,
  ""
].join("\n"));

console.log(`Synced Ashaar.js ${commit} into src/vendor`);
