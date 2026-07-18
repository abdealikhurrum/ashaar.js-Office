import { execFileSync } from "node:child_process";

const repos = ["vendor/citeproc-js", "vendor/csl-locales", "vendor/csl-styles"];

function run(command, args, options = {}) {
  console.log([command, ...args].join(" "));
  execFileSync(command, args, { stdio: "inherit", ...options });
}

for (const repo of repos) {
  run("git", ["-C", repo, "fetch", "origin", "master"]);
  run("git", ["-C", repo, "checkout", "origin/master"]);
}
run("node", ["scripts/sync-citeproc-vendor.mjs"]);
run("npm", ["test"]);
