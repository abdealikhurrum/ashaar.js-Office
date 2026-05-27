import { execFileSync } from "node:child_process";

const upstream = "vendor/ashaar-js";

function run(command, args, options = {}) {
  console.log([command, ...args].join(" "));
  execFileSync(command, args, { stdio: "inherit", ...options });
}

run("git", ["-C", upstream, "fetch", "origin", "master"]);
run("git", ["-C", upstream, "checkout", "origin/master"]);
run("node", ["scripts/sync-ashaar-vendor.mjs"]);
run("npm", ["test"]);
