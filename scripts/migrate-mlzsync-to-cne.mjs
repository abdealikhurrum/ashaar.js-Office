#!/usr/bin/env node
// One-time migration: Juris-M mlzsync (Extra) -> CNE cne-* (Extra) via the Zotero local API.
//
// Usage:
//   node scripts/migrate-mlzsync-to-cne.mjs                 # dry-run: print diff, write nothing
//   node scripts/migrate-mlzsync-to-cne.mjs --write         # apply (backs up first)
//   node scripts/migrate-mlzsync-to-cne.mjs --write --force # overwrite existing cne-* lines
//   node scripts/migrate-mlzsync-to-cne.mjs --write --strip-mlzsync  # drop the mlzsync block too
//   ...--base=http://localhost:23119                        # override the local-API base URL
//
// Requires Zotero running with the local API exposed (localhost:23119, no auth).
import { createRequire } from "module";
import { writeFileSync, mkdirSync } from "fs";
import { fileURLToPath } from "url";
import { dirname, join } from "path";

const require = createRequire(import.meta.url);
const CV = require("../src/taskpane/cite-variants.js");

const args = process.argv.slice(2);
const has = (f) => args.includes(f);
const WRITE = has("--write");
const FORCE = has("--force");
const STRIP = has("--strip-mlzsync");
const BASE = (args.find((a) => a.startsWith("--base=")) || "--base=http://localhost:23119").split("=")[1];
const LIB = "users/0"; // the local default library

async function getAllItems() {
  const out = [];
  let start = 0;
  for (;;) {
    const res = await fetch(`${BASE}/api/${LIB}/items?limit=100&start=${start}&format=json`);
    if (!res.ok) { throw new Error(`GET items failed: ${res.status} (is Zotero running with the local API?)`); }
    const batch = await res.json();
    if (!batch.length) { break; }
    out.push(...batch);
    start += batch.length;
    if (batch.length < 100) { break; }
  }
  return out;
}

// Merge cne-* lines into an extra string. Skips lines whose key already exists
// unless FORCE. Optionally strips the mlzsync block. Returns null when there is
// nothing to do (so already-migrated items are counted as skipped).
function mergeExtra(extra, cneLines) {
  let lines = String(extra || "").split(/\r?\n/);
  if (STRIP) { lines = lines.filter((l) => l.indexOf("mlzsync1:") === -1); }
  const existingKeys = new Set(
    lines.map((l) => (/^\s*(cne-[^:]+):/.exec(l) || [])[1]).filter(Boolean)
  );
  const additions = cneLines.filter((l) => {
    const k = (/^\s*(cne-[^:]+):/.exec(l) || [])[1];
    return FORCE || !existingKeys.has(k);
  });
  if (!additions.length && !STRIP) { return null; }
  return lines.concat(additions).filter((l) => l.length).join("\n");
}

async function patchExtra(item, extra) {
  const res = await fetch(`${BASE}/api/${LIB}/items/${item.key}`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json", "If-Unmodified-Since-Version": String(item.version) },
    body: JSON.stringify({ extra })
  });
  if (res.status === 412) { throw new Error("412 version conflict"); }
  if (!res.ok && res.status !== 204) { throw new Error(`PATCH failed: ${res.status}`); }
}

(async () => {
  const items = await getAllItems();
  const affected = [];
  let skipped = 0, converted = 0, failed = 0;

  for (const it of items) {
    const extra = (it.data && it.data.extra) || "";
    const parsed = CV.parseMlzsync(extra);
    if (!parsed) { continue; }
    const creators = (it.data && it.data.creators) || [];
    const cneLines = CV.mlzsyncToCneLines(parsed, creators);
    if (!cneLines.length) { continue; }
    const newExtra = mergeExtra(extra, cneLines);
    if (newExtra === null) { skipped++; continue; }

    console.log(`\n# ${it.key}  ${(it.data && it.data.title) || ""}`);
    cneLines.forEach((l) => console.log("  + " + l));
    affected.push({ key: it.key, after: newExtra, item: it });
  }

  console.log(`\n${affected.length} item(s) to convert, ${skipped} already migrated.`);

  if (!WRITE) {
    console.log("\nDRY RUN — no changes written. Re-run with --write to apply.");
    return;
  }

  // Backup the affected items before any write.
  const __dirname = dirname(fileURLToPath(import.meta.url));
  const backupDir = join(__dirname, "..", "scratch");
  try { mkdirSync(backupDir, { recursive: true }); } catch (e) { /* exists */ }
  const backupPath = join(backupDir, "mlzsync-backup.json");
  writeFileSync(backupPath, JSON.stringify(affected.map((a) => a.item), null, 2));
  console.log(`Backup written: ${backupPath}`);

  for (const a of affected) {
    try {
      await patchExtra(a.item, a.after);
      converted++;
      console.log(`  ✓ ${a.key}`);
    } catch (e) {
      failed++;
      console.error(`  ✗ ${a.key}: ${e.message}`);
    }
  }
  console.log(`\nDone. converted=${converted} skipped=${skipped} failed=${failed}`);
})().catch((e) => { console.error(e); process.exit(1); });
