#!/usr/bin/env node
// Reports the raw + gzip size of each real published entry file under
// dist/. Plugins (e.g. createBasicContextMenuPlugin) are tree-shaken
// exports inside `core` itself, not separate physical bundles, so "core and
// optional plugins" collapses to "these are the entry files that actually
// exist" — there's nothing per-plugin to report separately.
import { gzipSync } from "node:zlib";
import { readFile, stat } from "node:fs/promises";
import { fileURLToPath } from "node:url";

const distDir = fileURLToPath(new URL("../dist/", import.meta.url));

const entries = [
  { label: "cyexcel (aggregate)", file: "index.js" },
  { label: "cyexcel/core", file: "core.js" },
  { label: "cyexcel/register", file: "register.js" },
  { label: "cyexcel (UMD/CJS)", file: "cyexcel.umd.cjs" }
];

function formatBytes(bytes) {
  return `${(bytes / 1024).toFixed(1)} KB`;
}

async function main() {
  const rows = [];
  for (const entry of entries) {
    const path = new URL(entry.file, `file://${distDir}`);
    let content;
    try {
      content = await readFile(path);
    } catch {
      rows.push({ ...entry, missing: true });
      continue;
    }
    const rawSize = (await stat(path)).size;
    const gzipSize = gzipSync(content).length;
    rows.push({ ...entry, rawSize, gzipSize });
  }

  const missing = rows.filter((row) => row.missing);
  if (missing.length > 0) {
    console.error(
      `Missing dist/ files (run "pnpm build" first): ${missing.map((row) => row.file).join(", ")}`
    );
    process.exitCode = 1;
    return;
  }

  console.log("| Entry | File | Raw | Gzip |");
  console.log("| --- | --- | ---: | ---: |");
  for (const row of rows) {
    console.log(
      `| ${row.label} | \`dist/${row.file}\` | ${formatBytes(row.rawSize)} | ${formatBytes(row.gzipSize)} |`
    );
  }
}

await main();
