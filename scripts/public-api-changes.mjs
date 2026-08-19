#!/usr/bin/env node
// Derives public-API breakage from the diff of `src/core/index.ts`'s exports
// rather than from commit subjects.
//
// Why not `git log --grep '!'`: this repo already has two commits that removed
// public exports without marking the subject — `8de308d` alone removed 107 of
// 194 exports under the subject "feat(core): Phase 6 — Core as facade", and
// `5d68e75` renamed a field on an exported type. A release note built from
// subjects would have missed the largest break of the whole migration. Subjects
// are a claim; the export list is the fact, so this reads the fact.
//
//   node scripts/public-api-changes.mjs <range>     report every commit that
//                                                   changed the public surface
//   node scripts/public-api-changes.mjs --check <range>
//                                                   exit 1 if a commit removed
//                                                   exports without `!`
//
// `--check` is meant for new work (e.g. `HEAD~1..HEAD`), not for auditing
// history: the two commits above will always fail it, which is exactly why
// they are recorded in doc/scene-mutation-migration-plan.md §14.8 instead.
import { execFileSync } from "node:child_process";

const BARREL = "src/core/index.ts";

function git(args) {
  return execFileSync("git", args, { encoding: "utf8", maxBuffer: 64 * 1024 * 1024 });
}

/** Every identifier `src/core/index.ts` re-exports at `rev`, or null if absent. */
function exportsAt(rev) {
  let source;
  try {
    source = git(["show", `${rev}:${BARREL}`]);
  } catch {
    return null;
  }
  const withoutComments = source
    .replace(/\/\*[\s\S]*?\*\//g, "")
    .replace(/\/\/[^\n]*/g, "");
  const names = new Set();
  // `export { … }` and `export type { … }` alike — missing the second form is
  // what made an earlier version of this check report "0 removed" for a commit
  // that removed `Patch`.
  for (const [, block] of withoutComments.matchAll(/export\s+(?:type\s+)?\{([^}]*)\}/g)) {
    for (const entry of block.split(",")) {
      const name = entry
        .trim()
        .replace(/^type\s+/, "")
        .split(/\s+as\s+/)
        .pop()
        ?.trim();
      if (name && /^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) {
        names.add(name);
      }
    }
  }
  return names;
}

const args = process.argv.slice(2);
const check = args.includes("--check");
const range = args.find((arg) => !arg.startsWith("--"));
if (!range) {
  console.error("usage: public-api-changes.mjs [--check] <range>");
  process.exit(2);
}

const commits = git(["rev-list", "--reverse", range]).trim().split("\n").filter(Boolean);
let unmarked = 0;

for (const sha of commits) {
  const after = exportsAt(sha);
  const before = exportsAt(`${sha}^`);
  if (!after || !before) {
    continue;
  }
  const removed = [...before].filter((name) => !after.has(name)).sort();
  const added = [...after].filter((name) => !before.has(name)).sort();
  if (removed.length === 0 && added.length === 0) {
    continue;
  }

  const subject = git(["log", "-1", "--format=%s", sha]).trim();
  const marked = /^[^:]*!:/.test(subject);
  const breaking = removed.length > 0;
  const flag = breaking ? (marked ? "BREAKING" : "BREAKING (unmarked!)") : "additive";
  console.log(`${sha.slice(0, 7)}  ${flag}  +${added.length} -${removed.length}  ${subject}`);
  if (removed.length > 0) {
    console.log(`          removed: ${removed.join(", ")}`);
  }
  if (breaking && !marked) {
    unmarked += 1;
  }
}

if (commits.length > 0 && unmarked === 0) {
  console.log("no unmarked breaking changes in range");
}

if (check && unmarked > 0) {
  console.error(
    `\n${unmarked} commit(s) removed public exports without a \`!\` in the subject.\n` +
      "Mark the subject (e.g. `feat(core)!: …`) or split the removal out."
  );
  process.exit(1);
}
