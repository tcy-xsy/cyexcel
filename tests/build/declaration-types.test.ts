import { execFile } from "node:child_process";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import { describe, expect, it } from "vitest";

const execFileAsync = promisify(execFile);
const repoRoot = fileURLToPath(new URL("../../", import.meta.url));
const tsc = fileURLToPath(new URL("../../node_modules/typescript/bin/tsc", import.meta.url));

describe.each([
  ["ESNext", "Bundler"],
  ["Node16", "Node16"],
  ["NodeNext", "NodeNext"]
])("published declarations with module %s and resolution %s", (module, resolution) => {
  it.each([true, false])("preserves exports and rejects invalid assignments with skipLibCheck=%s", async (skipLibCheck) => {
    const result = await execFileAsync(process.execPath, [
      tsc, "-p", "tsconfig.consumer.json",
      "--module", module, "--moduleResolution", resolution,
      "--skipLibCheck", String(skipLibCheck)
    ], { cwd: repoRoot });
    expect(result.stdout).toBe("");
    expect(result.stderr).toBe("");
  }, 30_000);
});
