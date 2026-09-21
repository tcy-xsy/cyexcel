import { execFile } from "node:child_process";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import { describe, expect, it } from "vitest";

const execFileAsync = promisify(execFile);
// `URL.pathname` is percent-encoded (this repo's own directory name
// contains a space) — `fileURLToPath` decodes it back into a real
// filesystem path `spawn`'s `cwd` can actually use. Using the raw
// `.pathname` here silently gave `spawn` a nonexistent "...cy%20table..."
// directory and failed with a misleading ENOENT blaming the node binary.
const repoRoot = fileURLToPath(new URL("../../", import.meta.url));

/**
 * Real `node` child processes (not Vitest's own module resolver, which
 * doesn't implement Node's self-referencing-package resolution) — the only
 * faithful way to prove `package.json`'s `exports` map itself resolves
 * `"cyexcel"`/`"cyexcel/core"`/`"cyexcel/register"` the way an external
 * consumer's `node_modules` resolution would. `cwd` is the repo root so
 * Node treats this package as importable-by-its-own-name.
 *
 * `process.execPath` (the absolute path to the currently-running Node
 * binary) instead of the bare "node" — a spawned child doesn't necessarily
 * inherit a PATH that resolves "node" (e.g. under nvm-managed shells).
 */
function runNode(script: string): Promise<{ stdout: string }> {
  return execFileAsync(process.execPath, ["--input-type=module", "-e", script], { cwd: repoRoot });
}

function runNodeCjs(script: string): Promise<{ stdout: string }> {
  return execFileAsync(process.execPath, ["--input-type=commonjs", "-e", script], { cwd: repoRoot });
}

describe("package.json exports — ESM", () => {
  it("resolves \"cyexcel\" and exposes CyExcelCore", async () => {
    const { stdout } = await runNode(
      'import("cyexcel").then(m => console.log(typeof m.CyExcelCore))'
    );
    expect(stdout.trim()).toBe("function");
  });

  it("resolves \"cyexcel/core\" and exposes CyExcelCore", async () => {
    const { stdout } = await runNode(
      'import("cyexcel/core").then(m => console.log(typeof m.CyExcelCore))'
    );
    expect(stdout.trim()).toBe("function");
  });

  it("resolves \"cyexcel/register\" and exposes defineCyExcel", async () => {
    const { stdout } = await runNode(
      'import("cyexcel/register").then(m => console.log(typeof m.defineCyExcel))'
    );
    expect(stdout.trim()).toBe("function");
  });

  it("resolves \"cyexcel/toolbar\" and exposes createToolbar", async () => {
    const { stdout } = await runNode(
      'import("cyexcel/toolbar").then(m => console.log(typeof m.createToolbar))'
    );
    expect(stdout.trim()).toBe("function");
  });

  it("importing the toolbar entry in Node touches no DOM — it only defines functions", async () => {
    // 工具栏整个是 DOM 代码，但**模块顶层**不许碰 DOM：碰了的话 SSR 场景下光是
    // import 就会炸。真正的 document 调用全在 createToolbar() 的函数体里。
    const { stdout } = await runNode(
      'import("cyexcel/toolbar").then(() => console.log("imported"))'
    );
    expect(stdout.trim()).toBe("imported");
  });
});

describe("package.json exports — CJS/require", () => {
  it("resolves require(\"cyexcel\") and exposes CyExcelCore (the UMD bundle, not just defineCyExcel)", async () => {
    const { stdout } = await runNodeCjs('console.log(typeof require("cyexcel").CyExcelCore)');
    expect(stdout.trim()).toBe("function");
  });

  it("resolves require(\"cyexcel/register\") and exposes defineCyExcel", async () => {
    const { stdout } = await runNodeCjs('console.log(typeof require("cyexcel/register").defineCyExcel)');
    expect(stdout.trim()).toBe("function");
  });
});
