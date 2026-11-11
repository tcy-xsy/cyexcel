import { execFile } from "node:child_process";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import { beforeAll, describe, expect, it } from "vitest";

const execFileAsync = promisify(execFile);
const repoRoot = fileURLToPath(new URL("../../", import.meta.url));

interface PackedFile {
  readonly path: string;
}

let files: string[];

beforeAll(async () => {
  // --dry-run: reports exactly what `npm publish`/`pnpm pack` would include
  // without actually writing a .tgz to disk.
  const { stdout } = await execFileAsync("pnpm", ["pack", "--json", "--dry-run"], { cwd: repoRoot });
  const report = JSON.parse(stdout) as { files: readonly PackedFile[] };
  files = report.files.map((file) => file.path);
});

describe("published tarball contents", () => {
  it("includes the package manifest, README, and LICENSE", () => {
    expect(files).toContain("package.json");
    expect(files).toContain("README.md");
    expect(files).toContain("LICENSE");
    expect(files).toContain("THIRD_PARTY_NOTICES.txt");
  });

  it("includes the built dist/ output", () => {
    expect(files).toContain("dist/index.js");
    expect(files).toContain("dist/core.js");
    expect(files).toContain("dist/formula.js");
    expect(files).toContain("dist/types/formula/index.d.ts");
    expect(files).toContain("dist/cyexcel.umd.cjs");
    expect(files.some((file) => file.startsWith("dist/types/"))).toBe(true);
  });

  it("excludes source .ts files — only built dist/ output ships", () => {
    expect(files.some((file) => file.endsWith(".ts") && !file.startsWith("dist/types/"))).toBe(false);
  });

  it("excludes build caches, test artifacts, and local-only files", () => {
    const forbiddenPatterns = [
      /\.tsbuildinfo$/,
      /^coverage\//,
      /^test-results\//,
      /^playwright-report\//,
      /^blob-report\//,
      /^\.env/,
      /^examples\//,
      /^playground\//,
      /^tests\//,
      /^plugins\//
    ];
    const offenders = files.filter((file) => forbiddenPatterns.some((pattern) => pattern.test(file)));
    expect(offenders).toEqual([]);
  });
});
