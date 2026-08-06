import { readFile, readdir } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

describe("package.json has no runtime framework dependencies", () => {
  it("has no \"dependencies\" field at all", async () => {
    const raw = await readFile(new URL("../../package.json", import.meta.url), "utf-8");
    const pkg = JSON.parse(raw) as Record<string, unknown>;
    expect(pkg.dependencies).toBeUndefined();
  });
});

describe("dist/ output contains no bundled framework code", () => {
  it("no dist/*.js or *.cjs file mentions react/vue/jquery", async () => {
    const distDir = new URL("../../dist/", import.meta.url);
    const entries = await readdir(fileURLToPath(distDir));
    const bundleFiles = entries.filter((name) => name.endsWith(".js") || name.endsWith(".cjs"));
    expect(bundleFiles.length).toBeGreaterThan(0);

    const offenders: string[] = [];
    for (const file of bundleFiles) {
      const content = (await readFile(new URL(file, distDir), "utf-8")).toLowerCase();
      if (/\breact\b|\bvue\b|\bjquery\b/.test(content)) {
        offenders.push(file);
      }
    }
    expect(offenders).toEqual([]);
  });
});
