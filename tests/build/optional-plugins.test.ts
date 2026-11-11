import { readFile, readdir } from "node:fs/promises";
import { describe, expect, it } from "vitest";

describe("optional plugin isolation", () => {
  it("keeps the Excel parser out of main package dependencies and every JS bundle", async () => {
    const manifest = JSON.parse(await readFile(new URL("../../package.json", import.meta.url), "utf8"));
    expect(manifest.dependencies).not.toHaveProperty("exceljs");
    expect(manifest.dependencies).not.toHaveProperty("cyexcel-plugin-excel-import");
    const dist = new URL("../../dist/", import.meta.url);
    const files = (await readdir(dist)).filter(name => /\.(?:js|cjs)$/.test(name));
    expect(files.length).toBeGreaterThan(0);
    for (const file of files) {
      const code = await readFile(new URL(file, dist), "utf8");
      expect(code, file).not.toContain("exceljs");
      expect(code, file).not.toContain("createExcelImportPlugin");
    }
  });
});
