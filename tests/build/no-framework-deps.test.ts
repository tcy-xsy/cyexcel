import { readFile, readdir } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

// Step 24 之前这里还有一条"不得有任何 dependencies"的断言。numfmt（Excel 数字格式
// 解析，Univer 用的也是它）作为本项目第一个运行时依赖接入后，那条断言按设计删除——
// 见 doc/step24-style-plan.md §3.1。下面这条（产物里不得出现框架代码）留着：它在
// Step 22 Part B 刚抓到过一次误报，代码注释里写的 "react" 也会被打进产物。

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
