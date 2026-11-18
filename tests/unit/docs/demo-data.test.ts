import { afterEach, describe, expect, it } from "vitest";
import { getDemoData } from "../../../docs-site/src/demo/data";
import { buildPerformanceRows, getPerformanceColumns } from "../../../docs-site/src/demo/performance-data";
import { localizeDemoData } from "../../../docs-site/src/demo/localize";
import { setLocale } from "../../../docs-site/src/i18n";

afterEach(() => setLocale("zh"));

describe("localized documentation spreadsheet data", () => {
  it("contains English headers, product values and regional summaries on English pages", () => {
    setLocale("en");
    const data = getDemoData();
    expect(JSON.stringify(data)).not.toMatch(/\p{Script=Han}/u);
    expect(data.columns[1]!.header).toBe("Name");
    expect(data.rows[0]![1]).toBe("Industrial temperature sensor");
    expect(data.summaryRows[0]![0]).toBe("East China");
  });

  it("creates independent fixtures without mutating the Chinese source", () => {
    setLocale("en");
    const english = getDemoData();
    english.rows[0]![1] = "Edited";
    setLocale("zh");
    const chinese = getDemoData();
    expect(chinese.columns[1]!.header).toBe("名称");
    expect(chinese.rows[0]![1]).toBe("工业温度传感器");
    setLocale("en");
    expect(getDemoData().rows[0]![1]).toBe("Industrial temperature sensor");
  });

  it("localizes validation values and their color keys together", () => {
    setLocale("en");
    const criteria = localizeDemoData({
      type: "list", items: ["待处理", "进行中", "已完成"],
      colors: { 待处理: "#e8eaed", 进行中: "#d2e3fc", 已完成: "#ceead6" }
    });
    expect(criteria.items).toEqual(["Pending", "In progress", "Completed"]);
    expect(Object.keys(criteria.colors)).toEqual(criteria.items);
    expect(criteria.colors).toMatchObject({ "In progress": "#d2e3fc" });
  });

  it("preserves numeric data, formulas, dates and render functions", () => {
    setLocale("en");
    const date = new Date("2026-01-01T00:00:00Z");
    const render = () => {};
    const localized = localizeDemoData({ count: 42, formula: "=SUM(B1:B3)", date, render });
    expect(localized).toEqual({ count: 42, formula: "=SUM(B1:B3)", date, render });
    expect(localized.date).toBe(date);
    expect(localized.render).toBe(render);
  });

  it("uses English across generated performance rows, not only the visible first rows", () => {
    setLocale("en");
    expect(JSON.stringify([getPerformanceColumns(), buildPerformanceRows(40)])).not.toMatch(/\p{Script=Han}/u);
  });
});
