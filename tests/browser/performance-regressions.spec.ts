import { expect, test } from "@playwright/test";
import type { CyExcelCore } from "../../src/core";

declare global {
  interface Window {
    regressionSheet: CyExcelCore;
    regressionMetrics: { paints: number; lookups: number };
  }
}

test.beforeEach(async ({ page }) => {
  await page.goto("/tests/browser/performance-regressions.html");
  await page.waitForFunction(() => Boolean(window.regressionSheet));
  await page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
});

test("one scroll event produces one content paint", async ({ page }) => {
  const result = await page.evaluate(async () => {
    const scroller = document.querySelector("#app > div") as HTMLElement;
    let events = 0;
    const onScroll = () => { events += 1; };
    scroller.addEventListener("scroll", onScroll);
    window.regressionMetrics.paints = 0;
    scroller.scrollTop = 48;
    for (let i = 0; i < 4; i += 1) {
      await new Promise(resolve => requestAnimationFrame(resolve));
    }
    scroller.removeEventListener("scroll", onScroll);
    return { events, paints: window.regressionMetrics.paints };
  });
  expect(result).toEqual({ events: 1, paints: 1 });
});

test("filtering 100000 rows to two visible rows keeps drawing viewport-bound", async ({ page }) => {
  const result = await page.evaluate(async () => {
    const sheet = window.regressionSheet;
    sheet.filterByColumn(0, (_value, row) => row === 0 || row === 99999);
    await new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    window.regressionMetrics.paints = 0;
    window.regressionMetrics.lookups = 0;
    sheet.resize();
    return { ...window.regressionMetrics, lastValue: sheet.getCell(99999, 0) };
  });
  expect(result.lastValue).toBe(99999);
  expect(result.paints).toBe(1);
  expect(result.lookups).toBeLessThan(200);
});
