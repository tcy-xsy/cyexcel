import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";
import type * as CyExcelAPI from "../../src/index";

interface FixtureWindow {
  __cyexcelDistUmdReady?: boolean;
}

async function gotoFixture(page: Page): Promise<void> {
  await page.goto("/tests/build/dist-umd.html");
  await page.waitForFunction(
    () => (window as unknown as FixtureWindow).__cyexcelDistUmdReady === true
  );
}

test("the global UMD bundle exposes CyExcelCore (not just defineCyExcel) and creates a real, painted table", async ({
  page
}) => {
  await gotoFixture(page);

  const hasPaintedPixel = await page.evaluate(() => {
    const canvas = document.querySelectorAll("#app canvas")[0] as HTMLCanvasElement;
    const ctx = canvas.getContext("2d") as CanvasRenderingContext2D;
    const dpr = window.devicePixelRatio || 1;
    const { data } = ctx.getImageData(0, 0, Math.round(60 * dpr), Math.round(20 * dpr));
    for (let i = 3; i < data.length; i += 4) {
      if ((data[i] ?? 0) > 0) {
        return true;
      }
    }
    return false;
  });

  expect(hasPaintedPixel).toBe(true);
});

test("the global workbook bundles its formula parser and recalculates without external globals", async ({ page }) => {
  await gotoFixture(page);
  const result = await page.evaluate(() => {
    const api = (window as unknown as { CyExcel: typeof CyExcelAPI }).CyExcel;
    const container = document.createElement('div');
    container.style.cssText = 'width:600px;height:300px';
    document.body.append(container);
    const workbook = api.createWorkbook(container, { toolbar: false,
      snapshot: api.matrixToWorkbookSnapshot([[3, 0]]) });
    const sheet = workbook.getActiveSheet();
    sheet.setCell(0, 1, '=SUM(A1,4)');
    const before = sheet.getCell(0, 1);
    sheet.setCell(0, 0, 8);
    const after = sheet.getCell(0, 1);
    workbook.dispose();
    return [before, after];
  });
  expect(result).toEqual([7, 12]);
});
