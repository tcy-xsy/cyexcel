import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";

interface FixtureWindow {
  __cyexcelVariableRowHeightReady?: boolean;
  __rowLayoutRows?: number[];
  __sheet?: {
    readonly scene: {
      readonly overscanRows: number;
      readonly overscanColumns: number;
      getRowHeight(row: number): number;
      getVisibleRows(): { readonly start: number; readonly count: number };
      getVisibleColumns(): { readonly start: number; readonly count: number };
      getPanes(): readonly {
        readonly id: string;
        readonly rows: { readonly start: number; readonly count: number };
        readonly columns: { readonly start: number; readonly count: number };
      }[];
    };
    getCell(row: number, column: number): unknown;
    setCell(row: number, column: number, value: unknown): void;
    setRowHeight(row: number, height: number): void;
    startEdit(row: number, column: number): void;
    undo(): void;
    redo(): void;
  };
}

test("core options publish overscan through the Scene viewport and pane traversal", async ({ page }) => {
  await gotoFixture(page);

  const snapshot = await page.evaluate(() => {
    const scene = (window as unknown as FixtureWindow).__sheet?.scene;
    const body = scene?.getPanes().find((pane) => pane.id === "body");
    return {
      overscanRows: scene?.overscanRows,
      overscanColumns: scene?.overscanColumns,
      visibleRows: scene?.getVisibleRows(),
      visibleColumns: scene?.getVisibleColumns(),
      bodyRows: body?.rows,
      bodyColumns: body?.columns
    };
  });

  expect(snapshot.overscanRows).toBe(1);
  expect(snapshot.overscanColumns).toBe(2);
  expect(snapshot.bodyRows).toEqual(snapshot.visibleRows);
  expect(snapshot.bodyColumns).toEqual(snapshot.visibleColumns);
});

test("wrap measurement consumes the target Scene's overscanned pane rows", async ({ page }) => {
  await gotoFixture(page);

  expect(
    await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.scene.getRowHeight(11))
  ).toBeGreaterThan(24);
});

async function gotoFixture(page: Page): Promise<void> {
  await page.goto("/tests/browser/variable-row-height.html");
  await page.waitForFunction(
    () => (window as unknown as FixtureWindow).__cyexcelVariableRowHeightReady === true
  );
  await page.waitForTimeout(50);
}

async function editorRect(page: Page, row: number, column = 0): Promise<{ top: number; height: number }> {
  await page.evaluate(
    ({ row: r, column: c }) => (window as unknown as FixtureWindow).__sheet?.startEdit(r, c),
    { row, column }
  );
  const rect = await page.locator("#app input").evaluate((element: HTMLInputElement) => ({
    top: Number.parseFloat(element.style.top),
    height: Number.parseFloat(element.style.height)
  }));
  await page.keyboard.press("Escape");
  return rect;
}

test("programmatic row resizing keeps downstream editor geometry aligned", async ({ page }) => {
  await gotoFixture(page);

  const beforeRow2 = await editorRect(page, 2);
  const beforeRow3 = await editorRect(page, 3);
  await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.setRowHeight(2, 60));

  expect(await editorRect(page, 2)).toEqual({ top: beforeRow2.top, height: 60 });
  expect(await editorRect(page, 3)).toEqual({
    top: beforeRow3.top + (60 - beforeRow2.height),
    height: beforeRow3.height
  });
});

test("dragging a row boundary resizes that row and leaves the new height in place", async ({ page }) => {
  await gotoFixture(page);
  const rowHeader = await page.locator("#app canvas").nth(3).boundingBox();
  if (!rowHeader) throw new Error("missing row header canvas box");

  // Row 1 begins after the auto-grown frozen row 0. Read its current top
  // from the editor so the drag stays independent of platform font metrics.
  const row1 = await editorRect(page, 1);
  const boundaryY = rowHeader.y + 20 + row1.top + row1.height;
  await page.mouse.move(rowHeader.x + rowHeader.width - 3, boundaryY);
  await expect(page.locator("#app canvas").nth(3)).toHaveCSS("cursor", "row-resize");
  await page.mouse.down();
  await page.mouse.move(rowHeader.x + rowHeader.width - 3, boundaryY + 28);
  await page.mouse.up();

  const resizedHeight = row1.height + 28;
  expect((await editorRect(page, 1)).height).toBe(resizedHeight);
  expect((await editorRect(page, 2)).top).toBe(row1.top + resizedHeight);
});

test("narrowing a wrap column remeasures visible rows against the new width", async ({ page }) => {
  await gotoFixture(page);
  const before = await editorRect(page, 0);
  const header = await page.locator("#app canvas").nth(2).boundingBox();
  if (!header) throw new Error("missing column header canvas box");

  await page.mouse.move(header.x + 96, header.y + 10);
  await page.mouse.down();
  await page.mouse.move(header.x + 56, header.y + 10);
  await page.mouse.up();

  expect((await editorRect(page, 0)).height).toBeGreaterThan(before.height);
});

test("auto-owned wrap rows grow and shrink while manual heights stay locked", async ({ page }) => {
  await gotoFixture(page);

  const frozen = await editorRect(page, 0);
  expect(frozen.height).toBeGreaterThan(24);

  await page.evaluate(() =>
    (window as unknown as FixtureWindow).__sheet?.setCell(
      2,
      0,
      "编辑后写入的长中文内容会立即换行并把当前行增高到足够容纳所有文本"
    )
  );
  const grown = await editorRect(page, 2);
  expect(grown.height).toBeGreaterThan(24);

  await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.setCell(2, 0, "短文本"));
  expect((await editorRect(page, 2)).height).toBe(24);

  await page.evaluate(() => {
    const sheet = (window as unknown as FixtureWindow).__sheet;
    sheet?.setRowHeight(3, 60);
    sheet?.setCell(3, 0, "手工锁定高度以后，即使内容需要很多行也不能被自动测量覆盖");
  });
  expect((await editorRect(page, 3)).height).toBe(60);

  await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.setCell(3, 0, "短文本"));
  expect((await editorRect(page, 3)).height).toBe(60);
});

test("committing wrap text auto-grows the row and still completes the normal edit lifecycle", async ({ page }) => {
  await gotoFixture(page);

  const row2Before = await editorRect(page, 2);
  const contentCanvas = await page.locator("#app canvas").first().boundingBox();
  if (!contentCanvas) throw new Error("missing content canvas box");
  await page.mouse.click(contentCanvas.x + 20, contentCanvas.y + row2Before.top + row2Before.height / 2);
  await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.startEdit(2, 0));
  const input = page.locator("#app input");
  await input.fill("提交中的长中文内容触发自动增高时不能把这次正常提交误判成取消");
  await input.press("Enter");
  await expect(input).toHaveCount(0);

  // A successful commit moves the active cell down one row; Enter opens
  // that row. The old internal call to public setRowHeight() canceled the
  // editor mid-commit, so this second Enter reopened row 2 instead.
  const expectedRow3 = (await editorRect(page, 3)).top;
  await page.locator("#app canvas").first().press("Enter");
  await expect(input).toBeVisible();
  const row3 = await input.evaluate((element: HTMLInputElement) => Number.parseFloat(element.style.top));
  expect(row3).toBe(expectedRow3);
});

test("off-screen wrap rows stay at the default height until scrolling makes them visible", async ({ page }) => {
  await gotoFixture(page);

  expect(
    await page.evaluate(() => (window as unknown as FixtureWindow).__rowLayoutRows?.includes(50))
  ).toBe(false);

  await page.evaluate(() => {
    const scroller = document.querySelector("#app > div") as HTMLElement;
    scroller.scrollTop = 50 * 24;
  });
  await page.waitForTimeout(100);

  expect(
    await page.evaluate(() => (window as unknown as FixtureWindow).__rowLayoutRows?.includes(50))
  ).toBe(true);
  expect((await editorRect(page, 50)).height).toBeGreaterThan(24);
});

test("auto-height is skipped by history and preserves an existing redo entry", async ({ page }) => {
  await gotoFixture(page);

  await page.evaluate(() => {
    const fixture = window as unknown as FixtureWindow;
    const sheet = fixture.__sheet;
    const scroller = document.querySelector("#app > div") as HTMLElement;
    if (!sheet) throw new Error("missing sheet");

    sheet.setCell(1, 1, "recorded-value");
    scroller.scrollTop = 50 * 24;
  });
  await page.waitForFunction(
    () => {
      const sheet = (window as unknown as FixtureWindow).__sheet;
      return sheet !== undefined && sheet.scene.getRowHeight(50) > 24;
    }
  );

  const afterFirstMeasurement = await page.evaluate(() => {
    const fixture = window as unknown as FixtureWindow;
    const sheet = fixture.__sheet;
    if (!sheet) throw new Error("missing sheet");
    sheet.undo();
    return {
      afterUndo: sheet.getCell(1, 1),
      row50HeightAfterUndo: sheet.scene.getRowHeight(50)
    };
  });

  await page.evaluate(() => {
    const fixture = window as unknown as FixtureWindow;
    const sheet = fixture.__sheet;
    const scroller = document.querySelector("#app > div") as HTMLElement;
    if (!sheet) throw new Error("missing sheet");
    scroller.scrollTop = 70 * 24;
  });
  await page.waitForFunction(
    () => {
      const sheet = (window as unknown as FixtureWindow).__sheet;
      return sheet !== undefined && sheet.scene.getRowHeight(70) > 24;
    }
  );

  const redoResult = await page.evaluate(() => {
    const fixture = window as unknown as FixtureWindow;
    const sheet = fixture.__sheet;
    if (!sheet) throw new Error("missing sheet");
    const row70HeightBeforeRedo = sheet.scene.getRowHeight(70);
    sheet.redo();
    return {
      afterRedo: sheet.getCell(1, 1),
      row70HeightBeforeRedo,
      row70HeightAfterRedo: sheet.scene.getRowHeight(70)
    };
  });

  expect(afterFirstMeasurement.afterUndo).toBe("other-1");
  expect(afterFirstMeasurement.row50HeightAfterUndo).toBeGreaterThan(24);
  expect(redoResult.row70HeightBeforeRedo).toBeGreaterThan(24);
  expect(redoResult.afterRedo).toBe("recorded-value");
  expect(redoResult.row70HeightAfterRedo).toBe(redoResult.row70HeightBeforeRedo);
});
