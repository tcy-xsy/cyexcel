import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";

interface FixtureWindow {
  __cyexcelColumnStructureReady?: boolean;
  __sheet?: {
    getCell(row: number, column: number): unknown;
    getColumn(column: number): { readonly id: string };
    insertColumns(start: number, columns: readonly { readonly id: string; readonly field?: string; readonly width?: number }[]): void;
    deleteColumns(start: number, count?: number): void;
    moveColumns(start: number, count: number, target: number): void;
    hideColumn(column: number): void;
    showColumn(column: number): void;
    undo(): void;
    redo(): void;
  };
}

async function gotoFixture(page: Page): Promise<void> {
  await page.goto("/tests/browser/column-structure.html");
  await page.waitForFunction(
    () => (window as unknown as FixtureWindow).__cyexcelColumnStructureReady === true
  );
}

/** RenderScheduler batches repaints via requestAnimationFrame — pixel sampling must wait a frame after any mutation, same as headers.spec.ts's nextFrame(). */
async function nextFrame(page: Page): Promise<void> {
  await page.waitForFunction(
    () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)))
  );
}

async function getAppBox(page: Page) {
  const box = await page.locator("#app").boundingBox();
  if (!box) {
    throw new Error("missing #app bounding box");
  }
  return box;
}

function getCell(page: Page, row: number, column: number): Promise<unknown> {
  return page.evaluate(
    ({ row: r, column: c }) => (window as unknown as FixtureWindow).__sheet?.getCell(r, c),
    { row, column }
  );
}

function getColumnId(page: Page, column: number): Promise<unknown> {
  return page.evaluate(
    (c) => (window as unknown as FixtureWindow).__sheet?.getColumn(c).id,
    column
  );
}

/** Non-transparent pixel count for a logical-pixel box on the content canvas (index 0), DPR-aware — same technique as row-headers.spec.ts's sampleAlpha. */
function sampleAlpha(page: Page, x: number, y: number, width: number, height: number): Promise<number> {
  return page.evaluate(
    ({ x, y, width, height }) => {
      const canvas = document.querySelectorAll("#app canvas")[0] as HTMLCanvasElement;
      const ctx = canvas.getContext("2d") as CanvasRenderingContext2D;
      const dpr = window.devicePixelRatio || 1;
      const { data } = ctx.getImageData(
        Math.round(x * dpr),
        Math.round(y * dpr),
        Math.max(1, Math.round(width * dpr)),
        Math.max(1, Math.round(height * dpr))
      );
      let nonTransparentCount = 0;
      for (let i = 3; i < data.length; i += 4) {
        if ((data[i] ?? 0) > 0) {
          nonTransparentCount += 1;
        }
      }
      return nonTransparentCount;
    },
    { x, y, width, height }
  );
}

test("insertColumns inserts a blank column and shifts the rest right, and undo restores it", async ({ page }) => {
  await gotoFixture(page);

  expect(await getCell(page, 0, 1)).toBe("0b");

  await page.evaluate(() => {
    (window as unknown as FixtureWindow).__sheet?.insertColumns(1, [{ id: "x", field: "x", width: 60 }]);
  });

  expect(await getCell(page, 0, 1)).toBeUndefined(); // the new blank column
  expect(await getColumnId(page, 1)).toBe("x");
  expect(await getCell(page, 0, 2)).toBe("0b"); // original column b, shifted right

  await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.undo());

  expect(await getCell(page, 0, 1)).toBe("0b");
  expect(await getColumnId(page, 1)).toBe("b");

  await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.redo());
  expect(await getColumnId(page, 1)).toBe("x");
  expect(await getCell(page, 0, 2)).toBe("0b");
});

test("deleteColumns removes a column and shifts the rest left, and undo restores it", async ({ page }) => {
  await gotoFixture(page);

  expect(await getCell(page, 0, 2)).toBe("0c");

  await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.deleteColumns(1));

  expect(await getCell(page, 0, 1)).toBe("0c"); // column c shifted into column b's old slot
  expect(await getColumnId(page, 1)).toBe("c");

  await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.undo());

  expect(await getCell(page, 0, 1)).toBe("0b");
  expect(await getColumnId(page, 1)).toBe("b");
});

test("moveColumns moves a column to a new position, and undo restores it", async ({ page }) => {
  await gotoFixture(page);

  // [a, b, c, d] -> moveColumns(0, 1, 3) -> [b, c, a, d] (same pre-move
  // target-index convention as moveRows()/ArrayDataSource.moveColumns()).
  await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.moveColumns(0, 1, 3));

  expect(await getColumnId(page, 0)).toBe("b");
  expect(await getColumnId(page, 1)).toBe("c");
  expect(await getColumnId(page, 2)).toBe("a");
  expect(await getColumnId(page, 3)).toBe("d");
  expect(await getCell(page, 0, 2)).toBe("0a");

  await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.undo());

  expect(await getColumnId(page, 0)).toBe("a");
  expect(await getColumnId(page, 1)).toBe("b");
  expect(await getColumnId(page, 2)).toBe("c");
  expect(await getColumnId(page, 3)).toBe("d");
});

test("hideColumn collapses the column to zero width and shifts later columns left; showColumn restores it", async ({
  page
}) => {
  await gotoFixture(page);

  // Column d occupies x:[180,240) initially — its row-0 cell "0d" paints there.
  expect(await sampleAlpha(page, 180, 0, 60, 24)).toBeGreaterThan(0);

  await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.hideColumn(1)); // hide column b (width 60)
  await nextFrame(page);

  // Total content width shrinks by 60 (a:0-60, b:60-60, c:60-120, d:120-180)
  // — nothing paints past x=180 anymore.
  // Canvas strokes straddle the half-open content boundary by one device
  // pixel, so sample strictly after x=180 rather than counting d's border.
  expect(await sampleAlpha(page, 181, 0, 59, 24)).toBe(0);
  // Column c now paints where b used to start.
  expect(await sampleAlpha(page, 60, 0, 60, 24)).toBeGreaterThan(0);

  await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.showColumn(1));
  await nextFrame(page);

  expect(await sampleAlpha(page, 180, 0, 60, 24)).toBeGreaterThan(0);
});

test("right-clicking a column header opens a menu that inserts/deletes/hides that column", async ({ page }) => {
  await gotoFixture(page);
  const box = await getAppBox(page);

  // Column b spans x:[60,120), header band y:[0,24).
  await page.mouse.click(box.x + 90, box.y + 10, { button: "right" });
  await expect(page.locator('[data-item-id="cyexcel:insert-column-left"]')).toBeVisible();
  await page.locator('[data-item-id="cyexcel:insert-column-left"]').click();

  expect(await getCell(page, 0, 1)).toBeUndefined();
  expect(await getCell(page, 0, 2)).toBe("0b");

  await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.undo());

  // Column c now spans x:[120,180).
  await page.mouse.click(box.x + 150, box.y + 10, { button: "right" });
  await page.locator('[data-item-id="cyexcel:delete-column"]').click();

  expect(await getCell(page, 0, 2)).toBe("0d"); // d shifted into c's old slot

  await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.undo());

  await page.mouse.click(box.x + 210, box.y + 10, { button: "right" }); // column d
  await page.locator('[data-item-id="cyexcel:hide-column"]').click();
  await nextFrame(page);

  expect(await sampleAlpha(page, 181, 0, 59, 24)).toBe(0);
});
