import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";

interface CellPosition {
  readonly row: number;
  readonly column: number;
}

interface SelectionSnapshot {
  readonly active: CellPosition;
  readonly range: {
    readonly startRow: number;
    readonly startColumn: number;
    readonly endRow: number;
    readonly endColumn: number;
  };
}

interface FixtureWindow {
  __cyexcelSelectionReady?: boolean;
  __gridASelection?: SelectionSnapshot;
  __gridBSelection?: SelectionSnapshot;
}

async function gotoFixture(page: Page): Promise<void> {
  await page.goto("/tests/browser/selection-interaction.html");
  await page.waitForFunction(
    () => (window as unknown as FixtureWindow).__cyexcelSelectionReady === true
  );
}

function readSelection(page: Page, name: "__gridASelection" | "__gridBSelection") {
  return page.evaluate(
    (key) => (window as unknown as FixtureWindow)[key],
    name
  ) as Promise<SelectionSnapshot>;
}

test("dragging the mouse selects a normalized range, regardless of direction", async ({ page }) => {
  await gotoFixture(page);
  const box = await page.locator("#grid-a").boundingBox();
  if (!box) {
    throw new Error("missing bounding box for #grid-a");
  }

  await page.mouse.move(box.x + 10, box.y + 10);
  await page.mouse.down();
  await page.mouse.move(box.x + 260, box.y + 110);
  await page.mouse.up();

  const forward = await readSelection(page, "__gridASelection");
  expect(forward.range).toEqual({ startRow: 0, startColumn: 0, endRow: 5, endColumn: 5 });

  await page.mouse.move(box.x + 260, box.y + 110);
  await page.mouse.down();
  await page.mouse.move(box.x + 10, box.y + 10);
  await page.mouse.up();

  const backward = await readSelection(page, "__gridASelection");
  expect(backward.range).toEqual(forward.range);
});

test("a real mouse click with a couple pixels of jitter across a row boundary selects only one cell", async ({
  page
}) => {
  await gotoFixture(page);
  const box = await page.locator("#grid-a").boundingBox();
  if (!box) {
    throw new Error("missing bounding box for #grid-a");
  }

  // rowHeight is 20, so y=19 is row 0 and y=21 is row 1 — without a drag
  // threshold this 2px of incidental movement between mousedown and
  // mouseup reads as a drag and extends the selection into row 1, which is
  // exactly the reported "clicking sometimes selects multiple cells" bug.
  await page.mouse.move(box.x + 10, box.y + 19);
  await page.mouse.down();
  await page.mouse.move(box.x + 10, box.y + 21);
  await page.mouse.up();

  const selection = await readSelection(page, "__gridASelection");
  expect(selection.range).toEqual({ startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 });
});

test("keyboard navigation moves the active cell", async ({ page }) => {
  await gotoFixture(page);

  await page.locator("#grid-a").click({ position: { x: 10, y: 10 } });
  await page.keyboard.press("ArrowRight");
  await page.keyboard.press("ArrowDown");

  const selection = await readSelection(page, "__gridASelection");
  expect(selection.active).toEqual({ row: 1, column: 1 });
});

test("keyboard events only affect the focused instance", async ({ page }) => {
  await gotoFixture(page);

  await page.locator("#grid-a").click({ position: { x: 10, y: 10 } });
  const beforeB = await readSelection(page, "__gridBSelection");

  await page.keyboard.press("ArrowRight");
  await page.keyboard.press("ArrowRight");

  const afterA = await readSelection(page, "__gridASelection");
  const afterB = await readSelection(page, "__gridBSelection");

  expect(afterA.active).toEqual({ row: 0, column: 2 });
  expect(afterB.active).toEqual(beforeB.active);
});
