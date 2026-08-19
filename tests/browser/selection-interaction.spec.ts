import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";

interface CellPosition {
  readonly row: number;
  readonly column: number;
}

interface CellRange {
  readonly startRow: number;
  readonly startColumn: number;
  readonly endRow: number;
  readonly endColumn: number;
}

interface SelectionSnapshot {
  readonly active: CellPosition;
  readonly range: CellRange;
  readonly ranges: readonly CellRange[];
}

const MOD = process.platform === "darwin" ? "Meta" : "Control";

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

test(`${MOD}+click adds a disconnected region instead of replacing the selection`, async ({ page }) => {
  await gotoFixture(page);
  const box = await page.locator("#grid-a").boundingBox();
  if (!box) {
    throw new Error("missing bounding box for #grid-a");
  }

  await page.mouse.click(box.x + 10, box.y + 10); // selects (0, 0)
  await page.keyboard.down(MOD);
  await page.mouse.click(box.x + 260, box.y + 110); // adds (5, 5)
  await page.keyboard.up(MOD);

  const selection = await readSelection(page, "__gridASelection");
  expect(selection.ranges).toEqual([
    { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 },
    { startRow: 5, startColumn: 5, endRow: 5, endColumn: 5 }
  ]);
});

test("a plain click after a multi-region selection collapses back to a single region", async ({ page }) => {
  await gotoFixture(page);
  const box = await page.locator("#grid-a").boundingBox();
  if (!box) {
    throw new Error("missing bounding box for #grid-a");
  }

  await page.mouse.click(box.x + 10, box.y + 10);
  await page.keyboard.down(MOD);
  await page.mouse.click(box.x + 260, box.y + 110);
  await page.keyboard.up(MOD);
  await page.mouse.click(box.x + 60, box.y + 30); // plain click, no modifier

  const selection = await readSelection(page, "__gridASelection");
  expect(selection.ranges).toEqual([{ startRow: 1, startColumn: 1, endRow: 1, endColumn: 1 }]);
});

/**
 * Reported as "selecting one cell puts a border around the whole table".
 *
 * It was not a selection-rendering bug: the content canvas is `tabIndex = 0`
 * and gets `.focus()` on every click, so the browser painted its own default
 * focus ring around the entire canvas, behind the one-cell selection box.
 * Canvas-pixel assertions can never catch this — a CSS outline is not part of
 * the bitmap — which is why every existing selection test passed.
 *
 * The ring is still wanted for keyboard focus: the active-cell box looks the
 * same whether or not the grid is focused, so it is the only thing that tells
 * the two apart.
 */
function contentCanvasFocusState(page: Page) {
  return page.evaluate(() => {
    const canvas = document.querySelectorAll("#grid-a canvas")[0] as HTMLCanvasElement;
    return {
      focused: document.activeElement === canvas,
      outline: getComputedStyle(canvas).outlineStyle
    };
  });
}

test("clicking a cell does not leave a focus ring around the whole grid", async ({ page }) => {
  await gotoFixture(page);
  const box = await page.locator("#grid-a").boundingBox();
  if (!box) {
    throw new Error("missing bounding box for #grid-a");
  }

  await page.mouse.click(box.x + 50, box.y + 30);

  await expect.poll(() => contentCanvasFocusState(page)).toEqual({ focused: true, outline: "none" });
});

test("focusing the grid without a pointer keeps the focus ring", async ({ page }) => {
  await gotoFixture(page);

  await page.evaluate(() => {
    (document.querySelectorAll("#grid-a canvas")[0] as HTMLElement).focus();
  });

  await expect.poll(() => contentCanvasFocusState(page)).toEqual({ focused: true, outline: "auto" });
});
