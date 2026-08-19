import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";

interface FixtureWindow {
  __cyexcelTypedEditorsReady?: boolean;
  __sheet?: {
    getCell(row: number, column: number): unknown;
    setCell(row: number, column: number, value: unknown): void;
  };
}

async function gotoFixture(page: Page): Promise<void> {
  await page.goto("/tests/browser/typed-editors.html");
  await page.waitForFunction(
    () => (window as unknown as FixtureWindow).__cyexcelTypedEditorsReady === true
  );
}

async function getContentCanvasBox(page: Page) {
  const box = await page.locator("#app canvas").first().boundingBox();
  if (!box) {
    throw new Error("missing content canvas bounding box");
  }
  return box;
}

function getCell(page: Page, row: number, column: number): Promise<unknown> {
  return page.evaluate(
    ({ row: r, column: c }) => (window as unknown as FixtureWindow).__sheet?.getCell(r, c),
    { row, column }
  );
}

function setCell(page: Page, row: number, column: number, value: unknown): Promise<void> {
  return page.evaluate(
    ({ row: r, column: c, value: v }) => (window as unknown as FixtureWindow).__sheet?.setCell(r, c, v),
    { row, column, value }
  );
}

// Column x-centers, columns are active(90)/joinedAt(130)/status(110)/tags(180)/meta(200), rowHeight 28.
const COLUMN_X = { active: 45, joinedAt: 155, status: 275, tags: 420, meta: 610 };
const ROW_Y = { row0: 14, row1: 42 };

test("boolean column (explicit createEditor): checkbox reflects the value and toggling commits a real boolean", async ({
  page
}) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  expect(await getCell(page, 0, 0)).toBe(true);

  await page.mouse.dblclick(box.x + COLUMN_X.active, box.y + ROW_Y.row0);
  const checkbox = page.locator("#app input[type=checkbox]");
  await expect(checkbox).toBeChecked();

  await checkbox.click();
  await page.keyboard.press("Enter");
  await expect(page.locator("#app input[type=checkbox]")).toHaveCount(0);

  expect(await getCell(page, 0, 0)).toBe(false);
});

test("date column (explicit createEditor): native date input shows YYYY-MM-DD and commits a real Date", async ({
  page
}) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  await page.mouse.dblclick(box.x + COLUMN_X.joinedAt, box.y + ROW_Y.row0);
  const dateInput = page.locator("#app input[type=date]");
  await expect(dateInput).toHaveValue("2026-01-01");

  await dateInput.fill("2026-03-15");
  await page.keyboard.press("Enter");
  await expect(page.locator("#app input[type=date]")).toHaveCount(0);

  const isoDate = await page.evaluate(() => {
    const value = (window as unknown as FixtureWindow).__sheet?.getCell(0, 1);
    return value instanceof Date ? value.toISOString().slice(0, 10) : null;
  });
  expect(isoDate).toBe("2026-03-15");
});

test("select column (explicit createEditor): dropdown lists the configured options and preselects the current value", async ({
  page
}) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  await page.mouse.dblclick(box.x + COLUMN_X.status, box.y + ROW_Y.row0);
  const select = page.locator("#app select");
  await expect(select).toHaveValue("draft"); // row 0: index 0 % 3 -> "draft"

  const optionValues = await select.locator("option").evaluateAll((options) =>
    options.map((option) => (option as HTMLOptionElement).value)
  );
  expect(optionValues).toEqual(["draft", "active", "archived"]);

  await select.selectOption("active");
  await page.keyboard.press("Enter");
  await expect(page.locator("#app select")).toHaveCount(0);
  expect(await getCell(page, 0, 2)).toBe("active");
});

test("select column: a value outside the configured options round-trips instead of being silently replaced by the first option", async ({
  page
}) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  await setCell(page, 1, 2, "unknown-status");
  expect(await getCell(page, 1, 2)).toBe("unknown-status");

  await page.mouse.dblclick(box.x + COLUMN_X.status, box.y + ROW_Y.row1);
  await expect(page.locator("#app select")).toHaveValue("unknown-status");

  await page.keyboard.press("Enter"); // commit without touching anything
  await expect(page.locator("#app select")).toHaveCount(0);
  expect(await getCell(page, 1, 2)).toBe("unknown-status");
});

test("array column (no createEditor — default text editor): open-and-commit-unchanged round-trips instead of blanking the cell", async ({
  page
}) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  const before = await getCell(page, 0, 3);
  expect(before).toEqual(["tag-0", "group-0"]);

  await page.mouse.dblclick(box.x + COLUMN_X.tags, box.y + ROW_Y.row0);
  const input = page.locator("#app input[type=text]");
  await expect(input).toHaveValue(JSON.stringify(before)); // not "tag-0,group-0" or "[object Object]"

  await page.keyboard.press("Enter");
  await expect(input).toHaveCount(0);
  expect(await getCell(page, 0, 3)).toEqual(before);
});

test("json column (explicit createEditor + validate): malformed JSON is rejected on commit, valid JSON round-trips", async ({
  page
}) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  const before = await getCell(page, 0, 4);

  await page.mouse.dblclick(box.x + COLUMN_X.meta, box.y + ROW_Y.row0);
  const textarea = page.locator("#app textarea");
  await expect(textarea).toBeVisible();
  await textarea.fill("not json");
  await page.keyboard.press("Enter");

  // validate() rejects it — the editor must still be open, value untouched.
  await expect(page.locator("#app textarea")).toBeVisible();
  expect(await getCell(page, 0, 4)).toEqual(before);

  await textarea.fill('{"id":0,"note":"changed"}');
  await page.keyboard.press("Enter");
  await expect(page.locator("#app textarea")).toHaveCount(0);
  expect(await getCell(page, 0, 4)).toEqual({ id: 0, note: "changed" });
});

/**
 * A cell editor is an ordinary form control and must keep the browser's own
 * focus indication. The playground's `input:focus-visible` rule used to apply
 * to it as well, drawing a 2px red ring around a small checkbox nested inside
 * an already-bordered selected cell.
 *
 * Asserted on computed style rather than canvas pixels: an outline is not part
 * of the canvas bitmap, so a pixel assertion cannot see this class of bug at
 * all — which is exactly why it went unnoticed.
 */
test("the boolean editor's checkbox keeps the browser's default focus ring", async ({ page }) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  await page.mouse.dblclick(box.x + COLUMN_X.active, box.y + ROW_Y.row0);
  const checkbox = page.locator("#app input[type=checkbox]");
  await expect(checkbox).toBeChecked();
  await checkbox.click();

  const outline = await page.evaluate(() => {
    const el = document.querySelector("#app input[type=checkbox]") as HTMLInputElement;
    const style = getComputedStyle(el);
    return { style: style.outlineStyle, width: style.outlineWidth };
  });

  // "auto" is the engine's own ring; a host stylesheet overriding it would
  // report a concrete style such as "solid" with its own width.
  expect(outline.style).toBe("auto");
});
