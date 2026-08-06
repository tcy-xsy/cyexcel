import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";

import { escapeTsvCell } from "../../src/core/clipboard/tsv";

interface FixtureWindow {
  __cyexcelClipboardReady?: boolean;
  __sheet?: {
    getCell(row: number, column: number): unknown;
    undo(): void;
    copy(options?: unknown): Promise<{ status: string; plan: { estimatedBytes: number } }>;
    paste(options?: unknown): Promise<void>;
  };
}

const MOD = process.platform === "darwin" ? "Meta" : "Control";

async function gotoFixture(page: Page): Promise<void> {
  await page.goto("/tests/browser/clipboard.html");
  await page.waitForFunction(() => (window as unknown as FixtureWindow).__cyexcelClipboardReady === true);
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

test("real Ctrl/Cmd+C copies the selected range as TSV to the system clipboard", async ({ page }) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  // Drag-select name+age across both rows.
  await page.mouse.move(box.x + 10, box.y + 5);
  await page.mouse.down();
  await page.mouse.move(box.x + 150, box.y + 30);
  await page.mouse.up();
  await page.keyboard.press(`${MOD}+C`);

  await page.locator("#scratch").click();
  await page.keyboard.press(`${MOD}+V`);

  // A real <textarea>'s .value normalizes line endings to "\n" even though
  // the clipboard payload itself used "\r\n" (encodeTsv's row separator).
  await expect(page.locator("#scratch")).toHaveValue("Alice\t30\nBob\t40");
});

test("a clipboard.copy:false column is copied as a blank cell but keeps its column position", async ({
  page
}) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  // Select name+age+secret on row 0 only.
  await page.mouse.move(box.x + 10, box.y + 5);
  await page.mouse.down();
  await page.mouse.move(box.x + 200, box.y + 5);
  await page.mouse.up();
  await page.keyboard.press(`${MOD}+C`);

  await page.locator("#scratch").click();
  await page.keyboard.press(`${MOD}+V`);

  await expect(page.locator("#scratch")).toHaveValue("Alice\t30\t");
});

test("real Ctrl/Cmd+V pastes a system-clipboard TSV block and one undo reverts the whole paste", async ({
  page
}) => {
  await gotoFixture(page);

  await page.locator("#scratch").fill("Zed\t99\r\nYak\t88");
  await page.locator("#scratch").click();
  await page.keyboard.press(`${MOD}+A`);
  await page.keyboard.press(`${MOD}+C`);

  const box = await getContentCanvasBox(page);
  await page.mouse.click(box.x + 10, box.y + 5); // select cell (0, 0)
  await page.keyboard.press(`${MOD}+V`);

  await expect.poll(() => getCell(page, 0, 0)).toBe("Zed");
  expect(await getCell(page, 0, 1)).toBe(99);
  expect(await getCell(page, 1, 0)).toBe("Yak");
  expect(await getCell(page, 1, 1)).toBe(88);

  await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.undo());

  expect(await getCell(page, 0, 0)).toBe("Alice");
  expect(await getCell(page, 0, 1)).toBe(30);
  expect(await getCell(page, 1, 0)).toBe("Bob");
  expect(await getCell(page, 1, 1)).toBe(40);
});

test("pasting into a clipboard.paste:false column leaves its existing value unchanged", async ({ page }) => {
  await gotoFixture(page);

  await page.locator("#scratch").fill("Nova\t1\tignored\t999");
  await page.locator("#scratch").click();
  await page.keyboard.press(`${MOD}+A`);
  await page.keyboard.press(`${MOD}+C`);

  const box = await getContentCanvasBox(page);
  await page.mouse.click(box.x + 10, box.y + 5); // select cell (0, 0), row spans name/age/secret/locked
  await page.keyboard.press(`${MOD}+V`);

  await expect.poll(() => getCell(page, 0, 0)).toBe("Nova");
  expect(await getCell(page, 0, 3)).toBe(100); // "locked" column: unchanged, not the pasted 999
});

test("the programmatic copy() API fully serializes an array cell in raw mode, not just its preview", async ({
  page,
  context
}) => {
  await context.grantPermissions(["clipboard-read", "clipboard-write"]);
  await gotoFixture(page);

  const result = await page.evaluate(async () => {
    const sheet = (window as unknown as FixtureWindow).__sheet;
    await sheet?.copy({ mode: "raw", range: { startRow: 0, startColumn: 4, endRow: 0, endColumn: 4 } });
    return navigator.clipboard.readText();
  });

  expect(result).toBe(escapeTsvCell(JSON.stringify(["red", "green", "blue", "yellow"])));
});

test("the programmatic paste() API reads from the system clipboard and commits a single transaction", async ({
  page,
  context
}) => {
  await context.grantPermissions(["clipboard-read", "clipboard-write"]);
  await gotoFixture(page);

  await page.evaluate(async () => {
    await navigator.clipboard.writeText("Q\t5");
    await (window as unknown as FixtureWindow).__sheet?.paste({
      range: { startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 }
    });
  });

  expect(await getCell(page, 0, 0)).toBe("Q");
  expect(await getCell(page, 0, 1)).toBe(5);
});

test("copy({ download: true }) triggers a real file download instead of writing to the clipboard", async ({
  page
}) => {
  await gotoFixture(page);

  const [download] = await Promise.all([
    page.waitForEvent("download"),
    page.evaluate(() =>
      (window as unknown as FixtureWindow).__sheet?.copy({
        download: true,
        range: { startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 }
      })
    )
  ]);

  expect(download.suggestedFilename()).toMatch(/^cyexcel-export-.*\.tsv$/);
  const stream = await download.createReadStream();
  const chunks: Buffer[] = [];
  await new Promise<void>((resolve, reject) => {
    stream?.on("data", (chunk: Buffer) => chunks.push(chunk));
    stream?.on("end", () => resolve());
    stream?.on("error", reject);
  });
  expect(Buffer.concat(chunks).toString("utf-8")).toBe("Alice\t30\r\nBob\t40");
});
