import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";

interface FixtureWindow {
  __cyexcelRowReorderReady?: boolean;
  __sheet?: {
    getCell(row: number, column: number): unknown;
    undo(): void;
    readonly scene: {
      getRowHeight(row: number): number;
    };
    readonly selection: {
      readonly kind: string;
      readonly range: {
        readonly startRow: number;
        readonly endRow: number;
      };
    };
  };
  __rowReorderTargetAt?(clientY: number): RowReorderTargetProbe;
  __finishStationaryRowReorder?(clientY: number): RowReorderTargetProbe;
}

interface RowReorderTargetProbe {
  readonly target: number;
  readonly finalRow: number;
  readonly scrollTop: number;
}

async function gotoFixture(page: Page): Promise<void> {
  await page.goto("/tests/browser/row-reorder.html");
  await page.waitForFunction(
    () => (window as unknown as FixtureWindow).__cyexcelRowReorderReady === true
  );
}

function canvasBox(page: Page, index: number) {
  return page.locator("#app canvas").nth(index).boundingBox();
}

function getCell(page: Page, row: number, column: number): Promise<unknown> {
  return page.evaluate(
    ({ row: r, column: c }) => (window as unknown as FixtureWindow).__sheet?.getCell(r, c),
    { row, column }
  );
}

function undo(page: Page): Promise<void> {
  return page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.undo());
}

async function nextFrame(page: Page): Promise<void> {
  await page.waitForFunction(
    () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)))
  );
}

function rowHeight(page: Page, row: number): Promise<number | undefined> {
  return page.evaluate(
    (targetRow) => (window as unknown as FixtureWindow).__sheet?.scene.getRowHeight(targetRow),
    row
  );
}

function selection(page: Page) {
  return page.evaluate(() => {
    const snapshot = (window as unknown as FixtureWindow).__sheet?.selection;
    return snapshot
      ? {
          kind: snapshot.kind,
          startRow: snapshot.range.startRow,
          endRow: snapshot.range.endRow
        }
      : undefined;
  });
}

/**
 * Average red channel over a small logical-pixel box on the row-header
 * canvas (index 3), DPR-aware. The row header's own background
 * (HEADER_BACKGROUND_COLOR "#f3f4f6") is opaque everywhere a row is drawn,
 * so — unlike other canvases in this suite — a plain non-transparent-pixel
 * count can't distinguish "indicator drawn" from "just the ordinary
 * background"; the indicator's blue (ROW_REORDER_INDICATOR_COLOR
 * "#3b82f6", R=59) is what stands out against that near-white background
 * (R=243), so comparing the average red channel is the reliable signal.
 */
function averageRed(page: Page, x: number, y: number, width: number, height: number): Promise<number> {
  return page.evaluate(
    ({ x, y, width, height }) => {
      const canvas = document.querySelectorAll("#app canvas")[3] as HTMLCanvasElement;
      const ctx = canvas.getContext("2d") as CanvasRenderingContext2D;
      const dpr = window.devicePixelRatio || 1;
      const { data } = ctx.getImageData(
        Math.round(x * dpr),
        Math.round(y * dpr),
        Math.max(1, Math.round(width * dpr)),
        Math.max(1, Math.round(height * dpr))
      );
      let sum = 0;
      let count = 0;
      for (let i = 0; i < data.length; i += 4) {
        sum += data[i] ?? 0;
        count += 1;
      }
      return count > 0 ? sum / count : 0;
    },
    { x, y, width, height }
  );
}

test("dragging the top edge of a row selection to a new position reorders the row, as a single undo step", async ({
  page
}) => {
  await gotoFixture(page);
  const rowHeaderBox = await canvasBox(page, 3);
  if (!rowHeaderBox) throw new Error("missing row header canvas box");

  expect(await getCell(page, 5, 0)).toBe("5-0");
  expect(await getCell(page, 7, 0)).toBe("7-0");

  await page.mouse.click(rowHeaderBox.x + 20, rowHeaderBox.y + 130); // selects row 5
  await page.mouse.move(rowHeaderBox.x + 20, rowHeaderBox.y + 120); // row 5's top edge
  await page.mouse.down();
  await page.mouse.move(rowHeaderBox.x + 20, rowHeaderBox.y + 185); // row 8's top half -> target 8
  await page.mouse.up();
  await page.waitForTimeout(50);

  // moveRows(5, 1, 8) -> adjustedTarget 7: row 5's data now sits at index 7,
  // and everything in between shifted up by one.
  expect(await getCell(page, 7, 0)).toBe("5-0");
  expect(await getCell(page, 5, 0)).toBe("6-0");

  await undo(page);
  expect(await getCell(page, 5, 0)).toBe("5-0");
  expect(await getCell(page, 7, 0)).toBe("7-0");
});

test("row-header routing resolves overlapping gestures in resize, reorder, then selection priority", async ({
  page
}) => {
  await gotoFixture(page);
  const rowHeaderBox = await canvasBox(page, 3);
  if (!rowHeaderBox) throw new Error("missing row header canvas box");

  await page.mouse.click(rowHeaderBox.x + 20, rowHeaderBox.y + 130); // selects row 5

  // This point is both row 4's resize boundary and the selected row 5's
  // reorder edge. The trailing grip makes resize the highest-priority hit.
  await page.mouse.move(rowHeaderBox.x + rowHeaderBox.width - 3, rowHeaderBox.y + 120);
  await page.mouse.down();
  await page.mouse.move(rowHeaderBox.x + rowHeaderBox.width - 3, rowHeaderBox.y + 128);
  await page.mouse.up();
  await nextFrame(page);

  expect(await rowHeight(page, 4)).toBe(28);
  expect(await getCell(page, 5, 0)).toBe("5-0");

  await undo(page);
  expect(await rowHeight(page, 4)).toBe(20);

  // Away from the resize grip, the same selection edge is shared by
  // reorder and ordinary row selection. Reorder owns it.
  await page.mouse.click(rowHeaderBox.x + 20, rowHeaderBox.y + 130);
  await page.mouse.move(rowHeaderBox.x + 20, rowHeaderBox.y + 120);
  await page.mouse.down();
  await page.mouse.move(rowHeaderBox.x + 20, rowHeaderBox.y + 185);
  await page.mouse.up();
  await nextFrame(page);

  expect(await getCell(page, 7, 0)).toBe("5-0");
  await undo(page);

  // With neither higher-priority hit present, the route falls through to
  // row selection and extends the range without moving data.
  await page.mouse.click(rowHeaderBox.x + 20, rowHeaderBox.y + 130);
  await page.mouse.move(rowHeaderBox.x + 20, rowHeaderBox.y + 130);
  await page.mouse.down();
  await page.mouse.move(rowHeaderBox.x + 20, rowHeaderBox.y + 190);
  await page.mouse.up();
  await nextFrame(page);

  expect(await selection(page)).toEqual({ kind: "row", startRow: 5, endRow: 8 });
  expect(await getCell(page, 5, 0)).toBe("5-0");
  expect(await getCell(page, 8, 0)).toBe("8-0");
});

test("an insertion indicator line is visible on the row header during the drag", async ({ page }) => {
  await gotoFixture(page);
  const rowHeaderBox = await canvasBox(page, 3);
  if (!rowHeaderBox) throw new Error("missing row header canvas box");

  // Row 8's top edge (canvas-local y = headerHeight(20) + 8*20 = 180) is
  // just the plain (near-white) header background before any drag starts.
  // Sampled near the left edge (x=2..8), away from the horizontally
  // centered row-number text at x=20 — the indicator bar spans the full
  // width, so this box catches it without ever touching a digit's glyph.
  const restingRed = await averageRed(page, 2, 178, 6, 5);
  expect(restingRed).toBeGreaterThan(200);

  await page.mouse.click(rowHeaderBox.x + 20, rowHeaderBox.y + 130); // selects row 5
  await page.mouse.move(rowHeaderBox.x + 20, rowHeaderBox.y + 120);
  await page.mouse.down();
  await page.mouse.move(rowHeaderBox.x + 20, rowHeaderBox.y + 185); // targets boundary 8
  await page.waitForTimeout(50);

  // The indicator's blue (R=59) pulls the average red channel well below
  // the resting background's (R=243).
  const draggingRed = await averageRed(page, 2, 178, 6, 5);
  expect(draggingRed).toBeLessThan(restingRed - 50);

  await page.mouse.up();
  await page.waitForTimeout(50);

  // The indicator is cleared once the drag ends.
  expect(await averageRed(page, 2, 178, 6, 5)).toBeGreaterThan(200);
});

test("Escape cancels an in-progress reorder drag without moving anything", async ({ page }) => {
  await gotoFixture(page);
  const rowHeaderBox = await canvasBox(page, 3);
  if (!rowHeaderBox) throw new Error("missing row header canvas box");

  await page.mouse.click(rowHeaderBox.x + 20, rowHeaderBox.y + 130); // selects row 5
  await page.mouse.move(rowHeaderBox.x + 20, rowHeaderBox.y + 120);
  await page.mouse.down();
  await page.mouse.move(rowHeaderBox.x + 20, rowHeaderBox.y + 185);
  await page.waitForTimeout(50);

  await page.keyboard.press("Escape");
  await page.mouse.up();
  await page.waitForTimeout(50);

  expect(await getCell(page, 5, 0)).toBe("5-0");
  expect(await getCell(page, 7, 0)).toBe("7-0");
  expect(await averageRed(page, 2, 178, 6, 5)).toBeGreaterThan(200);
});

test("dragging from the middle of the row header (not an edge) still does ordinary multi-select", async ({
  page
}) => {
  await gotoFixture(page);
  const rowHeaderBox = await canvasBox(page, 3);
  if (!rowHeaderBox) throw new Error("missing row header canvas box");

  await page.mouse.click(rowHeaderBox.x + 20, rowHeaderBox.y + 130); // selects row 5 (center, not an edge)
  await page.mouse.move(rowHeaderBox.x + 20, rowHeaderBox.y + 130);
  await page.mouse.down();
  await page.mouse.move(rowHeaderBox.x + 20, rowHeaderBox.y + 190); // drag down toward row 8
  await page.mouse.up();
  await page.waitForTimeout(50);

  // Ordinary select-drag extended the selection, it did not reorder anything.
  expect(await getCell(page, 5, 0)).toBe("5-0");
  expect(await getCell(page, 8, 0)).toBe("8-0");
});

test("dragging a frozen row past the frozen boundary into scrollable rows reorders correctly", async ({
  page
}) => {
  await gotoFixture(page);
  const rowHeaderBox = await canvasBox(page, 3);
  if (!rowHeaderBox) throw new Error("missing row header canvas box");

  await page.mouse.click(rowHeaderBox.x + 20, rowHeaderBox.y + 30); // selects row 0 (frozen)
  await page.mouse.move(rowHeaderBox.x + 20, rowHeaderBox.y + 40); // row 0's bottom edge
  await page.mouse.down();
  await page.mouse.move(rowHeaderBox.x + 20, rowHeaderBox.y + 125); // row 5's top half -> target 5
  await page.mouse.up();
  await page.waitForTimeout(50);

  // moveRows(0, 1, 5) -> adjustedTarget 4: row 0's data now sits at index 4,
  // and row 1 (previously frozen alongside row 0) shifts up to become the
  // new frozen row 0.
  expect(await getCell(page, 4, 0)).toBe("0-0");
  expect(await getCell(page, 0, 0)).toBe("1-0");
});

test("dragging a scrollable row up past the frozen boundary reorders correctly", async ({ page }) => {
  await gotoFixture(page);
  const rowHeaderBox = await canvasBox(page, 3);
  if (!rowHeaderBox) throw new Error("missing row header canvas box");

  await page.mouse.click(rowHeaderBox.x + 20, rowHeaderBox.y + 130); // selects row 5
  await page.mouse.move(rowHeaderBox.x + 20, rowHeaderBox.y + 120); // row 5's top edge
  await page.mouse.down();
  await page.mouse.move(rowHeaderBox.x + 20, rowHeaderBox.y + 25); // row 0's top half -> target 0
  await page.mouse.up();
  await page.waitForTimeout(50);

  // moveRows(5, 1, 0): row 5's data now sits at index 0 (frozen), and row 0
  // shifts down to index 1.
  expect(await getCell(page, 0, 0)).toBe("5-0");
  expect(await getCell(page, 1, 0)).toBe("0-0");
});

test("hovering the reorder handle shows a grab cursor; hovering elsewhere in a selected row doesn't", async ({
  page
}) => {
  await gotoFixture(page);
  const rowHeaderBox = await canvasBox(page, 3);
  if (!rowHeaderBox) throw new Error("missing row header canvas box");

  const cursorOf = () =>
    page.evaluate(() => (document.querySelectorAll("#app canvas")[3] as HTMLCanvasElement).style.cursor);

  await page.mouse.click(rowHeaderBox.x + 20, rowHeaderBox.y + 130); // selects row 5

  await page.mouse.move(rowHeaderBox.x + 20, rowHeaderBox.y + 120); // row 5's top edge
  await page.waitForTimeout(50);
  expect(await cursorOf()).toBe("grab");

  await page.mouse.move(rowHeaderBox.x + 20, rowHeaderBox.y + 130); // well inside the selected row
  await page.waitForTimeout(50);
  expect(await cursorOf()).toBe("");

  await page.mouse.move(rowHeaderBox.x + 20, rowHeaderBox.y + 140); // row 5's bottom edge
  await page.waitForTimeout(50);
  expect(await cursorOf()).toBe("grab");
});

test("the grabbing cursor stays visible on document.body throughout the drag, even after the pointer leaves the narrow row-header column", async ({
  page
}) => {
  await gotoFixture(page);
  const rowHeaderBox = await canvasBox(page, 3);
  if (!rowHeaderBox) throw new Error("missing row header canvas box");

  const bodyCursor = () => page.evaluate(() => document.body.style.cursor);

  expect(await bodyCursor()).toBe("");

  await page.mouse.click(rowHeaderBox.x + 20, rowHeaderBox.y + 130); // selects row 5
  await page.mouse.move(rowHeaderBox.x + 20, rowHeaderBox.y + 120); // row 5's top edge
  await page.mouse.down();
  expect(await bodyCursor()).toBe("grabbing");

  // Drift well outside the row-header canvas's own box (it's only
  // rowHeaderWidth=40px wide) — a CSS cursor set on that element alone
  // would stop applying here, but document.body's does not.
  await page.mouse.move(rowHeaderBox.x + 300, rowHeaderBox.y + 500);
  expect(await bodyCursor()).toBe("grabbing");

  await page.mouse.up();
  expect(await bodyCursor()).toBe("");
});

test("Escape also clears the grabbing cursor on document.body", async ({ page }) => {
  await gotoFixture(page);
  const rowHeaderBox = await canvasBox(page, 3);
  if (!rowHeaderBox) throw new Error("missing row header canvas box");

  const bodyCursor = () => page.evaluate(() => document.body.style.cursor);

  await page.mouse.click(rowHeaderBox.x + 20, rowHeaderBox.y + 130); // selects row 5
  await page.mouse.move(rowHeaderBox.x + 20, rowHeaderBox.y + 120); // row 5's top edge
  await page.mouse.down();
  expect(await bodyCursor()).toBe("grabbing");

  await page.keyboard.press("Escape");
  expect(await bodyCursor()).toBe("");

  await page.mouse.up();
});

test("window blur cancels reorder, clears its indicator and cursor, and ignores later pointer movement", async ({
  page
}) => {
  await gotoFixture(page);
  const rowHeaderBox = await canvasBox(page, 3);
  if (!rowHeaderBox) throw new Error("missing row header canvas box");

  await page.mouse.click(rowHeaderBox.x + 20, rowHeaderBox.y + 130); // selects row 5
  await page.mouse.move(rowHeaderBox.x + 20, rowHeaderBox.y + 120);
  await page.mouse.down();
  await page.mouse.move(rowHeaderBox.x + 20, rowHeaderBox.y + 185);
  await nextFrame(page);

  expect(await averageRed(page, 2, 178, 6, 5)).toBeLessThan(190);
  expect(await page.evaluate(() => document.body.style.cursor)).toBe("grabbing");

  await page.evaluate(() => window.dispatchEvent(new Event("blur")));
  await nextFrame(page);

  expect(await averageRed(page, 2, 178, 6, 5)).toBeGreaterThan(200);
  expect(await page.evaluate(() => document.body.style.cursor)).toBe("");

  await page.mouse.move(rowHeaderBox.x + 20, rowHeaderBox.y + 225);
  await page.mouse.up();
  await nextFrame(page);

  expect(await getCell(page, 5, 0)).toBe("5-0");
  expect(await getCell(page, 7, 0)).toBe("7-0");
  expect(await averageRed(page, 2, 218, 6, 5)).toBeGreaterThan(200);
});

test("dragging correctly accounts for the current scroll position, not just the unscrolled layout", async ({
  page
}) => {
  await gotoFixture(page);
  const rowHeaderBox = await canvasBox(page, 3);
  if (!rowHeaderBox) throw new Error("missing row header canvas box");

  // Scroll the body 100px down first — row 7 (absolute y=140) is now the
  // first fully-visible scrollable row, right after the 2 frozen rows
  // (canvas-local y = headerHeight(20) + (140 - 100) = 60).
  await page.evaluate(() => {
    (document.querySelector("#app")!.firstElementChild as HTMLElement).scrollTop = 100;
  });
  await page.waitForTimeout(50);

  // Select row 7 at its on-screen center (canvas-local y = 20 + 40 + 10 = 70).
  await page.mouse.click(rowHeaderBox.x + 20, rowHeaderBox.y + 70);
  // Start the drag from row 7's true on-screen top edge (canvas-local y=60)
  // — with an unsynced viewport, RowReorderController would still think
  // row 7's edge is at the unscrolled y=140 and miss this hit entirely.
  await page.mouse.move(rowHeaderBox.x + 20, rowHeaderBox.y + 60);
  await page.mouse.down();
  // Drop on row 10's bottom half (canvas-local y = 20 + (200-100) + 15 = 135) -> target 11.
  await page.mouse.move(rowHeaderBox.x + 20, rowHeaderBox.y + 135);
  await page.mouse.up();
  await page.waitForTimeout(50);

  // moveRows(7, 1, 11) -> adjustedTarget 10: row 7's data now sits at index 10.
  expect(await getCell(page, 10, 0)).toBe("7-0");
  expect(await getCell(page, 7, 0)).toBe("8-0");
});

test("a stationary pointer recomputes the reorder target from each fresh auto-scroll frame", async ({
  page
}) => {
  await gotoFixture(page);
  const rowHeaderBox = await canvasBox(page, 3);
  if (!rowHeaderBox) throw new Error("missing row header canvas box");

  const scrollTop = () =>
    page.evaluate(() => (document.querySelector("#app")!.firstElementChild as HTMLElement).scrollTop);

  expect(await scrollTop()).toBe(0);

  await page.mouse.click(rowHeaderBox.x + 20, rowHeaderBox.y + 130); // selects row 5
  await page.mouse.move(rowHeaderBox.x + 20, rowHeaderBox.y + 120); // row 5's top edge
  await page.mouse.down();
  const stationaryClientY = rowHeaderBox.y + rowHeaderBox.height - 5;
  // This is the last pointermove in the gesture. All subsequent target
  // changes must come from auto-scroll ticks acquiring fresh frames.
  await page.mouse.move(rowHeaderBox.x + 20, stationaryClientY);

  const beforeScroll = await page.evaluate(
    (clientY) => (window as unknown as FixtureWindow).__rowReorderTargetAt?.(clientY),
    stationaryClientY
  );
  if (!beforeScroll) throw new Error("missing row reorder target probe");

  await page.waitForFunction(
    () => (document.querySelector("#app")!.firstElementChild as HTMLElement).scrollTop >= 80
  );
  expect(await scrollTop()).toBeGreaterThanOrEqual(80);

  // Probe the current public Scene and dispatch pointerup synchronously in
  // the same browser task, so another animation frame cannot race the
  // expected target captured here.
  const afterScroll = await page.evaluate(
    (clientY) => (window as unknown as FixtureWindow).__finishStationaryRowReorder?.(clientY),
    stationaryClientY
  );
  if (!afterScroll) throw new Error("missing stationary reorder completion probe");
  await nextFrame(page);

  expect(afterScroll.scrollTop).toBeGreaterThan(beforeScroll.scrollTop);
  expect(afterScroll.target).toBeGreaterThan(beforeScroll.target);
  expect(afterScroll.finalRow).toBeGreaterThan(beforeScroll.finalRow);
  expect(await getCell(page, afterScroll.finalRow, 0)).toBe("5-0");
  expect(await getCell(page, beforeScroll.finalRow, 0)).not.toBe("5-0");
});
