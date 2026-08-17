import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";

/**
 * Phase 7's core gate: a logical extent past what the browser can natively
 * scroll must still be fully reachable. The fixture is 800,000 rows x 20px =
 * 16,000,000 logical px against a 15,000,000px safe native extent, so it is
 * the one fixture in the suite that takes the segmented path — every other
 * one stays on the identity path, which is what keeps the rest of the browser
 * suite a regression net for "segmenting changed nothing for normal tables".
 *
 * Only the vertical axis is exercised here. The mapper is per-axis and both
 * axes run the identical class, so horizontal segmenting is covered by
 * ScrollAxisMapper's unit tests rather than by materialising 150,000 columns.
 */

interface Probe {
  readonly rowCount: number;
  readonly rowHeight: number;
  readonly headerHeight: number;
  readonly logicalExtent: number;
  resetProbe(): void;
  forceRepaint(): void;
  readonly requestedRows: { min: number; max: number };
  scrollTo(nativeTop: number): void;
  readonly nativeScrollTop: number;
  readonly nativeScrollHeight: number;
  readonly nativeMaxScrollTop: number;
}

declare global {
  interface Window {
    __cyexcelSegmentedScroll: Probe;
    __cyexcelSegmentedScrollReady?: boolean;
  }
}

async function gotoFixture(page: Page): Promise<void> {
  await page.goto("/tests/browser/segmented-scroll.html");
  await page.waitForFunction(() => window.__cyexcelSegmentedScrollReady === true);
}

/** Scrolls, then waits for the grid to actually ask about a fresh row range. */
async function scrollAndSettle(page: Page, nativeTop: number): Promise<{ min: number; max: number }> {
  await page.evaluate((top) => {
    const probe = window.__cyexcelSegmentedScroll;
    probe.resetProbe();
    const before = probe.nativeScrollTop;
    probe.scrollTo(top);
    // Assigning the offset it already has fires no scroll event and therefore
    // schedules no repaint, which would leave the probe empty forever.
    if (probe.nativeScrollTop === before) {
      probe.forceRepaint();
    }
  }, nativeTop);
  await expect
    .poll(() => page.evaluate(() => window.__cyexcelSegmentedScroll.requestedRows.max))
    .toBeGreaterThan(Number.NEGATIVE_INFINITY);
  return page.evaluate(() => window.__cyexcelSegmentedScroll.requestedRows);
}

test("the spacer is capped at the safe native extent instead of the logical one", async ({
  page
}) => {
  await gotoFixture(page);

  const { nativeScrollHeight, logicalExtent, headerHeight } = await page.evaluate(() => {
    const probe = window.__cyexcelSegmentedScroll;
    return {
      nativeScrollHeight: probe.nativeScrollHeight,
      logicalExtent: probe.logicalExtent,
      headerHeight: probe.headerHeight
    };
  });

  expect(logicalExtent).toBe(16_000_000);
  // Capped at 15,000,000 + the pinned header band, not the full 16,000,000.
  expect(nativeScrollHeight).toBe(15_000_000 + headerHeight);
  expect(nativeScrollHeight).toBeLessThan(logicalExtent);
});

test("scrolling to the native maximum reaches the very last logical row", async ({ page }) => {
  await gotoFixture(page);

  const max = await page.evaluate(() => window.__cyexcelSegmentedScroll.nativeMaxScrollTop);
  const requested = await scrollAndSettle(page, max);

  // Without segmenting the last reachable row would be around 15,000,000/20 =
  // 750,000 — 50,000 rows short of the end, permanently unreachable.
  expect(requested.max).toBe(799_999);
});

test("the last row stays reachable after the recentre settles, not just for one frame", async ({
  page
}) => {
  await gotoFixture(page);

  const max = await page.evaluate(() => window.__cyexcelSegmentedScroll.nativeMaxScrollTop);
  await scrollAndSettle(page, max);

  // Let the follow-up scroll event from the seek's own write land, then repaint
  // from scratch without moving. A mapping that only happened to be right for
  // the frame of the seek would show a different range here.
  await page.waitForTimeout(120);
  const settled = await page.evaluate(() => {
    const probe = window.__cyexcelSegmentedScroll;
    probe.resetProbe();
    probe.forceRepaint();
    return probe.requestedRows;
  });
  await expect
    .poll(() => page.evaluate(() => window.__cyexcelSegmentedScroll.requestedRows.max))
    .toBe(799_999);
  expect(settled.max).toBe(799_999);
});

test("a mid-axis position resolves to rows far past the native window", async ({ page }) => {
  await gotoFixture(page);

  const max = await page.evaluate(() => window.__cyexcelSegmentedScroll.nativeMaxScrollTop);
  const requested = await scrollAndSettle(page, max / 2);

  // Halfway down a 16,000,000px axis is logical row ~400,000. A non-segmented
  // mapping would land at half the native window instead — around row 375,000.
  expect(requested.min).toBeGreaterThan(380_000);
  expect(requested.max).toBeLessThan(420_000);
});

test("travelling down and back returns to row 0 without drift", async ({ page }) => {
  await gotoFixture(page);

  const max = await page.evaluate(() => window.__cyexcelSegmentedScroll.nativeMaxScrollTop);

  // Cross the axis in both directions several times, forcing repeated
  // recentres, then come back to the top.
  for (const fraction of [0.2, 0.6, 0.95, 0.4, 0.85, 0.1]) {
    await scrollAndSettle(page, max * fraction);
  }
  const backAtTop = await scrollAndSettle(page, 0);

  expect(backAtTop.min).toBe(0);
});

test("row ordering stays monotonic across a segment boundary", async ({ page }) => {
  await gotoFixture(page);

  const max = await page.evaluate(() => window.__cyexcelSegmentedScroll.nativeMaxScrollTop);
  const seen: number[] = [];
  for (const fraction of [0.1, 0.3, 0.5, 0.7, 0.9]) {
    const requested = await scrollAndSettle(page, max * fraction);
    seen.push(requested.min);
  }

  // Each stop must be strictly further down the logical axis than the last —
  // a recentre that mismapped would show up here as a jump backwards.
  for (let index = 1; index < seen.length; index += 1) {
    expect(seen[index]!).toBeGreaterThan(seen[index - 1]!);
  }
});

test("scrolling never leaves the native offset outside the scroller's own range", async ({
  page
}) => {
  await gotoFixture(page);

  const max = await page.evaluate(() => window.__cyexcelSegmentedScroll.nativeMaxScrollTop);
  for (const fraction of [0, 0.05, 0.5, 0.99, 1]) {
    await scrollAndSettle(page, max * fraction);
    const { nativeScrollTop, nativeMaxScrollTop } = await page.evaluate(() => {
      const probe = window.__cyexcelSegmentedScroll;
      return {
        nativeScrollTop: probe.nativeScrollTop,
        nativeMaxScrollTop: probe.nativeMaxScrollTop
      };
    });
    expect(nativeScrollTop).toBeGreaterThanOrEqual(0);
    expect(nativeScrollTop).toBeLessThanOrEqual(nativeMaxScrollTop);
  }
});
