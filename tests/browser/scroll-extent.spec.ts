import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";

interface FixtureWindow {
  __cyexcelScrollExtentReady?: boolean;
}

async function gotoFixture(page: Page): Promise<void> {
  await page.goto("/tests/browser/scroll-extent.html");
  await page.waitForFunction(
    () => (window as unknown as FixtureWindow).__cyexcelScrollExtentReady === true
  );
}

/**
 * Regression test for a real bug found via a user report against the
 * playground (a 10-column dataset with rowHeaderWidth/headerHeight set —
 * the last column's text was visibly clipped even after scrolling all the
 * way right). Root cause: #updateSpacerSize() sized the spacer to just
 * `layout.getTotalWidth()/getTotalHeight()`, but #resizeCanvases() treats
 * the row/column header bands as sharing the scroller's own box
 * (`rect.width - rowHeaderWidth`, `rect.height - headerHeight()`) rather
 * than adding to it — so the spacer (and therefore the scroller's native
 * scrollWidth/scrollHeight) came up short by exactly rowHeaderWidth/
 * headerHeight, permanently hiding that many pixels of the last column/row
 * behind nothing reachable by scrolling.
 *
 * Fixture: 4 columns × 100px = 400px total (container is 300px wide, so
 * horizontal scroll is required), rowHeaderWidth 32; 20 rows × 20px = 400px
 * total (container is 200px tall, so vertical scroll is required),
 * headerHeight 24.
 */
test("scrolling to the native max reveals the full last column and last row, accounting for the pinned row/column header bands", async ({
  page
}) => {
  await gotoFixture(page);

  const scrollerHandle = await page.evaluateHandle(
    () => document.querySelector("#app")!.firstElementChild as HTMLElement
  );

  await page.evaluate((el) => {
    el.scrollLeft = el.scrollWidth;
    el.scrollTop = el.scrollHeight;
  }, scrollerHandle);
  await page.waitForTimeout(50);

  const { scrollLeft, scrollTop, clientWidth, clientHeight } = await page.evaluate(
    (el) => ({
      scrollLeft: el.scrollLeft,
      scrollTop: el.scrollTop,
      clientWidth: el.clientWidth,
      clientHeight: el.clientHeight
    }),
    scrollerHandle
  );

  // Total column width (400) minus the body's own visible width
  // (clientWidth - rowHeaderWidth 32) is the scrollLeft needed to make the
  // last column's right edge land exactly at the viewport's right edge.
  // Before the fix this fell short by exactly 32px (rowHeaderWidth).
  const expectedScrollLeft = 400 - (clientWidth - 32);
  expect(scrollLeft).toBe(expectedScrollLeft);

  // Same reasoning vertically: total row height (400) minus the body's own
  // visible height (clientHeight - headerHeight 24). Before the fix this
  // fell short by exactly 24px (headerHeight).
  const expectedScrollTop = 400 - (clientHeight - 24);
  expect(scrollTop).toBe(expectedScrollTop);
});
