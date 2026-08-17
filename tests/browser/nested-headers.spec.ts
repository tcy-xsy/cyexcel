import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";

interface FixtureWindow {
  __cyexcelNestedHeadersReady?: boolean;
  __hitTest?: (point: { x: number; y: number }) => { level: number; group: { id: string } } | null;
  __sheet?: {
    readonly selection: {
      readonly kind: string;
      readonly active: { readonly row: number; readonly column: number };
    };
    readonly scene: {
      getCellRect(row: number, column: number): { readonly width: number };
    };
  };
}

async function gotoFixture(page: Page): Promise<void> {
  await page.goto("/tests/browser/nested-headers.html");
  await page.waitForFunction(
    () => (window as unknown as FixtureWindow).__cyexcelNestedHeadersReady === true
  );
}

/** nonTransparentCount for a small logical-pixel box on the header canvas (index 2), DPR-aware — same helper shape used throughout this suite. */
function sampleAlpha(
  page: Page,
  x: number,
  y: number,
  width: number,
  height: number
): Promise<number> {
  return page.evaluate(
    ({ x, y, width, height }) => {
      const canvas = document.querySelectorAll("#app canvas")[2] as HTMLCanvasElement;
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

function hitTest(page: Page, point: { x: number; y: number }) {
  return page.evaluate((p) => (window as unknown as FixtureWindow).__hitTest?.(p) ?? null, point);
}

// Bands, all headerHeight/group-row-height 24 (CyExcelCore reuses
// headerHeight as each group row's height too — no separate
// groupHeaderHeight option, see nested-headers.html): root (level 2)
// y=[0,24), grpA (level 1) y=[24,48), leafA1/leafA2 (level 0) y=[48,72),
// leaf title row y=[72,96).

test("all three group levels plus the leaf row paint real, non-transparent pixels", async ({ page }) => {
  await gotoFixture(page);

  expect(await sampleAlpha(page, 20, 8, 40, 10)).toBeGreaterThan(0); // level 2 "root"
  expect(await sampleAlpha(page, 20, 32, 40, 10)).toBeGreaterThan(0); // level 1 "grpA"
  expect(await sampleAlpha(page, 20, 56, 40, 10)).toBeGreaterThan(0); // level 0 "leafA1"
  expect(await sampleAlpha(page, 20, 80, 40, 10)).toBeGreaterThan(0); // leaf row "C0"
});

test("the header canvas's logical height equals columnHeaderHeight + groupHeaderHeight * groupDepth", async ({
  page
}) => {
  await gotoFixture(page);

  const box = await page.locator("#app canvas").nth(2).boundingBox();
  expect(box?.height).toBe(24 + 24 * 3); // 96
});

test("a group straddling the frozen boundary (root/grpA, columns 0-5 across the freeze at column 2) paints on both sides", async ({
  page
}) => {
  await gotoFixture(page);

  // x=50 is inside frozen column 0; x=300 is inside scrollable column 3 —
  // both are covered by "root" (level 2) and "grpA" (level 1), which span
  // the full [0,5] column range and therefore straddle the freeze boundary.
  expect(await sampleAlpha(page, 40, 8, 20, 10)).toBeGreaterThan(0); // root, frozen side
  expect(await sampleAlpha(page, 290, 8, 20, 10)).toBeGreaterThan(0); // root, scrollable side
  expect(await sampleAlpha(page, 40, 32, 20, 10)).toBeGreaterThan(0); // grpA, frozen side
  expect(await sampleAlpha(page, 290, 32, 20, 10)).toBeGreaterThan(0); // grpA, scrollable side
});

test("hitTestHeaderGroup resolves each level's group correctly from a real bundled build", async ({ page }) => {
  await gotoFixture(page);

  // x=50 -> column 0 (leafA1's range); y picks the level per the bands above.
  expect(await hitTest(page, { x: 50, y: 12 })).toMatchObject({ level: 2, group: { id: "root" } });
  expect(await hitTest(page, { x: 50, y: 36 })).toMatchObject({ level: 1, group: { id: "grpA" } });
  expect(await hitTest(page, { x: 50, y: 60 })).toMatchObject({ level: 0, group: { id: "leafA1" } });

  // x=300 -> column 3 (leafA2's range, columns [3,4,5]).
  expect(await hitTest(page, { x: 300, y: 60 })).toMatchObject({ level: 0, group: { id: "leafA2" } });

  // Below the group band (the leaf title row) resolves to nothing.
  expect(await hitTest(page, { x: 50, y: 80 })).toBeNull();
});

test("group bands do not select or resize columns, while the leaf band still selects", async ({ page }) => {
  await gotoFixture(page);
  const header = page.locator("#app canvas").nth(2);
  const box = await header.boundingBox();
  if (!box) throw new Error("missing header canvas box");

  await page.mouse.click(box.x + 280, box.y + 12);
  expect(
    await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.selection)
  ).toMatchObject({ kind: "cell", active: { row: 0, column: 0 } });

  await page.mouse.move(box.x + 80, box.y + 12);
  expect(await header.evaluate((canvas: HTMLCanvasElement) => canvas.style.cursor)).toBe("");
  await page.mouse.down();
  await page.mouse.move(box.x + 110, box.y + 12);
  await page.mouse.up();
  expect(
    await page.evaluate(() =>
      (window as unknown as FixtureWindow).__sheet?.scene.getCellRect(0, 0).width
    )
  ).toBe(80);

  await page.mouse.click(box.x + 280, box.y + 84);
  expect(
    await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.selection)
  ).toMatchObject({ kind: "column", active: { column: 3 } });
});
