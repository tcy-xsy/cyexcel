import { expect, test } from "@playwright/test";

test("registers and mounts the cy-excel element without errors", async ({ page }) => {
  const errors: string[] = [];
  page.on("pageerror", error => errors.push(error.message));

  await page.goto("/tests/browser/fixture.html");

  const sheet = page.locator("cy-excel");
  await expect(sheet).toHaveCount(1);
  await expect
    .poll(() =>
      sheet.evaluate(element => ({
        defined: element.constructor.name === "CyExcelElement",
        hasViewport: Boolean(element.shadowRoot?.querySelector('[part~="viewport"]'))
      }))
    )
    .toEqual({ defined: true, hasViewport: true });
  expect(errors).toEqual([]);
});
