import { expect, test } from "@playwright/test";

test("exposes the minimal element lifecycle through the public element", async ({ page }) => {
  await page.goto("/tests/browser/fixture.html");

  const state = await page.locator("cy-excel").evaluate(element => {
    const sheet = element as HTMLElement & {
      readonly core: { readonly disposed: boolean };
      dispose(): void;
    };

    const before = sheet.core.disposed;
    sheet.dispose();

    return { before, after: sheet.core.disposed };
  });

  expect(state).toEqual({ before: false, after: true });
});
