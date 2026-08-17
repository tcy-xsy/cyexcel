import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";

interface ConstructorFailureResult {
  readonly error: {
    readonly name: string;
    readonly message: string;
    readonly code: string | null;
  } | null;
  readonly cleanupCalls: number;
  readonly host: {
    readonly baselineChildCount: number;
    readonly childCount: number;
    readonly addedChildCount: number;
    readonly sentinelPresent: boolean;
  };
  readonly pointerdown: {
    readonly added: number;
    readonly pairedRemoved: number;
    readonly outstanding: number;
  };
}

interface ConstructorRollbackWindow {
  __cyexcelConstructorRollbackReady?: boolean;
  __runPluginSetupRollback?(): ConstructorFailureResult;
  __runUnknownCellTypeRollback?(): ConstructorFailureResult;
}

async function gotoFixture(page: Page): Promise<void> {
  await page.goto("/tests/browser/constructor-rollback.html");
  await page.waitForFunction(
    () =>
      (window as unknown as ConstructorRollbackWindow)
        .__cyexcelConstructorRollbackReady === true
  );
}

function expectConstructionResourcesRolledBack(result: ConstructorFailureResult): void {
  expect.soft(result.host).toEqual({
    baselineChildCount: 1,
    childCount: 1,
    addedChildCount: 0,
    sentinelPresent: true
  });
  // Deferring listener activation until construction succeeds is also valid (0/0/0).
  expect.soft(result.pointerdown).toEqual({
    added: result.pointerdown.added,
    pairedRemoved: result.pointerdown.added,
    outstanding: 0
  });
}

test("a later plugin setup failure rolls back earlier plugins, DOM, and global listeners", async ({
  page
}) => {
  await gotoFixture(page);

  const result = await page.evaluate(() =>
    (window as unknown as ConstructorRollbackWindow).__runPluginSetupRollback?.()
  );

  expect(result).toBeDefined();
  expect(result?.error).toEqual({
    name: "Error",
    message: "plugin-b-setup-failed",
    code: null
  });
  expect.soft(result?.cleanupCalls).toBe(1);
  expectConstructionResourcesRolledBack(result as ConstructorFailureResult);
});

test("an unknown cell type failure rolls back plugins, DOM, and global listeners", async ({
  page
}) => {
  await gotoFixture(page);

  const result = await page.evaluate(() =>
    (window as unknown as ConstructorRollbackWindow).__runUnknownCellTypeRollback?.()
  );

  expect(result).toBeDefined();
  expect(result?.error).toMatchObject({
    name: "CyExcelError",
    code: "unknown-cell-type"
  });
  expect(result?.error?.message).toContain('Unknown cell type "missing-cell-type"');
  expect.soft(result?.cleanupCalls).toBe(1);
  expectConstructionResourcesRolledBack(result as ConstructorFailureResult);
});
