import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";

type ResizeCommand = "resize-column" | "resize-row";
type ResizeKind = "column" | "row";

interface FixtureWindow {
  __cyexcelResizeGestureReady?: boolean;
  __resizeTimeline?: Array<Record<string, unknown>>;
  __rejectResizeCommand?(type: ResizeCommand | null): void;
  __rejectResizeTransaction?(type: ResizeCommand | null): void;
  __dispatchResizePointerCancel?(): void;
  __dispatchUnrelatedResizePointerCancel?(): void;
  __throwOnResizePointerCapture?(kind: ResizeKind): void;
  __resizeCanvasHasCapture?(kind: ResizeKind): boolean;
  __sceneIdentity?(): SceneIdentity;
  __lastPaintedWidth?(): number | undefined;
  __sheet?: {
    dispose(): void;
    insertRows(start: number, count?: number): void;
    startEdit(row: number, column: number): void;
    commitEdit(): Promise<void>;
    isEditing(): boolean;
    getCell(row: number, column: number): unknown;
    undo(): void;
    redo(): void;
  };
}

interface SceneIdentity {
  readonly generation: number;
  readonly viewportGeneration: number;
  readonly modelVersion: number;
  readonly layoutVersion: number;
}

interface CellRect {
  readonly left: number;
  readonly top: number;
  readonly width: number;
  readonly height: number;
}

async function gotoFixture(page: Page): Promise<void> {
  await page.goto("/tests/browser/resize-gesture.html");
  await page.waitForFunction(
    () => (window as unknown as FixtureWindow).__cyexcelResizeGestureReady === true
  );
}

async function nextFrame(page: Page): Promise<void> {
  await page.waitForFunction(
    () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)))
  );
}

async function editorRect(page: Page, row: number, column: number): Promise<CellRect> {
  await page.evaluate(
    ({ row: targetRow, column: targetColumn }) =>
      (window as unknown as FixtureWindow).__sheet?.startEdit(targetRow, targetColumn),
    { row, column }
  );
  const rect = await page.locator("#app input").evaluate((element: HTMLInputElement) => ({
    left: Number.parseFloat(element.style.left),
    top: Number.parseFloat(element.style.top),
    width: Number.parseFloat(element.style.width),
    height: Number.parseFloat(element.style.height)
  }));
  await page.keyboard.press("Escape");
  return rect;
}

async function canvasBox(page: Page, kind: ResizeKind) {
  return page.locator("#app canvas").nth(kind === "column" ? 2 : 3).boundingBox();
}

async function beginColumnResize(page: Page, delta: number): Promise<void> {
  const box = await canvasBox(page, "column");
  if (!box) throw new Error("missing column header canvas");
  await page.mouse.move(box.x + 80, box.y + 10);
  await page.mouse.down();
  await page.mouse.move(box.x + 80 + delta, box.y + 10);
  await nextFrame(page);
}

async function beginRowResize(page: Page, delta: number): Promise<void> {
  const box = await canvasBox(page, "row");
  if (!box) throw new Error("missing row header canvas");
  const boundaryY = box.y + 20 + 24;
  await page.mouse.move(box.x + box.width - 3, boundaryY);
  await page.mouse.down();
  await page.mouse.move(box.x + box.width - 3, boundaryY + delta);
  await nextFrame(page);
}

async function timeline(page: Page): Promise<Array<Record<string, unknown>>> {
  return page.evaluate(() => (window as unknown as FixtureWindow).__resizeTimeline ?? []);
}

async function sceneIdentity(page: Page): Promise<SceneIdentity> {
  return page.evaluate(() => {
    const identity = (window as unknown as FixtureWindow).__sceneIdentity?.();
    if (!identity) throw new Error("missing Scene identity probe");
    return identity;
  });
}

test("resize previews, cancellation and viewport sync publish strictly increasing Scene generations", async ({
  page
}) => {
  await gotoFixture(page);
  const identities = [await sceneIdentity(page)];
  const box = await canvasBox(page, "column");
  if (!box) throw new Error("missing column header canvas");

  await page.mouse.move(box.x + 80, box.y + 10);
  await page.mouse.down();
  await page.mouse.move(box.x + 100, box.y + 10);
  identities.push(await sceneIdentity(page));
  await page.mouse.move(box.x + 110, box.y + 10);
  identities.push(await sceneIdentity(page));
  await page.keyboard.press("Escape");
  identities.push(await sceneIdentity(page));
  await page.mouse.up();

  await page.locator("#app").evaluate((element) => {
    const scroller = element.firstElementChild as HTMLElement;
    scroller.scrollLeft = 20;
    scroller.dispatchEvent(new Event("scroll"));
  });
  identities.push(await sceneIdentity(page));

  expect(identities.map((identity) => identity.generation)).toEqual(
    [...identities].map((identity) => identity.generation).sort((left, right) => left - right)
  );
  expect(new Set(identities.map((identity) => identity.generation)).size).toBe(identities.length);
  expect(identities.every((identity) => identity.viewportGeneration === identity.generation)).toBe(true);

  // Preview/cancel/scroll do not masquerade as model or committed-layout changes.
  expect(new Set(identities.map((identity) => identity.modelVersion))).toEqual(new Set([0]));
  expect(identities[0]?.layoutVersion).toBe(identities[3]?.layoutVersion);
  expect(identities[4]?.layoutVersion).toBe(identities[0]?.layoutVersion);
  expect(identities[1]?.layoutVersion).toBe((identities[0]?.layoutVersion ?? 0) + 1);
  expect(identities[2]?.layoutVersion).toBe(identities[1]?.layoutVersion);
});

test("a committed resize publishes after its preview with committed model/layout identity", async ({ page }) => {
  await gotoFixture(page);
  const before = await sceneIdentity(page);
  await beginColumnResize(page, 30);
  const preview = await sceneIdentity(page);
  await page.mouse.up();
  const committed = await sceneIdentity(page);

  expect(preview.generation).toBeGreaterThan(before.generation);
  expect(committed.generation).toBeGreaterThan(preview.generation);
  expect(preview.modelVersion).toBe(before.modelVersion);
  expect(committed.modelVersion).toBe(before.modelVersion + 1);
  expect(committed.layoutVersion).toBe(preview.layoutVersion);
});

test("resize preview and cancellation synchronously repaint their published geometry", async ({ page }) => {
  await gotoFixture(page);
  const paintedWidths = await page.evaluate(() =>
    (window as unknown as FixtureWindow).__lastPaintedWidth?.()
  );
  expect(paintedWidths).toBe(80);

  const box = await canvasBox(page, "column");
  if (!box) throw new Error("missing column header canvas");
  await page.mouse.move(box.x + 80, box.y + 10);
  await page.mouse.down();
  await page.mouse.move(box.x + 110, box.y + 10);
  expect(await page.evaluate(() =>
    (window as unknown as FixtureWindow).__lastPaintedWidth?.()
  )).toBe(110);

  await page.keyboard.press("Escape");
  expect(await page.evaluate(() =>
    (window as unknown as FixtureWindow).__lastPaintedWidth?.()
  )).toBe(80);
  await page.mouse.up();
});

test("column resize captures its pointer and pointercancel restores preview geometry without history", async ({
  page
}) => {
  await gotoFixture(page);
  const before = await editorRect(page, 0, 1);

  await beginColumnResize(page, 30);
  expect(
    await page.evaluate(() =>
      (window as unknown as FixtureWindow).__resizeCanvasHasCapture?.("column")
    )
  ).toBe(true);
  await page.evaluate(() =>
    (window as unknown as FixtureWindow).__dispatchResizePointerCancel?.()
  );
  await page.mouse.up();
  await nextFrame(page);

  expect(await editorRect(page, 0, 1)).toEqual(before);
  expect(await timeline(page)).toEqual([]);
  await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.undo());
  expect(await editorRect(page, 0, 1)).toEqual(before);
});

test("Escape cancels row resize and releases capture", async ({ page }) => {
  await gotoFixture(page);
  const before = await editorRect(page, 1, 0);

  await beginRowResize(page, 20);
  expect(
    await page.evaluate(() =>
      (window as unknown as FixtureWindow).__resizeCanvasHasCapture?.("row")
    )
  ).toBe(true);
  await page.keyboard.press("Escape");
  expect(
    await page.evaluate(() =>
      (window as unknown as FixtureWindow).__resizeCanvasHasCapture?.("row")
    )
  ).toBe(false);
  await page.mouse.up();
  await nextFrame(page);

  expect(await editorRect(page, 1, 0)).toEqual(before);
  expect(await timeline(page)).toEqual([]);
});

test("window blur cancels resize, restores authoritative geometry and leaves no history", async ({
  page
}) => {
  await gotoFixture(page);
  const before = await editorRect(page, 0, 1);

  await beginColumnResize(page, 30);
  expect(
    await page.evaluate(() =>
      (window as unknown as FixtureWindow).__resizeCanvasHasCapture?.("column")
    )
  ).toBe(true);
  await page.evaluate(() => window.dispatchEvent(new Event("blur")));
  expect(
    await page.evaluate(() =>
      (window as unknown as FixtureWindow).__resizeCanvasHasCapture?.("column")
    )
  ).toBe(false);
  await page.mouse.up();
  await nextFrame(page);

  expect(await editorRect(page, 0, 1)).toEqual(before);
  expect(await timeline(page)).toEqual([]);
  await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.undo());
  expect(await editorRect(page, 0, 1)).toEqual(before);
});

test("a structural patch interrupts resize preview before pointerup can commit it", async ({
  page
}) => {
  await gotoFixture(page);
  const before = await editorRect(page, 0, 1);

  await beginColumnResize(page, 30);
  expect(
    await page.evaluate(() =>
      (window as unknown as FixtureWindow).__resizeCanvasHasCapture?.("column")
    )
  ).toBe(true);
  expect(
    await page.evaluate(() => (window as unknown as FixtureWindow).__lastPaintedWidth?.())
  ).toBe(110);

  await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.insertRows(0));
  expect(
    await page.evaluate(() =>
      (window as unknown as FixtureWindow).__resizeCanvasHasCapture?.("column")
    )
  ).toBe(false);
  expect(
    await page.evaluate(() => (window as unknown as FixtureWindow).__lastPaintedWidth?.())
  ).toBe(80);
  await page.mouse.up();
  await nextFrame(page);

  expect(await editorRect(page, 0, 1)).toEqual(before);
  expect(await timeline(page)).toEqual([{ kind: "change", changeTypes: ["row-structure"] }]);
  await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.undo());
  expect(await editorRect(page, 0, 1)).toEqual(before);
});

test("structural undo cancels an editor before its numeric anchor becomes stale", async ({ page }) => {
  await gotoFixture(page);
  await page.evaluate(() => {
    const sheet = (window as unknown as FixtureWindow).__sheet;
    sheet?.insertRows(0);
    sheet?.startEdit(1, 1);
  });
  await page.locator("#app input").fill("stale");

  await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.undo());

  expect(await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.isEditing())).toBe(false);
  await expect(page.locator("#app input")).toHaveCount(0);
  expect(await page.evaluate(async () => {
    const sheet = (window as unknown as FixtureWindow).__sheet;
    await sheet?.commitEdit();
    return [sheet?.getCell(0, 1), sheet?.getCell(1, 1)];
  })).toEqual(["0:1", "1:1"]);
});

test("structural redo cancels an editor before replay shifts its cell", async ({ page }) => {
  await gotoFixture(page);
  await page.evaluate(() => {
    const sheet = (window as unknown as FixtureWindow).__sheet;
    sheet?.insertRows(0);
    sheet?.undo();
    sheet?.startEdit(0, 1);
  });
  await page.locator("#app input").fill("stale");

  await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.redo());

  expect(await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.isEditing())).toBe(false);
  await expect(page.locator("#app input")).toHaveCount(0);
  expect(await page.evaluate(async () => {
    const sheet = (window as unknown as FixtureWindow).__sheet;
    await sheet?.commitEdit();
    return [sheet?.getCell(0, 1), sheet?.getCell(1, 1)];
  })).toEqual([undefined, "0:1"]);
});

test("scroll cancels an active column resize instead of committing preview geometry", async ({ page }) => {
  await gotoFixture(page);
  const before = await editorRect(page, 0, 1);

  await beginColumnResize(page, 30);
  await page.locator("#app").evaluate((element) => {
    const scroller = element.firstElementChild as HTMLElement;
    scroller.scrollLeft = 20;
    scroller.dispatchEvent(new Event("scroll"));
  });
  await page.mouse.up();
  await nextFrame(page);

  expect(await editorRect(page, 0, 1)).toEqual({ ...before, left: before.left - 20 });
  expect(await timeline(page)).toEqual([]);
});

test("dragging away and back to the original width is a no-op with no history entry", async ({ page }) => {
  await gotoFixture(page);
  const before = await editorRect(page, 0, 1);
  const box = await canvasBox(page, "column");
  if (!box) throw new Error("missing column header canvas");

  await page.mouse.move(box.x + 80, box.y + 10);
  await page.mouse.down();
  await page.mouse.move(box.x + 110, box.y + 10);
  await page.mouse.move(box.x + 80, box.y + 10);
  await page.mouse.up();
  await nextFrame(page);

  expect(await editorRect(page, 0, 1)).toEqual(before);
  expect(await timeline(page)).toEqual([]);
  await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.undo());
  expect(await editorRect(page, 0, 1)).toEqual(before);
});

test("a pointercancel for another pointer does not terminate the active resize", async ({ page }) => {
  await gotoFixture(page);
  const before = await editorRect(page, 0, 1);

  await beginColumnResize(page, 30);
  await page.evaluate(() =>
    (window as unknown as FixtureWindow).__dispatchUnrelatedResizePointerCancel?.()
  );
  expect(
    await page.evaluate(() =>
      (window as unknown as FixtureWindow).__resizeCanvasHasCapture?.("column")
    )
  ).toBe(true);
  await page.mouse.up();
  await nextFrame(page);

  expect(await editorRect(page, 0, 1)).toEqual({ ...before, left: before.left + 30 });
  expect(await timeline(page)).toEqual([{ kind: "change", changeTypes: ["column-layout"] }]);
});

for (const testCase of [
  {
    kind: "column" as const,
    row: 0,
    column: 1,
    delta: 30,
    changeType: "column-layout",
    changedRect: (before: CellRect) => ({ ...before, left: before.left + 30 })
  },
  {
    kind: "row" as const,
    row: 1,
    column: 0,
    delta: 20,
    changeType: "row-layout",
    changedRect: (before: CellRect) => ({ ...before, top: before.top + 20 })
  }
]) {
  test(`${testCase.kind} resize commits once and undo/redo restore the expected geometry`, async ({
    page
  }) => {
    await gotoFixture(page);
    const before = await editorRect(page, testCase.row, testCase.column);

    if (testCase.kind === "column") {
      await beginColumnResize(page, testCase.delta);
    } else {
      await beginRowResize(page, testCase.delta);
    }
    await page.mouse.up();
    await nextFrame(page);

    const changed = testCase.changedRect(before);
    expect(await editorRect(page, testCase.row, testCase.column)).toEqual(changed);
    expect(await timeline(page)).toEqual([{ kind: "change", changeTypes: [testCase.changeType] }]);

    await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.undo());
    expect(await editorRect(page, testCase.row, testCase.column)).toEqual(before);
    await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.redo());
    expect(await editorRect(page, testCase.row, testCase.column)).toEqual(changed);
  });
}

test("window pointer listeners complete resize when pointer capture throws", async ({ page }) => {
  await gotoFixture(page);
  const before = await editorRect(page, 0, 1);
  await page.evaluate(() =>
    (window as unknown as FixtureWindow).__throwOnResizePointerCapture?.("column")
  );

  await beginColumnResize(page, 30);
  expect(
    await page.evaluate(() =>
      (window as unknown as FixtureWindow).__resizeCanvasHasCapture?.("column")
    )
  ).toBe(false);
  await page.mouse.up();
  await nextFrame(page);

  expect(await editorRect(page, 0, 1)).toEqual({ ...before, left: before.left + 30 });
  expect(await timeline(page)).toEqual([
    {
      kind: "error",
      eventName: "interaction:header-pointer-capture",
      message: "forced pointer capture failure"
    },
    { kind: "change", changeTypes: ["column-layout"] }
  ]);
});

for (const testCase of [
  {
    kind: "column" as const,
    command: "resize-column" as const,
    errorEventName: "interaction:column-resize",
    row: 0,
    column: 1
  },
  {
    kind: "row" as const,
    command: "resize-row" as const,
    errorEventName: "interaction:row-resize",
    row: 1,
    column: 0
  }
]) {
  test(`a rejected ${testCase.command} restores authoritative geometry and reports one interaction error`, async ({
    page
  }) => {
    await gotoFixture(page);
    const before = await editorRect(page, testCase.row, testCase.column);
    await page.evaluate(
      (command) => (window as unknown as FixtureWindow).__rejectResizeCommand?.(command),
      testCase.command
    );

    if (testCase.kind === "column") {
      await beginColumnResize(page, 30);
    } else {
      await beginRowResize(page, 20);
    }
    await page.mouse.up();
    await nextFrame(page);

    expect(await editorRect(page, testCase.row, testCase.column)).toEqual(before);
    expect(await timeline(page)).toEqual([
      {
        kind: "error",
        eventName: testCase.errorEventName,
        message: `rejected ${testCase.command}`
      }
    ]);
    await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.undo());
    expect(await editorRect(page, testCase.row, testCase.column)).toEqual(before);
  });
}

for (const testCase of [
  { kind: "column" as const, command: "resize-column" as const, row: 0, column: 1 },
  { kind: "row" as const, command: "resize-row" as const, row: 1, column: 0 }
]) {
  test(`a rejected ${testCase.command} transaction restores authoritative geometry without history`, async ({
    page
  }) => {
    await gotoFixture(page);
    const before = await editorRect(page, testCase.row, testCase.column);
    await page.evaluate(
      (command) => (window as unknown as FixtureWindow).__rejectResizeTransaction?.(command),
      testCase.command
    );

    if (testCase.kind === "column") {
      await beginColumnResize(page, 30);
    } else {
      await beginRowResize(page, 20);
    }
    await page.mouse.up();
    await nextFrame(page);

    expect(await editorRect(page, testCase.row, testCase.column)).toEqual(before);
    expect(await timeline(page)).toEqual([
      {
        kind: "error",
        eventName: `interaction:${testCase.kind}-resize`,
        message: `rejected transaction ${testCase.command}`
      }
    ]);
    await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.undo());
    expect(await editorRect(page, testCase.row, testCase.column)).toEqual(before);
  });
}

test("dispose during resize releases capture and removes all gesture surfaces", async ({ page }) => {
  await gotoFixture(page);
  const pageErrors: string[] = [];
  page.on("pageerror", (error) => pageErrors.push(String(error)));

  await beginColumnResize(page, 30);
  expect(
    await page.evaluate(() =>
      (window as unknown as FixtureWindow).__resizeCanvasHasCapture?.("column")
    )
  ).toBe(true);
  const captureAfterDispose = await page.evaluate(() => {
    const fixture = window as unknown as FixtureWindow;
    fixture.__sheet?.dispose();
    return fixture.__resizeCanvasHasCapture?.("column");
  });
  await page.mouse.up();

  expect(captureAfterDispose).toBe(false);
  await expect(page.locator("#app canvas")).toHaveCount(0);
  expect(pageErrors).toEqual([]);
});
