import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";

interface FixtureWindow {
  __cyexcelFinalizerIsolationReady?: boolean;
  __timeline?: Array<Record<string, unknown>>;
  __armFailure?(kind: "cells" | "headers"): void;
  __throwFormat?(): void;
  __recordFormatSceneVersions?(): void;
  __formatSceneVersions?: number[];
  __armEditorGeometryFailure?(): void;
  __reenterFromCell?(): void;
  __mutateFromCanvasPaint?(): void;
  __sheet?: {
    readonly scene: { readonly modelVersion: number };
    getCell(row: number, column: number): unknown;
    getHeader(column: number): { readonly title: string };
    setCell(row: number, column: number, value: unknown): void;
    insertColumns(start: number, columns: readonly { readonly header: string; readonly width: number }[]): void;
    startEdit(row: number, column: number): void;
    undo(): void;
  };
}

async function gotoFixture(page: Page): Promise<void> {
  await page.goto("/tests/browser/finalizer-isolation.html");
  await page.waitForFunction(
    () => (window as unknown as FixtureWindow).__cyexcelFinalizerIsolationReady === true
  );
  await page.waitForFunction(
    () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)))
  );
}

test("a DOM cell finalizer failure cannot roll back or starve a column transaction", async ({ page }) => {
  await gotoFixture(page);
  const pageErrors: unknown[] = [];
  page.on("pageerror", (error) => pageErrors.push(error));

  const result = await page.evaluate(() => {
    const fixture = window as unknown as FixtureWindow;
    fixture.__armFailure?.("cells");
    fixture.__sheet?.insertColumns(0, [{ header: "Inserted", width: 40 }]);
    return {
      firstColumn: fixture.__sheet?.getHeader(0).title,
      shiftedCell: fixture.__sheet?.getCell(0, 1),
      timeline: fixture.__timeline
    };
  });

  expect(result.firstColumn).toBe("Inserted");
  expect(result.shiftedCell).toBe("outer");
  expect(result.timeline).toEqual([
    { kind: "error", eventName: "finalize:dom-cells", message: "cell update failed" },
    { kind: "change", modelVersion: 1, changeTypes: ["column-structure"] }
  ]);
  await expect(page.locator(".dom-header")).toHaveText("DOM");
  expect(pageErrors).toEqual([]);
});

test("a DOM header finalizer failure still publishes change and keeps history", async ({ page }) => {
  await gotoFixture(page);

  const result = await page.evaluate(() => {
    const fixture = window as unknown as FixtureWindow;
    fixture.__armFailure?.("headers");
    fixture.__sheet?.insertColumns(0, [{ header: "Inserted", width: 40 }]);
    const afterInsert = fixture.__sheet?.getHeader(0).title;
    fixture.__sheet?.undo();
    return { afterInsert, afterUndo: fixture.__sheet?.getHeader(0).title, timeline: fixture.__timeline };
  });

  expect(result).toMatchObject({ afterInsert: "Inserted", afterUndo: "DOM" });
  expect(result.timeline?.filter((entry) => entry.kind === "error")).toEqual([
    { kind: "error", eventName: "finalize:dom-headers", message: "header update failed" }
  ]);
  expect(result.timeline?.filter((entry) => entry.kind === "change")).toHaveLength(2);
  await expect(page.locator(".dom-cell")).toHaveText("outer");
});

test("a throwing wrap formatter reports an error after the data commit and preserves undo", async ({ page }) => {
  await gotoFixture(page);

  const result = await page.evaluate(() => {
    const fixture = window as unknown as FixtureWindow;
    fixture.__throwFormat?.();
    fixture.__sheet?.setCell(0, 1, "committed");
    const committed = fixture.__sheet?.getCell(0, 1);
    fixture.__sheet?.undo();
    return { committed, undone: fixture.__sheet?.getCell(0, 1), timeline: fixture.__timeline };
  });

  expect(result.committed).toBe("committed");
  expect(result.undone).toBe("short");
  expect(result.timeline).toEqual([
    { kind: "error", eventName: "finalize:wrap-measurement", message: "format failed" },
    { kind: "error", eventName: "finalize:canvas-cell:0:1", message: "format failed" },
    { kind: "change", modelVersion: 1, changeTypes: ["cell-value"] },
    { kind: "error", eventName: "finalize:wrap-measurement", message: "format failed" },
    { kind: "error", eventName: "finalize:canvas-cell:0:1", message: "format failed" },
    { kind: "change", modelVersion: 2, changeTypes: ["cell-value"] }
  ]);
});

test("wrap measurement observes the Scene published for its cell receipt", async ({ page }) => {
  await gotoFixture(page);

  const result = await page.evaluate(() => {
    const fixture = window as unknown as FixtureWindow;
    fixture.__recordFormatSceneVersions?.();
    fixture.__sheet?.setCell(
      0,
      1,
      "a long wrapped value whose measurement must read the newly published receipt Scene"
    );
    return {
      finalModelVersion: fixture.__sheet?.scene.modelVersion,
      formatSceneVersions: fixture.__formatSceneVersions
    };
  });

  expect(result.finalModelVersion).toBe(2);
  expect(result.formatSceneVersions?.[0]).toBe(1);
  expect(result.formatSceneVersions?.every((version) => version >= 1)).toBe(true);
});

test("an editor geometry failure cannot suppress DOM presentation", async ({ page }) => {
  await gotoFixture(page);
  const pageErrors: unknown[] = [];
  page.on("pageerror", (error) => pageErrors.push(error));

  const result = await page.evaluate(() => {
    const fixture = window as unknown as FixtureWindow;
    const sheet = fixture.__sheet;
    if (!sheet) throw new Error("missing sheet");
    sheet.startEdit(0, 1);
    fixture.__armEditorGeometryFailure?.();
    sheet.setCell(0, 0, "published-despite-editor-error");
    return {
      domText: document.querySelector(".dom-cell")?.textContent,
      editorMounted: document.querySelector(".throwing-editor") !== null,
      modelVersion: sheet.scene.modelVersion,
      timeline: fixture.__timeline
    };
  });

  expect(result).toMatchObject({
    domText: "published-despite-editor-error",
    editorMounted: true,
    modelVersion: 1
  });
  expect(result.timeline).toEqual([
    { kind: "error", eventName: "finalize:editor-geometry", message: "editor geometry failed" },
    { kind: "change", modelVersion: 1, changeTypes: ["cell-value"] }
  ]);
  expect(pageErrors).toEqual([]);
});

test("auto-height publishes its own row-layout receipt after the user cell receipt", async ({ page }) => {
  await gotoFixture(page);

  const timeline = await page.evaluate(() => {
    const fixture = window as unknown as FixtureWindow;
    fixture.__sheet?.setCell(0, 1, "a long wrapped value that must grow this narrow row substantially");
    return fixture.__timeline;
  });

  expect(timeline?.filter((entry) => entry.kind === "change")).toEqual([
    {
      kind: "change",
      modelVersion: 1,
      changeTypes: ["cell-value"]
    },
    {
      kind: "change",
      modelVersion: 2,
      changeTypes: ["row-layout"]
    }
  ]);
  expect(timeline?.filter((entry) => entry.kind === "layout-change")).toEqual([
    {
      kind: "layout-change",
      layoutVersion: 1,
      changeTypes: ["row-layout"]
    }
  ]);
  expect(timeline?.map((entry) => entry.kind)).toEqual([
    "change",
    "change",
    "layout-change"
  ]);
});

test("a committed change synchronously publishes DOM and canvas presentation before returning", async ({ page }) => {
  await gotoFixture(page);

  const result = await page.evaluate(() => {
    const fixture = window as unknown as FixtureWindow;
    fixture.__sheet?.setCell(0, 0, "published-now");
    return {
      domText: document.querySelector(".dom-cell")?.textContent,
      timeline: fixture.__timeline
    };
  });

  expect(result.domText).toBe("published-now");
  expect(result.timeline).toEqual([
    { kind: "change", modelVersion: 1, changeTypes: ["cell-value"] }
  ]);
});

test("a renderer mutation intent is replayed FIFO after the outer change", async ({ page }) => {
  await gotoFixture(page);

  const result = await page.evaluate(() => {
    const fixture = window as unknown as FixtureWindow;
    fixture.__reenterFromCell?.();
    fixture.__sheet?.setCell(0, 0, "outer-2");
    return {
      outer: fixture.__sheet?.getCell(0, 0),
      inner: fixture.__sheet?.getCell(0, 1),
      timeline: fixture.__timeline
    };
  });

  expect(result).toMatchObject({ outer: "outer-2", inner: "inner" });
  expect(result.timeline).toEqual([
    { kind: "change", modelVersion: 1, changeTypes: ["cell-value"] },
    { kind: "change", modelVersion: 2, changeTypes: ["cell-value"] }
  ]);
});

test("an asynchronous RAF renderer intent commits only after its read lease releases", async ({ page }) => {
  await gotoFixture(page);

  const result = await page.evaluate(async () => {
    const fixture = window as unknown as FixtureWindow;
    fixture.__mutateFromCanvasPaint?.();
    const scroller = document.querySelector("#app > div") as HTMLElement;
    scroller.dispatchEvent(new Event("scroll"));
    await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
    return {
      value: fixture.__sheet?.getCell(0, 1),
      timeline: fixture.__timeline
    };
  });

  expect(result.value).toBe("async-renderer-intent");
  expect(result.timeline).toEqual([
    {
      kind: "change",
      modelVersion: 1,
      changeTypes: ["cell-value"]
    },
    {
      kind: "change",
      modelVersion: 2,
      changeTypes: ["row-layout"]
    },
    {
      kind: "layout-change",
      layoutVersion: 1,
      changeTypes: ["row-layout"]
    }
  ]);
});
