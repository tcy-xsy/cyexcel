import { describe, expect, it } from "vitest";
import type * as CoreModule from "../../src/core";
import type * as IndexModule from "../../src/index";
import type * as RegisterModule from "../../src/register";

/**
 * `dist/*.js` has no co-located `.d.ts` (its declarations live under the
 * separate `dist/types/` tree), so a statically-analyzable
 * `import("../../dist/core.js")` trips TS7016 ("implicitly has an any
 * type") even behind an `as` cast — TS resolves literal specifiers before
 * the cast applies. Routing the specifier through a plain variable makes it
 * non-literal, so TS treats the whole expression as `any` (no diagnostic),
 * and the `TModule` type parameter borrows real types from the source
 * program (which tsconfig.json's main `include` already covers) for
 * whatever this dynamically-imported build artifact turns out to be.
 */
function importDist<TModule>(path: string): Promise<TModule> {
  return import(path) as Promise<TModule>;
}

describe("SSR safety — importing cyexcel/core touches no DOM", () => {
  it("runs in an environment with no document/window (sanity check on the test environment itself)", () => {
    expect(typeof document).toBe("undefined");
    expect(typeof window).toBe("undefined");
  });

  it("imports the built core entry without throwing", async () => {
    await expect(importDist("../../dist/core.js")).resolves.toBeDefined();
  });

  it("exercises the public importer and storage model with no DOM access", async () => {
    const { rowsToWorkbookSnapshot, WorkbookModel } =
      await importDist<typeof CoreModule>("../../dist/core.js");
    const sheet = WorkbookModel
      .fromSnapshot(rowsToWorkbookSnapshot([{ name: "a" }, { name: "b" }], ["name"]))
      .sheet();

    expect(sheet.rowCount()).toBe(2);
    expect(sheet.columnCount()).toBe(1);
    expect(sheet.getCell(0, 0)).toBe("a");
    expect(sheet.getCell(1, 0)).toBe("b");
  });
});

describe("SSR safety — importing the bare \"cyexcel\" aggregate entry touches no DOM", () => {
  it("imports without throwing and does not auto-register any custom element", async () => {
    const mod = await importDist<typeof IndexModule>("../../dist/index.js");
    expect(typeof mod.CyExcelCore).toBe("function");
    expect(typeof customElements).toBe("undefined");
  });
});

describe("cyexcel/register in a DOM-less environment", () => {
  it("no-ops instead of throwing when customElements is unavailable", async () => {
    await expect(importDist<typeof RegisterModule>("../../dist/register.js")).resolves.toBeDefined();
  });
});
