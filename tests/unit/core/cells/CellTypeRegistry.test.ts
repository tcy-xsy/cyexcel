import { describe, expect, it } from "vitest";

import { CellTypeRegistry } from "../../../../src/core/cells";
import { CyExcelError, CyExcelErrorCode } from "../../../../src/core/errors";

describe("CellTypeRegistry.resolve — defaults", () => {
  it("falls back to a base format() when nothing registered", () => {
    const registry = new CellTypeRegistry();
    registry.register("bare", {});

    const type = registry.resolve("bare");

    expect(type.format(undefined, { row: 0, column: {} as never })).toBe("");
    expect(type.format("x", { row: 0, column: {} as never })).toBe("x");
  });
});

describe("CellTypeRegistry.resolve — extends chain", () => {
  it("merges base defaults, the parent type, then the type's own overrides, in that order", () => {
    const registry = new CellTypeRegistry();
    registry.register("number", {
      format: (value) => `#${String(value)}`,
      compare: (a, b) => (a as number) - (b as number)
    });
    registry.register("money", {
      extends: "number",
      format: (value) => `¥${String(value)}`
    });

    const money = registry.resolve("money");

    // Own override wins over the parent's format...
    expect(money.format(10, { row: 0, column: {} as never })).toBe("¥10");
    // ...but fields the child never touched are still inherited from the parent.
    expect(money.compare?.(1, 2)).toBe(-1);
  });

  it("supports multi-level extends chains", () => {
    const registry = new CellTypeRegistry();
    registry.register("number", { format: (value) => `n:${String(value)}` });
    registry.register("money", { extends: "number" });
    registry.register("usd", { extends: "money", format: (value) => `$${String(value)}` });

    expect(registry.resolve("usd").format(5, { row: 0, column: {} as never })).toBe("$5");
    // "money" itself never overrode format, so it still uses "number"'s.
    expect(registry.resolve("money").format(5, { row: 0, column: {} as never })).toBe("n:5");
  });
});

describe("CellTypeRegistry.resolve — caching", () => {
  it("returns the same resolved object on repeated calls instead of re-merging", () => {
    const registry = new CellTypeRegistry();
    registry.register("money", { extends: "number" });
    registry.register("number", { format: (value) => String(value) });

    const first = registry.resolve("money");
    const second = registry.resolve("money");

    expect(first).toBe(second);
  });
});

describe("CellTypeRegistry.resolve — unknown type", () => {
  it("throws a stable error code for an unregistered type name", () => {
    const registry = new CellTypeRegistry();

    try {
      registry.resolve("does-not-exist");
      expect.unreachable("expected resolve() to throw");
    } catch (error) {
      expect(error).toBeInstanceOf(CyExcelError);
      expect((error as CyExcelError).code).toBe(CyExcelErrorCode.UnknownCellType);
    }
  });

  it("throws the same error code when extends references an unregistered type", () => {
    const registry = new CellTypeRegistry();
    registry.register("money", { extends: "does-not-exist" });

    try {
      registry.resolve("money");
      expect.unreachable("expected resolve() to throw");
    } catch (error) {
      expect(error).toBeInstanceOf(CyExcelError);
      expect((error as CyExcelError).code).toBe(CyExcelErrorCode.UnknownCellType);
    }
  });
});
