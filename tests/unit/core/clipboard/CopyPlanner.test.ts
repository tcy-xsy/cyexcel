import { describe, expect, it } from "vitest";

import { CopyPlanner } from "../../../../src/core";
import type { RuntimeColumn } from "../../../../src/core";

function columnWithPolicy(copy = true): Pick<RuntimeColumn, "clipboard"> {
  return { clipboard: { copy, paste: true } };
}

describe("CopyPlanner", () => {
  it("reports rows/columns/cells for a plain rectangular range", () => {
    const planner = new CopyPlanner({
      getCellValue: () => "x",
      getColumn: () => columnWithPolicy(),
      maxCells: 1_000_000,
      maxBytes: 1_000_000
    });

    const plan = planner.plan({ startRow: 0, startColumn: 0, endRow: 2, endColumn: 3 });

    expect(plan.rows).toBe(3);
    expect(plan.columns).toBe(4);
    expect(plan.cells).toBe(12);
    expect(plan.containsComplexValues).toBe(false);
    expect(plan.recommendedMode).toBe("display");
  });

  it("flags containsComplexValues when any visited cell is an array or plain object", () => {
    const planner = new CopyPlanner({
      getCellValue: (row, column) => (row === 0 && column === 1 ? ["a", "b"] : "plain"),
      getColumn: () => columnWithPolicy(),
      maxCells: 1_000_000,
      maxBytes: 1_000_000
    });

    const plan = planner.plan({ startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 });

    expect(plan.containsComplexValues).toBe(true);
    expect(plan.recommendedMode).toBe("raw");
  });

  it("does not treat a Date value as a complex value", () => {
    const planner = new CopyPlanner({
      getCellValue: () => new Date(),
      getColumn: () => columnWithPolicy(),
      maxCells: 1_000_000,
      maxBytes: 1_000_000
    });

    const plan = planner.plan({ startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 });

    expect(plan.containsComplexValues).toBe(false);
  });

  it("recommends download once cell count exceeds maxCells", () => {
    const planner = new CopyPlanner({
      getCellValue: () => "x",
      getColumn: () => columnWithPolicy(),
      maxCells: 5,
      maxBytes: 1_000_000
    });

    const plan = planner.plan({ startRow: 0, startColumn: 0, endRow: 0, endColumn: 9 });

    expect(plan.cells).toBe(10);
    expect(plan.recommendedMode).toBe("download");
  });

  it("recommends download once estimated bytes exceeds maxBytes", () => {
    const planner = new CopyPlanner({
      getCellValue: () => "x".repeat(100),
      getColumn: () => columnWithPolicy(),
      maxCells: 1_000_000,
      maxBytes: 50
    });

    const plan = planner.plan({ startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 });

    expect(plan.recommendedMode).toBe("download");
  });

  it("skips columns whose clipboard.copy policy is false — never reads their cell values", () => {
    const reads: number[] = [];
    const planner = new CopyPlanner({
      getCellValue: (_row, column) => {
        reads.push(column);
        return "value";
      },
      getColumn: (column) => columnWithPolicy(column !== 1),
      maxCells: 1_000_000,
      maxBytes: 1_000_000
    });

    const plan = planner.plan({ startRow: 0, startColumn: 0, endRow: 0, endColumn: 2 });

    expect(reads).toEqual([0, 2]);
    expect(plan.estimatedBytes).toBe("value".length * 2);
  });

  it("plans a single cell containing a 100k-item array without a long task", () => {
    const hugeArray = Array.from({ length: 100_000 }, (_, i) => i);
    const planner = new CopyPlanner({
      getCellValue: () => hugeArray,
      getColumn: () => columnWithPolicy(),
      maxCells: 1_000_000,
      maxBytes: 1_000_000_000
    });

    const start = performance.now();
    const plan = planner.plan({ startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 });
    const elapsed = performance.now() - start;

    expect(elapsed).toBeLessThan(50);
    expect(plan.containsComplexValues).toBe(true);
  });
});
