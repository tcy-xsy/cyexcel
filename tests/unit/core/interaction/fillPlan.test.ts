import { describe, expect, it } from "vitest";

import {
  computeAutoFillDownTargetRow,
  computeFillPlan,
  type FillPlanSources
} from "../../../../src/core/interaction";
import type { CellRange } from "../../../../src/core/types";

/** A grid addressed as `${row}:${column}`; missing entries read as undefined. */
function gridSources(
  cells: Readonly<Record<string, unknown>>,
  pasteBlockedColumns: readonly number[] = []
): FillPlanSources {
  return {
    getCell: (row, column) => cells[`${row}:${column}`],
    isPasteBlocked: (column) => pasteBlockedColumns.includes(column)
  };
}

const range = (
  startRow: number,
  startColumn: number,
  endRow: number,
  endColumn: number
): CellRange => ({ startRow, startColumn, endRow, endColumn });

describe("computeFillPlan downward", () => {
  it("copies a single source value down the extension", () => {
    const values = computeFillPlan(
      range(0, 0, 0, 0),
      range(0, 0, 3, 0),
      gridSources({ "0:0": "x" })
    );
    expect(values).toEqual([["x"], ["x"], ["x"], ["x"]]);
  });

  it("extrapolates a numeric run by its linear step", () => {
    const values = computeFillPlan(
      range(0, 0, 1, 0),
      range(0, 0, 4, 0),
      gridSources({ "0:0": 2, "1:0": 4 })
    );
    expect(values).toEqual([[2], [4], [6], [8], [10]]);
  });

  it("repeats a non-numeric sequence cyclically", () => {
    const values = computeFillPlan(
      range(0, 0, 1, 0),
      range(0, 0, 4, 0),
      gridSources({ "0:0": "a", "1:0": "b" })
    );
    expect(values).toEqual([["a"], ["b"], ["a"], ["b"], ["a"]]);
  });

  it("always copies the source region verbatim, never extrapolated", () => {
    const values = computeFillPlan(
      range(1, 0, 2, 0),
      range(1, 0, 3, 0),
      gridSources({ "1:0": 10, "2:0": 20 })
    );
    expect(values[0]).toEqual([10]);
    expect(values[1]).toEqual([20]);
  });

  it("keeps existing values in a paste-blocked column instead of filling it", () => {
    const values = computeFillPlan(
      range(0, 0, 0, 1),
      range(0, 0, 2, 1),
      gridSources(
        { "0:0": 1, "0:1": "keep", "1:1": "old-1", "2:1": "old-2" },
        [1]
      )
    );
    expect(values.map((row) => row[0])).toEqual([1, 1, 1]);
    expect(values.map((row) => row[1])).toEqual(["keep", "old-1", "old-2"]);
  });
});

describe("computeFillPlan rightward", () => {
  it("extrapolates across columns", () => {
    const values = computeFillPlan(
      range(0, 0, 0, 1),
      range(0, 0, 0, 3),
      gridSources({ "0:0": 5, "0:1": 10 })
    );
    expect(values).toEqual([[5, 10, 15, 20]]);
  });

  it("keeps existing values in a paste-blocked target column", () => {
    const values = computeFillPlan(
      range(0, 0, 0, 0),
      range(0, 0, 0, 2),
      gridSources({ "0:0": "src", "0:1": "old-1", "0:2": "old-2" }, [2])
    );
    expect(values).toEqual([["src", "src", "old-2"]]);
  });
});

describe("computeFillPlan degenerate cases", () => {
  it("returns just the copied source when the target adds nothing", () => {
    const values = computeFillPlan(
      range(0, 0, 1, 1),
      range(0, 0, 1, 1),
      gridSources({ "0:0": 1, "0:1": 2, "1:0": 3, "1:1": 4 })
    );
    expect(values).toEqual([[1, 2], [3, 4]]);
  });

  it("fills unknown source cells with undefined rather than throwing", () => {
    const values = computeFillPlan(range(0, 0, 0, 0), range(0, 0, 1, 0), gridSources({}));
    expect(values).toEqual([[undefined], [undefined]]);
  });
});

describe("computeAutoFillDownTargetRow", () => {
  const bounds = { columnCount: 3, rowCount: 10 };

  it("follows the left neighbor column's run", () => {
    const cells = { "0:0": "a", "1:0": "b", "2:0": "c" };
    expect(
      computeAutoFillDownTargetRow(range(0, 1, 0, 1), bounds, (row, column) =>
        (cells as Record<string, unknown>)[`${row}:${column}`]
      )
    ).toBe(2);
  });

  it("uses the right neighbor when the selection starts at column 0", () => {
    const cells = { "0:1": "a", "1:1": "b" };
    expect(
      computeAutoFillDownTargetRow(range(0, 0, 0, 0), bounds, (row, column) =>
        (cells as Record<string, unknown>)[`${row}:${column}`]
      )
    ).toBe(1);
  });

  it("stops at the first empty neighbor cell", () => {
    const cells = { "0:0": "a", "1:0": "b", "2:0": "", "3:0": "d" };
    expect(
      computeAutoFillDownTargetRow(range(0, 1, 0, 1), bounds, (row, column) =>
        (cells as Record<string, unknown>)[`${row}:${column}`]
      )
    ).toBe(1);
  });

  it("returns null when the run does not reach past the selection", () => {
    const cells = { "0:0": "a" };
    expect(
      computeAutoFillDownTargetRow(range(0, 1, 0, 1), bounds, (row, column) =>
        (cells as Record<string, unknown>)[`${row}:${column}`]
      )
    ).toBe(null);
  });

  it("returns null when there is no neighbor column at all", () => {
    expect(
      computeAutoFillDownTargetRow(range(0, 0, 0, 0), { columnCount: 1, rowCount: 10 }, () => "x")
    ).toBe(null);
  });

  it("never scans past the row count", () => {
    expect(
      computeAutoFillDownTargetRow(range(0, 1, 0, 1), { columnCount: 3, rowCount: 4 }, () => "x")
    ).toBe(3);
  });
});
