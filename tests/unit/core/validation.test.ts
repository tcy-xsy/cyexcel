import { describe, expect, it } from "vitest";

import {
  assertValidCoordinate,
  assertValidCyExcelOptions,
  assertValidRange,
  CyExcelError,
  CyExcelErrorCode
} from "../../../src/core";

function expectSync(result: unknown): void {
  expect(result).toBeUndefined();
  expect(typeof (result as { then?: unknown })?.then).not.toBe("function");
}

describe("assertValidCoordinate", () => {
  it("accepts non-negative integer coordinates", () => {
    expectSync(assertValidCoordinate(0, 0));
    expectSync(assertValidCoordinate(999_999, 49));
  });

  it.each([
    ["negative row", -1, 0],
    ["negative column", 0, -1],
    ["non-integer row", 1.5, 0],
    ["non-integer column", 0, 1.5],
    ["NaN row", Number.NaN, 0],
    ["NaN column", 0, Number.NaN]
  ])("rejects %s with a stable error code", (_label, row, column) => {
    try {
      assertValidCoordinate(row, column);
      expect.unreachable("expected assertValidCoordinate to throw");
    } catch (error) {
      expect(error).toBeInstanceOf(CyExcelError);
      expect((error as CyExcelError).code).toBe(
        CyExcelErrorCode.InvalidCoordinate
      );
    }
  });
});

describe("assertValidRange", () => {
  it("accepts a normalized non-negative integer range", () => {
    expectSync(
      assertValidRange({ startRow: 0, startColumn: 0, endRow: 10, endColumn: 5 })
    );
  });

  it("rejects a range whose start is after its end", () => {
    try {
      assertValidRange({ startRow: 10, startColumn: 0, endRow: 0, endColumn: 0 });
      expect.unreachable("expected assertValidRange to throw");
    } catch (error) {
      expect(error).toBeInstanceOf(CyExcelError);
      expect((error as CyExcelError).code).toBe(CyExcelErrorCode.InvalidRange);
    }
  });

  it("rejects a range with a negative bound", () => {
    try {
      assertValidRange({ startRow: 0, startColumn: -1, endRow: 0, endColumn: 0 });
      expect.unreachable("expected assertValidRange to throw");
    } catch (error) {
      expect(error).toBeInstanceOf(CyExcelError);
      expect((error as CyExcelError).code).toBe(CyExcelErrorCode.InvalidRange);
    }
  });
});

describe("assertValidCyExcelOptions", () => {
  it("accepts options with no constrained fields set", () => {
    expectSync(assertValidCyExcelOptions({}));
  });

  it("accepts valid rowHeight, overscan, and freeze values", () => {
    expectSync(
      assertValidCyExcelOptions({
        rowHeight: 32,
        headerHeight: 40,
        overscanRows: 5,
        overscanColumns: 2,
        freeze: { rows: 1, columns: 1 }
      })
    );
  });

  it.each([
    ["negative rowHeight", { rowHeight: -1 }],
    ["non-finite headerHeight", { headerHeight: Number.POSITIVE_INFINITY }],
    ["non-integer overscanRows", { overscanRows: 1.5 }],
    ["negative overscanColumns", { overscanColumns: -1 }],
    ["negative freeze.rows", { freeze: { rows: -1 } }],
    ["negative freeze.columns", { freeze: { columns: -1 } }]
  ])("rejects %s with a stable error code", (_label, options) => {
    try {
      assertValidCyExcelOptions(options);
      expect.unreachable("expected assertValidCyExcelOptions to throw");
    } catch (error) {
      expect(error).toBeInstanceOf(CyExcelError);
      expect((error as CyExcelError).code).toBe(CyExcelErrorCode.InvalidOptions);
    }
  });
});
