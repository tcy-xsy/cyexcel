import { describe, expect, it } from "vitest";

import {
  assertValidCoordinate,
  assertValidRange
} from "../../../src/core/validation";
import { CyExcelError, CyExcelErrorCode } from "../../../src/core/errors";

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
