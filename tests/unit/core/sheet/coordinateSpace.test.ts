import { describe, expect, it } from "vitest";

import { assertValidBlankCount, resolveCoordinateSpace } from "../../../../src/core/sheet";
import { CyExcelError, CyExcelErrorCode } from "../../../../src/core/errors";

describe("resolveCoordinateSpace", () => {
  it("is blank count alone when nothing is occupied", () => {
    // 空表：没有占用行，坐标空间就是空白区本身
    expect(resolveCoordinateSpace(undefined, 1000)).toBe(1000);
  });

  it("is maxOccupied + 1 + blank when something is occupied", () => {
    // 占用到第 19 行（0 基）=> 20 行数据 + 1000 空白
    expect(resolveCoordinateSpace(19, 1000)).toBe(1020);
  });

  it("still spans the data when blank is zero", () => {
    expect(resolveCoordinateSpace(19, 0)).toBe(20);
  });

  it("counts a single occupied cell at row 0 as one row", () => {
    expect(resolveCoordinateSpace(0, 0)).toBe(1);
  });

  it("handles a high sparse index without materialising anything", () => {
    expect(resolveCoordinateSpace(1_000_000, 1000)).toBe(1_001_001);
  });
});

describe("assertValidBlankCount", () => {
  it("accepts zero", () => {
    expect(() => assertValidBlankCount(0, "blankRows")).not.toThrow();
  });

  it("rejects a negative count", () => {
    expect(() => assertValidBlankCount(-1, "blankRows")).toThrow(CyExcelError);
    try {
      assertValidBlankCount(-1, "blankRows");
    } catch (error) {
      expect((error as CyExcelError).code).toBe(CyExcelErrorCode.InvalidOptions);
    }
  });

  it("rejects a non-integer count", () => {
    expect(() => assertValidBlankCount(1.5, "blankRows")).toThrow(CyExcelError);
  });

  it("rejects NaN and Infinity", () => {
    expect(() => assertValidBlankCount(Number.NaN, "blankRows")).toThrow(CyExcelError);
    expect(() => assertValidBlankCount(Number.POSITIVE_INFINITY, "blankRows")).toThrow(CyExcelError);
  });

  it("names the offending option in the message", () => {
    try {
      assertValidBlankCount(-1, "blankColumns");
    } catch (error) {
      expect((error as CyExcelError).message).toContain("blankColumns");
    }
  });
});
