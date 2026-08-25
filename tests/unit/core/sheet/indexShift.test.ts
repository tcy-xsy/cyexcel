import { describe, expect, it } from "vitest";

import { ROW_PLAN_HOLE, remapByColumnShift, remapByRowPlan } from "../../../../src/core/sheet";

const H = ROW_PLAN_HOLE;

describe("remapByRowPlan", () => {
  it("moves entries to their new index", () => {
    const source = { 0: "a", 2: "c" };
    expect(remapByRowPlan(source, Int32Array.from([2, 1, 0]))).toStrictEqual({ 0: "c", 2: "a" });
  });

  it("leaves a hole row with no entry", () => {
    expect(remapByRowPlan({ 0: "a" }, Int32Array.from([H, 0]))).toStrictEqual({ 1: "a" });
  });

  it("drops entries for rows the plan does not name", () => {
    expect(remapByRowPlan({ 0: "a", 1: "b" }, Int32Array.from([1]))).toStrictEqual({ 0: "b" });
  });

  it("returns an empty map for an empty plan", () => {
    expect(remapByRowPlan({ 0: "a" }, Int32Array.from([]))).toStrictEqual({});
  });

  it("does not carry an entry for a source index that had none", () => {
    // 第 1 行没有元数据，搬完之后新的第 0 行也不该凭空有一条
    const result = remapByRowPlan({ 0: "a" }, Int32Array.from([1, 0]));
    expect(result).toStrictEqual({ 1: "a" });
    expect(0 in result).toBe(false);
  });
});

describe("remapByColumnShift", () => {
  it("shifts entries at and after an insertion point", () => {
    const shifted = remapByColumnShift({ 0: "a", 1: "b" }, (i) => (i < 1 ? i : i + 2));
    expect(shifted).toStrictEqual({ 0: "a", 3: "b" });
  });

  it("drops entries the remap sends nowhere", () => {
    const shifted = remapByColumnShift({ 0: "a", 1: "gone", 2: "c" }, (i) => {
      if (i < 1) return i;
      if (i < 2) return undefined;
      return i - 1;
    });
    expect(shifted).toStrictEqual({ 0: "a", 1: "c" });
  });

  it("returns an empty map when everything is dropped", () => {
    expect(remapByColumnShift({ 0: "a" }, () => undefined)).toStrictEqual({});
  });
});
