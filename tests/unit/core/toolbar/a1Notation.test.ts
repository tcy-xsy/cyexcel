import { describe, expect, it } from "vitest";
import { parseA1, rangeToA1, toA1 } from "../../../../src/core/toolbar/a1Notation";

describe("toA1", () => {
  it("下标是 0 基，A1 记法是 1 基", () => {
    expect(toA1(0, 0)).toBe("A1");
    expect(toA1(6, 1)).toBe("B7");
  });

  it("超过 Z 之后进位成两个字母", () => {
    expect(toA1(0, 25)).toBe("Z1");
    expect(toA1(0, 26)).toBe("AA1");
  });
});

describe("rangeToA1", () => {
  it("单格只显示一个地址", () => {
    expect(rangeToA1({ startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 })).toBe("A1");
  });

  it("多格显示区间，与 Excel 名称框一致", () => {
    expect(rangeToA1({ startRow: 0, startColumn: 0, endRow: 2, endColumn: 2 })).toBe("A1:C3");
  });
});

describe("parseA1", () => {
  it("认得出常规写法", () => {
    expect(parseA1("A1")).toEqual({ row: 0, column: 0 });
    expect(parseA1("B7")).toEqual({ row: 6, column: 1 });
    expect(parseA1("AA1")).toEqual({ row: 0, column: 26 });
  });

  it("大小写不敏感，前后空白忽略", () => {
    expect(parseA1("  b7 ")).toEqual({ row: 6, column: 1 });
  });

  it("认不出来返回 null 而不是抛——名称框里输入任意文本是常态", () => {
    expect(parseA1("")).toBeNull();
    expect(parseA1("hello")).toBeNull();
    expect(parseA1("7B")).toBeNull();
    expect(parseA1("A0")).toBeNull(); // 行号从 1 开始
    expect(parseA1("$A$1")).toBeNull(); // 绝对引用要等公式
    expect(parseA1("Sheet1!A1")).toBeNull(); // 跨表要等多 Sheet
  });

  it("与 toA1 互为逆运算", () => {
    for (const [row, column] of [[0, 0], [6, 1], [99, 26], [12345, 701]] as const) {
      expect(parseA1(toA1(row, column))).toEqual({ row, column });
    }
  });
});
