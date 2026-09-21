import { describe, expect, it } from "vitest";
import { CyExcelError } from "../../../../src/core/errors";
import {
  assertValidNumFmt,
  formatCellText
} from "../../../../src/core/style/formatCellText";

const fallback = (text: string) => () => text;

describe("assertValidNumFmt", () => {
  it("接受常见 pattern", () => {
    expect(() => assertValidNumFmt("#,##0.00")).not.toThrow();
    expect(() => assertValidNumFmt("0%")).not.toThrow();
    expect(() => assertValidNumFmt("yyyy-mm-dd")).not.toThrow();
  });

  it("非法 pattern 抛 invalid-number-format，不静默降级", () => {
    try {
      assertValidNumFmt("[[[");
      throw new Error("should have thrown");
    } catch (error) {
      expect(error).toBeInstanceOf(CyExcelError);
      expect((error as CyExcelError).code).toBe("invalid-number-format");
    }
  });
});

describe("formatCellText", () => {
  it("没有 numFmt 时退回 fallback——这是 CellType.format 的向后兼容路径", () => {
    expect(formatCellText(1234, undefined, fallback("由列格式化"))).toBe("由列格式化");
    expect(formatCellText(1234, { bold: true }, fallback("由列格式化"))).toBe("由列格式化");
  });

  it("数字按 pattern 格式化", () => {
    expect(formatCellText(1234.5, { numFmt: "#,##0.00" }, fallback("原样"))).toBe("1,234.50");
  });

  it("百分比 pattern", () => {
    expect(formatCellText(0.25, { numFmt: "0%" }, fallback("原样"))).toBe("25%");
  });

  it("Date 直接交给 numfmt——它自己认 Date，不需要先转序列号", () => {
    const date = new Date(2026, 8, 21);
    expect(formatCellText(date, { numFmt: "yyyy-mm-dd" }, fallback("原样"))).toBe("2026-09-21");
  });

  it("字符串不进 numfmt，走 fallback——不做隐式的字符串转数字", () => {
    // 存的是 "0012"，设了数字格式也不该被画成 12：那没法解释给用户听。
    expect(formatCellText("0012", { numFmt: "#,##0" }, fallback("0012"))).toBe("0012");
  });

  it("布尔、null、undefined 都走 fallback", () => {
    expect(formatCellText(true, { numFmt: "#,##0" }, fallback("TRUE"))).toBe("TRUE");
    expect(formatCellText(null, { numFmt: "#,##0" }, fallback(""))).toBe("");
    expect(formatCellText(undefined, { numFmt: "#,##0" }, fallback(""))).toBe("");
  });

  it("格式化过程抛错时兜回 fallback，不让渲染整帧崩掉", () => {
    const exploding = {
      get numFmt(): string {
        return "#,##0";
      }
    };
    // numfmt 对极端值域可能抛；这里直接构造一个会让它抛的输入替身。
    const hostile = {
      valueOf() {
        throw new Error("boom");
      }
    } as unknown as number;
    expect(formatCellText(hostile, exploding, fallback("兜底"))).toBe("兜底");
  });
});
