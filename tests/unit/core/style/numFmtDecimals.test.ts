import { describe, expect, it } from "vitest";
import { addDecimal, subtractDecimal } from "../../../../src/core/style/numFmtDecimals";

describe("addDecimal", () => {
  it("没设过格式时从整数基线开始", () => {
    expect(addDecimal(undefined)).toBe("0.0");
    expect(addDecimal("")).toBe("0.0");
  });

  it("逐次加一位", () => {
    expect(addDecimal("0")).toBe("0.0");
    expect(addDecimal("0.0")).toBe("0.00");
    expect(addDecimal("#,##0.00")).toBe("#,##0.000");
  });

  it("保留 pattern 的其余部分——千分位、货币符号都不能丢", () => {
    expect(addDecimal("¥#,##0")).toBe("¥#,##0.0");
  });

  it("百分号在小数点之后，加位不会把它挤掉", () => {
    expect(addDecimal("0%")).toBe("0.0%");
  });

  it("封顶 10 位", () => {
    expect(addDecimal("0.0000000000")).toBe("0.0000000000");
  });

  it("分段、科学计数、文本段、日期一律原样返回——那些的“小数位”没有确定含义", () => {
    expect(addDecimal("0.00;[Red]-0.00")).toBe("0.00;[Red]-0.00");
    expect(addDecimal("0.00E+00")).toBe("0.00E+00");
    expect(addDecimal("@")).toBe("@");
    expect(addDecimal("yyyy-mm-dd")).toBe("yyyy-mm-dd");
  });
});

describe("subtractDecimal", () => {
  it("逐次减一位", () => {
    expect(subtractDecimal("#,##0.000")).toBe("#,##0.00");
    expect(subtractDecimal("0.0")).toBe("0");
  });

  it("减到 0 位就停，并且不留下一个光秃秃的小数点", () => {
    expect(subtractDecimal("0")).toBe("0");
    expect(subtractDecimal("#,##0")).toBe("#,##0");
  });

  it("保留 pattern 的其余部分", () => {
    expect(subtractDecimal("¥#,##0.00")).toBe("¥#,##0.0");
    expect(subtractDecimal("0.00%")).toBe("0.0%");
  });

  it("不可调整的 pattern 原样返回", () => {
    expect(subtractDecimal("yyyy-mm-dd hh:mm")).toBe("yyyy-mm-dd hh:mm");
  });
});
