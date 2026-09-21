import { describe, expect, it } from "vitest";
import { createColumnAwarePainter } from "../../../../src/core/cells/createColumnAwarePainter";
import { ColumnManager } from "../../../../src/core/cells/ColumnManager";
import { CellTypeRegistry } from "../../../../src/core/cells/CellTypeRegistry";
import { registerBuiltinCellTypes } from "../../../../src/core/cells/builtins";
import { CELL_FONT, TEXT_COLOR } from "../../../../src/core/render/theme";
import type { CellStyle } from "../../../../src/core/style/types";
import { RecordingCanvasContext } from "../../../helpers/RecordingCanvasContext";
import type { CellValue } from "../../../../src/core/sheet/types";

const rect = { x: 100, y: 50, width: 80, height: 24 };

/**
 * 画一格，把 ctx 上发生的事记下来。
 *
 * 这是渲染层在这个环境里能被验证的那一半——像素长什么样只能靠眼睛（见
 * doc/step24-style-plan.md §6），但"用了哪个 font 串、填了什么色、文字落在哪个
 * x 上"是确定的，而恰恰是这些最容易写错。
 */
function paint(
  style: CellStyle | undefined,
  value: CellValue = "文本",
  cellRect: typeof rect = rect
) {
  const ctx = new RecordingCanvasContext();
  const registry = new CellTypeRegistry();
  registerBuiltinCellTypes(registry);
  const columns = new ColumnManager({ columns: [{}], registry });
  const painter = createColumnAwarePainter(() => columns.getColumns());
  painter({
    ctx,
    rect: cellRect,
    row: 0,
    column: 0,
    value,
    ...(style !== undefined ? { style } : {})
  });
  return ctx;
}

function fillTextCall(ctx: RecordingCanvasContext) {
  return ctx.calls.find((call) => call.method === "fillText");
}

describe("样式绘制", () => {
  it("没有样式时，font 与颜色和 Step 24 之前逐字相同", () => {
    const ctx = paint(undefined);
    expect(ctx.font).toBe(CELL_FONT);
    expect(ctx.fillStyle).toBe(TEXT_COLOR);
  });

  it("bold/italic 拼进 font，顺序是 CSS 要求的 style weight size family", () => {
    // 顺序反了（"600 italic ..."）整条 font 声明会被浏览器丢弃，
    // 表现为这一格悄悄变回默认字体——所以顺序要锁死。
    expect(paint({ bold: true }).font).toBe("600 13px " + CELL_FONT.split("px ")[1]);
    expect(paint({ italic: true }).font).toBe("italic 13px " + CELL_FONT.split("px ")[1]);
    expect(paint({ bold: true, italic: true }).font).toBe(
      "italic 600 13px " + CELL_FONT.split("px ")[1]
    );
  });

  it("fontSize/fontFamily 各自可覆盖", () => {
    expect(paint({ fontSize: 20 }).font).toBe("20px " + CELL_FONT.split("px ")[1]);
    expect(paint({ fontFamily: "Courier" }).font).toBe("13px Courier");
  });

  it("color 改文字颜色", () => {
    expect(paint({ color: "#ff0000" }).fillStyle).toBe("#ff0000");
  });

  it("居中对齐时文字的 x 是格子中线，不是左内边距", () => {
    // textAlign 决定"给定的 x 算左边还是中间"，两者必须一起改，
    // 只改其中一个会让文字偏到看不见的地方。
    const ctx = paint({ align: "center" });
    expect(ctx.textAlign).toBe("center");
    expect(fillTextCall(ctx)?.args[1]).toBe(rect.x + rect.width / 2);
  });

  it("右对齐时 x 是右边界减去内边距", () => {
    const ctx = paint({ align: "right" });
    expect(fillTextCall(ctx)?.args[1]).toBe(rect.x + rect.width - 4);
  });

  it("默认左对齐，x 是左边界加内边距", () => {
    expect(fillTextCall(paint(undefined))?.args[1]).toBe(rect.x + 4);
  });

  it("padding 覆盖默认内边距", () => {
    expect(fillTextCall(paint({ padding: 10 }))?.args[1]).toBe(rect.x + 10);
  });

  it("纵向对齐改的是基线 y", () => {
    // 必须用高行才看得出区别：默认 24px 行高里 16px 行高 + 4px 内边距正好填满，
    // top 与 middle 本来就重合——这不是 bug，是那个高度下唯一合理的结果。
    const tall = { ...rect, height: 60 };
    const middle = fillTextCall(paint(undefined, "文本", tall))?.args[2];
    const top = fillTextCall(paint({ verticalAlign: "top" }, "文本", tall))?.args[2];
    const bottom = fillTextCall(paint({ verticalAlign: "bottom" }, "文本", tall))?.args[2];
    expect(middle).toBe(tall.y + tall.height / 2);
    expect(top).toBeLessThan(middle as number);
    expect(bottom).toBeGreaterThan(middle as number);
  });

  it("下划线量出文字宽度后自己画线——Canvas 没有原生下划线", () => {
    const ctx = paint({ underline: true });
    expect(ctx.calls.some((call) => call.method === "measureText")).toBe(true);
    expect(ctx.calls.some((call) => call.method === "stroke")).toBe(true);
  });

  it("不设下划线/删除线就不会有多余的描边调用", () => {
    expect(paint(undefined).calls.some((call) => call.method === "stroke")).toBe(false);
  });

  it("numFmt 生效时画的是格式化后的文本", () => {
    const ctx = paint({ numFmt: "#,##0.00" }, 1234.5);
    expect(fillTextCall(ctx)?.args[0]).toBe("1,234.50");
  });

  it("没有 numFmt 时画的是列格式化的结果", () => {
    expect(fillTextCall(paint(undefined, 1234.5))?.args[0]).toBe("1234.5");
  });
});
