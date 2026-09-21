import { describe, expect, it, vi } from "vitest";
import { defaultToolbarSchema } from "../../../../src/core/toolbar/schema";
import {
  commonValue,
  everyCellHas,
  MAX_PROBED_CELLS,
  resolveItemState
} from "../../../../src/core/toolbar/state";
import { TOOLBAR_GROUPS, type ToolbarContext } from "../../../../src/core/toolbar/types";
import type { CellStyle } from "../../../../src/core/style/types";
import type { CyExcelCore } from "../../../../src/core/CyExcelCore";

/**
 * 假的 ToolbarContext——`state.ts` 是纯函数，测它不需要真的 CyExcelCore。
 * 这正是计划 §1 把逻辑与 DOM 分层的收益：按钮状态是最容易写错的部分，而它在 node
 * 里可以完整测。
 */
function contextOf(
  styles: Record<string, CellStyle | undefined>,
  range = { startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 },
  sheet: Partial<CyExcelCore> = {}
): ToolbarContext {
  return {
    sheet: sheet as CyExcelCore,
    range,
    styleOf: (row, column) => styles[`${row},${column}`]
  };
}

describe("默认 schema 的结构不变量", () => {
  it("id 唯一", () => {
    const ids = defaultToolbarSchema.map((item) => item.id);
    expect(new Set(ids).size).toBe(ids.length);
  });

  it("每一项的 group 都是已知分组", () => {
    for (const item of defaultToolbarSchema) {
      expect(TOOLBAR_GROUPS).toContain(item.group);
    }
  });

  it("同一组内 order 不重复——否则渲染顺序取决于数组写法，不稳定", () => {
    for (const group of TOOLBAR_GROUPS) {
      const orders = defaultToolbarSchema.filter((item) => item.group === group).map((item) => item.order);
      expect(new Set(orders).size).toBe(orders.length);
    }
  });

  it("select 项必须带 options 和 apply，否则渲染出来是个空下拉", () => {
    for (const item of defaultToolbarSchema.filter((entry) => entry.type === "select")) {
      expect(item.options, item.id).toBeDefined();
      expect(item.options?.length, item.id).toBeGreaterThan(0);
      expect(item.apply, item.id).toBeTypeOf("function");
    }
  });

  it("panel/color 项必须带 panel 函数", () => {
    for (const item of defaultToolbarSchema.filter((entry) => entry.type === "panel" || entry.type === "color")) {
      expect(item.panel, item.id).toBeTypeOf("function");
    }
  });

  it("button/toggle 项必须带 run", () => {
    for (const item of defaultToolbarSchema.filter((entry) => entry.type === "button" || entry.type === "toggle")) {
      expect(item.run, item.id).toBeTypeOf("function");
    }
  });

  it("每一项都有 label——图标按钮也要有，读屏靠它", () => {
    for (const item of defaultToolbarSchema) {
      expect(item.label.length, item.id).toBeGreaterThan(0);
    }
  });

  it("四组能力都在：撤销、加粗、排序、合并各能找到", () => {
    const ids = defaultToolbarSchema.map((item) => item.id);
    expect(ids).toContain("undo");
    expect(ids).toContain("bold");
    expect(ids).toContain("sort-asc");
    expect(ids).toContain("merge");
  });
});

describe("everyCellHas", () => {
  it("整片选区都为真才算激活——与 Excel 一致", () => {
    const all = contextOf({
      "0,0": { bold: true },
      "0,1": { bold: true },
      "1,0": { bold: true },
      "1,1": { bold: true }
    });
    expect(everyCellHas(all, "bold")).toBe(true);
  });

  it("有一格不是就不激活——半粗的选区，按钮不该亮", () => {
    const partial = contextOf({
      "0,0": { bold: true },
      "0,1": { bold: true },
      "1,0": { bold: true },
      "1,1": undefined
    });
    expect(everyCellHas(partial, "bold")).toBe(false);
  });

  it("超大选区只看左上角那一格——状态是提示，不值得为它扫一百万格", () => {
    const styleOf = vi.fn(() => ({ bold: true }) as CellStyle);
    const huge: ToolbarContext = {
      sheet: {} as CyExcelCore,
      range: { startRow: 0, startColumn: 0, endRow: MAX_PROBED_CELLS, endColumn: 10 },
      styleOf
    };
    expect(everyCellHas(huge, "bold")).toBe(true);
    expect(styleOf).toHaveBeenCalledTimes(1);
  });
});

describe("commonValue", () => {
  it("选区里一致时返回那个值", () => {
    const ctx = contextOf({
      "0,0": { fontSize: 16 },
      "0,1": { fontSize: 16 },
      "1,0": { fontSize: 16 },
      "1,1": { fontSize: 16 }
    });
    expect(commonValue(ctx, "fontSize")).toBe(16);
  });

  it("不一致时是 undefined——下拉框该显示空白，不是随便挑一格的值", () => {
    const ctx = contextOf({
      "0,0": { fontSize: 16 },
      "0,1": { fontSize: 12 },
      "1,0": { fontSize: 16 },
      "1,1": { fontSize: 16 }
    });
    expect(commonValue(ctx, "fontSize")).toBeUndefined();
  });

  it("都没设时是 undefined", () => {
    expect(commonValue(contextOf({}), "fontSize")).toBeUndefined();
  });
});

describe("resolveItemState", () => {
  it("没声明 activated/disabled/value 的项拿到确定的默认值", () => {
    const item = defaultToolbarSchema.find((entry) => entry.id === "undo");
    expect(resolveItemState(item!, contextOf({}))).toEqual({
      activated: false,
      disabled: false,
      value: undefined
    });
  });

  it("加粗项在全粗的选区上是激活的", () => {
    const bold = defaultToolbarSchema.find((entry) => entry.id === "bold");
    const ctx = contextOf({
      "0,0": { bold: true },
      "0,1": { bold: true },
      "1,0": { bold: true },
      "1,1": { bold: true }
    });
    expect(resolveItemState(bold!, ctx).activated).toBe(true);
  });

  it("合并项在单格选区上禁用——单格不构成合并，禁用好过让它抛错", () => {
    const merge = defaultToolbarSchema.find((entry) => entry.id === "merge");
    const single = contextOf({}, { startRow: 2, startColumn: 3, endRow: 2, endColumn: 3 });
    expect(resolveItemState(merge!, single).disabled).toBe(true);

    const multi = contextOf({}, { startRow: 2, startColumn: 3, endRow: 3, endColumn: 4 });
    expect(resolveItemState(merge!, multi).disabled).toBe(false);
  });
});

describe("toggle 的 run", () => {
  it("全粗时再点一次是清掉，不是重复设上", () => {
    const setRangeStyle = vi.fn();
    const ctx = contextOf(
      { "0,0": { bold: true }, "0,1": { bold: true }, "1,0": { bold: true }, "1,1": { bold: true } },
      { startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 },
      { setRangeStyle } as unknown as Partial<CyExcelCore>
    );
    defaultToolbarSchema.find((item) => item.id === "bold")?.run?.(ctx);
    expect(setRangeStyle).toHaveBeenCalledWith(ctx.range, { bold: null });
  });

  it("半粗时点一次是全部变粗", () => {
    const setRangeStyle = vi.fn();
    const ctx = contextOf(
      { "0,0": { bold: true } },
      { startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 },
      { setRangeStyle } as unknown as Partial<CyExcelCore>
    );
    defaultToolbarSchema.find((item) => item.id === "bold")?.run?.(ctx);
    expect(setRangeStyle).toHaveBeenCalledWith(ctx.range, { bold: true });
  });
});
