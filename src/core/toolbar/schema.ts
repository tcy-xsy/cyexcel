import { STYLE_KEYS, type CellStyle } from "../style/types";
import { addDecimal, subtractDecimal } from "../style/numFmtDecimals";
import { commonValue, everyCellHas } from "./state";
import { DEFAULT_FONT_SIZE } from "../render/theme";
import type { ToolbarItem, ToolbarSchema } from "./types";
import * as icons from "./icons";
import { applyBorderPreset, borderPanel, colorPanel, filterPanel, findPanel } from "./panels";

/**
 * toggle 一个布尔样式键：整片选区已经全是真就清掉，否则全部设上。
 *
 * 这条"全真才算激活"的语义在 `everyCellHas` 上有更长的解释——与 Excel 一致，
 * 也是唯一能让"再点一次"有确定含义的定义。
 */
function toggleStyleItem(
  id: string,
  key: "bold" | "italic" | "underline" | "strike",
  label: string,
  order: number,
  icon: string
): ToolbarItem {
  return {
    id,
    group: "format",
    order,
    type: "toggle",
    icon,
    label,
    tooltip: label,
    activated: (ctx) => everyCellHas(ctx, key),
    run: (ctx) => {
      const on = everyCellHas(ctx, key);
      ctx.sheet.setRangeStyle(ctx.range, { [key]: on ? null : true });
    }
  };
}

/** 对齐是一组互斥项：点已激活的那个不会取消，与 Excel 一样（总得有个对齐方式）。 */
function alignItem(
  id: string,
  key: "align" | "verticalAlign",
  value: string,
  label: string,
  order: number,
  icon: string
): ToolbarItem {
  return {
    id,
    group: "align",
    order,
    type: "toggle",
    icon,
    label,
    tooltip: label,
    activated: (ctx) => commonValue(ctx, key) === value,
    run: (ctx) => {
      ctx.sheet.setRangeStyle(ctx.range, { [key]: value });
    }
  };
}

const FONT_FAMILIES = [
  { value: "", label: "默认字体" },
  { value: "SimSun", label: "宋体" },
  { value: "SimHei", label: "黑体" },
  { value: "Microsoft YaHei", label: "微软雅黑" },
  { value: "Arial", label: "Arial" },
  { value: "Courier New", label: "Courier New" }
] as const;

const FONT_SIZES = [9, 10, 11, 12, 13, 14, 16, 18, 20, 24, 28, 32, 48].map((size) => ({
  value: size,
  label: String(size)
}));

/**
 * 数字格式给的是**预设**，但存下去的是 Excel pattern 串本身——预设只是几个常用的
 * 入口，不是另一套命名体系（ADR 0004 §5④ 提醒过：自造预设名会在 xlsx 导出时欠一张
 * 映射表）。想要别的格式，宿主直接 `setRangeStyle(range, { numFmt: "…" })`。
 */
const NUMBER_FORMATS = [
  { value: "", label: "常规" },
  { value: "0", label: "整数" },
  { value: "#,##0.00", label: "千分位两位小数" },
  { value: "0%", label: "百分比" },
  { value: "0.00%", label: "百分比两位小数" },
  { value: "¥#,##0.00", label: "货币" },
  { value: "yyyy-mm-dd", label: "日期" },
  { value: "yyyy-mm-dd hh:mm", label: "日期时间" }
] as const;

/**
 * 默认工具栏 schema。
 *
 * `order` 留了空隙（10、20、30…），宿主想往两项中间插东西不必改我们的数字。
 */
export const defaultToolbarSchema: ToolbarSchema = Object.freeze([
  // ---- history ----
  {
    id: "undo",
    group: "history",
    order: 10,
    type: "button",
    icon: icons.undo,
    label: "撤销",
    tooltip: "撤销 (Ctrl/Cmd+Z)",
    run: (ctx) => ctx.sheet.undo()
  },
  {
    id: "redo",
    group: "history",
    order: 20,
    type: "button",
    icon: icons.redo,
    label: "重做",
    tooltip: "重做 (Ctrl/Cmd+Shift+Z)",
    run: (ctx) => ctx.sheet.redo()
  },
  {
    id: "copy",
    group: "history",
    order: 30,
    type: "button",
    icon: icons.copy,
    label: "复制",
    tooltip: "复制 (Ctrl/Cmd+C)",
    run: (ctx) => ctx.sheet.copy()
  },
  {
    id: "paste",
    group: "history",
    order: 40,
    type: "button",
    icon: icons.paste,
    label: "粘贴",
    tooltip: "粘贴 (Ctrl/Cmd+V)",
    run: (ctx) => ctx.sheet.paste()
  },
  {
    id: "clear-style",
    group: "history",
    order: 50,
    type: "button",
    icon: icons.clearStyle,
    label: "清除样式",
    tooltip: "清除选区的全部样式，保留值",
    run: (ctx) => ctx.sheet.clearRangeStyle(ctx.range)
  },

  // ---- format ----
  {
    id: "font-family",
    group: "format",
    order: 10,
    type: "select",
    label: "字体",
    options: FONT_FAMILIES,
    // fallback ""：没设字体 = 默认字体，两者不该报成"不一致"。不一致时仍然
    // 返回 undefined，下拉框显示空白。
    value: (ctx) => commonValue(ctx, "fontFamily", ""),
    apply: (ctx, value) => {
      ctx.sheet.setRangeStyle(ctx.range, { fontFamily: value === "" ? null : String(value) });
    }
  },
  {
    id: "font-size",
    group: "format",
    order: 20,
    type: "select",
    label: "字号",
    options: FONT_SIZES,
    // 显示**实际生效**的字号而不是空白（同 Excel：点任何格子字号框都有数字）。
    // DEFAULT_FONT_SIZE 正是 resolveDrawStyle 对缺省 fontSize 取的值，也在 FONT_SIZES 里。
    value: (ctx) => commonValue(ctx, "fontSize", DEFAULT_FONT_SIZE),
    apply: (ctx, value) => {
      ctx.sheet.setRangeStyle(ctx.range, { fontSize: Number(value) });
    }
  },
  toggleStyleItem("bold", "bold", "加粗", 30, icons.bold),
  toggleStyleItem("italic", "italic", "斜体", 40, icons.italic),
  toggleStyleItem("underline", "underline", "下划线", 50, icons.underline),
  toggleStyleItem("strike", "strike", "删除线", 60, icons.strike),
  // 颜色与边框是**分段按钮**：点图标直接套用上次选过的那个，点箭头才展开面板。
  // 同时有 panel 和 apply 的项会被渲染成这种形态（见 createToolbar.buildControl）。
  {
    id: "text-color",
    group: "format",
    order: 70,
    type: "color",
    icon: icons.textColor,
    label: "字体颜色",
    tooltip: "字体颜色（点图标套用上次的颜色）",
    defaultValue: "#ef4444",
    value: (ctx) => commonValue(ctx, "color") ?? "",
    apply: (ctx, value) => {
      ctx.sheet.setRangeStyle(ctx.range, { color: String(value) });
    },
    panel: (ctx, host, onChosen) =>
      colorPanel(host, {
        current: commonValue(ctx, "color"),
        onPick: (color) => ctx.sheet.setRangeStyle(ctx.range, { color: color ?? null }),
        ...(onChosen !== undefined ? { onChosen } : {})
      })
  },
  {
    id: "background-color",
    group: "format",
    order: 80,
    type: "color",
    icon: icons.background,
    label: "背景颜色",
    tooltip: "背景颜色（点图标套用上次的颜色）",
    defaultValue: "#eab308",
    value: (ctx) => commonValue(ctx, "background") ?? "",
    apply: (ctx, value) => {
      ctx.sheet.setRangeStyle(ctx.range, { background: String(value) });
    },
    panel: (ctx, host, onChosen) =>
      colorPanel(host, {
        current: commonValue(ctx, "background"),
        onPick: (color) => ctx.sheet.setRangeStyle(ctx.range, { background: color ?? null }),
        ...(onChosen !== undefined ? { onChosen } : {})
      })
  },
  {
    id: "border",
    group: "format",
    order: 90,
    type: "panel",
    icon: icons.border,
    label: "边框",
    tooltip: "边框（点图标套用上次的边框）",
    defaultValue: "all",
    apply: (ctx, value) => applyBorderPreset(ctx, String(value)),
    panel: (ctx, host, onChosen) => borderPanel(host, ctx, onChosen)
  },
  {
    id: "number-format",
    group: "format",
    order: 100,
    type: "select",
    label: "数字格式",
    tooltip: "数字格式",
    options: NUMBER_FORMATS,
    value: (ctx) => commonValue(ctx, "numFmt") ?? "",
    apply: (ctx, value) => {
      ctx.sheet.setRangeStyle(ctx.range, { numFmt: value === "" ? null : String(value) });
    }
  },
  // Excel 与 Univer 工具栏上都有的四颗数字格式快捷键（Univer 对应
  // SetPercentCommand / AddDecimalCommand / SubtractDecimalCommand）。
  {
    id: "format-percent",
    group: "format",
    order: 110,
    type: "button",
    icon: icons.percent,
    label: "百分比格式",
    tooltip: "百分比格式（0%）",
    run: (ctx) => ctx.sheet.setRangeStyle(ctx.range, { numFmt: "0%" })
  },
  {
    id: "format-thousands",
    group: "format",
    order: 120,
    type: "button",
    icon: icons.thousands,
    label: "千分位",
    tooltip: "千分位样式（#,##0.00）",
    run: (ctx) => ctx.sheet.setRangeStyle(ctx.range, { numFmt: "#,##0.00" })
  },
  {
    id: "add-decimal",
    group: "format",
    order: 130,
    type: "button",
    icon: icons.addDecimalIcon,
    label: "增加小数位",
    tooltip: "增加小数位",
    run: (ctx) => {
      ctx.sheet.setRangeStyle(ctx.range, { numFmt: addDecimal(commonValue(ctx, "numFmt")) });
    }
  },
  {
    id: "subtract-decimal",
    group: "format",
    order: 140,
    type: "button",
    icon: icons.subtractDecimalIcon,
    label: "减少小数位",
    tooltip: "减少小数位",
    run: (ctx) => {
      ctx.sheet.setRangeStyle(ctx.range, { numFmt: subtractDecimal(commonValue(ctx, "numFmt")) });
    }
  },

  // ---- align ----
  alignItem("align-left", "align", "left", "左对齐", 10, icons.alignLeft),
  alignItem("align-center", "align", "center", "居中", 20, icons.alignCenter),
  alignItem("align-right", "align", "right", "右对齐", 30, icons.alignRight),
  alignItem("align-top", "verticalAlign", "top", "顶端对齐", 40, icons.alignTop),
  alignItem("align-middle", "verticalAlign", "middle", "垂直居中", 50, icons.alignMiddle),
  alignItem("align-bottom", "verticalAlign", "bottom", "底端对齐", 60, icons.alignBottom),

  // ---- data ----
  {
    id: "sort-asc",
    group: "data",
    order: 10,
    type: "button",
    icon: icons.sortAsc,
    label: "升序",
    tooltip: "按当前列升序排序",
    run: (ctx) => {
      ctx.sheet.sortByColumn(ctx.range.startColumn, "asc");
    }
  },
  {
    id: "sort-desc",
    group: "data",
    order: 20,
    type: "button",
    icon: icons.sortDesc,
    label: "降序",
    tooltip: "按当前列降序排序",
    run: (ctx) => {
      ctx.sheet.sortByColumn(ctx.range.startColumn, "desc");
    }
  },
  {
    id: "filter",
    group: "data",
    order: 30,
    type: "panel",
    icon: icons.filter,
    label: "过滤",
    panel: (ctx, host) => filterPanel(host, ctx)
  },
  {
    id: "find",
    group: "data",
    order: 40,
    type: "panel",
    icon: icons.find,
    label: "查找替换",
    tooltip: "查找与替换",
    panel: (ctx, host) => findPanel(host, ctx)
  },

  // ---- structure ----
  {
    id: "merge",
    group: "structure",
    order: 10,
    type: "button",
    icon: icons.merge,
    label: "合并单元格",
    // 单格不构成合并，禁用而不是让它抛错
    disabled: (ctx) =>
      ctx.range.startRow === ctx.range.endRow && ctx.range.startColumn === ctx.range.endColumn,
    run: (ctx) => ctx.sheet.mergeCells(ctx.range)
  },
  {
    id: "unmerge",
    group: "structure",
    order: 20,
    type: "button",
    icon: icons.unmerge,
    label: "取消合并",
    run: (ctx) => ctx.sheet.unmergeCells(ctx.range)
  },
  {
    id: "insert-row",
    group: "structure",
    order: 30,
    type: "button",
    icon: icons.insertRow,
    label: "插入行",
    tooltip: "在选区上方插入一行",
    run: (ctx) => ctx.sheet.insertRows(ctx.range.startRow, 1)
  },
  {
    id: "delete-row",
    group: "structure",
    order: 40,
    type: "button",
    icon: icons.deleteRow,
    label: "删除行",
    run: (ctx) =>
      ctx.sheet.deleteRows(ctx.range.startRow, ctx.range.endRow - ctx.range.startRow + 1)
  },
  {
    id: "insert-column",
    group: "structure",
    order: 50,
    type: "button",
    icon: icons.insertColumn,
    label: "插入列",
    tooltip: "在选区左侧插入一列",
    run: (ctx) => ctx.sheet.insertColumns(ctx.range.startColumn, [{}])
  },
  {
    id: "delete-column",
    group: "structure",
    order: 60,
    type: "button",
    icon: icons.deleteColumn,
    label: "删除列",
    run: (ctx) =>
      ctx.sheet.deleteColumns(ctx.range.startColumn, ctx.range.endColumn - ctx.range.startColumn + 1)
  },
  {
    id: "hide-column",
    group: "structure",
    order: 70,
    type: "button",
    icon: icons.hideColumn,
    label: "隐藏列",
    run: (ctx) => ctx.sheet.hideColumn(ctx.range.startColumn)
  }
]);

/** `clearRangeStyle` 用的全键 patch，schema 之外也能用，所以单独导出一份。 */
export const CLEAR_ALL_STYLE_PATCH = Object.freeze(
  Object.fromEntries(STYLE_KEYS.map((key) => [key, null]))
) as { readonly [K in keyof CellStyle]: null };
