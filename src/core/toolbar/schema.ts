import { defaultTranslator, type Translator } from "../locale/index.js";
import { STYLE_KEYS, type CellStyle } from "../style/types.js";
import { addDecimal, subtractDecimal } from "../style/numFmtDecimals.js";
import { commonValue, everyCellHas } from "./state.js";
import { DEFAULT_FONT_SIZE } from "../render/theme.js";
import type { ToolbarContext, ToolbarItem, ToolbarSchema } from "./types.js";
import * as icons from "./icons.js";
import { applyBorderPreset, borderPanel, colorPanel, filterPanel } from "./panels.js";
import { dataValidationDialog, findReplaceDialog } from "./dialogs.js";
import { exportSheetToExcel } from "./exportExcel.js";

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

/**
 * 弹出系统文件选择器，把选中的图片读成 `data:` URI 写进当前格。用一次性 `<input>`
 * 而不是常驻 DOM 节点——工具栏项没有专属挂载点，用完即扔最简单。
 */
function pickImageFile(ctx: ToolbarContext): Promise<void> {
  return new Promise((resolve) => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = "image/*";
    input.style.display = "none";
    input.addEventListener("change", () => {
      const file = input.files?.[0];
      input.remove();
      if (!file) {
        resolve();
        return;
      }
      const reader = new FileReader();
      reader.onload = () => {
        ctx.sheet.setCellImage(ctx.range.startRow, ctx.range.startColumn, String(reader.result));
        resolve();
      };
      reader.onerror = () => resolve();
      reader.readAsDataURL(file);
    });
    document.body.appendChild(input);
    input.click();
  });
}

export function createDefaultToolbarSchema(t: Translator = defaultTranslator): ToolbarSchema {
  const FONT_FAMILIES = [
    { value: "", label: t("Default font") },
    { value: "SimSun", label: t("SimSun") },
    { value: "SimHei", label: t("SimHei") },
    { value: "Microsoft YaHei", label: t("Microsoft YaHei") },
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
    { value: "", label: t("General") },
    { value: "0", label: t("Whole number") },
    { value: "#,##0.00", label: t("Number with thousands separator and two decimals") },
    { value: "0%", label: t("Percentage") },
    { value: "0.00%", label: t("Percentage with two decimals") },
    { value: "¥#,##0.00", label: t("Currency") },
    { value: "yyyy-mm-dd", label: t("Date") },
    { value: "yyyy-mm-dd hh:mm", label: t("Date and time") }
  ] as const;

  /**
   * 默认工具栏 schema。
   *
   * `order` 留了空隙（10、20、30…），宿主想往两项中间插东西不必改我们的数字。
   */
  return Object.freeze([
    // ---- history ----
    {
      id: "undo",
      group: "history",
      order: 10,
      type: "button",
      icon: icons.undo,
      label: t("Undo"),
      tooltip: t("Undo (Ctrl/Cmd+Z)"),
      run: (ctx) => ctx.sheet.undo()
    },
    {
      id: "redo",
      group: "history",
      order: 20,
      type: "button",
      icon: icons.redo,
      label: t("Redo"),
      tooltip: t("Redo (Ctrl/Cmd+Shift+Z)"),
      run: (ctx) => ctx.sheet.redo()
    },
    {
      id: "copy",
      group: "history",
      order: 30,
      type: "button",
      icon: icons.copy,
      label: t("Copy"),
      tooltip: t("Copy (Ctrl/Cmd+C)"),
      run: (ctx) => ctx.sheet.copy()
    },
    {
      id: "paste",
      group: "history",
      order: 40,
      type: "button",
      icon: icons.paste,
      label: t("Paste"),
      tooltip: t("Paste (Ctrl/Cmd+V)"),
      run: (ctx) => ctx.sheet.paste()
    },
    {
      id: "clear-style",
      group: "history",
      order: 50,
      type: "button",
      icon: icons.clearStyle,
      label: t("Clear styles"),
      tooltip: t("Clear all styles in the selection and keep values"),
      run: (ctx) => ctx.sheet.clearRangeStyle(ctx.range)
    },

    // ---- format ----
    {
      id: "font-family",
      group: "format",
      order: 10,
      type: "select",
      label: t("Font family"),
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
      label: t("Font size"),
      options: FONT_SIZES,
      // 显示**实际生效**的字号而不是空白（同 Excel：点任何格子字号框都有数字）。
      // DEFAULT_FONT_SIZE 正是 resolveDrawStyle 对缺省 fontSize 取的值，也在 FONT_SIZES 里。
      value: (ctx) => commonValue(ctx, "fontSize", DEFAULT_FONT_SIZE),
      apply: (ctx, value) => {
        ctx.sheet.setRangeStyle(ctx.range, { fontSize: Number(value) });
      }
    },
    toggleStyleItem("bold", "bold", t("Bold"), 30, icons.bold),
    toggleStyleItem("italic", "italic", t("Italic"), 40, icons.italic),
    toggleStyleItem("underline", "underline", t("Underline"), 50, icons.underline),
    toggleStyleItem("strike", "strike", t("Strikethrough"), 60, icons.strike),
    // 颜色与边框是**分段按钮**：点图标直接套用上次选过的那个，点箭头才展开面板。
    // 同时有 panel 和 apply 的项会被渲染成这种形态（见 createToolbar.buildControl）。
    {
      id: "text-color",
      group: "format",
      order: 70,
      type: "color",
      icon: icons.textColor,
      label: t("Text color"),
      tooltip: t("Text color (click the icon to reuse the last color)"),
      defaultValue: "#ef4444",
      value: (ctx) => commonValue(ctx, "color") ?? "",
      apply: (ctx, value) => {
        ctx.sheet.setRangeStyle(ctx.range, { color: String(value) });
      },
      panel: (ctx, host, onChosen) =>
        colorPanel(host, {
          t,
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
      label: t("Fill color"),
      tooltip: t("Fill color (click the icon to reuse the last color)"),
      defaultValue: "#eab308",
      value: (ctx) => commonValue(ctx, "background") ?? "",
      apply: (ctx, value) => {
        ctx.sheet.setRangeStyle(ctx.range, { background: String(value) });
      },
      panel: (ctx, host, onChosen) =>
        colorPanel(host, {
          t,
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
      label: t("Borders"),
      tooltip: t("Borders (click the icon to reuse the last borders)"),
      defaultValue: "all",
      apply: (ctx, value) => applyBorderPreset(ctx, String(value)),
      panel: (ctx, host, onChosen) => borderPanel(host, ctx, onChosen)
    },
    {
      id: "number-format",
      group: "format",
      order: 100,
      type: "select",
      label: t("Number format"),
      tooltip: t("Number format"),
      options: NUMBER_FORMATS,
      // 同字体：fallback "" 表示"没设 = 常规"，而选区内格式不一致时仍然返回
      // undefined、下拉框显示空白，不会误报成"常规"。
      value: (ctx) => commonValue(ctx, "numFmt", ""),
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
      label: t("Percentage format"),
      tooltip: t("Percentage format (0%)"),
      run: (ctx) => ctx.sheet.setRangeStyle(ctx.range, { numFmt: "0%" })
    },
    {
      id: "format-thousands",
      group: "format",
      order: 120,
      type: "button",
      icon: icons.thousands,
      label: t("Thousands separator"),
      tooltip: t("Thousands separator (#,##0.00)"),
      run: (ctx) => ctx.sheet.setRangeStyle(ctx.range, { numFmt: "#,##0.00" })
    },
    {
      id: "add-decimal",
      group: "format",
      order: 130,
      type: "button",
      icon: icons.addDecimalIcon,
      label: t("Increase decimals"),
      tooltip: t("Increase decimals"),
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
      label: t("Decrease decimals"),
      tooltip: t("Decrease decimals"),
      run: (ctx) => {
        ctx.sheet.setRangeStyle(ctx.range, { numFmt: subtractDecimal(commonValue(ctx, "numFmt")) });
      }
    },

    // ---- align ----
    alignItem("align-left", "align", "left", t("Align left"), 10, icons.alignLeft),
    alignItem("align-center", "align", "center", t("Center"), 20, icons.alignCenter),
    alignItem("align-right", "align", "right", t("Align right"), 30, icons.alignRight),
    alignItem("align-top", "verticalAlign", "top", t("Align top"), 40, icons.alignTop),
    alignItem("align-middle", "verticalAlign", "middle", t("Align middle"), 50, icons.alignMiddle),
    alignItem("align-bottom", "verticalAlign", "bottom", t("Align bottom"), 60, icons.alignBottom),

    // ---- data ----
    {
      id: "sort-asc",
      group: "data",
      order: 10,
      type: "button",
      icon: icons.sortAsc,
      label: t("Sort ascending"),
      tooltip: t("Sort ascending by the current column"),
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
      label: t("Sort descending"),
      tooltip: t("Sort descending by the current column"),
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
      label: t("Filter"),
      panel: (ctx, host) => filterPanel(host, ctx)
    },
    {
      id: "find",
      group: "data",
      order: 40,
      // 对话框而不是面板：两个输入框 + 两个选项 + 五个按钮塞不进 160px 的小浮层，
      // 而且「下一个」要求开着框还能看见表格跳到哪一格（同 Excel 的 Ctrl+H）。
      type: "dialog",
      icon: icons.find,
      label: t("Find and replace"),
      tooltip: t("Find and replace"),
      dialog: (ctx, body, close) => findReplaceDialog(body, close, ctx)
    },
    {
      id: "data-validation",
      group: "data",
      order: 45,
      // 对话框：允许 / 数据 / 值 / 显示方式 / 错误警告，一屏放得下，也要能边开边看表格
      type: "dialog",
      icon: icons.dataValidation,
      label: t("Data validation"),
      tooltip: t("Set a data validation rule for the selection"),
      dialog: (ctx, body, close) => dataValidationDialog(body, close, ctx)
    },
    {
      id: "export-excel",
      group: "data",
      order: 50,
      type: "button",
      icon: icons.exportExcel,
      label: t("Export Excel"),
      tooltip: t("Export an Excel file (.xlsx)"),
      run: async (ctx) => {
        if (ctx.exportExcel) await ctx.exportExcel();
        else await exportSheetToExcel(ctx.sheet, {}, (row, column) => ctx.styleOf(row, column));
      }
    },

    // ---- structure ----
    {
      id: "merge",
      group: "structure",
      order: 10,
      type: "button",
      icon: icons.merge,
      label: t("Merge cells"),
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
      label: t("Unmerge cells"),
      run: (ctx) => ctx.sheet.unmergeCells(ctx.range)
    },
    {
      id: "insert-row",
      group: "structure",
      order: 30,
      type: "button",
      icon: icons.insertRow,
      label: t("Insert row"),
      tooltip: t("Insert a row above the selection"),
      run: (ctx) => ctx.sheet.insertRows(ctx.range.startRow, 1)
    },
    {
      id: "delete-row",
      group: "structure",
      order: 40,
      type: "button",
      icon: icons.deleteRow,
      label: t("Delete row"),
      run: (ctx) =>
        ctx.sheet.deleteRows(ctx.range.startRow, ctx.range.endRow - ctx.range.startRow + 1)
    },
    {
      id: "insert-column",
      group: "structure",
      order: 50,
      type: "button",
      icon: icons.insertColumn,
      label: t("Insert column"),
      tooltip: t("Insert a column to the left of the selection"),
      run: (ctx) => ctx.sheet.insertColumns(ctx.range.startColumn, [{}])
    },
    {
      id: "delete-column",
      group: "structure",
      order: 60,
      type: "button",
      icon: icons.deleteColumn,
      label: t("Delete column"),
      run: (ctx) =>
        ctx.sheet.deleteColumns(ctx.range.startColumn, ctx.range.endColumn - ctx.range.startColumn + 1)
    },
    {
      id: "hide-column",
      group: "structure",
      order: 70,
      type: "button",
      icon: icons.hideColumn,
      label: t("Hide column"),
      run: (ctx) => ctx.sheet.hideColumn(ctx.range.startColumn)
    },
    {
      id: "insert-image",
      group: "structure",
      order: 80,
      type: "button",
      icon: icons.image,
      label: t("Insert image"),
      tooltip: t("Upload an image into this cell"),
      // 图片是单元格级的，选区必须是单格——跟 merge 反过来。
      disabled: (ctx) =>
        ctx.range.startRow !== ctx.range.endRow || ctx.range.startColumn !== ctx.range.endColumn,
      run: (ctx) => pickImageFile(ctx)
    }
  ]);

}

/** `clearRangeStyle` 用的全键 patch，schema 之外也能用，所以单独导出一份。 */

export const CLEAR_ALL_STYLE_PATCH = Object.freeze(
  Object.fromEntries(STYLE_KEYS.map((key) => [key, null]))
) as { readonly [K in keyof CellStyle]: null };

export const defaultToolbarSchema: ToolbarSchema = createDefaultToolbarSchema();
