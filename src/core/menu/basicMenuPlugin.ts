import type { CyExcelPlugin, PluginContext } from "../plugins";
import type { ContextMenuItem, SheetMenuAction } from "./types";

type Action = Extract<ContextMenuItem, { type: "action" }>;
const separator = (id: string): ContextMenuItem => ({ type: "separator", id: `cyexcel:${id}` });
const submenu = (id: string, label: string, children: ContextMenuItem[]): ContextMenuItem =>
  ({ type: "submenu", id: `cyexcel:${id}`, label, children });

/** Optional core plugin; the Sheet and Workbook presets install it by default. */
export function createBasicContextMenuPlugin(): CyExcelPlugin {
  return {
    name: "cyexcel:basic-context-menu",
    setup(context: PluginContext) {
      return context.contextMenu.registerProvider({ name: "cyexcel:basic-context-menu" }, menu => {
        if (menu.target === "blank" || menu.target === "corner") return [];
        if (menu.target === "row-header" && menu.row === undefined) return [];
        if (menu.target === "column-header" && menu.column === undefined) return [];
        const range = menu.selection.range;
        const column = menu.column ?? range.startColumn;
        const multiple = menu.selection.ranges.length > 1;
        const state = context.sheetMenu?.getState(range, column);
        const structureDisabled = multiple || state?.structureEnabled === false;
        const structureDescription = multiple ? "请先选择一个连续区域" : state?.structureEnabled === false
          ? "当前工作表含公式，暂不支持调整行列引用" : undefined;
        const action = (id: string, label: string, run: Action["run"], disabled = false, description?: string): Action => ({
          type: "action", id: `cyexcel:${id}`, label, run, disabled,
          ...(description === undefined ? {} : { description })
        });
        const sheetAction = (id: SheetMenuAction, label: string, disabled = false): Action => action(id, label,
          (_menu, signal) => context.sheetMenu?.run(id,
            id === "filter-value" && menu.row !== undefined
              ? { startRow: menu.row, endRow: menu.row, startColumn: column, endColumn: column } : range,
            column, signal), multiple || disabled,
          multiple ? "请先选择一个连续区域" : undefined);
        const rows = range.endRow - range.startRow + 1;
        const columns = range.endColumn - range.startColumn + 1;
        const rowActions: ContextMenuItem[] = [
          action("insert-row-above", rows > 1 ? `在上方插入 ${rows} 行` : "在上方插入行", () => context.insertRows(range.startRow, rows), structureDisabled, structureDescription),
          action("insert-row-below", rows > 1 ? `在下方插入 ${rows} 行` : "在下方插入行", () => context.insertRows(range.endRow + 1, rows), structureDisabled, structureDescription),
          action("delete-row", state?.hasFilter ? "删除选中的可见行" : rows > 1 ? `删除 ${rows} 行` : "删除行", (_menu, signal) => {
            if (state?.hasFilter) return context.sheetMenu?.run("delete-visible-rows", range, column, signal);
            context.deleteRows(range.startRow, rows);
          }, structureDisabled, structureDescription)
        ];
        const columnActions: ContextMenuItem[] = [
          action("insert-column-left", columns > 1 ? `在左侧插入 ${columns} 列` : "在左侧插入列", () => context.insertColumns(range.startColumn, Array.from({ length: columns }, () => ({}))), structureDisabled, structureDescription),
          action("insert-column-right", columns > 1 ? `在右侧插入 ${columns} 列` : "在右侧插入列", () => context.insertColumns(range.endColumn + 1, Array.from({ length: columns }, () => ({}))), structureDisabled, structureDescription),
          action("delete-column", columns > 1 ? `删除 ${columns} 列` : "删除列", () => context.deleteColumns(range.startColumn, columns), structureDisabled, structureDescription)
        ];
        const mod = typeof navigator !== "undefined" && /Mac|iPhone|iPad/.test(navigator.platform) ? "⌘" : "Ctrl+";
        const items: ContextMenuItem[] = [];
        if (state) items.push({ ...sheetAction("cut", "剪切", state.canCut === false), shortcut: `${mod}X` });
        items.push(
          { ...action("copy", "复制", async (_menu, signal) => {
            const result = await context.copy({ range, signal });
            if (result.status === "oversized") throw new Error("所选区域过大，请缩小选区后复制。");
          }, multiple), shortcut: `${mod}C` },
          { ...action("paste", "粘贴", (_menu, signal) => context.paste({ range, signal }), multiple), shortcut: `${mod}V` }
        );
        if (state) items.push(separator("clear-separator"),
          { ...sheetAction("clear-contents", "清除内容"), shortcut: "Delete" },
          sheetAction("clear-formats", "清除格式"));
        items.push(separator("structure-separator"));
        if (menu.target === "row-header") items.push(...rowActions);
        else if (menu.target === "column-header") items.push(...columnActions,
          action("hide-column", columns > 1 ? `隐藏 ${columns} 列` : "隐藏列", (_menu, signal) => {
            if (context.sheetMenu) return context.sheetMenu.run("hide-columns", range, column, signal);
            for (let index = range.startColumn; index <= range.endColumn; index++) context.hideColumn(index);
          }, multiple || state?.canHideColumns === false));
        else items.push(submenu("insert", "插入", [rowActions[0]!, rowActions[1]!, columnActions[0]!, columnActions[1]!]),
          submenu("delete", "删除", [rowActions[2]!, columnActions[2]!]));
        if (state) {
          if (menu.target === "column-header") items.push(sheetAction("show-columns", "取消隐藏所有列", !state.hasHiddenColumns));
          if (menu.target === "cell") items.push(separator("merge-separator"),
            sheetAction("merge", "合并单元格", !state.canMerge), sheetAction("unmerge", "取消合并", !state.hasMerge));
          items.push(separator("format-separator"), submenu("format", "单元格格式", [
            { ...sheetAction("bold", "加粗"), checked: state.bold },
            { ...sheetAction("italic", "斜体"), checked: state.italic },
            separator("alignment-separator"), sheetAction("align-left", "左对齐"),
            sheetAction("align-center", "居中对齐"), sheetAction("align-right", "右对齐")
          ]));
          if (menu.target !== "row-header") items.push(submenu("sort-filter", "排序和筛选", [
            { ...sheetAction("sort-asc", "升序排列（整表）", !state.canSort || !state.structureEnabled), ...(structureDescription === undefined ? {} : { description: structureDescription }) },
            { ...sheetAction("sort-desc", "降序排列（整表）", !state.canSort || !state.structureEnabled), ...(structureDescription === undefined ? {} : { description: structureDescription }) },
            separator("filter-separator"),
            sheetAction("filter-value", "按此单元格的值筛选", menu.target !== "cell" || !state.canFilter),
            sheetAction("clear-filter", "清除筛选", !state.hasFilter)
          ]));
        }
        return items;
      });
    }
  };
}
