import { defaultTranslator, type Translator } from "../locale/index.js";
import type { CyExcelPlugin, PluginContext } from "../plugins/index.js";
import type { ContextMenuItem, SheetMenuAction } from "./types.js";

type Action = Extract<ContextMenuItem, { type: "action" }>;
const separator = (id: string): ContextMenuItem => ({ type: "separator", id: `cyexcel:${id}` });
const submenu = (id: string, label: string, children: ContextMenuItem[]): ContextMenuItem =>
  ({ type: "submenu", id: `cyexcel:${id}`, label, children });

/** Optional core plugin; the Sheet and Workbook presets install it by default. */
export function createBasicContextMenuPlugin(translator?: Translator): CyExcelPlugin {
  return {
    name: "cyexcel:basic-context-menu",
    setup(context: PluginContext) {
      const t = translator ?? context.t ?? defaultTranslator;
      return context.contextMenu.registerProvider({ name: "cyexcel:basic-context-menu" }, menu => {
        if (menu.target === "blank" || menu.target === "corner") return [];
        if (menu.target === "row-header" && menu.row === undefined) return [];
        if (menu.target === "column-header" && menu.column === undefined) return [];
        const range = menu.selection.range;
        const column = menu.column ?? range.startColumn;
        const multiple = menu.selection.ranges.length > 1;
        const state = context.sheetMenu?.getState(range, column);
        const structureDisabled = multiple || state?.structureEnabled === false;
        const structureDescription = multiple ? t("Select a single contiguous range first.") : state?.structureEnabled === false
          ? t("Row and column changes are unavailable on sheets containing formulas.") : undefined;
        const action = (id: string, label: string, run: Action["run"], disabled = false, description?: string): Action => ({
          type: "action", id: `cyexcel:${id}`, label, run, disabled,
          ...(description === undefined ? {} : { description })
        });
        const sheetAction = (id: SheetMenuAction, label: string, disabled = false): Action => action(id, label,
          (_menu, signal) => context.sheetMenu?.run(id,
            id === "filter-value" && menu.row !== undefined
              ? { startRow: menu.row, endRow: menu.row, startColumn: column, endColumn: column } : range,
            column, signal), multiple || disabled,
          multiple ? t("Select a single contiguous range first.") : undefined);
        const rows = range.endRow - range.startRow + 1;
        const columns = range.endColumn - range.startColumn + 1;
        const rowActions: ContextMenuItem[] = [
          action("insert-row-above", rows > 1 ? t("Insert {0} rows above", rows) : t("Insert row above"), () => context.insertRows(range.startRow, rows), structureDisabled, structureDescription),
          action("insert-row-below", rows > 1 ? t("Insert {0} rows below", rows) : t("Insert row below"), () => context.insertRows(range.endRow + 1, rows), structureDisabled, structureDescription),
          action("delete-row", state?.hasFilter ? t("Delete selected visible rows") : rows > 1 ? t("Delete {0} rows", rows) : t("Delete row"), (_menu, signal) => {
            if (state?.hasFilter) return context.sheetMenu?.run("delete-visible-rows", range, column, signal);
            context.deleteRows(range.startRow, rows);
          }, structureDisabled, structureDescription)
        ];
        const columnActions: ContextMenuItem[] = [
          action("insert-column-left", columns > 1 ? t("Insert {0} columns left", columns) : t("Insert column left"), () => context.insertColumns(range.startColumn, Array.from({ length: columns }, () => ({}))), structureDisabled, structureDescription),
          action("insert-column-right", columns > 1 ? t("Insert {0} columns right", columns) : t("Insert column right"), () => context.insertColumns(range.endColumn + 1, Array.from({ length: columns }, () => ({}))), structureDisabled, structureDescription),
          action("delete-column", columns > 1 ? t("Delete {0} columns", columns) : t("Delete column"), () => context.deleteColumns(range.startColumn, columns), structureDisabled, structureDescription)
        ];
        const mod = typeof navigator !== "undefined" && /Mac|iPhone|iPad/.test(navigator.platform) ? "⌘" : "Ctrl+";
        const items: ContextMenuItem[] = [];
        if (state) items.push({ ...sheetAction("cut", t("Cut"), state.canCut === false), shortcut: `${mod}X` });
        items.push(
          { ...action("copy", t("Copy"), async (_menu, signal) => {
            const result = await context.copy({ range, signal });
            if (result.status === "oversized") throw new Error(t("The selection is too large. Select a smaller range to copy."));
          }, multiple), shortcut: `${mod}C` },
          { ...action("paste", t("Paste"), (_menu, signal) => context.paste({ range, signal }), multiple), shortcut: `${mod}V` }
        );
        if (state) items.push(separator("clear-separator"),
          { ...sheetAction("clear-contents", t("Clear contents")), shortcut: "Delete" },
          sheetAction("clear-formats", t("Clear formatting")));
        items.push(separator("structure-separator"));
        if (menu.target === "row-header") items.push(...rowActions);
        else if (menu.target === "column-header") items.push(...columnActions,
          action("hide-column", columns > 1 ? t("Hide {0} columns", columns) : t("Hide column"), (_menu, signal) => {
            if (context.sheetMenu) return context.sheetMenu.run("hide-columns", range, column, signal);
            for (let index = range.startColumn; index <= range.endColumn; index++) context.hideColumn(index);
          }, multiple || state?.canHideColumns === false));
        else items.push(submenu("insert", t("Insert"), [rowActions[0]!, rowActions[1]!, columnActions[0]!, columnActions[1]!]),
          submenu("delete", t("Delete"), [rowActions[2]!, columnActions[2]!]));
        if (state) {
          if (menu.target === "column-header") items.push(sheetAction("show-columns", t("Unhide all columns"), !state.hasHiddenColumns));
          if (menu.target === "cell") items.push(separator("merge-separator"),
            sheetAction("merge", t("Merge cells"), !state.canMerge), sheetAction("unmerge", t("Unmerge cells"), !state.hasMerge));
          items.push(separator("format-separator"), submenu("format", t("Format cells"), [
            { ...sheetAction("bold", t("Bold")), checked: state.bold },
            { ...sheetAction("italic", t("Italic")), checked: state.italic },
            separator("alignment-separator"), sheetAction("align-left", t("Align left")),
            sheetAction("align-center", t("Align center")), sheetAction("align-right", t("Align right"))
          ]));
          if (menu.target !== "row-header") items.push(submenu("sort-filter", t("Sort and filter"), [
            { ...sheetAction("sort-asc", t("Sort sheet ascending"), !state.canSort || !state.structureEnabled), ...(structureDescription === undefined ? {} : { description: structureDescription }) },
            { ...sheetAction("sort-desc", t("Sort sheet descending"), !state.canSort || !state.structureEnabled), ...(structureDescription === undefined ? {} : { description: structureDescription }) },
            separator("filter-separator"),
            sheetAction("filter-value", t("Filter by this cell's value"), menu.target !== "cell" || !state.canFilter),
            sheetAction("clear-filter", t("Clear filter"), !state.hasFilter)
          ]));
        }
        return items;
      });
    }
  };
}
