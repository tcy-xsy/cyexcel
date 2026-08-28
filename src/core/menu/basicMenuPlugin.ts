import type { CyExcelPlugin, PluginContext } from "../plugins";

/**
 * The "official basic menu" doc/customization-and-hooks.md's §10 asks for —
 * deliberately an *optional* plugin, not baked into ContextMenuService or
 * CyExcelCore itself ("官方基础菜单作为可选插件，不把业务菜单写入 core").
 * A consumer wanting Copy/Paste in the right-click menu passes this into
 * `CyExcelCoreOptions.plugins`; nothing installs it automatically.
 */
export function createBasicContextMenuPlugin(): CyExcelPlugin {
  return {
    name: "cyexcel:basic-context-menu",
    setup(context: PluginContext) {
      return context.contextMenu.registerProvider({ name: "cyexcel:basic-context-menu", priority: 0 }, (menuContext) => {
        if (menuContext.target === "cell") {
          return [
            {
              type: "action",
              id: "cyexcel:copy",
              label: "复制",
              run: async () => {
                await context.copy();
              }
            },
            {
              type: "action",
              id: "cyexcel:paste",
              label: "粘贴",
              run: async () => {
                await context.paste();
              }
            }
          ];
        }

        // Row structural actions only make sense with a real row to anchor
        // to — same reasoning as the cell-only Copy/Paste above, just for
        // the row-header target (Step 12). These call the same public
        // insertRows()/deleteRows() every other caller uses, never the
        // Store directly (doc/customization-and-hooks.md §11.2's "不直接
        // 操作 Store" requirement).
        if (menuContext.target === "row-header" && menuContext.row !== undefined) {
          const row = menuContext.row;
          return [
            {
              type: "action",
              id: "cyexcel:insert-row-above",
              label: "在上方插入行",
              run: () => {
                context.insertRows(row);
              }
            },
            {
              type: "action",
              id: "cyexcel:insert-row-below",
              label: "在下方插入行",
              run: () => {
                context.insertRows(row + 1);
              }
            },
            {
              type: "action",
              id: "cyexcel:delete-row",
              label: "删除行",
              run: () => {
                context.deleteRows(row);
              }
            }
          ];
        }

        // Mirrors the row-header branch above, one axis over — same
        // public insertColumns()/deleteColumns()/hideColumn() path every
        // other caller uses. Unlike a new row (every cell defaults to
        // `undefined`), a new column needs an `id` to exist at all; there's
        // no "the user's own data" to derive one from here, so a
        // timestamp-based id is generated, same convention already used for
        // `ClipboardService`'s exported-file naming.
        if (menuContext.target === "column-header" && menuContext.column !== undefined) {
          const column = menuContext.column;
          // insertColumns()/deleteColumns()/hideColumn() all throw while
          // headerGroups is configured (grouped headers reference raw
          // column indices a structural change would invalidate) — disabled
          // here instead of left to throw, so the menu shows *why* the
          // action can't run instead of silently closing with only a
          // console.error (see doc/progress.md's Step 19 playground entry).
          const blocked = context.hasHeaderGroups();
          const disabledLabel = blocked ? "（已禁用：存在分组表头）" : "";
          return [
            {
              type: "action",
              id: "cyexcel:insert-column-left",
              label: `在左侧插入列${disabledLabel}`,
              disabled: blocked,
              run: () => {
                // 空定义就是一列：列没有身份可给（ADR 0004 §3.8.2），宽度、
                // 类型、表头字母全部走默认。
                context.insertColumns(column, [{}]);
              }
            },
            {
              type: "action",
              id: "cyexcel:insert-column-right",
              label: `在右侧插入列${disabledLabel}`,
              disabled: blocked,
              run: () => {
                context.insertColumns(column + 1, [{}]);
              }
            },
            {
              type: "action",
              id: "cyexcel:delete-column",
              label: `删除列${disabledLabel}`,
              disabled: blocked,
              run: () => {
                context.deleteColumns(column);
              }
            },
            {
              type: "action",
              id: "cyexcel:hide-column",
              label: `隐藏列${disabledLabel}`,
              disabled: blocked,
              run: () => {
                context.hideColumn(column);
              }
            }
          ];
        }

        return [];
      });
    }
  };
}
