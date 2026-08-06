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

        return [];
      });
    }
  };
}
