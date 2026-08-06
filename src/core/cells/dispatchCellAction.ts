import { SetCellCommand } from "../commands";
import type { TableStore } from "../data";
import type { Command } from "../types";
import type { CellAction } from "./types";

export type CellActionDispatch =
  | { readonly kind: "command"; readonly command: Command }
  | { readonly kind: "start-edit" };

/**
 * Translates a CanvasCellRenderer.hitTest() result into something the
 * caller can act on — Step 8's 4th pre-written test ("hitTest 返回
 * CellAction，再转换为 Command"). Deliberately store-agnostic beyond
 * constructing the Command: it never calls store.setCell()/executes
 * anything itself, so the caller still decides whether/how to run it
 * (through CommandManager, inside a transaction, etc). "start-edit" is
 * returned as a sentinel rather than calling EditorManager directly — this
 * function has no EditorManager reference, keeping it independent of real
 * DOM/pointer events for Node testing.
 */
export function dispatchCellAction(
  action: CellAction,
  store: TableStore,
  row: number,
  column: number
): CellActionDispatch {
  switch (action.type) {
    case "set-value":
      return { kind: "command", command: new SetCellCommand(store, row, column, action.value) };
    case "start-edit":
      return { kind: "start-edit" };
  }
}
