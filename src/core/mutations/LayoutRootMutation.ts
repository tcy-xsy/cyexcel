import { CyExcelError, CyExcelErrorCode } from "../errors";
import { invertRowOrder, type LayoutEngine, type LayoutRoot, type RowHeightProvenance } from "../layout";
import type { SheetModel } from "../sheet/SheetModel";
import type { PreparedSheetMutation } from "../sheet/SheetTransaction";
import type { RowMeta } from "../sheet/types";
import type { Patch } from "../types";
import { rowHeightOf, rowHeightProvenanceOf, withRowHeight } from "./rowHeightMeta";
import type { Mutation, MutationPrepareContext, PreparedMutationParticipant } from "./types";

export type LayoutRootAction =
  | {
      readonly type: "set-row-height";
      readonly row: number;
      readonly height: number;
      readonly provenance: RowHeightProvenance;
    }
  | { readonly type: "set-column-width"; readonly column: number; readonly width: number }
  | { readonly type: "hide-column"; readonly column: number }
  | { readonly type: "show-column"; readonly column: number }
  /**
   * A whole-axis row permutation (`order[newPosition] = oldPosition`).
   *
   * It belongs here rather than behind an `AxisChangeDescriptor` because a
   * permutation has no `start`/`count`/`target` to put in one, and because a
   * reorder is exactly what this union is for: a root replacement whose
   * inverse is another action of the same kind (here, the inverse
   * permutation) rather than something that has to be read back off the
   * previous root.
   */
  | { readonly type: "reorder-rows"; readonly order: Int32Array };

function patchFor(action: LayoutRootAction, root: LayoutRoot): Patch {
  switch (action.type) {
    case "set-row-height":
      return {
        type: "row-layout",
        rows: Object.freeze([action.row]),
        range: {
          startRow: action.row,
          startColumn: 0,
          endRow: root.rowCount - 1,
          endColumn: Math.max(0, root.columnCount - 1)
        }
      };
    case "reorder-rows":
      return { type: "row-reorder", order: action.order };
    case "set-column-width":
    case "hide-column":
    case "show-column":
      return { type: "column-layout", column: action.column };
  }
}

/** History budgeting: only a permutation carries a payload worth counting. */
function estimatedBytesFor(action: LayoutRootAction): number {
  return action.type === "reorder-rows" ? 96 + action.order.length * 4 : 96;
}

function prepareAction(layout: LayoutEngine, root: LayoutRoot, action: LayoutRootAction): LayoutRoot {
  switch (action.type) {
    case "set-row-height":
      // 只喂高度：来源写进 `rowData[row].hp`，那半边由 prepare() 自己发出去
      // （设计稿 §4.3）。
      return layout.prepareSetRowHeights(new Map([[action.row, action.height]]), root);
    case "set-column-width":
      return layout.prepareSetColumnWidth(action.column, action.width, root);
    case "hide-column":
      return layout.prepareHideColumn(action.column, root);
    case "show-column":
      return layout.prepareShowColumn(action.column, root);
    case "reorder-rows":
      return layout.prepareReorderRows(action.order, root);
  }
}

function inverseAction(
  root: LayoutRoot,
  action: LayoutRootAction,
  previousRowMeta: RowMeta | undefined,
  defaultRowHeight: number
): LayoutRootAction {
  switch (action.type) {
    case "set-row-height":
      // 高度和来源都从**权威**（影子里的 `rowData` 条目）读，不再从根上读：根只是
      // 索引，而且它已经不存来源了（ADR 0004 阶段 4B T8）。
      return {
        type: "set-row-height",
        row: action.row,
        height: rowHeightOf(previousRowMeta, defaultRowHeight),
        provenance: rowHeightProvenanceOf(previousRowMeta)
      };
    case "set-column-width":
      return {
        type: "set-column-width",
        column: action.column,
        width: root.columnAxis.getColumnRestoreWidth(action.column)
      };
    case "hide-column":
      return root.columnAxis.isColumnHidden(action.column)
        ? { type: "hide-column", column: action.column }
        : { type: "show-column", column: action.column };
    case "show-column":
      return root.columnAxis.isColumnHidden(action.column)
        ? { type: "hide-column", column: action.column }
        : { type: "show-column", column: action.column };
    case "reorder-rows":
      // The only action whose inverse needs nothing from `root`: undoing a
      // permutation is applying the inverse permutation.
      return { type: "reorder-rows", order: invertRowOrder(action.order) };
  }
}

/**
 * Declarative Layout root replacement whose commit/rollback are reference swaps only.
 *
 * `set-row-height` 是唯一一个还要写模型的动作：行高的权威在 `rowData`，索引在
 * `RowMetrics`，两半必须同进同退，所以它在 prepare 期顺带暂存一条 `set-row-meta`
 * （设计稿 §4.2/§4.3）。其余动作仍然只是换一个根的引用——列宽/隐藏要等 T8b。
 */
export class LayoutRootMutation implements Mutation {
  readonly type = "layout-root";
  readonly changes: readonly Patch[];
  readonly #layout: LayoutEngine;
  readonly #sheet: SheetModel;
  readonly #action: LayoutRootAction;

  constructor(
    layout: LayoutEngine,
    sheet: SheetModel,
    action: LayoutRootAction,
    changes?: readonly Patch[]
  ) {
    this.#layout = layout;
    this.#sheet = sheet;
    this.#action = Object.freeze({ ...action });
    this.changes = Object.freeze([...(changes ?? [patchFor(action, layout.root)])]);
  }

  prepare(context: MutationPrepareContext): PreparedMutationParticipant {
    const action = this.#action;
    const previousRoot = context.roots.get(this.#layout, () => this.#layout.root);
    const previousRowMeta = action.type === "set-row-height"
      ? context.readRowMeta(this.#sheet, action.row)
      : undefined;

    // 高度和来源都没变就一条也不暂存：否则一次「设成它已经是的那个高度」会变成
    // 一条有效果的变更，进历史、发补丁。这与旧实现「根没换就没有效果」等价，只是
    // 判据换成了权威。
    const rowMetaChanges = action.type === "set-row-height" && (
      rowHeightOf(previousRowMeta, this.#layout.rowHeight) !== action.height ||
      rowHeightProvenanceOf(previousRowMeta) !== action.provenance
    );

    const staged: PreparedSheetMutation[] = [];
    let nextRoot: LayoutRoot;
    try {
      if (action.type === "set-row-height" && rowMetaChanges) {
        staged.push(context.prepareData(this.#sheet, {
          type: "set-row-meta",
          row: action.row,
          meta: withRowHeight(previousRowMeta, action.height, action.provenance)
        }));
      }
      nextRoot = prepareAction(this.#layout, previousRoot, action);
    } catch (error) {
      for (const prepared of staged) {
        prepared.dispose();
      }
      throw error;
    }
    context.roots.set(this.#layout, nextRoot);

    const changes = Object.freeze([patchFor(action, nextRoot)]);
    const forward = new LayoutRootMutation(this.#layout, this.#sheet, action, changes);
    // No `changes` for the inverse: `changes` describes the *forward* patch.
    // Every other action's patch happens to be direction-independent, and
    // prepare() recomputes `changes` before the inverse's own is ever read,
    // so passing it was harmless — but a reorder's patch carries the
    // permutation itself, so handing the inverse the forward one loads a gun
    // for any future code that reads `mutation.changes` on an unprepared
    // inverse and gets the wrong direction for a sort undo. The
    // constructor's `patchFor(action, layout.root)` fallback derives it from
    // the inverse action instead.
    const inverse = new LayoutRootMutation(
      this.#layout,
      this.#sheet,
      inverseAction(previousRoot, action, previousRowMeta, this.#layout.rowHeight)
    );
    let commitStarted = false;
    let rolledBack = false;

    return {
      // 一次只改来源的行高写不动根，但它确实改了权威里的 `hp`。
      hasEffect: nextRoot !== previousRoot || staged.length > 0,
      forward,
      inverse,
      estimatedBytes: estimatedBytesFor(action),
      commit: () => {
        if (this.#layout.root !== previousRoot) {
          throw new CyExcelError(CyExcelErrorCode.MutationConflict, "Layout root changed after mutation prepare");
        }
        commitStarted = true;
        for (const prepared of staged) {
          prepared.commit();
        }
        this.#layout.replaceRoot(nextRoot);
      },
      rollback: () => {
        if (!commitStarted || rolledBack) {
          return;
        }
        rolledBack = true;
        this.#layout.replaceRoot(previousRoot);
        for (let index = staged.length - 1; index >= 0; index -= 1) {
          staged[index]!.rollback();
        }
      },
      dispose: () => {
        for (const prepared of staged) {
          prepared.dispose();
        }
      }
    };
  }
}
