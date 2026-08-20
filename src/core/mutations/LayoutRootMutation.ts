import { CyExcelError, CyExcelErrorCode } from "../errors";
import { invertRowOrder, type LayoutEngine, type LayoutRoot, type RowHeightProvenance } from "../layout";
import type { Patch } from "../types";
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
      return layout.prepareSetRowHeights(new Map([[
        action.row,
        { height: action.height, provenance: action.provenance }
      ]]), root);
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

function inverseAction(root: LayoutRoot, action: LayoutRootAction): LayoutRootAction {
  switch (action.type) {
    case "set-row-height": {
      const height = root.rowMetrics.heightAt(action.row);
      const provenance = root.rowHeightProvenance.get(action.row) ?? "default";
      return { type: "set-row-height", row: action.row, height, provenance };
    }
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

/** Declarative Layout root replacement whose commit/rollback are reference swaps only. */
export class LayoutRootMutation implements Mutation {
  readonly type = "layout-root";
  readonly changes: readonly Patch[];
  readonly #layout: LayoutEngine;
  readonly #action: LayoutRootAction;

  constructor(layout: LayoutEngine, action: LayoutRootAction, changes?: readonly Patch[]) {
    this.#layout = layout;
    this.#action = Object.freeze({ ...action });
    this.changes = Object.freeze([...(changes ?? [patchFor(action, layout.root)])]);
  }

  prepare(context: MutationPrepareContext): PreparedMutationParticipant {
    const previousRoot = context.roots.get(this.#layout, () => this.#layout.root);
    const nextRoot = prepareAction(this.#layout, previousRoot, this.#action);
    context.roots.set(this.#layout, nextRoot);
    const changes = Object.freeze([patchFor(this.#action, nextRoot)]);
    const forward = new LayoutRootMutation(this.#layout, this.#action, changes);
    const inverse = new LayoutRootMutation(
      this.#layout,
      inverseAction(previousRoot, this.#action),
      changes
    );
    let commitStarted = false;
    let rolledBack = false;

    return {
      hasEffect: nextRoot !== previousRoot,
      forward,
      inverse,
      estimatedBytes: estimatedBytesFor(this.#action),
      commit: () => {
        if (this.#layout.root !== previousRoot) {
          throw new CyExcelError(CyExcelErrorCode.MutationConflict, "Layout root changed after mutation prepare");
        }
        commitStarted = true;
        this.#layout.replaceRoot(nextRoot);
      },
      rollback: () => {
        if (!commitStarted || rolledBack) {
          return;
        }
        rolledBack = true;
        this.#layout.replaceRoot(previousRoot);
      },
      dispose: () => {}
    };
  }
}
