import type { Patch } from "../types";
import type { CellRange } from "../types";
import type { SheetModel } from "../sheet/SheetModel";
import type { PreparedSheetMutation } from "../sheet/SheetTransaction";
import type { CellRecord } from "../sheet/types";
import type { Mutation, MutationPrepareContext, PreparedMutationParticipant } from "./types";

export interface CellRecordWrite {
  readonly row: number;
  readonly column: number;
  readonly record: CellRecord | undefined;
}

function boundsOf(writes: readonly CellRecordWrite[]): CellRange {
  let startRow = writes[0]!.row;
  let endRow = startRow;
  let startColumn = writes[0]!.column;
  let endColumn = startColumn;
  for (const write of writes) {
    startRow = Math.min(startRow, write.row);
    endRow = Math.max(endRow, write.row);
    startColumn = Math.min(startColumn, write.column);
    endColumn = Math.max(endColumn, write.column);
  }
  return Object.freeze({ startRow, startColumn, endRow, endColumn });
}

/**
 * 把一组算好的记录写回去。`changed` 走调用方给定的 Patch 类型（样式给
 * `"cell-style"`/`"range-style"`，图片给 `"cell-image"`），单格与多格各自的字面量
 * 补丁数组由 `patchesFor` 拼。
 *
 * forward / inverse 都是它——**History 里存的必须是固化好的记录，不能是"再合并一次"
 * 的指令**。重放一条合并语义的变更会重新读当时的状态，而 undo 栈里的每一步都是在
 * 不同状态上重放的，第二次算出来的结果可以和第一次不同。`RowHeightMutation` 把
 * inverse 做成"具体的高度值"而不是"再算一次"，理由一模一样。
 */
export class CellRecordWriteMutation implements Mutation {
  readonly type = "cell-record-write";
  readonly changes: readonly Patch[];
  readonly #sheet: SheetModel;
  readonly #writes: readonly CellRecordWrite[];
  readonly #changedTag: "style" | "image";

  constructor(
    sheet: SheetModel,
    writes: readonly CellRecordWrite[],
    changes: readonly Patch[],
    changedTag: "style" | "image"
  ) {
    this.#sheet = sheet;
    this.#writes = Object.freeze([...writes]);
    this.changes = Object.freeze([...changes]);
    this.#changedTag = changedTag;
    Object.freeze(this);
  }

  prepare(context: MutationPrepareContext): PreparedMutationParticipant {
    const staged: PreparedSheetMutation[] = [];
    try {
      for (const write of this.#writes) {
        staged.push(context.prepareData(this.#sheet, {
          type: "set-cell",
          row: write.row,
          column: write.column,
          record: write.record,
          changed: this.#changedTag
        }));
      }
    } catch (error) {
      for (const prepared of staged) {
        prepared.dispose();
      }
      throw error;
    }
    return participantOf(this.#sheet, this.#writes, staged, this.changes, this.#changedTag);
  }
}

/**
 * 把一组已暂存的单格写入包成一个参与者：commit/rollback/dispose 逐条转发，
 * forward / inverse 用各自的记录固化成两条 `CellRecordWriteMutation`。
 */
export function participantOf(
  sheet: SheetModel,
  writes: readonly CellRecordWrite[],
  staged: readonly PreparedSheetMutation[],
  changes: readonly Patch[],
  changedTag: "style" | "image"
): PreparedMutationParticipant {
  const inverseWrites = staged.map((prepared, index) => {
    const inverse = prepared.inverse;
    // 单格写入的逆必然还是单格写入——`invertSheetMutation` 对 `set-cell` 的逆是
    // 「写回原记录」。这里断言住，免得以后逆的形状变了还在悄悄按 set-cell 读。
    if (inverse.type !== "set-cell") {
      throw new Error(`Expected set-cell inverse, got ${inverse.type}`);
    }
    return {
      row: writes[index]!.row,
      column: writes[index]!.column,
      record: inverse.record
    };
  });

  return {
    hasEffect: staged.length > 0,
    forward: new CellRecordWriteMutation(sheet, writes, changes, changedTag),
    inverse: new CellRecordWriteMutation(sheet, inverseWrites, changes, changedTag),
    estimatedBytes:
      staged.length === 0
        ? 0
        : 64 + staged.reduce((total, prepared) => total + prepared.estimatedBytes, 0),
    commit: () => {
      for (const prepared of staged) {
        prepared.commit();
      }
    },
    rollback: () => {
      for (const prepared of staged) {
        prepared.rollback();
      }
    },
    dispose: () => {
      for (const prepared of staged) {
        prepared.dispose();
      }
    }
  };
}

export function boundsOfWrites(writes: readonly CellRecordWrite[]): CellRange {
  return boundsOf(writes);
}
