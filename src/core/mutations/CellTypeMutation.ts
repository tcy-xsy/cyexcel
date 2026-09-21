import type { CellRange, Patch } from "../types";
import type { SheetModel } from "../sheet/SheetModel";
import type { PreparedSheetMutation } from "../sheet/SheetTransaction";
import type { CellRecord } from "../sheet/types";
import type { Mutation, MutationPrepareContext, PreparedMutationParticipant } from "./types";
import { CellRecordWriteMutation } from "./CellStyleMutation";

/**
 * 给一片区域设格级语义类型（`CellRecord.t`），`null` 表示回到列的默认类型。
 *
 * `t` 这个字段从 ADR 0004 起就在 schema 里，但一直没有任何代码读它——列类型决定
 * 一切。接上它是为了让"这一列是文本，但这一格是日期"成为可能，与 Excel 的
 * "类型跟着格子走"靠近一步；**但不做格式驱动类型的自动推断**（设日期 numFmt 不会
 * 顺手改类型），那对结构化表格的使用者太意外。
 *
 * 与 `CellStyleMutation` 同形：prepare 期读影子记录，写回整条 record。
 */
export class CellTypeMutation implements Mutation {
  readonly type = "cell-type";
  readonly changes: readonly Patch[];
  readonly #sheet: SheetModel;
  readonly #range: CellRange;
  readonly #typeName: string | null;

  constructor(sheet: SheetModel, range: CellRange, typeName: string | null) {
    this.#sheet = sheet;
    this.#range = Object.freeze({ ...range });
    this.#typeName = typeName;
    const single = range.startRow === range.endRow && range.startColumn === range.endColumn;
    this.changes = Object.freeze([
      single
        ? Object.freeze({ type: "cell-value" as const, row: range.startRow, column: range.startColumn })
        : Object.freeze({ type: "range-value" as const, range: this.#range })
    ]);
    Object.freeze(this);
  }

  prepare(context: MutationPrepareContext): PreparedMutationParticipant {
    const writes: { row: number; column: number; record: CellRecord | undefined }[] = [];
    for (let row = this.#range.startRow; row <= this.#range.endRow; row += 1) {
      for (let column = this.#range.startColumn; column <= this.#range.endColumn; column += 1) {
        const previous = context.readRecord(this.#sheet, row, column);
        const next = this.#nextRecord(previous);
        if (next !== previous) {
          writes.push({ row, column, record: next });
        }
      }
    }

    const staged: PreparedSheetMutation[] = [];
    try {
      for (const write of writes) {
        staged.push(
          context.prepareData(this.#sheet, {
            type: "set-cell",
            row: write.row,
            column: write.column,
            record: write.record
          })
        );
      }
    } catch (error) {
      for (const prepared of staged) {
        prepared.dispose();
      }
      throw error;
    }

    const inverseWrites = staged.map((prepared, index) => {
      const inverse = prepared.inverse;
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
      forward: new CellRecordWriteMutation(this.#sheet, writes, this.changes),
      inverse: new CellRecordWriteMutation(this.#sheet, inverseWrites, this.changes),
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

  /** 返回原对象本身表示"没变"。清掉类型后若整条记录空了，记录一并删除。 */
  #nextRecord(previous: CellRecord | undefined): CellRecord | undefined {
    if (this.#typeName === null) {
      if (previous?.t === undefined) {
        return previous;
      }
      const { t: _dropped, ...rest } = previous;
      return rest.v === undefined && rest.s === undefined && rest.f === undefined && rest.si === undefined
        ? undefined
        : (rest as CellRecord);
    }
    if (previous?.t === this.#typeName) {
      return previous;
    }
    return { ...(previous ?? { v: undefined }), t: this.#typeName };
  }
}
