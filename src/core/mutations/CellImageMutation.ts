import type { CellRange, Patch } from "../types";
import type { SheetModel } from "../sheet/SheetModel";
import type { PreparedSheetMutation } from "../sheet/SheetTransaction";
import type { CellRecord } from "../sheet/types";
import type { Mutation, MutationPrepareContext, PreparedMutationParticipant } from "./types";
import { participantOf, type CellRecordWrite } from "./CellRecordWrite";

/**
 * 把一格的图片 URL 设成 `url`（`undefined` 清掉）。只支持单格——`Patch` 里
 * `cell-image` 没有 range 版本，见该类型的文档。
 *
 * 形状照抄 `CellStyleMutation`：prepare 期读影子记录、算出下一条记录、固化
 * forward/inverse——理由同样是"同一事务里先设值/样式、再设图片"不能互相覆盖
 * （必须读的是这次事务里已经生效的影子模型，不是活模型）。
 *
 * 不需要新的 `SheetMutation` 类型：`set-cell` 载的就是整条 `CellRecord`，`img`
 * 只是它的一个字段，整条替换天然带上——`apply.ts`/`invert.ts` 一行不用改。
 */
export class CellImageMutation implements Mutation {
  readonly type = "cell-image";
  readonly changes: readonly Patch[];
  readonly #sheet: SheetModel;
  readonly #row: number;
  readonly #column: number;
  readonly #url: string | undefined;

  constructor(sheet: SheetModel, row: number, column: number, url: string | undefined) {
    this.#sheet = sheet;
    this.#row = row;
    this.#column = column;
    this.#url = url;
    this.changes = Object.freeze([Object.freeze({ type: "cell-image" as const, row, column })]);
    Object.freeze(this);
  }

  prepare(context: MutationPrepareContext): PreparedMutationParticipant {
    const previous = context.readRecord(this.#sheet, this.#row, this.#column);
    const next = this.#nextRecord(previous);
    const writes: CellRecordWrite[] = next === previous ? [] : [{ row: this.#row, column: this.#column, record: next }];

    const staged: PreparedSheetMutation[] = [];
    try {
      for (const write of writes) {
        staged.push(context.prepareData(this.#sheet, {
          type: "set-cell",
          row: write.row,
          column: write.column,
          record: write.record,
          changed: "image"
        }));
      }
    } catch (error) {
      for (const prepared of staged) {
        prepared.dispose();
      }
      throw error;
    }

    return participantOf(this.#sheet, writes, staged, this.changes, "image");
  }

  /**
   * 算出这一格的新记录。返回**原对象本身**表示"没变"。
   *
   * 空格子设图片会造出一条只有 `img` 的记录——这是刻意的：跟"空格子刷背景色"
   * （`CellStyleMutation`）同一个道理，格子可以只有图片没有值。反过来，清掉 `img`
   * 时如果这一格除了 `img` 什么都没有，整条记录一并删掉，回到"未占用"。
   */
  #nextRecord(previous: CellRecord | undefined): CellRecord | undefined {
    if (this.#url === previous?.img) {
      return previous;
    }
    if (this.#url === undefined) {
      if (previous === undefined) {
        return undefined;
      }
      const { img: _dropped, ...rest } = previous;
      return rest.v === undefined &&
        rest.t === undefined &&
        rest.s === undefined &&
        rest.f === undefined &&
        rest.si === undefined
        ? undefined
        : (rest as CellRecord);
    }
    return { ...(previous ?? { v: undefined }), img: this.#url };
  }
}
