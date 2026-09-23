import type { CellRange } from "../types";
import type { Patch } from "../types";
import type { SheetModel } from "../sheet/SheetModel";
import type { PreparedSheetMutation } from "../sheet/SheetTransaction";
import type { CellRecord } from "../sheet/types";
import type { StyleRegistry } from "../style/StyleRegistry";
import { mergeStyle, type CellStylePatch } from "../style/styleKey";
import { assertValidNumFmt } from "../style/formatCellText";
import type { Mutation, MutationPrepareContext, PreparedMutationParticipant } from "./types";
import { participantOf, type CellRecordWrite } from "./CellRecordWrite";

/**
 * 给一片区域改样式（合并语义）。
 *
 * # 为什么是一个自定义 Mutation，而不是 planner 里直接 stage 一条 set-cell
 *
 * 样式是**合并**语义（`{bold:true}` 只改 bold），所以要先读到这一格现在的 `s`，
 * 才算得出新的 `s`。这个读必须发生在 prepare 期、读的必须是影子模型——同一个事务里
 * 先写值再改样式的两条变更，第二条若在 plan 期读活模型，就会拿到没被第一条更新过的
 * 旧记录，把刚写的值盖回去。`RowHeightMutation` 因为同一个理由长成同样的形状
 * （那边读 `RowMeta`，这边读 `CellRecord`）。
 *
 * # 为什么不新增 SheetMutation 类型
 *
 * `set-cell` 载的就是整条 `CellRecord`，而 `s` 是它的一个字段——整条替换天然把样式
 * 带上了。于是 `apply.ts` 与 `invert.ts` 一行都不用改，撤销重做直接成立：逆操作是
 * 写回旧 record，旧 record 带着旧的 `s`。新增一条 `set-cell-style` 只会让那两个
 * switch 各多一个做同样事情的分支。
 */
export class CellStyleMutation implements Mutation {
  readonly type = "cell-style";
  readonly changes: readonly Patch[];
  readonly #sheet: SheetModel;
  readonly #styles: StyleRegistry;
  readonly #range: CellRange;
  readonly #patch: CellStylePatch;

  constructor(sheet: SheetModel, styles: StyleRegistry, range: CellRange, patch: CellStylePatch) {
    // numFmt 在**构造期**校验，所以非法 pattern 根本不会让事务开始
    // （Step 24 退出条件第 5 条：设置时报错，不是等那一格滚进视口才发现）。
    if (typeof patch.numFmt === "string") {
      assertValidNumFmt(patch.numFmt);
    }
    this.#sheet = sheet;
    this.#styles = styles;
    this.#range = Object.freeze({ ...range });
    this.#patch = Object.freeze({ ...patch });
    const single =
      range.startRow === range.endRow && range.startColumn === range.endColumn;
    this.changes = Object.freeze([
      single
        ? Object.freeze({ type: "cell-style" as const, row: range.startRow, column: range.startColumn })
        : Object.freeze({ type: "range-style" as const, range: this.#range })
    ]);
    Object.freeze(this);
  }

  prepare(context: MutationPrepareContext): PreparedMutationParticipant {
    const writes: CellRecordWrite[] = [];
    for (let row = this.#range.startRow; row <= this.#range.endRow; row += 1) {
      for (let column = this.#range.startColumn; column <= this.#range.endColumn; column += 1) {
        const previous = context.readRecord(this.#sheet, row, column);
        const next = this.#nextRecord(previous);
        if (next !== previous) {
          // 样式没变的格子不发变更——省掉一条没有意义的撤销步骤，也省掉重绘。
          writes.push({ row, column, record: next });
        }
      }
    }

    const staged: PreparedSheetMutation[] = [];
    try {
      for (const write of writes) {
        staged.push(context.prepareData(this.#sheet, {
          type: "set-cell",
          row: write.row,
          column: write.column,
          record: write.record,
          changed: "style"
        }));
      }
    } catch (error) {
      for (const prepared of staged) {
        prepared.dispose();
      }
      throw error;
    }

    return participantOf(this.#sheet, writes, staged, this.changes, "style");
  }

  /**
   * 算出这一格的新记录。返回**原对象本身**表示"没变"。
   *
   * 空格子设样式会造出一条只有 `s` 的记录——这是刻意的：Excel 里给空格子刷背景色，
   * 格子就有了样式却仍然没有值。反过来，清掉最后一个样式键时，如果这一格除了 `s`
   * 什么都没有，整条记录一并删掉，回到"未占用"（不变量 1：未占用的格子不存空记录）。
   */
  #nextRecord(previous: CellRecord | undefined): CellRecord | undefined {
    const previousStyle = previous?.s === undefined ? undefined : this.#styles.get(previous.s);
    const merged = mergeStyle(previousStyle, this.#patch);
    const nextId = merged === undefined ? undefined : this.#styles.intern(merged);
    if (nextId === previous?.s) {
      return previous;
    }
    if (nextId === undefined) {
      if (previous === undefined) {
        return undefined;
      }
      const { s: _dropped, ...rest } = previous;
      return rest.v === undefined &&
        rest.t === undefined &&
        rest.f === undefined &&
        rest.si === undefined &&
        rest.img === undefined
        ? undefined
        : (rest as CellRecord);
    }
    return { ...(previous ?? { v: undefined }), s: nextId };
  }
}
