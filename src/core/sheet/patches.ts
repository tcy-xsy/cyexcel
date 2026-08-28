import type { Patch } from "../types";
import type { SheetMutation } from "./mutation";

/**
 * 一次变更**之后**的坐标空间，`row-layout` 那条补丁的 `range` 要它来定下界。
 *
 * 为什么是"之后"而不是"之前"：这个 `range` 表达的是"从这一行往下的几何全都失效了"，
 * 而失效范围要盖住新状态里的每一行。仓库里两个既有的 `row-layout` 生产者都是这个口径
 * （`RowHeightMutation` 传的是 `nextRoot`、`LayoutRootMutation` 传的是施加后的 root），
 * 这里不另起一个。在 `SheetTransaction` 上取这两个数天然就是"之后"——`prepareMutation`
 * 返回时影子已经吃过这条变更了。
 */
export interface SheetAxisSpace {
  readonly rowCount: number;
  readonly columnCount: number;
}

/**
 * 一条 `SheetMutation` 对应哪几条 `Patch`（设计稿 §5 的映射表）。
 *
 * 它替代 `TableStore.changesForDataMutation`。**放在 `sheet/` 而不是留在 `TableStore`**，
 * 是因为往 `SheetMutation` 加一个成员会同时欠下三笔债——求逆、施加、这张映射表——
 * 放在同一个目录里，加成员时三处都在眼前；散在两个模块里，漏掉的那一处 TypeScript
 * 只在 `switch` 不穷尽时才拦得住，而"映射成一条错的补丁"它拦不住。
 *
 * ## 两处不是机械对应的口径
 *
 * 1. **`restore-*` 映射成 `insert`，不是自成一类。** 它们**确实**把行/列插回去，
 *    下游的 `remapIndex()` 必须照插入处理——否则撤销一次删除之后，选区与视口停在
 *    旧下标上，而模型已经整体下移。
 * 2. **`reorder-rows` 用它自己的 `row-reorder`，不近似成 `row-structure`。**
 *    理由是 `types.ts:134` 已经写下的那条：`row-structure` 的每一个消费者都按
 *    `start`/`count` 理解下标位移，而一个排列没有这个形状，硬塞进去会编译通过、
 *    然后静默地把下标重映射错。
 *
 * ## 为什么 `space` 是必填的
 *
 * 十三个成员里只有 `set-row-meta` 用得上它。做成可选的话，调用方漏传就会得到一条
 * `range` 不对的补丁——那是渲染层"该重画哪一段"的依据，错了是**画少了**，不报错。
 * 必填让编译器把这件事挡在前面。
 */
export function changesForSheetMutation(
  mutation: SheetMutation,
  space: SheetAxisSpace
): readonly Patch[] {
  switch (mutation.type) {
    case "set-cell":
      return [{ type: "cell-value", row: mutation.row, column: mutation.column }];

    case "set-range":
      return [{ type: "range-value", range: mutation.range }];

    case "insert-rows":
      return [
        { type: "row-structure", operation: "insert", start: mutation.start, count: mutation.rows.length }
      ];

    case "restore-rows":
      // 见上面口径 1：这是一次插入，不是一个新类别。
      return [
        { type: "row-structure", operation: "insert", start: mutation.start, count: mutation.rows.length }
      ];

    case "delete-rows":
      return [
        { type: "row-structure", operation: "delete", start: mutation.start, count: mutation.count }
      ];

    case "move-rows":
      return [
        {
          type: "row-structure",
          operation: "move",
          start: mutation.start,
          count: mutation.count,
          target: mutation.target
        }
      ];

    case "reorder-rows":
      // 见上面口径 2：带着 `order` 走，不折成 start/count。
      return [{ type: "row-reorder", order: mutation.order }];

    case "insert-columns":
      return [
        { type: "column-structure", operation: "insert", start: mutation.start, count: mutation.count }
      ];

    case "restore-columns":
      // 口径 1 在列轴上的同一条。
      return [
        { type: "column-structure", operation: "insert", start: mutation.start, count: mutation.count }
      ];

    case "delete-columns":
      return [
        { type: "column-structure", operation: "delete", start: mutation.start, count: mutation.count }
      ];

    case "move-columns":
      return [
        {
          type: "column-structure",
          operation: "move",
          start: mutation.start,
          count: mutation.count,
          target: mutation.target
        }
      ];

    case "set-row-meta":
      // `range` 从这一行一直铺到最后一行：改一行的高度会把它**下面每一行**的
      // 纵向偏移全部推走。`endColumn` 的 `Math.max(0, …)` 是空表的下界保护，与
      // `RowHeightMutation` / `LayoutRootMutation` 逐字一致。
      return [
        {
          type: "row-layout",
          rows: [mutation.row],
          range: {
            startRow: mutation.row,
            startColumn: 0,
            endRow: space.rowCount - 1,
            endColumn: Math.max(0, space.columnCount - 1)
          }
        }
      ];

    case "set-column-meta":
      // `column-layout` 这个变体本来就不带 `range`——列宽变化的失效范围由下游按
      // 整条列轴处理（`changes.ts` 的 `hasColumnLayout`）。这里不替它补一个。
      return [{ type: "column-layout", column: mutation.column }];
  }
}
