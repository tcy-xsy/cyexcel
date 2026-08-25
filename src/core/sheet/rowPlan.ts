import { CyExcelError, CyExcelErrorCode } from "../errors";

/**
 * 行计划里表示"这是一行新的空行"的标记。
 *
 * ADR §3.1 的实测结论：在按行分桶的稀疏矩阵上，insert/delete/move/reorder
 * 坍缩成同一个原语——一个容洞的排列。`plan[newRow]` 是该行的旧下标，
 * `ROW_PLAN_HOLE` 表示无来源；没有被任何 `plan[i]` 提到的旧行被丢弃。
 */
export const ROW_PLAN_HOLE = -1;

function rangeError(message: string): CyExcelError {
  return new CyExcelError(CyExcelErrorCode.InvalidRange, message);
}

/**
 * `occupiedExtent` 是最大占用行号 + 1（空矩阵为 0），不是坐标空间。
 *
 * **插入点落在空白区时排列不需要加长。** 坐标空间是 `maxOccupied + 1 + blankRows`
 * （spec §3.3），`start` 可以远在占用范围之外——但那种情况下没有任何占用行
 * 需要移动，`plan` 就该是恒等的。而长度 `occupiedExtent + count` 的排列在
 * `start > occupiedExtent` 时最大下标是 `occupiedExtent + count - 1`，永远够
 * 不到 `newRow >= start + count` 那一支，于是天然就是恒等——**正是要的东西**。
 * 铺到 `start` 会让分配量随调用方配的 `blankRows` 走（无上限，spec §3.3 只要求
 * ≥ 0），换不到任何行为差别。见本文件末尾"排列长度的界"。
 */
export function buildInsertRowsPlan(
  occupiedExtent: number,
  start: number,
  count: number
): Int32Array {
  const plan = new Int32Array(occupiedExtent + count);
  for (let newRow = 0; newRow < plan.length; newRow += 1) {
    if (newRow < start) {
      plan[newRow] = newRow;
    } else if (newRow < start + count) {
      plan[newRow] = ROW_PLAN_HOLE;
    } else {
      plan[newRow] = newRow - count;
    }
  }
  return plan;
}

export function buildDeleteRowsPlan(
  occupiedExtent: number,
  start: number,
  count: number
): Int32Array {
  // 删除范围可以整段或部分落在占用范围之外——那是坐标空间里合法的空白区
  // （spec §3.3）。只有落在占用范围**之内**的那一段才真的让行数变少；
  // 按整个 `count` 收缩排列会把范围之外的行一起丢掉，而且是静默的。
  const deletedWithin = Math.max(0, Math.min(count, occupiedExtent - start));
  const remaining = occupiedExtent - deletedWithin;
  const plan = new Int32Array(remaining);
  for (let newRow = 0; newRow < remaining; newRow += 1) {
    plan[newRow] = newRow < start ? newRow : newRow + count;
  }
  return plan;
}

export function buildMoveRowsPlan(
  occupiedExtent: number,
  start: number,
  count: number,
  target: number
): Int32Array {
  // 块或落点可以伸进空白区（spec §3.3）——排列必须覆盖到那里，否则空白
  // 间隙被压掉，被移动的数据会落到错误的下标上。
  //
  // 这里的加长与插入不同，**是必需的**：把数据搬到第 N 行之后，这张表的占用
  // 范围真的变成了 N + 1，排列长度只是如实反映这个新范围（见下方"排列长度的
  // 界"）。代价随**数据搬到了多远**走，不随 `blankRows` 走。
  const extent = Math.max(occupiedExtent, start + count, target);
  const moved: number[] = [];
  const rest: number[] = [];
  for (let row = 0; row < extent; row += 1) {
    if (row >= start && row < start + count) {
      moved.push(row);
    } else {
      rest.push(row);
    }
  }
  // target 是移动后该块应落在的新下标，先把块摘出去再按 target 插回
  const insertAt = target > start ? target - count : target;
  rest.splice(insertAt, 0, ...moved);
  return Int32Array.from(rest);
}

/**
 * ## 排列长度的界
 *
 * 三个 builder 共同遵守一条规则：**排列的长度是这张表操作后的占用范围**
 * （或它的一个紧上界），永远不随坐标空间走。
 *
 * | 操作 | 长度 | 由什么决定 |
 * | --- | --- | --- |
 * | 删除 | `occupiedExtent - deletedWithin` | 只会变短 |
 * | 插入 | `occupiedExtent + count` | 占用范围 + 插入量 |
 * | 移动 | `max(occupiedExtent, start + count, target)` | 数据被搬到了多远 |
 *
 * 这条规则是有代价的约束，不是巧合：`blankRows` 由调用方配置且**没有上限**
 * （spec §3.3 只要求 ≥ 0），而 `SheetModel` 的入口按坐标空间校验参数，所以
 * `start` / `target` 可以是很大的数。任何一处让长度跟着坐标空间走，都会让
 * 一次只搬一行的操作付出与 `blankRows` 成正比的分配——`applyRowPlan` 还要
 * 在这个长度上跑一趟 O(plan.length) 的循环。
 *
 * 移动那一档随 `target` 增长是**如实计价**：数据真的落到了那么远的地方，
 * 表的占用范围真的变那么大了。插入那一档曾经也这么写过，那是**白付**——
 * 插入点在空白区时没有任何占用行移动。见 `buildInsertRowsPlan` 的注释。
 */

/**
 * 计划必须是从旧行到新行的**单射**：同一个旧行被搬到两处会凭空复制数据。
 * 允许有洞（新空行），也允许旧行不被提到（那就是删除）。
 */
export function assertValidRowPlan(plan: Int32Array, occupiedExtent: number): void {
  const seen = new Set<number>();
  for (let newRow = 0; newRow < plan.length; newRow += 1) {
    const source = plan[newRow] as number;
    if (source === ROW_PLAN_HOLE) {
      continue;
    }
    if (!Number.isInteger(source) || source < 0 || source >= occupiedExtent) {
      throw rangeError(
        `row plan entry ${newRow} names source row ${source}, outside occupied extent ${occupiedExtent}`
      );
    }
    if (seen.has(source)) {
      throw rangeError(`row plan names source row ${source} more than once`);
    }
    seen.add(source);
  }
}
