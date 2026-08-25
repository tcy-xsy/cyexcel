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

/** `occupiedExtent` 是最大占用行号 + 1（空矩阵为 0），不是坐标空间。 */
export function buildInsertRowsPlan(
  occupiedExtent: number,
  start: number,
  count: number
): Int32Array {
  // 插入点可以落在空白区——坐标空间是 `maxOccupied + 1 + blankRows`
  // （spec §3.3），比占用范围大得多。排列必须一直铺到落点，否则它压根没
  // 描述被插入的那一段。
  const extent = Math.max(occupiedExtent, start);
  const plan = new Int32Array(extent + count);
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
