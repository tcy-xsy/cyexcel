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
  const remaining = Math.max(0, occupiedExtent - count);
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
  const moved: number[] = [];
  const rest: number[] = [];
  for (let row = 0; row < occupiedExtent; row += 1) {
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
