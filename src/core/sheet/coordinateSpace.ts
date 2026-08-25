import { CyExcelError, CyExcelErrorCode } from "../errors";

/**
 * 坐标空间是派生量：`maxOccupied + 1 + blank`。
 *
 * `rowCount` / `columnCount` 不进快照——存下来会与 `blank*` 打架（数据增删后
 * 该跟着变还是不变？没有不别扭的答案）。见 spec §3.3。
 *
 * @param maxOccupied 最大占用下标（0 基）；`undefined` 表示什么都没占用
 * @param blank 数据之后保留的空白数量
 */
export function resolveCoordinateSpace(maxOccupied: number | undefined, blank: number): number {
  return maxOccupied === undefined ? blank : maxOccupied + 1 + blank;
}

/** `blank*` 必须是非负整数（spec §3.3 硬约束）。 */
export function assertValidBlankCount(value: number, label: string): void {
  if (!Number.isInteger(value) || value < 0) {
    throw new CyExcelError(
      CyExcelErrorCode.InvalidOptions,
      `${label} must be a non-negative integer, received ${String(value)}`
    );
  }
}
