import { CyExcelError, CyExcelErrorCode } from "../errors";
import type { CellRange } from "../types";
import type { CellValue } from "./types";

/**
 * 合并区几何搬移的纯函数集合（设计稿 §0.2/§2/§3）。这个文件只做几何/数组运算，
 * 不碰 `SheetModel`、`TableStore` 或任何 Command 层类型——`readonly CellRange[]`
 * 进，`readonly CellRange[]`（或表示结果的判别联合）出。行/列 insert、delete、
 * move 三种结构操作对一组合并区的影响，以及 `mergeCells(range)` 的重叠决策，
 * 全部落在这里，供 T2（SheetModel 集成）、`invert.ts`、Command 层、渲染/交互
 * 层复用。
 */

/** 单轴上的闭区间 `[start, end]`，`start <= end`。 */
type AxisInterval = readonly [start: number, end: number];

function axisInterval(range: CellRange, axis: "row" | "column"): AxisInterval {
  return axis === "row" ? [range.startRow, range.endRow] : [range.startColumn, range.endColumn];
}

function withAxisInterval(range: CellRange, axis: "row" | "column", interval: AxisInterval): CellRange {
  const [start, end] = interval;
  return axis === "row"
    ? { ...range, startRow: start, endRow: end }
    : { ...range, startColumn: start, endColumn: end };
}

export function findMergeAt(
  merges: readonly CellRange[],
  row: number,
  column: number
): CellRange | undefined {
  return merges.find(
    (merge) =>
      row >= merge.startRow &&
      row <= merge.endRow &&
      column >= merge.startColumn &&
      column <= merge.endColumn
  );
}

export function rangesIntersect(a: CellRange, b: CellRange): boolean {
  return (
    a.startRow <= b.endRow &&
    a.endRow >= b.startRow &&
    a.startColumn <= b.endColumn &&
    a.endColumn >= b.startColumn
  );
}

export function rangeContains(outer: CellRange, inner: CellRange): boolean {
  return (
    outer.startRow <= inner.startRow &&
    outer.endRow >= inner.endRow &&
    outer.startColumn <= inner.startColumn &&
    outer.endColumn >= inner.endColumn
  );
}

/**
 * 返回与 `range` 相交的所有合并区。给渲染/DOM 宿主这类"每帧都要问一遍"的调用方
 * 用——按范围问一次，比逐格调 `findMergeAt` 便宜得多：合并区为空或都不相交时
 * 这里直接是一次 `Array.filter` 扫过（成本只随合并区数量而不是格子数量增长），
 * 不会像逐格点查询那样把"零合并区"也摊成一次 O(可见格子数) 的固定开销。
 */
export function mergesIntersectingRange(
  merges: readonly CellRange[],
  range: CellRange
): readonly CellRange[] {
  return merges.filter((merge) => rangesIntersect(merge, range));
}

/**
 * 把 `range` 按它相交到的合并区整体外扩，直到不动点——外扩本身可能新覆盖到另一
 * 条此前不相交的合并区，所以要收敛到没有合并区再"跨"到 `expanded` 外面为止。
 * 专供填充柄拖拽预览框使用：预览范围可以是任意 N×M，边缘格子可能恰好落在某个
 * 合并区上，需要把预览框整体撑大到包住它，不能只处理单格选区（那是
 * `SelectionLayerRenderer#expandForMerge` 的既定 YAGNI 范围，这里不重复它）。
 */
export function expandRangeForMerges(
  merges: readonly CellRange[],
  range: CellRange
): CellRange {
  let expanded = range;
  let grew = true;
  while (grew) {
    grew = false;
    for (const merge of merges) {
      if (rangesIntersect(merge, expanded) && !rangeContains(expanded, merge)) {
        expanded = {
          startRow: Math.min(expanded.startRow, merge.startRow),
          startColumn: Math.min(expanded.startColumn, merge.startColumn),
          endRow: Math.max(expanded.endRow, merge.endRow),
          endColumn: Math.max(expanded.endColumn, merge.endColumn)
        };
        grew = true;
      }
    }
  }
  return expanded;
}

/**
 * `mergeCells(range)` 重叠校验：只拒绝**部分**相交。完全包含关系（含 `candidate`
 * 与 `existing` 相等）不算违规相交——那种情形要放过，交给 `resolveMergeCells` 去
 * 吸收，不能在这一步就抛错。
 */
export function assertMergesDoNotOverlap(merges: readonly CellRange[], candidate: CellRange): void {
  for (const existing of merges) {
    if (rangeContains(candidate, existing) || rangeContains(existing, candidate)) {
      continue; // 完全包含关系（含相等）交给 resolveMergeCells 处理，不算违规
    }
    if (rangesIntersect(existing, candidate)) {
      throw new CyExcelError(
        CyExcelErrorCode.MergeOverlapsExistingMerge,
        "New merge range partially overlaps an existing merged range"
      );
    }
  }
}

/**
 * `mergeCells(range)` 的完整重叠决策（设计稿 §0.2/§2），是 mergeCells 的唯一
 * 重叠入口：与 `existing` 中任何一条**部分**相交（非完全包含关系）则抛
 * `MergeOverlapsExistingMerge`；**完全包住**某条已有合并区，则那条从结果中移除
 * （被新 `range` 取代，即"吸收"）；不相交的原样保留；最终返回值追加 `range` 本身。
 */
export function resolveMergeCells(
  existing: readonly CellRange[],
  range: CellRange
): readonly CellRange[] {
  assertMergesDoNotOverlap(existing, range);
  const survivors = existing.filter((merge) => !rangeContains(range, merge));
  return [...survivors, range];
}

/**
 * `mergeCells(range)` 清空合并区用的取值网格：整张全是 `undefined`（"清空这一格"）。
 * 锚点不由这张网格保护——`stageMergeClear` 把锚点排除在写入范围之外。
 */
export function buildMergeClearGrid(range: CellRange): readonly (readonly (CellValue | undefined)[])[] {
  const rows = range.endRow - range.startRow + 1;
  const columns = range.endColumn - range.startColumn + 1;
  const grid: (CellValue | undefined)[][] = [];
  for (let row = 0; row < rows; row += 1) {
    const line: (CellValue | undefined)[] = [];
    for (let column = 0; column < columns; column += 1) {
      // 全部是 undefined（"这一格清空"）。`null` 曾经用来区分"清空"与"跳过锚点"，
      // 但 `CellValue` 收紧之后 null 不再是合法值；而调用方（stageMergeClear）本来
      // 就把锚点排除在写入范围之外，所以这里不需要再表达"跳过"。
      line.push(undefined);
    }
    grid.push(line);
  }
  return grid;
}

/**
 * 插入 `insert(start, count)` 对单轴区间 `[mStart, mEnd]` 的搬移（语义与
 * `buildInsertRowsPlan` 一致：新内容出现在原下标 `start`，原 `start` 及之后整体
 * 后移）。三路互斥，不留"落在边界算哪边"的二义分支：
 * - `start <= mStart`：插入点在合并区起点或之前（相等时视为"之前"），整体平移，
 *   不撑大——否则插入点恰好等于 `mStart` 时会把新插入的整行/列也划进合并区。
 * - `start > mEnd`：插入点在合并区结束之后，不受影响。
 * - 否则（`mStart < start <= mEnd`）：插入点严格落在合并区内部，撑大合并区。
 */
function relocateIntervalForInsert(
  mStart: number,
  mEnd: number,
  start: number,
  count: number
): AxisInterval {
  if (start <= mStart) {
    return [mStart + count, mEnd + count];
  }
  if (start > mEnd) {
    return [mStart, mEnd];
  }
  return [mStart, mEnd + count];
}

/**
 * 删除 `delete(start, count)`（删除区间 `[delStart, delEnd] = [start, start+count-1]`）
 * 对单轴区间 `[mStart, mEnd]` 的搬移。四路互斥，`null` 表示删除区间**全覆盖**了
 * 该合并区（该区间应从结果里整体移除）：
 * - `delEnd < mStart`：删除区间整体在合并区之前，整体平移。
 * - `delStart > mEnd`：删除区间整体在合并区之后，不受影响。
 * - `delStart <= mStart && delEnd >= mEnd`：全覆盖，合并区消失。
 * - 否则：部分重叠（删上沿/删下沿/删中段三种子情形之一），收缩、闭合空洞。
 */
function relocateIntervalForDelete(
  mStart: number,
  mEnd: number,
  start: number,
  count: number
): AxisInterval | null {
  const delStart = start;
  const delEnd = start + count - 1;
  if (delEnd < mStart) {
    return [mStart - count, mEnd - count];
  }
  if (delStart > mEnd) {
    return [mStart, mEnd];
  }
  if (delStart <= mStart && delEnd >= mEnd) {
    return null;
  }
  const newStart = mStart < delStart ? mStart : delStart;
  const newEnd = mEnd <= delEnd ? delStart - 1 : mEnd - count;
  return [newStart, newEnd];
}

export function relocateMergesForInsert(
  merges: readonly CellRange[],
  axis: "row" | "column",
  start: number,
  count: number
): readonly CellRange[] {
  return merges.map((merge) => {
    const [mStart, mEnd] = axisInterval(merge, axis);
    return withAxisInterval(merge, axis, relocateIntervalForInsert(mStart, mEnd, start, count));
  });
}

export function relocateMergesForDelete(
  merges: readonly CellRange[],
  axis: "row" | "column",
  start: number,
  count: number
): readonly CellRange[] {
  const result: CellRange[] = [];
  for (const merge of merges) {
    const [mStart, mEnd] = axisInterval(merge, axis);
    const relocated = relocateIntervalForDelete(mStart, mEnd, start, count);
    if (relocated !== null) {
      result.push(withAxisInterval(merge, axis, relocated));
    }
  }
  return result;
}

export type MergeRelocateResult =
  | { readonly ok: true; readonly merges: readonly CellRange[] }
  | { readonly ok: false; readonly reason: "straddles-move-boundary" };

/**
 * 移动 `move(start, count, target)`：移动块 `[start, start+count-1]`，落点用
 * `invert.ts`/`rowPlan.ts`/`columnShift.ts` 已有的同一口径
 * `insertAt = target > start ? target - count : target`（摘块之后的插入位）。
 * 单轴上三路互斥：
 * - 合并区**整体落在块内**：跟着块走，`delta = insertAt - start` 平移。
 * - 合并区**与块相交但不完全落在块内**（跨越边界）：拒绝，`ok: false`。
 * - 合并区**完全在块外**：复合两步——先把块当作一次删除摘掉
 *   （`relocateIntervalForDelete`）得到中间坐标，再把块当作插回
 *   （`relocateIntervalForInsert`）得到最终坐标。这与 `buildMoveRowsPlan` 自己
 *   "摘出再插入"是同一个组合，不另起第三套公式，保证与既有的行/列搬移同构。
 *
 * **"完全在块外"这一支还有第四种结局，同样必须拒绝。** 摘除之后的插入
 * （`relocateIntervalForInsert`）自己也有三路：插入点在合并区起点或之前——
 * 平移；插入点在合并区结束之后——不受影响；插入点**严格落在合并区内部**——
 * 撑大合并区。前两种是合并区的**等距搬移**（长度不变），移动与它的逆（同一个
 * 复合两步，参数对称）互为精确逆。第三种不是等距的：块被"塞进"了这条外部合并
 * 区里，合并区被撑大，块本身此后落在这条（已经变大的）合并区内部——那正是
 * "跨越移动块边界"的定义，逆向搬回去时会被本函数自己拒绝（"合并区与块相交但不
 * 完全落在块内"），undo 因此在 commit 之前就已经不可达，`CommandManager.undo()`
 * 却已经先跑了 `#replay(entry.inverse)`（撤销条目在这一步之前不会从栈里摘除），
 * 于是这一条 undo 记录永久卡死撤销栈——这是撤销时才炸的错误，必须在正向
 * `move` 这一步就拒绝，不能等它把不可逆的状态造出来了才发现。
 *
 * 因此"完全在块外"这一支不能对 `relocateIntervalForInsert` 的结果照单全收：
 * 先看插入点是否严格落在摘除后的区间内部（`afterDelete[0] < insertAt <=
 * afterDelete[1]`，与 `relocateIntervalForInsert` 判"撑大"用的是同一个条件），
 * 是就当作跨越边界处理，`ok: false`；否则才真正调用 `relocateIntervalForInsert`
 * 取那个平移或不变的结果。
 */
export function relocateMergesForMove(
  merges: readonly CellRange[],
  axis: "row" | "column",
  start: number,
  count: number,
  target: number
): MergeRelocateResult {
  const blockStart = start;
  const blockEnd = start + count - 1;
  const insertAt = target > start ? target - count : target;
  const result: CellRange[] = [];
  for (const merge of merges) {
    const [mStart, mEnd] = axisInterval(merge, axis);
    const containedInBlock = mStart >= blockStart && mEnd <= blockEnd;
    const outsideBlock = mEnd < blockStart || mStart > blockEnd;
    if (containedInBlock) {
      const delta = insertAt - start;
      result.push(withAxisInterval(merge, axis, [mStart + delta, mEnd + delta]));
      continue;
    }
    if (outsideBlock) {
      const afterDelete = relocateIntervalForDelete(mStart, mEnd, start, count);
      if (afterDelete === null) {
        // 不可达：outsideBlock 与"删除区间全覆盖 [mStart, mEnd]"互斥
        // （全覆盖要求 blockStart <= mStart && blockEnd >= mEnd，那正是
        // containedInBlock 的条件），两者不会同时成立。
        throw new Error(
          "unreachable: a merge classified as outside the move block cannot be fully covered by it"
        );
      }
      const [afterDeleteStart, afterDeleteEnd] = afterDelete;
      const wouldExpand = insertAt > afterDeleteStart && insertAt <= afterDeleteEnd;
      if (wouldExpand) {
        // 插入点严格落在这条外部合并区（摘除后的坐标）内部：块被塞进了它里面，
        // 撑大而非平移，逆向搬回去必然撞上"跨越移动块边界"——见本函数文档注释
        // 倒数第二段。拒绝在这里发生，正向 move 就不会造出不可逆的状态。
        return { ok: false, reason: "straddles-move-boundary" };
      }
      const afterInsert = relocateIntervalForInsert(afterDeleteStart, afterDeleteEnd, insertAt, count);
      result.push(withAxisInterval(merge, axis, afterInsert));
      continue;
    }
    return { ok: false, reason: "straddles-move-boundary" };
  }
  return { ok: true, merges: result };
}
