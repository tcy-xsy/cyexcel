import type { RowHeightProvenance } from "../../src/core/layout";
import type { RowHeightProvenanceView } from "../../src/core/measurement";
import type { SheetModel } from "../../src/core/sheet/SheetModel";

/**
 * 行高与来源的**权威**读法。
 *
 * ADR 0004 阶段 4B T8 之后 `LayoutEngine` 不再存来源，`getRowHeightProvenance`
 * 也随之删掉；行高的权威是 `rowData[row].h`/`.hp`，布局层的 `RowMetrics` 只是它
 * 的稠密索引（设计稿 §4.1/§4.2）。测试里想问"这一行现在多高、谁定的"，问的就是
 * 这两个函数。
 */
export function rowHeightProvenanceOf(sheet: SheetModel, row: number): RowHeightProvenance {
  return sheet.getRowMeta(row)?.hp ?? "default";
}

export function modelRowHeightOf(sheet: SheetModel, row: number, defaultRowHeight: number): number {
  return sheet.getRowMeta(row)?.h ?? defaultRowHeight;
}

/**
 * 测量服务要的那个来源视图，由一张显式的表拼出来。
 *
 * 真身是 `CyExcelCore` 里读 `rowData` 的那个（`autoRows()` 只走稀疏表里真实存在的
 * 条目），但测量的用例关心的是"给定这些行是 auto / manual 会怎样"，不该为此建一张
 * 表——所以这里留一个同形状的替身。
 */
export function provenanceViewOf(
  entries: ReadonlyMap<number, Exclude<RowHeightProvenance, "default">>
): RowHeightProvenanceView {
  return {
    provenanceOf: (row) => entries.get(row) ?? "default",
    autoRows: () => [...entries]
      .filter(([, provenance]) => provenance === "auto")
      .map(([row]) => row)
  };
}
