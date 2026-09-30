import type { ColumnDefinition } from "../cells/types";
import type { LayoutRootAction } from "../mutations/LayoutRootMutation";
import type { Mutation } from "../mutations/types";
import type { CellRange } from "../types";
import type { CellStylePatch } from "../style/styleKey";
import type { CellValue } from "../sheet/types";
import type { DataValidationRule } from "../sheet/validations";
import type {
  RowHeightCommandProvenance,
  SetRowHeightUpdate
} from "./SetRowHeightsCommand";

export interface DataCommandMutationFactory {
  createSetCell(row: number, column: number, value: CellValue | null | undefined): Mutation;
  createSetRange(range: CellRange, values: readonly (readonly (CellValue | null | undefined)[])[]): Mutation;
  /** 一片区域改样式（合并语义）。Step 24。 */
  createSetStyle(range: CellRange, patch: CellStylePatch): Mutation;
  /** 单格设置/清空图片 URL。与 `v`/`s` 独立——单元格级图片，不依赖列配置。 */
  createSetCellImage(row: number, column: number, url: string | undefined): Mutation;
  createSetMerges(merges: readonly CellRange[]): Mutation;
  /** 整体替换数据验证规则列表（区域几何与重叠决策已在 plan 期做完）。 */
  createSetValidations(validations: readonly DataValidationRule[]): Mutation;
  createInsertRows(start: number, rows: readonly (readonly (CellValue | null | undefined)[])[]): Mutation;
  createDeleteRows(start: number, count: number): Mutation;
  createMoveRows(start: number, count: number, target: number): Mutation;
  /** Whole-axis permutation, `order[newPosition] = oldPosition`. */
  createReorderRows(order: Int32Array): Mutation;
  /** 插列只要个数：列不再有身份（ADR 0004 §3.8.2），没有字段映射可带。 */
  createInsertColumns(start: number, count: number): Mutation;
  createDeleteColumns(start: number, count: number): Mutation;
  createMoveColumns(start: number, count: number, target: number): Mutation;
}

export interface StructureCommandMutationFactory {
  createInsertRows(start: number, count: number): Mutation;
  createDeleteRows(start: number, count: number): Mutation;
  createMoveRows(start: number, count: number, target: number): Mutation;
  createInsertColumns(start: number, columns: readonly ColumnDefinition[]): Mutation;
  createDeleteColumns(start: number, count: number): Mutation;
  createMoveColumns(start: number, count: number, target: number): Mutation;
}

export interface LayoutCommandMutationFactory {
  createLayoutMutation(action: LayoutRootAction): Mutation;
  createRowHeightMutation(
    updates: readonly SetRowHeightUpdate[],
    provenance: RowHeightCommandProvenance
  ): Mutation;
}

export interface StructuralCommandGuards {
  checkRows(change: {
    readonly operation: "insert" | "delete" | "move";
    readonly start: number;
    readonly count: number;
    readonly target?: number;
  }): void;
  checkColumns(change: {
    readonly operation: "insert" | "delete" | "move";
    readonly start: number;
    readonly count: number;
    readonly target?: number;
  }): void;
}
