import type { ColumnDefinition } from "../cells/types";
import type { LayoutRootAction } from "../mutations/LayoutRootMutation";
import type { Mutation } from "../mutations/types";
import type { CellRange, ColumnFieldMapping } from "../types";
import type {
  RowHeightCommandProvenance,
  SetRowHeightUpdate
} from "./SetRowHeightsCommand";

export interface DataCommandMutationFactory {
  createSetCell(row: number, column: number, value: unknown): Mutation;
  createSetRange(range: CellRange, values: readonly (readonly unknown[])[]): Mutation;
  createInsertRows(start: number, rows: readonly (readonly unknown[])[]): Mutation;
  createDeleteRows(start: number, count: number): Mutation;
  createMoveRows(start: number, count: number, target: number): Mutation;
  createInsertColumns(start: number, columns: readonly ColumnFieldMapping[]): Mutation;
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
