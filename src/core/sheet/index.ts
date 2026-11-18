export type {
  CellRecord,
  ColumnMeta,
  ResourceEntry,
  RowHeightSource,
  RowMeta,
  SheetSnapshot,
  WorkbookSnapshot
} from "./types.js";
export { DEFAULT_BLANK_COLUMNS, DEFAULT_BLANK_ROWS, SNAPSHOT_VERSION } from "./types.js";
export { assertValidBlankCount, resolveCoordinateSpace } from "./coordinateSpace.js";
export { CellMatrix } from "./CellMatrix.js";
export {
  assertValidWorkbookSnapshot,
  deserializeSheet,
  serializeSheet,
  type DeserializeOutput,
  type SerializeInput
} from "./serialize.js";
export { SheetModel, type SheetModelOptions } from "./SheetModel.js";
export {
  ROW_PLAN_HOLE,
  assertValidRowPlan,
  buildDeleteRowsPlan,
  buildInsertRowsPlan,
  buildMoveRowsPlan
} from "./rowPlan.js";
export { remapByColumnShift, remapByRowPlan } from "./indexShift.js";
export {
  buildDeleteColumnsShift,
  buildInsertColumnsShift,
  buildMoveColumnsShift,
  type ColumnShift
} from "./columnShift.js";
export { IndexedTable, type IndexAxis } from "./IndexedTable.js";
export { IndexedTableRegistry } from "./IndexedTableRegistry.js";
export {
  assertResourceSheetId,
  assertValidResourceEntry,
  resourceToTable,
  tableToResource
} from "./resources.js";
export { WorkbookModel, type WorkbookModelOptions } from "./WorkbookModel.js";
export {
  assertCanonicalIndexKeys,
  assertIndexIsNonNegativeInteger,
  isCanonicalIndexKey
} from "./indexKey.js";
export {
  isStructuralSheetMutation,
  type ColumnCells,
  type RowCells,
  type SheetMutation,
  type SideTableEntries
} from "./mutation.js";
export { invertSheetMutation } from "./invert.js";
export {
  applySheetMutation,
  assertRestoreTablesAreWritable,
  estimateSheetMutationBytes
} from "./apply.js";
export { SheetTransaction, type PreparedSheetMutation } from "./SheetTransaction.js";
export { createSheetPlanningResources } from "./planning.js";
export { changesForSheetMutation, type SheetAxisSpace } from "./patches.js";
export {
  matrixToWorkbookSnapshot,
  rowValuesToCells,
  rowsToWorkbookSnapshot,
  valueToRecord,
  type ImportRow,
  type MatrixToWorkbookSnapshotOptions,
  type RowsToWorkbookSnapshotOptions
} from "./import.js";
