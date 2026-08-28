export type {
  CellRecord,
  ColumnMeta,
  ResourceEntry,
  RowHeightSource,
  RowMeta,
  SheetSnapshot,
  WorkbookSnapshot
} from "./types";
export { DEFAULT_BLANK_COLUMNS, DEFAULT_BLANK_ROWS, SNAPSHOT_VERSION } from "./types";
export { assertValidBlankCount, resolveCoordinateSpace } from "./coordinateSpace";
export { CellMatrix } from "./CellMatrix";
export {
  assertValidWorkbookSnapshot,
  deserializeSheet,
  serializeSheet,
  type DeserializeOutput,
  type SerializeInput
} from "./serialize";
export { SheetModel, type SheetModelOptions } from "./SheetModel";
export {
  ROW_PLAN_HOLE,
  assertValidRowPlan,
  buildDeleteRowsPlan,
  buildInsertRowsPlan,
  buildMoveRowsPlan
} from "./rowPlan";
export { remapByColumnShift, remapByRowPlan } from "./indexShift";
export {
  buildDeleteColumnsShift,
  buildInsertColumnsShift,
  buildMoveColumnsShift,
  type ColumnShift
} from "./columnShift";
export { IndexedTable, type IndexAxis } from "./IndexedTable";
export { IndexedTableRegistry } from "./IndexedTableRegistry";
export {
  assertResourceSheetId,
  assertValidResourceEntry,
  resourceToTable,
  tableToResource
} from "./resources";
export { WorkbookModel, type WorkbookModelOptions } from "./WorkbookModel";
export {
  assertCanonicalIndexKeys,
  assertIndexIsNonNegativeInteger,
  isCanonicalIndexKey
} from "./indexKey";
export {
  isStructuralSheetMutation,
  type ColumnCells,
  type RowCells,
  type SheetMutation,
  type SideTableEntries
} from "./mutation";
export { invertSheetMutation } from "./invert";
export {
  applySheetMutation,
  assertRestoreTablesAreWritable,
  estimateSheetMutationBytes
} from "./apply";
export { SheetTransaction, type PreparedSheetMutation } from "./SheetTransaction";
export { changesForSheetMutation, type SheetAxisSpace } from "./patches";
export {
  rowValuesToCells,
  rowsToWorkbookSnapshot,
  valueToRecord,
  type ImportRow,
  type RowsToWorkbookSnapshotOptions
} from "./import";
