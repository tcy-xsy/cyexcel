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
