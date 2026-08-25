import { CyExcelError, CyExcelErrorCode } from "../errors";
import type { CellRange } from "../types";
import { assertValidBlankCount } from "./coordinateSpace";
import { CellMatrix } from "./CellMatrix";
import type { CellRecord, ColumnMeta, RowMeta, SheetSnapshot, WorkbookSnapshot } from "./types";
import { SNAPSHOT_VERSION } from "./types";

export interface SerializeInput {
  readonly id: string;
  readonly name: string;
  readonly blankRows: number;
  readonly blankColumns: number;
  readonly defaultRowHeight: number;
  readonly defaultColumnWidth: number;
  readonly freeze?: { readonly rows: number; readonly columns: number };
  readonly matrix: CellMatrix;
  readonly rowData: Readonly<Record<number, RowMeta>>;
  readonly columnData: Readonly<Record<number, ColumnMeta>>;
  readonly mergeData: readonly CellRange[];
}

export interface DeserializeOutput {
  readonly id: string;
  readonly name: string;
  readonly blankRows: number;
  readonly blankColumns: number;
  readonly defaultRowHeight: number;
  readonly defaultColumnWidth: number;
  readonly freeze?: { readonly rows: number; readonly columns: number };
  readonly matrix: CellMatrix;
  readonly rowData: Readonly<Record<number, RowMeta>>;
  readonly columnData: Readonly<Record<number, ColumnMeta>>;
  readonly mergeData: readonly CellRange[];
}

export function serializeSheet(input: SerializeInput): SheetSnapshot {
  const cellData: Record<number, Record<number, CellRecord>> = {};
  for (const [row, bucket] of input.matrix.rows()) {
    // 行桶原样拷一层，避免快照与活矩阵共享可变引用
    cellData[row] = { ...bucket };
  }
  return {
    id: input.id,
    name: input.name,
    blankRows: input.blankRows,
    blankColumns: input.blankColumns,
    defaultRowHeight: input.defaultRowHeight,
    defaultColumnWidth: input.defaultColumnWidth,
    ...(input.freeze !== undefined ? { freeze: input.freeze } : {}),
    cellData,
    rowData: { ...input.rowData },
    columnData: { ...input.columnData },
    mergeData: [...input.mergeData]
  };
}

export function deserializeSheet(snapshot: SheetSnapshot): DeserializeOutput {
  const matrix = new CellMatrix();
  for (const rowKey of Object.keys(snapshot.cellData)) {
    const row = Number(rowKey);
    const bucket = snapshot.cellData[row];
    if (!bucket) {
      continue;
    }
    for (const columnKey of Object.keys(bucket)) {
      const column = Number(columnKey);
      const record = bucket[column];
      if (record !== undefined) {
        matrix.set(row, column, record);
      }
    }
  }
  return {
    id: snapshot.id,
    name: snapshot.name,
    blankRows: snapshot.blankRows,
    blankColumns: snapshot.blankColumns,
    defaultRowHeight: snapshot.defaultRowHeight,
    defaultColumnWidth: snapshot.defaultColumnWidth,
    ...(snapshot.freeze !== undefined ? { freeze: snapshot.freeze } : {}),
    matrix,
    rowData: { ...snapshot.rowData },
    columnData: { ...snapshot.columnData },
    mergeData: [...snapshot.mergeData]
  };
}

export function assertValidWorkbookSnapshot(value: WorkbookSnapshot): void {
  if (value.version !== SNAPSHOT_VERSION) {
    throw new CyExcelError(
      CyExcelErrorCode.InvalidOptions,
      `Unsupported snapshot version ${String(value.version)}, expected ${SNAPSHOT_VERSION}`
    );
  }
  // spec §11 不变量 6：多 Sheet 在 M3 之前不放开
  if (value.sheetOrder.length !== 1) {
    throw new CyExcelError(
      CyExcelErrorCode.InvalidOptions,
      `Exactly one sheet is supported for now, received ${value.sheetOrder.length}`
    );
  }
  for (const id of value.sheetOrder) {
    const sheet = value.sheets[id];
    if (!sheet) {
      throw new CyExcelError(
        CyExcelErrorCode.InvalidOptions,
        `sheetOrder references "${id}" which is not present in sheets`
      );
    }
    assertValidBlankCount(sheet.blankRows, `sheets.${id}.blankRows`);
    assertValidBlankCount(sheet.blankColumns, `sheets.${id}.blankColumns`);
  }
}
