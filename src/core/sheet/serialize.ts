import { CyExcelError, CyExcelErrorCode } from "../errors";
import type { CellRange } from "../types";
import { assertValidBlankCount } from "./coordinateSpace";
import { CellMatrix } from "./CellMatrix";
import type { CellRecord, ColumnMeta, RowMeta, SheetSnapshot, WorkbookSnapshot } from "./types";
import { SNAPSHOT_VERSION } from "./types";
import { assertValidResourceEntry } from "./resources";

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
    // 行桶拷一层新对象，但桶内的 CellRecord 仍按引用共享——不深拷贝格子。
    // 因此 CellRecord 必须被当作不可变对象（见 CellRecord 的文档注释）：
    // 就地修改一条记录会通过这份共享引用污染快照或活矩阵中的另一份。
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
  // 快照是公开 API 的入口（spec §1），deserializeSheet 因此不能信任调用方已经
  // 校验过——即便 assertValidWorkbookSnapshot 从未被调用（例如 SheetModel.fromSnapshot
  // 就直接走到这里），格式错误的键也必须在这里被拒绝，而不是被 Number() 强转
  // 成 NaN/负数后悄悄丢行或污染坐标空间。
  const label = `sheets.${snapshot.id}`;
  assertValidSheetKeys(snapshot, label);
  const matrix = new CellMatrix();
  for (const rowKey of Object.keys(snapshot.cellData)) {
    const row = Number(rowKey);
    const bucket = snapshot.cellData[row];
    if (!bucket) {
      // 校验已经保证 rowKey 是规范形式的非负整数字符串，因此这里查不到桶
      // 说明校验逻辑本身或快照在校验之后被篡改——是编程错误，必须炸出来，
      // 而不是像旧代码那样 continue 静默丢行。
      throw new CyExcelError(
        CyExcelErrorCode.InvalidOptions,
        `${label}.cellData is missing the bucket for row key "${rowKey}" despite passing key validation`
      );
    }
    for (const columnKey of Object.keys(bucket)) {
      const column = Number(columnKey);
      const record = bucket[column];
      if (record === undefined) {
        // 同上：columnKey 已通过校验，查不到记录说明状态在校验后被篡改。
        throw new CyExcelError(
          CyExcelErrorCode.InvalidOptions,
          `${label}.cellData[${rowKey}] is missing the record for column key "${columnKey}" despite passing key validation`
        );
      }
      matrix.set(row, column, record);
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

/** 规范形式的非负整数字符串："0" 或不带前导零的正整数。 */
const NON_NEGATIVE_INTEGER_KEY = /^(0|[1-9]\d*)$/;

/**
 * `cellData` / `rowData` / `columnData` 都按行号或列号做稀疏下标索引
 * （spec §3）。这些下标经 `JSON.parse` 读入后只是字符串键，`Number()` 强转
 * 对非法输入（非数字、负数、小数）不会报错而是产出 NaN 或垃圾值——继续用它
 * 索引矩阵会静默丢数据或产出负的 occupied 计数（spec §11 不变量 5 依赖的
 * 派生坐标空间算术因此被破坏）。所以键的合法性必须在这里显式校验。
 */
function assertValidIndexKeys(record: Readonly<Record<number, unknown>>, label: string): void {
  for (const key of Object.keys(record)) {
    if (!NON_NEGATIVE_INTEGER_KEY.test(key)) {
      throw new CyExcelError(
        CyExcelErrorCode.InvalidOptions,
        `${label} must be a non-negative integer key, received "${key}"`
      );
    }
  }
}

/** 校验一张 sheet 的 `cellData`/`rowData`/`columnData` 键（spec §3、§11 不变量 5）。 */
function assertValidSheetKeys(sheet: SheetSnapshot, label: string): void {
  assertValidIndexKeys(sheet.cellData, `${label}.cellData row key`);
  for (const rowKey of Object.keys(sheet.cellData)) {
    const bucket = sheet.cellData[Number(rowKey)];
    if (bucket) {
      assertValidIndexKeys(bucket, `${label}.cellData[${rowKey}] column key`);
    }
  }
  assertValidIndexKeys(sheet.rowData, `${label}.rowData row key`);
  assertValidIndexKeys(sheet.columnData, `${label}.columnData column key`);
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
    assertValidSheetKeys(sheet, `sheets.${id}`);
  }
  const seenResources = new Set<string>();
  value.resources.forEach((entry, index) => {
    assertValidResourceEntry(entry, `resources[${index}]`);
    if (seenResources.has(entry.name)) {
      throw new CyExcelError(
        CyExcelErrorCode.InvalidOptions,
        `resources[${index}] repeats the name "${entry.name}"; names identify resources in the snapshot`
      );
    }
    seenResources.add(entry.name);
  });
}
