import { CyExcelError, CyExcelErrorCode } from "../errors";
import type { CellRange } from "../types";
import { assertValidBlankCount } from "./coordinateSpace";
import { CellMatrix } from "./CellMatrix";
import { assertCanonicalIndexKeys } from "./indexKey";
import type { CellRecord, ColumnMeta, RowMeta, SheetSnapshot, WorkbookSnapshot } from "./types";
import { SNAPSHOT_VERSION } from "./types";
import { assertResourceSheetId, assertValidResourceEntry } from "./resources";

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

/**
 * 校验一张 sheet 的 `cellData`/`rowData`/`columnData` 键（spec §3、§11 不变量 5）。
 *
 * 这些下标经 `JSON.parse` 读入后只是字符串键，`Number()` 强转对非法输入（非数字、
 * 负数、小数）不会报错而是产出 NaN 或垃圾值——继续用它索引矩阵会静默丢数据或产出
 * 负的 occupied 计数（spec §11 不变量 5 依赖的派生坐标空间算术因此被破坏）。
 * 规则本身在 `indexKey.ts`：本文件此前有一份自己的正则，`resources.ts` 有另一份，
 * 同一条规则写三遍而消息已经开始分叉。
 */
function assertValidSheetKeys(sheet: SheetSnapshot, label: string): void {
  assertCanonicalIndexKeys(sheet.cellData, `${label}.cellData row key`);
  for (const rowKey of Object.keys(sheet.cellData)) {
    const bucket = sheet.cellData[Number(rowKey)];
    if (bucket) {
      assertCanonicalIndexKeys(bucket, `${label}.cellData[${rowKey}] column key`);
    }
  }
  assertCanonicalIndexKeys(sheet.rowData, `${label}.rowData row key`);
  assertCanonicalIndexKeys(sheet.columnData, `${label}.columnData column key`);
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
    // 键与 `sheet.id` 不一致会静默改键：`sheetOrder` 与 `sheets` 的键都用外面这个
    // id，而 `WorkbookModel.toSnapshot()` 输出的是 `sheet.toSnapshot().id`，即里面
    // 那个。实测 { sheetOrder: ["alias"], sheets: { alias: { …, id: "renamed" } } }
    // 往返之后 sheetOrder 变成 ["renamed"]——一份快照被加载再保存就换了 id。
    if (sheet.id !== id) {
      throw new CyExcelError(
        CyExcelErrorCode.InvalidOptions,
        `sheets["${id}"] has id "${sheet.id}"; the key and the sheet's own id must agree`
      );
    }
    assertValidBlankCount(sheet.blankRows, `sheets.${id}.blankRows`);
    assertValidBlankCount(sheet.blankColumns, `sheets.${id}.blankColumns`);
    assertValidSheetKeys(sheet, `sheets.${id}`);
  }
  const seenResources = new Set<string>();
  value.resources.forEach((entry, index) => {
    assertValidResourceEntry(entry, `resources[${index}]`);
    // 形状校验之后再查引用完整性：sheetId 指向的 sheet 必须在 sheetOrder 里，
    // 否则下标失去所指就是静默错位（spec §7.2"Sheet 作用域"一段）。
    assertResourceSheetId(entry, `resources[${index}]`, value.sheetOrder);
    if (seenResources.has(entry.name)) {
      throw new CyExcelError(
        CyExcelErrorCode.InvalidOptions,
        `resources[${index}] repeats the name "${entry.name}"; names identify resources in the snapshot`
      );
    }
    seenResources.add(entry.name);
  });
}
