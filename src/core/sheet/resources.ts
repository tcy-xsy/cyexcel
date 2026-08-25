import { CyExcelError, CyExcelErrorCode } from "../errors";
import { IndexedTable } from "./IndexedTable";
import { assertCanonicalIndexKeys } from "./indexKey";
import type { ResourceEntry } from "./types";

function invalid(message: string): CyExcelError {
  return new CyExcelError(CyExcelErrorCode.InvalidOptions, message);
}

/**
 * 校验一条插件资源（spec §7）。
 *
 * `indexedBy: "row" | "column"` 的资源，其 `data` **必须**是稀疏下标映射——
 * 核心要按下标重键它（§7.2），键不是规范的非负整数就无法重键，而 `Number()`
 * 对非法键不报错、产出 NaN 或垃圾值，会静默丢条目。`"none"` 是不透明 blob，
 * 核心不碰，因此 `data` 不作任何约束。
 *
 * 键值规则本身在 `indexKey.ts`（含符号键的说明）：本文件此前抄了一份自己的正则，
 * 而 `IndexedTable` 的公开构造函数根本没走到这条路——校验器造好了却从未到达容器。
 */
export function assertValidResourceEntry(entry: ResourceEntry, label: string): void {
  if (typeof entry.name !== "string" || entry.name.length === 0) {
    throw invalid(`${label}.name must be a non-empty string`);
  }
  if (entry.indexedBy !== "row" && entry.indexedBy !== "column" && entry.indexedBy !== "none") {
    throw invalid(
      `${label}.indexedBy must be "row", "column" or "none", received ${JSON.stringify(entry.indexedBy)}`
    );
  }
  if (entry.indexedBy === "none") {
    // "none" 是 workbook 级的不透明 blob，没有轴也就没有所属的 sheet
    // （types.ts 里 ResourceEntry.sheetId 的文档注释）。这是形状检查，不需要
    // 知道有哪些 sheet 存在，因此和下面的键校验一样留在本函数里，而不是挪去
    // assertResourceSheetId。
    if (entry.sheetId !== undefined) {
      throw invalid(`${label}.sheetId must not be set because indexedBy is "none"`);
    }
    return;
  }
  const data: unknown = entry.data;
  if (typeof data !== "object" || data === null || Array.isArray(data)) {
    throw invalid(
      `${label}.data must be a sparse index map because indexedBy is "${entry.indexedBy}"`
    );
  }
  assertCanonicalIndexKeys(data as Readonly<Record<number, unknown>>, `${label}.data key`);
}

/**
 * 校验一条资源的 `sheetId` 指向的 sheet 确实存在——给出了却指向未知 sheet，
 * 下标就失去了所指（spec §7.2"Sheet 作用域"一段）。
 *
 * **为什么这是一个独立的函数，而不是给 `assertValidResourceEntry` 加第三参数。**
 * 两者要的上下文不一样：`assertValidResourceEntry` 校验的是形状（名字、
 * `indexedBy`、`data` 的键、"none" 不得带 `sheetId`），这些不需要知道有哪些
 * sheet。而"`sheetId` 指向的 sheet 存不存在"必须有 sheet 列表才判得出来。
 * 本文件的 `resourceToTable`（下方）只把一条条目转成一张裸表，它**根本无从
 * 知道**有哪些 sheet 存在——那不是给 `sheetIds` 补一个默认值 `[]` 能解决的，
 * 补了默认值只会让它把任何带 `sheetId` 的条目都判成"指向未知 sheet"，
 * 那是它做不了这项检查，不是缺一个入参。因此调用点按"手上有没有 sheet
 * 列表"分成两类：`resourceToTable` 没有，不调；`assertValidWorkbookSnapshot`
 * 与 `WorkbookModel.addResource` 有，两处都在各自的 `assertValidResourceEntry`
 * 之后追加一次本函数的调用。
 */
export function assertResourceSheetId(
  entry: ResourceEntry,
  label: string,
  sheetIds: readonly string[]
): void {
  if (entry.indexedBy === "none") {
    // 这一情形已经由 assertValidResourceEntry 拒绝了（"none" 不得带 sheetId），
    // 这里直接放行，不重复判一遍。
    return;
  }
  if (entry.sheetId !== undefined && !sheetIds.includes(entry.sheetId)) {
    throw invalid(`${label}.sheetId "${entry.sheetId}" does not match any known sheet`);
  }
}

/**
 * 把一条按下标索引的资源变成登记处能搬的表。`"none"` 不在此列——核心不碰它，
 * 转成表就等于承诺会重键它，那是错的承诺。
 */
export function resourceToTable(entry: ResourceEntry): IndexedTable<unknown> {
  assertValidResourceEntry(entry, `resource "${entry.name}"`);
  if (entry.indexedBy === "none") {
    throw invalid(
      `resource "${entry.name}" is indexedBy "none"; the core does not rekey it and it has no table form`
    );
  }
  return new IndexedTable<unknown>(
    entry.name,
    entry.indexedBy,
    entry.data as Readonly<Record<number, unknown>>
  );
}

/**
 * `sheetId` 是可选的第二参数，而不是 `WorkbookModel` 在收拢结果时另行拼接。
 * `IndexedTable` 本身不存 `sheetId`（它只管重键，不管条目原来长什么样），
 * `WorkbookModel` 要把它另存一份才能在这里传进来；但"按 `entry.sheetId` 是否
 * 给出决定要不要在结果里放这个键"是 `exactOptionalPropertyTypes` 下构造
 * `ResourceEntry` 的通用写法（本文件里 `resourceToTable`/上面的校验都在处理同一个
 * `sheetId` 字段），放在这里能和其余的条目构造逻辑挨在一起，不用在
 * `WorkbookModel` 里再写一遍"给出就展开、没给出就不展开"的判断。
 */
export function tableToResource(table: IndexedTable<unknown>, sheetId?: string): ResourceEntry {
  return {
    name: table.name,
    indexedBy: table.indexedBy,
    ...(sheetId !== undefined ? { sheetId } : {}),
    data: { ...table.read() }
  };
}
