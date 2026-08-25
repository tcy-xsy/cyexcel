import { CyExcelError, CyExcelErrorCode } from "../errors";
import { IndexedTable } from "./IndexedTable";
import type { ResourceEntry } from "./types";

/** 规范形式的非负整数字符串："0" 或不带前导零的正整数。与 serialize.ts 同一条。 */
const NON_NEGATIVE_INTEGER_KEY = /^(0|[1-9]\d*)$/;

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
    return;
  }
  const data: unknown = entry.data;
  if (typeof data !== "object" || data === null || Array.isArray(data)) {
    throw invalid(
      `${label}.data must be a sparse index map because indexedBy is "${entry.indexedBy}"`
    );
  }
  for (const key of Object.keys(data)) {
    if (!NON_NEGATIVE_INTEGER_KEY.test(key)) {
      throw invalid(`${label}.data key must be a non-negative integer, received "${key}"`);
    }
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

export function tableToResource(table: IndexedTable<unknown>): ResourceEntry {
  return { name: table.name, indexedBy: table.indexedBy, data: { ...table.read() } };
}
