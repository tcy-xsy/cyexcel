import { CyExcelError, CyExcelErrorCode } from "../errors";

/**
 * 规范形式的非负整数字符串："0" 或不带前导零的正整数。
 *
 * **这条规则此前在两个文件里各写了一遍**（`serialize.ts` 与 `resources.ts`），
 * 正则相同而消息已经开始分叉。它现在是按下标索引的存储唯一的键值规则：
 * `IndexedTable`（活表）、`resources`（插件条目）、`serialize`（快照闸门）
 * 全部走这里。
 *
 * 为什么必须是**规范**形式而不是"能转成数字就行"：`Number()` 对非法输入不抛错，
 * `" 1"` / `"1e3"` / `"1.5"` 都能转出个数来，转完之后重键会把条目搬到错误的
 * 下标上，或者搬到一个排列永远不会点名的位置——静默丢数据。
 */
const CANONICAL_INDEX_KEY = /^(0|[1-9]\d*)$/;

export function isCanonicalIndexKey(key: string): boolean {
  return CANONICAL_INDEX_KEY.test(key);
}

/**
 * 校验一张按下标索引的映射的全部键。`label` 用于拼出可定位的错误消息。
 *
 * `Object.keys()` 对符号键不可见，因此符号键绕过这里的校验但仍被对象展开操作
 * 复制进表；惰性但不失效——JSON 无法产生符号键，而重键只遍历 `Object.keys()`，
 * 永不触及它。
 */
export function assertCanonicalIndexKeys(
  record: Readonly<Record<number, unknown>>,
  label: string
): void {
  for (const key of Object.keys(record)) {
    if (!isCanonicalIndexKey(key)) {
      throw new CyExcelError(
        CyExcelErrorCode.InvalidOptions,
        `${label} must be a non-negative integer key, received "${key}"`
      );
    }
  }
}

/** 校验一个用作下标的数字。`set(index)` 这类以 number 入参的路径用它。 */
export function assertIndexIsNonNegativeInteger(index: number, label: string): void {
  if (!Number.isInteger(index) || index < 0) {
    throw new CyExcelError(
      CyExcelErrorCode.InvalidOptions,
      `${label} must be a non-negative integer, received ${String(index)}`
    );
  }
}
