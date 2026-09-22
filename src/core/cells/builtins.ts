import type { CellTypeRegistry } from "./CellTypeRegistry";
import { serialFromISODate, serialToISODate } from "../style/dateSerial";


function formatText(value: unknown): string {
  return value === undefined || value === null ? "" : String(value);
}

/** 序列号画成 `yyyy-mm-dd`；不是序列号（空格、脏数据）就原样走文本兜底。 */
function formatSerialDate(value: unknown): string {
  if (typeof value !== "number") {
    return formatText(value);
  }
  return serialToISODate(value) ?? formatText(value);
}

/**
 * Registers the built-in CellTypes listed in doc/architecture.md §8. Each
 * has a distinct, deterministic `format()` (no Intl locale formatting, to
 * keep output stable across test environments).
 *
 * 没有 `array` / `json` 类型了：`CellValue` 收紧成 `string | number | boolean`
 * 之后，格子里根本放不进数组和对象，这两个类型没有可格式化的输入（见
 * `sheet/types.ts` 的 `CellValue`）。
 *
 * `compare` is defined by exactly the types whose natural order differs from
 * the order their values' string forms give, because that string order is
 * what `createDefaultSortCompare()` falls back to when a type defines no
 * `compare`. `text`/`select` are already in string order; `date` is a serial
 * number, so plain numeric order is already chronological.
 */
export function registerBuiltinCellTypes(registry: CellTypeRegistry): void {
  registry.register("text", {
    format: (value) => formatText(value)
  });

  registry.register("number", {
    format: (value) => (typeof value === "number" && Number.isFinite(value) ? String(value) : ""),
    // Editing an empty cell must clear it, not coerce the empty string to 0.
    parse: (text) => (text.trim() === "" ? undefined : Number(text)),
    compare: (left, right) => (left as number) - (right as number)
  });

  registry.register("boolean", {
    format: (value) => (value ? "true" : "false"),
    // Editor-agnostic: whatever text a text-based editor produces (default
    // or otherwise) has to round-trip correctly even if no dedicated
    // checkbox editor is ever wired up for this column — see doc/progress.md's
    // editor-pluggability patch.
    parse: (text) => text.trim().toLowerCase() === "true",
    serialize: (value) => (value ? "true" : "false"),
    // By truthiness, matching format()/parse(), not by the value's string
    // form: a "boolean" column may hold 1/0 as readily as true/false, and
    // sorting has to agree with what is drawn. String order would read
    // "0" < "1" < "false" < "true" and interleave the two groups.
    compare: (left, right) => Number(Boolean(left)) - Number(Boolean(right))
  });

  // 日期 = **序列号**（自 1900-01-01 起的天数）+ 日期 numFmt，与 Excel / xlsx /
  // Univer 一致。这个类型本身只负责把用户敲的日期串转成序列号；显示成什么样由
  // numFmt 决定，所以这里的 format 只是没有格式时的兜底。
  registry.register("date", {
    format: (value) => formatSerialDate(value),
    // 走 numfmt 的解析器而不是 `new Date(text)`：后者把 "2026-01-15" 当 UTC 午夜，
    // 在东八区往返出来是 46037.333。序列号是墙上时钟，不该带时区。
    parse: (text) => serialFromISODate(text) ?? text,
    serialize: (value) => formatSerialDate(value)
    // 不需要 compare：序列号就是数字，默认的数字比较天然按时间序。
    // （Step 22 那段处理 "Fri < Mon < Sat" 的代码被这个数据模型消解掉了。）
  });

  registry.register("select", {
    // `options` is column/type metadata (see CellType.options), not an
    // editor choice — a column shows the matching label when configured,
    // and degrades to the raw value otherwise (no options configured, or a
    // value that isn't in the configured list).
    format: (value, context) => {
      const match = context.column.options?.find((option) => option.value === value);
      return match ? (match.label ?? match.value) : formatText(value);
    }
  });

}
