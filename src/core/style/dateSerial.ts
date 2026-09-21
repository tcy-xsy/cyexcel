import { dateFromSerial, dateToSerial, parseDate } from "numfmt";

/**
 * Excel 序列号与外部日期表示的互转。
 *
 * **日期在这个模型里不是一种值类型**——它是数字（自 1900-01-01 起的天数）加一个
 * 日期 `numFmt`，与 Excel、Univer、xlsx 完全一致。好处是值永远是 JSON 原生类型，
 * 快照往返无损、导出 xlsx 零转换；代价是宿主拿到的是数字，要显示成日期得自己转
 * ——这几个函数就是为此提供的。
 *
 * 序列号是**墙上时钟**，没有时区：46037 就是 2026-01-15，在哪个时区读都一样。
 * 所以 ISO 串那两个函数（`serialFromISODate` / `serialToISODate`）不走 `Date`
 * ——`new Date("2026-01-15")` 按 UTC 解析，在东八区会变成当天 08:00，往返出来是
 * 46037.333 而不是 46037。要跟 `Date` 打交道的调用方才用下面那对，它们按**本地**
 * 日历字段换算，语义是"这个时刻在本地是哪一天"。
 */

/** `Date` → 序列号，按本地日历字段。无效日期返回 `null`，由调用方决定怎么处理。 */
export function toSerial(date: Date): number | null {
  const serial = dateToSerial(date);
  return typeof serial === "number" ? serial : null;
}

/** 序列号 → 本地时区的 `Date`。不是有效序列号时返回 `null`。 */
export function fromSerial(serial: number): Date | null {
  const parts = dateFromSerial(serial);
  if (!Array.isArray(parts)) {
    return null;
  }
  const [year, month, day, hours = 0, minutes = 0, seconds = 0] = parts as number[];
  if (year === undefined || month === undefined || day === undefined) {
    return null;
  }
  // numfmt 的月份是 1-12，`Date` 构造器要 0-11。
  return new Date(year, month - 1, day, hours, minutes, seconds);
}

/**
 * `yyyy-mm-dd`（以及 numfmt 认得的其它日期写法）→ 序列号，不经过 `Date`，因此
 * 与时区无关。不是日期时返回 `null`。
 */
export function serialFromISODate(text: string): number | null {
  const parsed = parseDate(text);
  return parsed && typeof parsed.v === "number" ? parsed.v : null;
}

/** 序列号 → `yyyy-mm-dd`。不是有效序列号时返回 `null`。 */
export function serialToISODate(serial: number): string | null {
  const parts = dateFromSerial(serial);
  if (!Array.isArray(parts)) {
    return null;
  }
  const [year, month, day] = parts as number[];
  if (year === undefined || month === undefined || day === undefined) {
    return null;
  }
  const pad = (n: number): string => String(n).padStart(2, "0");
  return `${String(year).padStart(4, "0")}-${pad(month)}-${pad(day)}`;
}
