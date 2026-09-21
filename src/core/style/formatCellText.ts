import { format as formatNumber, isValidFormat } from "numfmt";
import { CyExcelError, CyExcelErrorCode } from "../errors";
import type { CellStyle } from "./types";

/**
 * 校验 Excel 数字格式 pattern，非法就抛——**不静默降级成"当文本画出来"**
 * （Step 24 退出条件第 5 条）。在 `setCellStyle` 入口调用，错误在设置时报出来，
 * 而不是等到某一格滚进视口才露出来。
 */
export function assertValidNumFmt(pattern: string): void {
  if (!isValidFormat(pattern)) {
    throw new CyExcelError(
      CyExcelErrorCode.InvalidNumberFormat,
      `Invalid number format pattern: ${JSON.stringify(pattern)}`
    );
  }
}

/**
 * 一格最终显示成什么文本。
 *
 * 优先级：格样式的 `numFmt` → 调用方给的 `fallback`（列的 `CellType.format()`）。
 * 这是 Step 24 "未设置时退回 `CellType.format()`，向后兼容"那条的落点。
 *
 * # 哪些值进 numfmt
 *
 * 只有数字和 `Date`。numfmt 自己就认 `Date`（内部转 Excel 序列号，并按时区处理），
 * 所以这里不需要先 `dateToSerial`。
 *
 * 字符串、布尔、`null`/`undefined` 一律走 `fallback`，**不做隐式的字符串转数字**：
 * "看起来像数字的字符串"要不要当数字格式化，是编辑链路解释用户输入的事
 * （ADR 0004 §3.6 把那条线划在 `#parseEditedValue`），不是渲染时猜。一格存着
 * `"0012"` 却因为设了数字格式被画成 `12`，是没法解释给用户听的。
 *
 * # 为什么格式化失败要兜住
 *
 * pattern 在设置时已经 `assertValidNumFmt` 过，但"合法的 pattern" 与"这个值能用这个
 * pattern 格式化"是两回事（numfmt 对某些值域组合会抛）。兜回 `fallback` 而不是让
 * 渲染整帧崩掉——一格画错远好过整张表白屏。
 */
export function formatCellText(
  value: unknown,
  style: CellStyle | undefined,
  fallback: () => string
): string {
  const pattern = style?.numFmt;
  if (pattern === undefined) {
    return fallback();
  }
  if (typeof value !== "number" && !(value instanceof Date)) {
    return fallback();
  }
  try {
    return formatNumber(pattern, value);
  } catch {
    return fallback();
  }
}
