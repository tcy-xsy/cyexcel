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
 * 一格最终显示成什么文本：格样式的 `numFmt` 优先，否则退回 `fallback`（列的
 * `CellType.format()`）。
 *
 * 只有数字和 `Date` 进 numfmt（它自己认 Date）。字符串一律走 fallback，**不做隐式的
 * 字符串转数字**——存着 `"0012"` 却被画成 `12` 没法解释给用户听。格式化抛错时也兜回
 * fallback：一格画错好过整张表白屏。
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
