import { columnIndexToLetter } from "../headers";

/**
 * A1 记法与 `(row, column)` 的互转。
 *
 * 编辑栏的名称框要用它（显示 `E1`、输入 `B7` 跳过去），但它本身是纯字符串处理，
 * 所以放在这里而不是 `formulaBar.ts` 里——那个文件在 node 里测不了，这个能测。
 *
 * **下标是 0 基，A1 记法是 1 基**：`A1` 是 `(0, 0)`。这个差一位是整个文件里唯一
 * 值得留神的地方。
 */

/** `(0, 0)` → `"A1"`。 */
export function toA1(row: number, column: number): string {
  return `${columnIndexToLetter(column)}${row + 1}`;
}

/** 一片区域的显示文本：单格就是 `A1`，多格是 `A1:C3`（与 Excel 名称框一致）。 */
export function rangeToA1(range: {
  startRow: number;
  startColumn: number;
  endRow: number;
  endColumn: number;
}): string {
  const start = toA1(range.startRow, range.startColumn);
  return range.startRow === range.endRow && range.startColumn === range.endColumn
    ? start
    : `${start}:${toA1(range.endRow, range.endColumn)}`;
}

/**
 * `"B7"` → `{ row: 6, column: 1 }`。认不出来返回 `null`——名称框里输入的是任意
 * 文本，解析失败是常态，不是异常，所以不抛。
 *
 * 大小写不敏感（`b7` 和 `B7` 一样），前后空白忽略。不支持 `$A$1` 绝对引用、
 * 不支持跨表 `Sheet1!A1`：前者要等公式，后者要等多 Sheet，现在都还没有。
 */
export function parseA1(text: string): { row: number; column: number } | null {
  const match = /^\s*([A-Za-z]+)\s*(\d+)\s*$/.exec(text);
  if (match === null) {
    return null;
  }
  const letters = (match[1] as string).toUpperCase();
  const digits = Number(match[2]);
  if (!Number.isSafeInteger(digits) || digits < 1) {
    return null;
  }
  let column = 0;
  for (const char of letters) {
    // 26 进制但没有 0：A=1…Z=26，AA=27。所以每一位都要先乘 26 再加，
    // 不能当成普通的 0 基进制来算。
    column = column * 26 + (char.charCodeAt(0) - 64);
  }
  return { row: digits - 1, column: column - 1 };
}
