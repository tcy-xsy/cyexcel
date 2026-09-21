/**
 * 数字格式的小数位增减。
 *
 * Excel 与 Univer 工具栏上那两颗 `.00` / `.0` 按钮（Univer 的
 * `AddDecimalCommand`/`SubtractDecimalCommand`）。看着简单，但它改的是 **pattern
 * 字符串本身**，而不是某个"小数位数"字段——我们的存储层只有 Excel pattern 串
 * （ADR 0004 §3.4），所以这里得把串拆开再拼回去。
 *
 * 纯字符串处理、不碰 DOM 也不碰 core，所以能在 node 里完整单测——这正是把它放在
 * `style/` 而不是 `toolbar/` 的理由。
 */

/** 没设过格式的格子，第一次点"增加小数位"时从这个基线开始。 */
const DEFAULT_BASE = "0";

/**
 * 只认 pattern 最后一段里的小数部分。分段格式（`正;负;零`）、科学计数、日期一律
 * 原样返回——那些的"小数位"不是一个有确定含义的概念，猜错了不如不动。
 */
function isAdjustable(pattern: string): boolean {
  return !/[;eE@]/.test(pattern) && !/[yYmMdDhHsS]/.test(pattern);
}

function decimalsOf(pattern: string): number {
  const match = /\.(0+)/.exec(pattern);
  return match === null ? 0 : (match[1] as string).length;
}

/**
 * 小数部分要插在**最后一个数字占位符之后**，不是追加到串尾。
 *
 * `0.00%` 的百分号就在尾巴上：追加到尾巴会得到 `0%.0`，那是个连 numfmt 都认不出
 * 的串，而且它会顺着 setRangeStyle 一路写进样式表里去。货币后缀、单位后缀同理。
 */
function withDecimals(pattern: string, decimals: number): string {
  const base = pattern.replace(/\.0+/, "");
  if (decimals <= 0) {
    return base;
  }
  let lastDigit = -1;
  for (let index = base.length - 1; index >= 0; index -= 1) {
    const char = base[index] as string;
    if (char === "0" || char === "#") {
      lastDigit = index;
      break;
    }
  }
  const fraction = `.${"0".repeat(decimals)}`;
  // 没有任何数字占位符的 pattern（理论上进不来，isAdjustable 已经挡掉了大部分）
  // 就退回追加，至少不丢内容。
  return lastDigit < 0
    ? `${base}${fraction}`
    : `${base.slice(0, lastDigit + 1)}${fraction}${base.slice(lastDigit + 1)}`;
}

/** 小数位 +1（上限 10 位——再多在 Excel 里也显示不出意义）。 */
export function addDecimal(pattern: string | undefined): string {
  const current = pattern === undefined || pattern === "" ? DEFAULT_BASE : pattern;
  if (!isAdjustable(current)) {
    return current;
  }
  return withDecimals(current, Math.min(10, decimalsOf(current) + 1));
}

/** 小数位 -1，减到 0 位就不再减（此时 pattern 里没有小数点）。 */
export function subtractDecimal(pattern: string | undefined): string {
  const current = pattern === undefined || pattern === "" ? DEFAULT_BASE : pattern;
  if (!isAdjustable(current)) {
    return current;
  }
  return withDecimals(current, Math.max(0, decimalsOf(current) - 1));
}
