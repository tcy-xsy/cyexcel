/**
 * Converts a 0-based column index to its spreadsheet-style letter label
 * (0 -> "A", 25 -> "Z", 26 -> "AA", 27 -> "AB", ...) — the same base-26,
 * no-zero-digit algorithm Excel/Google Sheets/Univer use for column
 * references. Used as HeaderManager's optional fallback title for a column
 * with no explicit `header` declared (see `columnLetterFallback`).
 */
export function columnIndexToLetter(index: number): string {
  let remaining = index + 1;
  let label = "";
  while (remaining > 0) {
    const digit = (remaining - 1) % 26;
    label = String.fromCharCode(65 + digit) + label;
    remaining = Math.floor((remaining - 1) / 26);
  }
  return label;
}
