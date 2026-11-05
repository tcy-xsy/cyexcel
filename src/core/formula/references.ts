/** Shift A1 row references when an exported sheet gains a header row.
 * Quoted strings, sheet names and function names such as LOG10 stay intact.
 */
export function shiftFormulaRows(formula: string, offset: number): string {
  if (offset === 0) return formula;
  return formula.replace(/"(?:[^"]|"")*"|'(?:[^']|'')*'|(\$?[A-Za-z]{1,3})(\$?)([1-9][0-9]*)(?![A-Za-z0-9_])/g,
    (match: string, column: string | undefined, absolute: string | undefined, row: string | undefined, index: number) => {
      if (column === undefined || row === undefined) return match;
      if (index > 0 && /[A-Za-z0-9_.]/.test(formula[index - 1]!)) return match;
      if (/^\s*[(!]/.test(formula.slice(index + match.length))) return match;
      const shifted = Number(row) + offset;
      return shifted > 1048576 ? "#REF!" : `${column}${absolute ?? ""}${shifted}`;
    });
}
