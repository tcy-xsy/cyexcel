/**
 * Excel/Google Sheets-compatible TSV clipboard format: cells joined by a
 * tab, rows joined by "\r\n". A cell that contains a tab, newline, or
 * double quote gets wrapped in double quotes with internal quotes doubled
 * — matches how real spreadsheet apps escape clipboard TSV, so a real
 * cross-app paste round-trips correctly.
 */
const NEEDS_QUOTING = /[\t\n\r"]/;

export function escapeTsvCell(text: string): string {
  if (!NEEDS_QUOTING.test(text)) {
    return text;
  }
  return `"${text.replace(/"/g, '""')}"`;
}

export function encodeTsv(grid: readonly (readonly string[])[]): string {
  return grid.map((row) => row.map(escapeTsvCell).join("\t")).join("\r\n");
}

function unescapeTsvCell(text: string): string {
  if (text.length >= 2 && text.startsWith('"') && text.endsWith('"')) {
    return text.slice(1, -1).replace(/""/g, '"');
  }
  return text;
}

/**
 * Parses a single TSV row, respecting quoted cells that may themselves
 * contain tabs or (via an embedded CRLF/LF) newlines — those can't be split
 * on naively by `\t`, so this walks the string character by character.
 */
function splitTsvRow(row: string): string[] {
  const cells: string[] = [];
  let current = "";
  let inQuotes = false;

  for (let i = 0; i < row.length; i += 1) {
    const char = row[i];
    if (inQuotes) {
      if (char === '"' && row[i + 1] === '"') {
        current += '"';
        i += 1;
      } else if (char === '"') {
        inQuotes = false;
      } else {
        current += char;
      }
    } else if (char === '"' && current === "") {
      inQuotes = true;
    } else if (char === "\t") {
      cells.push(current);
      current = "";
    } else {
      current += char;
    }
  }
  cells.push(current);
  return cells;
}

/**
 * Splits TSV text into rows on "\r\n" or bare "\n" (real clipboard input —
 * e.g. from another app, or a browser normalizing line endings — isn't
 * guaranteed to use our own encodeTsv()'s exact separator), but not on a
 * newline embedded inside a quoted cell.
 */
export function decodeTsv(text: string): string[][] {
  if (text === "") {
    return [[""]];
  }

  const rows: string[] = [];
  let current = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    if (char === '"') {
      inQuotes = !inQuotes;
      current += char;
    } else if (!inQuotes && (char === "\n" || char === "\r")) {
      if (char === "\r" && text[i + 1] === "\n") {
        i += 1;
      }
      rows.push(current);
      current = "";
    } else {
      current += char;
    }
  }
  rows.push(current);

  return rows.map((row) => splitTsvRow(row).map(unescapeTsvCell));
}
