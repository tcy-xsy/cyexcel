# cyexcel-plugin-excel-import

Optional XLSX import for cyexcel. This is a separate workspace package: the main package does not depend on ExcelJS or include the parser in its ESM or UMD builds. ExcelJS loads when a file is read.

## Installation

Requires cyexcel 0.1.1 or later, before 1.0. The plugin targets ESM applications using a bundler such as Vite or Webpack.

```sh
npm install cyexcel@^0.1.1 cyexcel-plugin-excel-import@^0.1.0
```

## Development

Run from the repository root:

```sh
pnpm install
pnpm build
pnpm build:plugins
pnpm test:plugins
```

Output: `plugins/excel-import/dist/`.

## Integration

```ts
import { createWorkbook } from 'cyexcel';
import { createExcelImportPlugin } from 'cyexcel-plugin-excel-import';

const excelImport = createExcelImportPlugin({
  onWarning: message => console.warn(message)
});
const workbook = createWorkbook(container, {
  toolbar: { items: excelImport.toolbarItems }
});
excelImport.attach(workbook);

// Optional: import directly without opening the toolbar dialog.
const result = await excelImport.importExcel(file);
console.log(result.warnings);

// On unmount:
excelImport.dispose();
workbook.dispose();
```

The toolbar requires confirmation before replacing all tabs and clearing their undo history. Errors appear inside the dialog. Closing the dialog cancels application of pending data. Parse failures keep the existing workbook. Only one import may run at a time per plugin instance.

For parsing without replacing an existing workbook:

```ts
import { readExcelWorkbook } from 'cyexcel-plugin-excel-import';

const { sheets, warnings } = await readExcelWorkbook(file);
const workbook = createWorkbook(container, { sheets });
```

Each entry contains a separate single-sheet snapshot. Inputs accept `Blob` (including `File`), `ArrayBuffer`, and `Uint8Array`. Options: `maxFileSize` (default 25 MiB), `maxCells` (default 200,000 populated or styled cells), `signal`, and `onWarning`.

Files are parsed locally. ExcelJS decompresses and parses the workbook in memory before cell counting; these limits do not make parsing streaming. Cancellation prevents data application but does not interrupt synchronous parsing.

## Supported Content

- Worksheet names and order; hidden sheets become visible tabs.
- Text, numbers, booleans, dates as Excel serial numbers, formula text and cached results, including expanded shared formulas.
- Common font, fill, alignment, number format and border properties. Unsupported border patterns become solid lines; indexed colors are not imported.
- Merges, frozen rows/columns, row heights, column widths and hidden rows/columns.
- Wrapping is applied at column level because cyexcel does not expose per-cell wrapping.

Only unencrypted `.xlsx` is supported. Legacy `.xls`, macros, charts, images, data validation, conditional formatting, named ranges and print settings are not imported. Hyperlinks retain display text; rich text becomes plain text. Recalculation follows cyexcel's formula support; cross-sheet references are preserved but cannot calculate. Warnings are returned in the result and sent to `onWarning` when provided.

## License

MIT. ExcelJS is distributed separately under its own MIT license.
