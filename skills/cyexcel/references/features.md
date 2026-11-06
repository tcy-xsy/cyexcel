# Features and boundaries

## Workbook management

```ts
import { createWorkbook, matrixToWorkbookSnapshot } from "cyexcel";

const workbook = createWorkbook(container, {
  sheets: [{ name: "Orders", snapshot: matrixToWorkbookSnapshot([[3, 120, 0]]) }]
});
const sheet = workbook.getActiveSheet();
sheet.setCell(0, 2, "=A1*B1");
const id = workbook.addSheet({ name: "Notes" });
workbook.renameSheet(id, "Remarks");
workbook.setActiveSheet(id);
workbook.deleteSheet(id);
// Call workbook.dispose() on unmount.
```

The handle exposes `getActiveSheet`, `setActiveSheet`, `addSheet`, `deleteSheet`, `renameSheet`, `exportExcel`, and `dispose`. Row and cell operations belong to the active `CyExcelCore`, not the workbook handle. Sheet deletion is not undoable and cannot remove the final sheet. Invalid sheet IDs and blank programmatic names throw `RangeError`. Do not assume cell edits and workbook-tab operations share one history stack.

## Formulas

`createWorkbook()` enables formulas by default. `createSheet()` and direct core construction require explicit configuration:

```ts
import { createSheet, matrixToWorkbookSnapshot } from "cyexcel";
import { createFormulaEngine } from "cyexcel/formula";

const sheet = createSheet(container, {
  snapshot: matrixToWorkbookSnapshot([[3, 120, 0]]),
  formulaEngine: createFormulaEngine()
});
sheet.setCell(0, 2, "=A1*B1");
sheet.getCell(0, 2); // 360
sheet.getCellFormula(0, 2); // "=A1*B1"
sheet.getCellFormulaError(0, 2); // undefined when valid
```

Supported functions: `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTIF`, `SUMIF`, `IF`, `IFERROR`, `ROUND`, `ABS`, `CONCATENATE`.

Same-sheet A1 references, absolute references, and ranges are supported. Cross-sheet references, named ranges, dynamic references, spill arrays, and volatile functions are not. Formula-containing sheets block row/column insert, delete, move, and sort because reference rewriting is unavailable. Copy and fill use values without formula translation. Default limits are 100,000 cells per range and dependency depth 256; `createFormulaEngine({ maxRangeCells, maxDepth })` configures them.

## Data validation

```ts
sheet.setDataValidation(
  { startRow: 0, startColumn: 1, endRow: 99, endColumn: 1 },
  {
    criteria: { type: "whole", operator: "between", value1: 1, value2: 100 },
    errorStyle: "stop",
    errorMessage: "Quantity must be an integer from 1 to 100"
  }
);
```

Criteria types: `whole`, `decimal`, `textLength`, `date`, `time`, `list`, `custom`. List rules use `items` and optionally `renderMode: "arrow" | "text" | "custom"`. A cell belongs to at most one rule; a new rule replaces overlapping coverage.

Only editor submission with `errorStyle: "stop"` rejects invalid input. The rejection dialog discards the attempted edit on confirmation or Escape. Warning and information styles allow submission. Paste, fill, `setCell()`, and replacement allow writes and mark invalid cells. Do not use validation as an authorization or storage constraint. `findInvalidCells(range?)` and `validateCell(row, column)` provide explicit checks. Custom validators are synchronous pure functions registered per instance; functions are not serialized, and unregistered validator names are treated as valid.

## Clipboard and clearing

Presets enable the context menu by default; core requires `contextMenu: true`. `clearContents(range?)` clears values, formulas, and cell images while retaining styles and validation in one undoable operation. `clearRangeStyle(range)` clears formatting.

Clipboard access depends on browser permissions and secure-context rules. Preserve a keyboard paste path when programmatic reads are denied. Cut clears source data only after a successful clipboard write; it is not deferred movement on paste. Multi-range menu operations are restricted to supported actions.

## XLSX export

```ts
await workbook.exportExcel({ fileName: "report.xlsx" });
const blob = await workbook.exportExcel({ download: false, includeHeaders: false });
```

Exports include all workbook tabs, including unvisited sheets, using current names. Values, formulas with cached results, styles, merges, and built-in validation rules are supported. Images and JavaScript custom validators are omitted; freeze panes, filter conditions, conditional formats, and print settings are not exported. XLSX import is not implemented. Export does not establish server persistence or collaborative editing.
