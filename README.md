# cyexcel

A high-performance, framework-agnostic spreadsheet component. Canvas-virtualized
rendering handles 100K-1M rows, a Command/Transaction layer drives every
mutation (so undo/redo and clipboard paste are never special-cased), and the
sheet owns its data outright in a sparse cell matrix — your rows are *imported*
into it once, not adapted behind an interface.

`cyexcel/core` never touches the Custom Elements
registry, so it works equally well as a plain library call or embedded inside
any framework's own component model.

Workbook shells enable formulas using the MIT-licensed `fast-formula-parser`.
The core accepts an optional calculator through `formulaEngine`, so importing
`cyexcel/core` does not import the parser. See [formula support](doc/formulas.md)
for setup, supported functions, and current limitations.

## Install

```bash
npm install cyexcel
```

## Quick start

```ts
import { CyExcelCore, matrixToWorkbookSnapshot } from "cyexcel/core";

// Row i, column j becomes cell (i, j). With no explicit headers, cyexcel
// uses Excel-style A/B/C column letters.
const data = [
  ["A100", "Widget", 9.99],
  ["A101", "Gadget", 14.5]
];

const sheet = new CyExcelCore(document.querySelector("#sheet"), {
  snapshot: matrixToWorkbookSnapshot(data)
});

sheet.setCell(0, 2, 12.99); // one undoable Command
sheet.undo();
```

`matrixToWorkbookSnapshot` is a one-shot import, not a binding: once the sheet is
constructed, `data` and the sheet share nothing. Editing a cell never writes
back into your array, and mutating your array never shows up in the table.

If your source data is a list of objects, `rowsToWorkbookSnapshot(objects,
fields)` remains available as the field-mapped importer.

`CyExcelCore` paints into whatever container element you give it — there's no
required Web Component wrapper. See `examples/` for this same pattern wired
into plain HTML, native ESM, a custom cell renderer, an async data source,
and minimal React/Vue integrations (both load React/Vue from a CDN via a
native `<script type="importmap">`, not a bundler — run `pnpm build` first,
then serve `examples/` with **any plain static file server** — e.g.
`npx serve examples`, not this repo's own `pnpm dev`/`pnpm playground`. Vite's
dev server tries to resolve every module's imports itself and doesn't know
about `<script type="importmap">`, so it 500s on the bare `"react"`/`"vue"`
specifiers those two examples rely on the *browser* to resolve).

## Entry points

`createWorkbook(container, options)` provides sheet tabs with add and switch
controls. Click a tab's close icon to delete its sheet after confirmation, or call
`workbook.deleteSheet(sheetId)` directly. Deletion cannot be undone and always
leaves at least one sheet. Deleting the active sheet selects its right neighbor,
or its left neighbor when there is no sheet to the right. An unknown ID or an
attempt to delete the last sheet throws a `RangeError`.
Double-click a tab (or press F2 while it is focused) to rename it inline. Enter
or blur saves the name; Escape cancels. Blank input keeps the previous name.
`workbook.renameSheet(sheetId, name)` also supports renaming programmatically;
it trims whitespace and throws a `RangeError` for a blank name or unknown ID.

The default workbook toolbar exports all sheets to `.xlsx`, including sheets
that have not been opened. `await workbook.exportExcel({ fileName: "report.xlsx" })`
uses the current tab names, styles, merges and built-in validation rules. Pass
`download: false` to receive a Blob, or `includeHeaders: false` to omit custom
column headers. See [Excel export](doc/excel-export.md) for supported features and limitations.

| Import | Use it for |
| --- | --- |
| `cyexcel` | `core`'s full export surface, aggregated in one place |
| `cyexcel/core` | Just the logic/rendering layer — no Custom Elements registry access at all |
| `cyexcel/register` | Side-effecting: auto-registers `<cy-excel>` on import (currently an empty, unconfigured shell — see `doc/implementation-steps.md` Step 13) |

## Documentation

- [`doc/architecture.md`](doc/architecture.md) — module layering and rendering pipeline
- [`doc/api-design.md`](doc/api-design.md) — public API surface
- [`doc/customization-and-hooks.md`](doc/customization-and-hooks.md) — Hooks, Events, plugins, custom cell renderers
- [`doc/roadmap.md`](doc/roadmap.md) — milestones and what's implemented so far

## License

[MIT](LICENSE)
