# cyexcel

English | [简体中文](README.zh-CN.md)

[Documentation](https://tcy-xsy.github.io/cyexcel/en/) · [Interactive examples](https://tcy-xsy.github.io/cyexcel/en/examples/basic/)

cyexcel is a framework-independent spreadsheet component for the browser. It uses virtualized Canvas rendering, stores cell data in a sparse matrix, and manages edits through commands, transactions, and undo history.

## Features

- Cell editing, styles, merging, and unmerging.
- Row and column insertion, deletion, movement, and resizing.
- Undo and redo, clipboard operations, find and replace, sorting, and filtering.
- Data validation, dropdown options, links, and cell images.
- Multiple worksheets with add, delete, switch, and rename operations.
- Basic formula calculation and XLSX export.

Initial data is imported into a snapshot. Edits to the spreadsheet and changes to the original array are independent. Drawing and hit testing are virtualized; the full dataset remains in memory. The documentation includes a 100,000-row, five-column example. Actual capacity depends on data contents, device performance, and available memory, and should be measured with representative data.

## Installation

```sh
npm install cyexcel
```

The initial release is `0.1.0`. To use a local source build, run `pnpm install`, `pnpm build`, and `pnpm pack` in a checkout, then install the generated `.tgz` in the consuming project.

## Quick start

The mount element must have measurable width and height.

```html
<div id="sheet" style="width: 100%; height: 480px;"></div>
```

```ts
import { createSheet, matrixToWorkbookSnapshot } from "cyexcel";

const container = document.querySelector<HTMLElement>("#sheet");
if (container === null) {
  throw new Error("Missing sheet container");
}

const sheet = createSheet(container, {
  snapshot: matrixToWorkbookSnapshot([
    ["A100", "Sensor", 9.99],
    ["A101", "Power module", 14.5]
  ])
});

sheet.setCell(0, 2, 12.99);
sheet.undo();

// Call sheet.dispose() when the page or host component unmounts.
```

`createSheet()` enables the toolbar, formula bar, and context menu by default. Disable them individually with `toolbar: false`, `formulaBar: false`, and `contextMenu: false`.

`matrixToWorkbookSnapshot()` maps row i, column j to cell `(i, j)`. Without custom headers, column labels use A, B, C, and so on. Use `rowsToWorkbookSnapshot(rows, fields)` to import object arrays in the specified field order.

## Entry points

| Import | Purpose |
| --- | --- |
| `cyexcel` | Sheets, workbooks, snapshot importers, and commonly used toolbar APIs |
| `cyexcel/core` | Core classes, types, and extension APIs; does not access the Custom Elements registry or import the formula parser |
| `cyexcel/toolbar` | Toolbar and formula-bar APIs |
| `cyexcel/formula` | Optional formula engine |
| `cyexcel/register` | Registers `<cy-excel>` on import; the element currently does not provide the full spreadsheet configuration interface |

`CyExcelCore` mounts into an ordinary HTML container. It can also be created and disposed within React, Vue, or another framework's component lifecycle.

## Workbooks

`createWorkbook(container, options)` provides worksheet tabs and operations to add, switch, delete, and rename sheets.

- Click a tab's close button to delete the sheet after confirmation. Deletion is not undoable, and at least one sheet must remain.
- Deleting the active sheet selects its right neighbor, or its left neighbor when no sheet exists to the right.
- Double-click a tab, or press F2 while it is focused, to edit its name. Enter or blur saves, Escape cancels, and blank input keeps the previous name.
- `workbook.deleteSheet(sheetId)` and `workbook.renameSheet(sheetId, name)` provide programmatic operations. Unknown IDs, deleting the final sheet, and blank programmatic names throw `RangeError`.

The default workbook toolbar exports every sheet, including sheets that have not been opened:

```ts
await workbook.exportExcel({ fileName: "report.xlsx" });
```

Export uses current sheet names and supports styles, merged ranges, and built-in validation rules. Set `download: false` to return a Blob, or `includeHeaders: false` to omit custom column headers. Images, custom JavaScript validators, and some advanced settings are not exported.

## Formulas and limitations

Workbooks enable formulas using the MIT-licensed `fast-formula-parser`. The core accepts an optional calculation engine through `formulaEngine`. See the documentation's Formulas example for configuration.

- Supports 12 basic functions and same-sheet references. Cross-sheet references are not supported.
- Sheets containing formulas currently restrict row/column insertion, deletion, movement, and sorting because structural reference rewriting is not implemented.
- Copy and fill use values and do not translate formula references.
- XLSX export is built in. XLSX import is available in the optional [Excel import workspace plugin](plugins/excel-import/README.md), which is not yet published and requires the workspace version of cyexcel. The parser is excluded from the main package.

## Documentation and examples

Start the documentation site locally:

```sh
pnpm install
pnpm docs:dev
```

The site supports English and Simplified Chinese, with guides, interactive examples, API documentation, data structures, and extension topics. Use the language switcher or open the `/en/` path for English. Existing hash links remain supported. Source files are in [docs-site/](docs-site/).

Build a static deployment with:

```sh
pnpm docs:build
```

The build prerenders every Chinese and English documentation page into HTML, including page metadata, canonical and language links, and structured data. It also generates `sitemap.xml` and `robots.txt`. Documentation text is available without JavaScript; interactive examples require JavaScript.

Deploy the generated `dist-docs/` directory to a static host such as GitHub Pages, Cloudflare Pages, or Nginx. No backend service is needed. The default public URL is `https://tcy-xsy.github.io/cyexcel/`. For another host or subdirectory, set `DOCS_SITE_URL` to the full public URL when building. The host must serve each directory's `index.html`. The documentation build runs in Node.js and does not require Playwright or a browser installation.

The `examples/` directory includes plain HTML, ESM, custom renderers, and React/Vue integrations. Run `pnpm build`, then serve the directory with `npx serve examples`. React and Vue examples load their dependencies from a CDN using browser import maps.

## AI coding assistant Skill

The [cyexcel Skill](skills/cyexcel/SKILL.md) guides Claude Code and Codex through integration, configuration, and extension tasks, with API examples, lifecycle guidance, and current limitations.

Install it from the consuming project's directory:

```sh
npx skills add tcy-xsy/cyexcel --skill cyexcel --agent claude-code codex
```

Follow the installation prompts. Add `--global` for a user-level installation. Start a new assistant session after installation:

- Claude Code: enter `/cyexcel`, followed by the task.
- Codex: enter `$cyexcel`, followed by the task.

For example: "Create a purchasing workbook with quantity validation, amount formulas, and XLSX export."

The Skill provides coding instructions; it does not install the component runtime or provide a remote spreadsheet-control service. Install `cyexcel` separately in the consuming project.

## License

[MIT](LICENSE)
