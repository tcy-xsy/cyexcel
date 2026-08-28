# cyexcel

A high-performance, framework-agnostic spreadsheet component. Canvas-virtualized
rendering handles 100K-1M rows, a Command/Transaction layer drives every
mutation (so undo/redo and clipboard paste are never special-cased), and the
sheet owns its data outright in a sparse cell matrix — your rows are *imported*
into it once, not adapted behind an interface.

Zero runtime dependencies. `cyexcel/core` never touches the Custom Elements
registry, so it works equally well as a plain library call or embedded inside
any framework's own component model.

## Install

```bash
npm install cyexcel
```

## Quick start

```ts
import { CyExcelCore, rowsToWorkbookSnapshot } from "cyexcel/core";

const columns = [
  { id: "sku", field: "sku", width: 140, header: "SKU" },
  { id: "name", field: "name", width: 240, header: "Name" },
  { id: "price", field: "price", width: 100, type: "number", header: "Price" }
];

// The i-th field is the i-th column — a parameter of the import, not of
// `columns`. It's the only place field names are ever consulted.
const fields = ["sku", "name", "price"];

const data = [
  { sku: "A100", name: "Widget", price: 9.99 },
  { sku: "A101", name: "Gadget", price: 14.5 }
];

const sheet = new CyExcelCore(document.querySelector("#sheet"), {
  columns,
  snapshot: rowsToWorkbookSnapshot(data, fields)
});

sheet.setCell(0, 2, 12.99); // one undoable Command
sheet.undo();
```

`rowsToWorkbookSnapshot` is a one-shot import, not a binding: once the sheet is
constructed, `data` and the sheet share nothing. Editing a cell never writes
back into your array, and mutating your array never shows up in the table.

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
