# Integration

## Package availability

The initial package release is `0.1.0`. The Skill contains instructions, not the component runtime. First inspect the consumer's package manifest and installed version. Use an existing workspace dependency or a built local package when registry installation is unavailable.

For a source checkout, Node.js 20+ and the declared pnpm version are required:

```sh
pnpm install
pnpm build
pnpm pack --out /absolute/path/to/artifacts
```

Install the generated `.tgz` into the consumer with its package manager. Do not recommend a raw Git dependency that lacks built `dist/` files. Use `pnpm add cyexcel` only after verifying an appropriate registry release exists.

## Single sheet

Give the mount element an explicit height, such as `height: 480px`, and a nonzero width. The component does not require a separate CSS import.

```ts
import { createSheet, matrixToWorkbookSnapshot } from "cyexcel";

const sheet = createSheet(container, {
  snapshot: matrixToWorkbookSnapshot([
    ["Sensor", 3, 120],
    ["Power module", 5, 80]
  ], { blankRows: 100 }),
  columns: [
    { header: "Item", width: 180 },
    { header: "Quantity", width: 120, type: "number" },
    { header: "Unit price", width: 120, type: "number" }
  ]
});

sheet.setCell(0, 1, 4);
sheet.undo();
const snapshot = sheet.toSnapshot();
// Restore later with createSheet(container, { snapshot }).
// Call sheet.dispose() when the host component unmounts.
```

`toolbar`, `formulaBar`, and `contextMenu` can each be set to `false`. An explicit `topWidgets` configuration takes precedence over the toolbar and formula-bar switches.

`matrixToWorkbookSnapshot()` does not interpret an initial string as a formula. Enable an engine and write formulas with `setCell()` after creation, or supply actual formula records through the supported snapshot schema. `null` and `undefined` do not create value records. Object arrays can be imported with `rowsToWorkbookSnapshot(rows, fields)`; `fields` controls import order, not live field binding.

## Framework lifecycle

React: create the instance inside an effect after the mount ref exists, and return `() => sheet.dispose()` from that effect. Keep the instance in a ref when event handlers need it. Strict Mode can mount, clean up, and mount again; disposal must occur for each created instance.

Vue: create the instance in `onMounted()` using a template ref, and dispose it in `onBeforeUnmount()`. Do not make the entire instance deeply reactive.

For server-rendered applications, put creation in a client-only lifecycle; use a dynamic import there when the framework requires it. Updating the original input array does not update an existing sheet. Use public setters for incremental changes or intentionally recreate the instance when replacing the document.

## Entry points

| Import | Purpose |
| --- | --- |
| `cyexcel` | Presets, selected core exports, importers, toolbar helpers |
| `cyexcel/core` | Full core API and types, without importing the bundled formula parser |
| `cyexcel/formula` | `createFormulaEngine` and supported-function metadata |
| `cyexcel/toolbar` | Toolbar and formula-bar helpers |
| `cyexcel/register` | Side-effect registration of the custom element; not needed for container integration |
