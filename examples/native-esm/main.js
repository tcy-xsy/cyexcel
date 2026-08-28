// In a real project this would be `import { CyExcelCore } from "cyexcel/core"`,
// resolved by Node/the browser's own import map against node_modules. This
// repo demos against its own freshly-built dist/ instead of a published
// package, so the relative path stands in for that.
import { CyExcelCore, rowsToWorkbookSnapshot } from "../../dist/core.js";

const columns = [
  { id: "sku", field: "sku", width: 140, header: "SKU" },
  { id: "name", field: "name", width: 240, header: "Name" },
  { id: "price", field: "price", width: 100, type: "number", header: "Price" }
];

// The i-th field is the i-th column: `fields` is a parameter of the *import*,
// not part of the column config (doc/adr/0004-single-data-model.md §3.2).
const fields = ["sku", "name", "price"];

const data = [
  { sku: "A100", name: "Widget", price: 9.99 },
  { sku: "A101", name: "Gadget", price: 14.5 },
  { sku: "A102", name: "Doohickey", price: 3.25 }
];

// The sheet takes a snapshot, not a row array. This is a one-shot import: once
// it returns, `data` and the sheet share nothing — editing one never moves the
// other.
new CyExcelCore(document.getElementById("sheet"), {
  columns,
  snapshot: rowsToWorkbookSnapshot(data, fields)
});
