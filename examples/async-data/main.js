import { CyExcelCore, rowsToWorkbookSnapshot } from "../../dist/core.js";

const columns = [
  { width: 140, header: "SKU" },
  { width: 240, header: "Name" },
  { width: 100, type: "number", header: "Price" }
];

/**
 * The i-th field is the i-th column. This is a parameter of the *import* — the
 * one moment the server's field names matter — not part of the column config
 * (doc/adr/0004-single-data-model.md §3.2, §4.5). After the import the sheet
 * addresses everything by (row, column) and never sees a field name again.
 */
const fields = ["sku", "name", "price"];

/** Stands in for a network call — real code would `await fetch(...)` here. */
function fetchRowsFromServer() {
  return new Promise((resolve) => {
    setTimeout(
      () =>
        resolve([
          { sku: "A100", name: "Widget", price: 9.99 },
          { sku: "A101", name: "Gadget", price: 14.5 },
          { sku: "A102", name: "Doohickey", price: 3.25 }
        ]),
      400
    );
  });
}

const rows = await fetchRowsFromServer();

document.getElementById("sheet").textContent = "";
// Everything async happened *before* this line. The sheet is constructed with
// a snapshot already in hand, so there is no loading state inside the table
// itself — and no reference kept to `rows`, which is free to be collected the
// moment the import returns.
new CyExcelCore(document.getElementById("sheet"), {
  columns,
  snapshot: rowsToWorkbookSnapshot(rows, fields)
});
