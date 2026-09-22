import { CyExcelCore, matrixToWorkbookSnapshot } from "../../dist/core.js";

/** Stands in for a network call — real code would `await fetch(...)` here. */
function fetchRowsFromServer() {
  return new Promise((resolve) => {
    setTimeout(
      () =>
        resolve([
          ["A100", "Widget", 9.99],
          ["A101", "Gadget", 14.5],
          ["A102", "Doohickey", 3.25]
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
  snapshot: matrixToWorkbookSnapshot(rows)
});
