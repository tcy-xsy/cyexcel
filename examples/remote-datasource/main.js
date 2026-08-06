import { CyExcelCore } from "../../dist/core.js";

const columns = [
  { id: "sku", field: "sku", width: 140, header: "SKU" },
  { id: "name", field: "name", width: 240, header: "Name" },
  { id: "price", field: "price", width: 100, type: "number", header: "Price" }
];

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

/**
 * A minimal custom DataSource — getCell/setCell are synchronous (required:
 * they're called during paint/edit-commit), backed by a plain in-memory
 * array that's already been fully populated by the time this is
 * constructed. `columns` closed over here just to map a column index to
 * the row object's field name.
 */
class InMemoryDataSource {
  #rows;
  #columns;

  constructor(rows, cols) {
    this.#rows = rows;
    this.#columns = cols;
  }

  getRowCount() {
    return this.#rows.length;
  }

  getColumnCount() {
    return this.#columns.length;
  }

  getCell(row, column) {
    return this.#rows[row]?.[this.#columns[column].field];
  }

  setCell(row, column, value) {
    this.#rows[row][this.#columns[column].field] = value;
  }
}

const rows = await fetchRowsFromServer();
const dataSource = new InMemoryDataSource(rows, columns);

document.getElementById("sheet").textContent = "";
new CyExcelCore(document.getElementById("sheet"), { columns, dataSource });
