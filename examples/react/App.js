import React, { useEffect, useRef } from "react";
import { createRoot } from "react-dom/client";
import { CyExcelCore, rowsToWorkbookSnapshot } from "../../dist/core.js";

const columns = [
  { width: 140, header: "SKU" },
  { width: 240, header: "Name" },
  { width: 100, type: "number", header: "Price" }
];

/** The i-th field is the i-th column — a parameter of the import, not of `columns`. */
const fields = ["sku", "name", "price"];

const data = [
  { sku: "A100", name: "Widget", price: 9.99 },
  { sku: "A101", name: "Gadget", price: 14.5 },
  { sku: "A102", name: "Doohickey", price: 3.25 }
];

function CyExcelSheet() {
  const hostRef = useRef(null);

  useEffect(() => {
    // Imported once, on mount. The snapshot is a value handed to the sheet,
    // not a binding: React re-renders never flow back into it, and edits made
    // in the sheet never touch `data`.
    const sheet = new CyExcelCore(hostRef.current, {
      columns,
      snapshot: rowsToWorkbookSnapshot(data, fields)
    });
    return () => sheet.dispose();
  }, []);

  return React.createElement("div", { className: "sheet-host", ref: hostRef });
}

createRoot(document.getElementById("root")).render(React.createElement(CyExcelSheet));
