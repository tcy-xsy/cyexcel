import React, { useEffect, useRef } from "react";
import { createRoot } from "react-dom/client";
import { CyExcelCore } from "../../dist/core.js";

const columns = [
  { id: "sku", field: "sku", width: 140, header: "SKU" },
  { id: "name", field: "name", width: 240, header: "Name" },
  { id: "price", field: "price", width: 100, type: "number", header: "Price" }
];

const data = [
  { sku: "A100", name: "Widget", price: 9.99 },
  { sku: "A101", name: "Gadget", price: 14.5 },
  { sku: "A102", name: "Doohickey", price: 3.25 }
];

function CyExcelSheet() {
  const hostRef = useRef(null);

  useEffect(() => {
    const sheet = new CyExcelCore(hostRef.current, { columns, data });
    return () => sheet.dispose();
  }, []);

  return React.createElement("div", { className: "sheet-host", ref: hostRef });
}

createRoot(document.getElementById("root")).render(React.createElement(CyExcelSheet));
