import React, { useEffect, useRef } from "react";
import { createRoot } from "react-dom/client";
import { CyExcelCore, matrixToWorkbookSnapshot } from "../../dist/core.js";

// No explicit headers means Excel-style A/B/C column letters.
const data = [
  ["A100", "Widget", 9.99],
  ["A101", "Gadget", 14.5],
  ["A102", "Doohickey", 3.25]
];

function CyExcelSheet() {
  const hostRef = useRef(null);

  useEffect(() => {
    // Imported once, on mount. The snapshot is a value handed to the sheet,
    // not a binding: React re-renders never flow back into it, and edits made
    // in the sheet never touch `data`.
    const sheet = new CyExcelCore(hostRef.current, {
      snapshot: matrixToWorkbookSnapshot(data)
    });
    return () => sheet.dispose();
  }, []);

  return React.createElement("div", { className: "sheet-host", ref: hostRef });
}

createRoot(document.getElementById("root")).render(React.createElement(CyExcelSheet));
