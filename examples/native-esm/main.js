// In a real project this would be `import { CyExcelCore } from "cyexcel/core"`,
// resolved by Node/the browser's own import map against node_modules. This
// repo demos against its own freshly-built dist/ instead of a published
// package, so the relative path stands in for that.
import { CyExcelCore, matrixToWorkbookSnapshot } from "../../dist/core.js";

// No explicit headers means Excel-style A/B/C column letters.
const data = [
  ["A100", "Widget", 9.99],
  ["A101", "Gadget", 14.5],
  ["A102", "Doohickey", 3.25]
];

// The sheet takes a snapshot, not a row array. This is a one-shot import: once
// it returns, `data` and the sheet share nothing — editing one never moves the
// other.
new CyExcelCore(document.getElementById("sheet"), {
  snapshot: matrixToWorkbookSnapshot(data)
});
