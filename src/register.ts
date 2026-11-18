import { defineCyExcel } from "./element/defineCyExcel.js";

if (typeof globalThis.customElements !== "undefined") {
  defineCyExcel();
}

export { defineCyExcel } from "./element/defineCyExcel.js";
