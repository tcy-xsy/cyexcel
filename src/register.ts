import { defineCyExcel } from "./element/defineCyExcel";

if (typeof globalThis.customElements !== "undefined") {
  defineCyExcel();
}

export { defineCyExcel } from "./element/defineCyExcel";
