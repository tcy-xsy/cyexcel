import {
  createCyExcelCore,
  defineCyExcel,
  type CyExcelCoreOptions
} from "../../src";
import type { CyExcelCore } from "../../src/core";

const options = { id: "orders" } satisfies CyExcelCoreOptions;
const core: CyExcelCore = createCyExcelCore(options);
const defineResult: CustomElementConstructor | undefined =
  typeof customElements === "undefined" ? undefined : defineCyExcel();

void core;
void defineResult;
