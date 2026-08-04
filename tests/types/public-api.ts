import {
  createCyExcelCore,
  defineCyExcel,
  type CyExcelCoreOptions
} from "../../src";
import type { CyExcelCore } from "../../src/core";

declare const container: HTMLElement;

const options = { rowHeight: 24 } satisfies CyExcelCoreOptions;
const core: CyExcelCore = createCyExcelCore(container, options);
const defineResult: CustomElementConstructor | undefined =
  typeof customElements === "undefined" ? undefined : defineCyExcel();

void core;
void defineResult;
