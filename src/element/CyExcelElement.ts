import { createCyExcelCore, type CyExcelCore } from "../core";

export function createCyExcelElementConstructor(): CustomElementConstructor {
  return class CyExcelElement extends HTMLElement {
    readonly #core: CyExcelCore;

    constructor() {
      super();
      const root = this.attachShadow({ mode: "open" });
      const viewport = document.createElement("div");
      viewport.part.add("viewport");
      viewport.setAttribute("role", "grid");
      root.append(viewport);
      // No columns/data yet — the attribute/property configuration entry
      // point is Step 13's job. This renders an empty, but fully wired,
      // table until then.
      this.#core = createCyExcelCore(viewport);
    }

    get core(): CyExcelCore {
      return this.#core;
    }

    dispose(): void {
      this.#core.dispose();
    }
  };
}
