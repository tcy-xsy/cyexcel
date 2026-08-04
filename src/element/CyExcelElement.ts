import { createCyExcelCore, type CyExcelCore } from "../core";

export function createCyExcelElementConstructor(): CustomElementConstructor {
  return class CyExcelElement extends HTMLElement {
    readonly #core: CyExcelCore;

    constructor() {
      super();
      this.#core = createCyExcelCore();
      const root = this.attachShadow({ mode: "open" });
      const viewport = document.createElement("div");
      viewport.part.add("viewport");
      viewport.setAttribute("role", "grid");
      root.append(viewport);
    }

    get core(): CyExcelCore {
      return this.#core;
    }

    dispose(): void {
      this.#core.dispose();
    }
  };
}
