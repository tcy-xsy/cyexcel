import { createCyExcelElementConstructor } from "./CyExcelElement";

export const DEFAULT_CYEXCEL_TAG_NAME = "cy-excel";

const constructors = new Map<string, CustomElementConstructor>();

export function defineCyExcel(
  tagName = DEFAULT_CYEXCEL_TAG_NAME
): CustomElementConstructor {
  if (typeof customElements === "undefined" || typeof HTMLElement === "undefined") {
    throw new Error("CYEXCEL_DOM_UNAVAILABLE");
  }

  const knownConstructor = constructors.get(tagName);
  const registeredConstructor = customElements.get(tagName);

  if (registeredConstructor) {
    if (knownConstructor && registeredConstructor === knownConstructor) {
      return registeredConstructor;
    }
    throw new Error(`CYEXCEL_TAG_ALREADY_DEFINED: ${tagName}`);
  }

  const elementConstructor = knownConstructor ?? createCyExcelElementConstructor();
  constructors.set(tagName, elementConstructor);
  customElements.define(tagName, elementConstructor);
  return elementConstructor;
}
