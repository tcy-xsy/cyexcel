import { SELECTION_STROKE_COLOR, SELECTION_STROKE_WIDTH } from "../render/theme";
import type { CellEditor, CellEditorContext } from "./CellEditor";

/**
 * Optional built-in editor for boolean-shaped values — a native
 * `<input type="checkbox">`. Not wired to any CellType automatically; a
 * column opts into it explicitly via `createEditor()` (see doc/architecture.md
 * §8). Reads `context.initialValue` directly (not `initialText`) since a
 * checkbox needs a real boolean to set its checked state, not text.
 */
export function createBooleanCellEditor(): CellEditor {
  let wrapper: HTMLDivElement | null = null;
  let checkbox: HTMLInputElement | null = null;

  return {
    mount(container: HTMLElement, context: CellEditorContext): void {
      const wrapperElement = document.createElement("div");
      wrapperElement.style.position = "absolute";
      wrapperElement.style.left = `${context.rect.x}px`;
      wrapperElement.style.top = `${context.rect.y}px`;
      wrapperElement.style.width = `${context.rect.width}px`;
      wrapperElement.style.height = `${context.rect.height}px`;
      wrapperElement.style.boxSizing = "border-box";
      wrapperElement.style.display = "flex";
      wrapperElement.style.alignItems = "center";
      wrapperElement.style.justifyContent = "flex-start";
      wrapperElement.style.paddingLeft = "4px";
      wrapperElement.style.background = "#ffffff";
      wrapperElement.style.border = `${SELECTION_STROKE_WIDTH}px solid ${SELECTION_STROKE_COLOR}`;

      const checkboxElement = document.createElement("input");
      checkboxElement.type = "checkbox";
      checkboxElement.checked = Boolean(context.initialValue);
      wrapperElement.appendChild(checkboxElement);

      container.appendChild(wrapperElement);
      wrapper = wrapperElement;
      checkbox = checkboxElement;
    },
    focus(): void {
      checkbox?.focus();
    },
    getValue(): unknown {
      return checkbox?.checked ?? false;
    },
    destroy(): void {
      wrapper?.remove();
      wrapper = null;
      checkbox = null;
    }
  };
}
