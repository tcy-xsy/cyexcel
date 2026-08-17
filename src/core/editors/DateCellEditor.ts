import { CELL_FONT, CELL_PADDING, SELECTION_STROKE_COLOR, SELECTION_STROKE_WIDTH } from "../render/theme";
import type { CellEditor, CellEditorContext } from "./CellEditor";
import { positionEditorElement } from "./positionEditorElement";

function toDateInputValue(value: unknown): string {
  return value instanceof Date && !Number.isNaN(value.getTime()) ? value.toISOString().slice(0, 10) : "";
}

/**
 * Optional built-in editor for date-shaped values — a native
 * `<input type="date">`, using the browser's own date picker instead of a
 * hand-built calendar widget. Not wired to any CellType automatically; a
 * column opts into it explicitly via `createEditor()`.
 */
export function createDateCellEditor(): CellEditor {
  let input: HTMLInputElement | null = null;

  return {
    mount(container: HTMLElement, context: CellEditorContext): void {
      const element = document.createElement("input");
      element.type = "date";
      element.value = toDateInputValue(context.initialValue);
      element.style.position = "absolute";
      positionEditorElement(element, context.rect);
      element.style.boxSizing = "border-box";
      element.style.margin = "0";
      element.style.padding = `0 ${CELL_PADDING}px`;
      element.style.border = `${SELECTION_STROKE_WIDTH}px solid ${SELECTION_STROKE_COLOR}`;
      element.style.outline = "none";
      element.style.background = "#ffffff";
      element.style.font = CELL_FONT;
      container.appendChild(element);
      input = element;
    },
    updateGeometry(rect): void {
      if (input) {
        positionEditorElement(input, rect);
      }
    },
    focus(): void {
      input?.focus();
    },
    getValue(): unknown {
      return input?.value ? new Date(input.value) : null;
    },
    destroy(): void {
      input?.remove();
      input = null;
    }
  };
}
