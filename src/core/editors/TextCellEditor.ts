import { CELL_FONT, CELL_PADDING, SELECTION_STROKE_COLOR, SELECTION_STROKE_WIDTH } from "../render/theme";
import type { CellEditor, CellEditorContext } from "./CellEditor";

/**
 * Built-in default editor: a plain absolutely-positioned <input>, styled to
 * sit flush over the cell it replaces — same border color/weight as the
 * selection box (SelectionLayerRenderer hides its own box while editing, so
 * this border is the only selection indicator visible during edit), same
 * padding/font as ContentLayerRenderer's default painter so the text
 * doesn't visibly shift when entering/exiting edit mode.
 */
export function createTextCellEditor(): CellEditor {
  let input: HTMLInputElement | null = null;

  return {
    mount(container: HTMLElement, context: CellEditorContext): void {
      const element = document.createElement("input");
      element.type = "text";
      // initialInput wins when editing started by typing a character — that
      // keystroke should replace the display entirely, not be layered on
      // top of the old value's text. initialText (from the column's
      // CellType.serialize(), when the integration layer provides one) is
      // next — a plain String(initialValue) mangles non-string values like
      // Date/array/object (see doc/progress.md's patch for this bug).
      element.value =
        context.initialInput ??
        context.initialText ??
        (context.initialValue === undefined || context.initialValue === null
          ? ""
          : String(context.initialValue));
      element.style.position = "absolute";
      element.style.left = `${context.rect.x}px`;
      element.style.top = `${context.rect.y}px`;
      element.style.width = `${context.rect.width}px`;
      element.style.height = `${context.rect.height}px`;
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
    focus(): void {
      // Deliberately not input?.select() — full-selecting the existing
      // value meant double-click/Enter into an existing value required
      // first deselecting before you could append or make a precise edit.
      // Leaving the cursor at the end (the native post-assignment default)
      // matches Excel's double-click/F2 behavior: typing continues the
      // value instead of replacing all of it.
      input?.focus();
    },
    getValue(): unknown {
      return input?.value ?? "";
    },
    destroy(): void {
      input?.remove();
      input = null;
    }
  };
}
