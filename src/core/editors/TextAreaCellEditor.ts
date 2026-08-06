import { CELL_FONT, CELL_PADDING, SELECTION_STROKE_COLOR, SELECTION_STROKE_WIDTH } from "../render/theme";
import type { CellEditor, CellEditorContext, ValidationResult } from "./CellEditor";

export interface TextAreaCellEditorOptions {
  /**
   * Optional syntax check run on commit — this editor has no built-in
   * assumption about what the text should look like (JSON, CSV, anything
   * else); the caller decides. Not providing one means no validation at all.
   */
  readonly validate?: (text: string) => ValidationResult;
}

/**
 * Optional built-in editor for values too long/structured to comfortably
 * edit in a single-line `<input>` — a native multi-line `<textarea>`. Not
 * wired to any CellType automatically.
 */
export function createTextAreaCellEditor(options: TextAreaCellEditorOptions = {}): CellEditor {
  let textarea: HTMLTextAreaElement | null = null;

  const base = {
    mount(container: HTMLElement, context: CellEditorContext): void {
      const element = document.createElement("textarea");
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
      element.style.resize = "none";
      element.style.background = "#ffffff";
      element.style.font = CELL_FONT;
      container.appendChild(element);
      textarea = element;
    },
    focus(): void {
      textarea?.focus();
      textarea?.select();
    },
    getValue(): unknown {
      return textarea?.value ?? "";
    },
    destroy(): void {
      textarea?.remove();
      textarea = null;
    }
  };

  const editor: CellEditor = options.validate
    ? { ...base, validate: () => options.validate?.(textarea?.value ?? "") as ValidationResult }
    : base;

  return editor;
}
