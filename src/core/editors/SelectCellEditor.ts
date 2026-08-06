import { CELL_FONT, SELECTION_STROKE_COLOR, SELECTION_STROKE_WIDTH } from "../render/theme";
import type { CellEditor, CellEditorContext } from "./CellEditor";

export interface SelectCellEditorOption {
  readonly value: string;
  /** Defaults to `value` when not given. */
  readonly label?: string;
}

/**
 * Optional built-in editor for enumerated values — a native `<select>`.
 * Takes its option list as a plain constructor argument rather than reading
 * `context.column.options` itself, so this stays independent of the `cells`
 * module and of any particular CellType — whoever calls `createEditor()`
 * decides where the options come from. Not wired to any CellType
 * automatically.
 */
export function createSelectCellEditor(options: readonly SelectCellEditorOption[]): CellEditor {
  let select: HTMLSelectElement | null = null;

  return {
    mount(container: HTMLElement, context: CellEditorContext): void {
      const element = document.createElement("select");
      element.style.position = "absolute";
      element.style.left = `${context.rect.x}px`;
      element.style.top = `${context.rect.y}px`;
      element.style.width = `${context.rect.width}px`;
      element.style.height = `${context.rect.height}px`;
      element.style.boxSizing = "border-box";
      element.style.border = `${SELECTION_STROKE_WIDTH}px solid ${SELECTION_STROKE_COLOR}`;
      element.style.outline = "none";
      element.style.background = "#ffffff";
      element.style.font = CELL_FONT;

      for (const option of options) {
        const optionElement = document.createElement("option");
        optionElement.value = option.value;
        optionElement.textContent = option.label ?? option.value;
        element.appendChild(optionElement);
      }

      const currentValue =
        context.initialValue === undefined || context.initialValue === null
          ? ""
          : String(context.initialValue);
      // A native <select> with no matching <option selected> defaults to
      // the first option — silently changing the value on an "open, don't
      // touch anything, confirm" edit. If the current value isn't in the
      // configured list, add a temporary option for it so it round-trips.
      if (!options.some((option) => option.value === currentValue)) {
        const fallbackOption = document.createElement("option");
        fallbackOption.value = currentValue;
        fallbackOption.textContent = currentValue;
        element.appendChild(fallbackOption);
      }
      element.value = currentValue;

      container.appendChild(element);
      select = element;
    },
    focus(): void {
      select?.focus();
    },
    getValue(): unknown {
      return select?.value ?? "";
    },
    destroy(): void {
      select?.remove();
      select = null;
    }
  };
}
