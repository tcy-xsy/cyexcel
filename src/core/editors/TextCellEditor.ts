import type { CellEditor, CellEditorContext } from "./CellEditor";

/** Built-in default editor: a plain absolutely-positioned <input>. */
export function createTextCellEditor(): CellEditor {
  let input: HTMLInputElement | null = null;

  return {
    mount(container: HTMLElement, context: CellEditorContext): void {
      const element = document.createElement("input");
      element.type = "text";
      element.value =
        context.initialInput ??
        (context.initialValue === undefined || context.initialValue === null
          ? ""
          : String(context.initialValue));
      element.style.position = "absolute";
      element.style.left = `${context.rect.x}px`;
      element.style.top = `${context.rect.y}px`;
      element.style.width = `${context.rect.width}px`;
      element.style.height = `${context.rect.height}px`;
      element.style.boxSizing = "border-box";
      container.appendChild(element);
      input = element;
    },
    focus(): void {
      input?.focus();
      input?.select();
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
