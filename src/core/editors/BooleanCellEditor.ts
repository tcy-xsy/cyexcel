import { SELECTION_STROKE_COLOR, SELECTION_STROKE_WIDTH } from "../render/theme";
import type { CellEditor, CellEditorContext } from "./CellEditor";
import { positionEditorElement } from "./positionEditorElement";

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
      positionEditorElement(wrapperElement, context.rect);
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
      // Toggle without letting the press move focus out of the editor.
      //
      // On macOS, clicking a checkbox does not focus it — WebKit follows that
      // platform convention and blurs instead. EditorManager treats a focusout
      // that leaves the editor container as "commit", so on Safari merely
      // clicking this checkbox committed the *old* value and tore the editor
      // down; the trace was pointerdown@INPUT -> focusout@INPUT -> the grid
      // canvas taking focus, with the click event never completing on the
      // input at all. Chromium focuses the control on press, which is why the
      // bug only surfaced once the browser matrix was actually run.
      //
      // preventDefault() on mousedown suppresses only the focus change; the
      // checkbox's activation behaviour runs on click and still toggles
      // normally, so the state is left to the platform. (Toggling by hand here
      // as well double-toggles: the manual flip and the native click cancel
      // out, and the cell commits its original value.)
      checkboxElement.addEventListener("mousedown", (event) => {
        event.preventDefault();
        checkboxElement.focus();
      });
      // No focus ring on the checkbox itself. Unlike the grid canvas — where
      // suppressing the ring would have left nothing indicating focus — the
      // editor wrapper above already draws a solid selection-coloured border
      // for as long as the editor is open, so the composite still shows both
      // that the editor has focus and which cell it belongs to. A second ring
      // inside that border only reads as visual noise in a cell-sized box.
      checkboxElement.style.outline = "none";
      wrapperElement.appendChild(checkboxElement);

      container.appendChild(wrapperElement);
      wrapper = wrapperElement;
      checkbox = checkboxElement;
    },
    updateGeometry(rect): void {
      if (wrapper) {
        positionEditorElement(wrapper, rect);
      }
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
