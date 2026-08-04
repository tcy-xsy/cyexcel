import type { LayoutEngine } from "../layout";
import type { CellEditor, CellEditorFactory } from "./CellEditor";
import { createTextCellEditor } from "./TextCellEditor";

export interface EditorManagerOptions {
  /** DOM node editors mount into; must already be connected to the document. */
  readonly container: HTMLElement;
  readonly layout: LayoutEngine;
  readonly getCellValue: (row: number, column: number) => unknown;
  readonly commitValue: (
    row: number,
    column: number,
    value: unknown
  ) => void | Promise<void>;
  readonly editorFactory?: CellEditorFactory;
  readonly onEditStart?: (row: number, column: number) => void;
  readonly onEditCommit?: (row: number, column: number) => void;
  readonly onEditCancel?: (row: number, column: number) => void;
}

interface ActiveEdit {
  readonly row: number;
  readonly column: number;
  readonly editor: CellEditor;
}

/**
 * Owns the single active DOM cell editor. Knows nothing about
 * TableStore/CommandManager — commitValue() is how the integration layer
 * turns a commit into a SetCellCommand.
 */
export class EditorManager {
  readonly #container: HTMLElement;
  readonly #layout: LayoutEngine;
  readonly #getCellValue: (row: number, column: number) => unknown;
  readonly #commitValue: (
    row: number,
    column: number,
    value: unknown
  ) => void | Promise<void>;
  readonly #editorFactory: CellEditorFactory;
  readonly #onEditStart: ((row: number, column: number) => void) | undefined;
  readonly #onEditCommit: ((row: number, column: number) => void) | undefined;
  readonly #onEditCancel: ((row: number, column: number) => void) | undefined;

  #active: ActiveEdit | null = null;
  #composing = false;

  constructor(options: EditorManagerOptions) {
    this.#container = options.container;
    this.#layout = options.layout;
    this.#getCellValue = options.getCellValue;
    this.#commitValue = options.commitValue;
    this.#editorFactory = options.editorFactory ?? createTextCellEditor;
    this.#onEditStart = options.onEditStart;
    this.#onEditCommit = options.onEditCommit;
    this.#onEditCancel = options.onEditCancel;

    this.#container.addEventListener("keydown", this.#handleKeyDown);
    this.#container.addEventListener("compositionstart", this.#handleCompositionStart);
    this.#container.addEventListener("compositionend", this.#handleCompositionEnd);
    this.#container.addEventListener("focusout", this.#handleFocusOut);
  }

  isEditing(): boolean {
    return this.#active !== null;
  }

  startEdit(row: number, column: number, initialInput?: string): void {
    if (this.#active) {
      return;
    }

    const editor = this.#editorFactory();
    const rect = this.#layout.getCellRect(row, column);
    const initialValue = this.#getCellValue(row, column);

    editor.mount(this.#container, { row, column, initialValue, initialInput, rect });
    editor.focus();

    this.#active = { row, column, editor };
    this.#onEditStart?.(row, column);
  }

  cancelEdit(): void {
    if (!this.#active) {
      return;
    }
    const { row, column, editor } = this.#active;
    this.#active = null;
    this.#composing = false;
    editor.destroy();
    this.#onEditCancel?.(row, column);
  }

  async commitEdit(): Promise<void> {
    if (!this.#active) {
      return;
    }
    const { row, column, editor } = this.#active;

    const validation = editor.validate?.();
    if (validation && !validation.valid) {
      // Keep the editor, its value, and focus — the user needs to fix it.
      return;
    }

    const value = editor.getValue();
    await this.#commitValue(row, column, value);

    // #active may have been cleared by a concurrent cancelEdit()/dispose()
    // while awaiting commitValue — only tear down if this is still current.
    if (this.#active?.editor === editor) {
      this.#active = null;
      this.#composing = false;
      editor.destroy();
      this.#onEditCommit?.(row, column);
    }
  }

  dispose(): void {
    if (this.#active) {
      const { row, column, editor } = this.#active;
      this.#active = null;
      this.#composing = false;
      editor.destroy();
      this.#onEditCancel?.(row, column);
    }
    this.#container.removeEventListener("keydown", this.#handleKeyDown);
    this.#container.removeEventListener("compositionstart", this.#handleCompositionStart);
    this.#container.removeEventListener("compositionend", this.#handleCompositionEnd);
    this.#container.removeEventListener("focusout", this.#handleFocusOut);
  }

  #handleKeyDown = (event: KeyboardEvent): void => {
    if (!this.#active) {
      return;
    }
    if (event.key === "Escape") {
      event.preventDefault();
      this.cancelEdit();
      return;
    }
    if (event.key === "Enter" && !this.#composing) {
      event.preventDefault();
      void this.commitEdit();
    }
  };

  #handleCompositionStart = (): void => {
    this.#composing = true;
  };

  #handleCompositionEnd = (): void => {
    this.#composing = false;
  };

  #handleFocusOut = (event: FocusEvent): void => {
    if (!this.#active) {
      return;
    }
    const next = event.relatedTarget as Node | null;
    if (next && this.#container.contains(next)) {
      return;
    }
    void this.commitEdit();
  };
}
