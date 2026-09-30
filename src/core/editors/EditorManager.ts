import type { SceneSnapshot } from "../scene";
import type { CellRange } from "../types";
import type { CellEditor, CellEditorFactory } from "./CellEditor";
import { createTextCellEditor } from "./TextCellEditor";
import type { CellValue } from "../sheet/types";
import { ACCEPT, type EditValidationOutcome, type ValidationPresenter } from "./editValidation";
import { createDomValidationPresenter } from "./ValidationDialog";
import { intersectRect } from "./visibleCellRect";

export interface EditorManagerOptions {
  /** DOM node editors mount into; must already be connected to the document. */
  readonly container: HTMLElement;
  readonly getScene: () => SceneSnapshot;
  /** Resolves the merged range (if any) owning a cell — the editor overlay spans the whole merge. */
  readonly findMergeAt?: (row: number, column: number) => CellRange | undefined;
  readonly getCellValue: (row: number, column: number) => CellValue | undefined;
  readonly commitValue: (
    row: number,
    column: number,
    value: CellValue | undefined
  ) => void | Promise<void>;
  readonly editorFactory?: CellEditorFactory;
  /**
   * Per-cell editor override, checked before falling back to editorFactory.
   * Returning undefined for a given (row, column) means "use the default" —
   * this is how a column opts into a specific editor (checkbox, date input,
   * ...) without EditorManager itself knowing anything about column types.
   */
  readonly resolveEditor?: (row: number, column: number) => CellEditor | undefined;
  /** Text-editor-only initial value, e.g. from the column's CellType.serialize(). */
  readonly serializeCellValue?: (row: number, column: number, value: CellValue | undefined) => string;
  readonly onEditStart?: (row: number, column: number) => void;
  readonly onEditCommit?: (row: number, column: number) => void;
  readonly onEditCancel?: (row: number, column: number) => void;
  /**
   * Called with the editor's value just before it is committed; the answer
   * decides whether to commit or refuse (数据验证设计稿 §7.1). Only a
   * stop-level rule refuses, like Univer — warning/information values commit
   * and are merely marked. Omitted means "always accept". Only this path is
   * gated — paste, fill and `setCell()` mark invalid cells but never refuse.
   */
  readonly validateCommit?: (row: number, column: number, value: CellValue | undefined) => EditValidationOutcome;
  /** How a refusal is shown. Defaults to a modal dialog on `document.body`, created on first use. */
  readonly validationPresenter?: ValidationPresenter;
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
  readonly #getScene: () => SceneSnapshot;
  readonly #findMergeAt: ((row: number, column: number) => CellRange | undefined) | undefined;
  readonly #getCellValue: (row: number, column: number) => CellValue | undefined;
  readonly #commitValue: (
    row: number,
    column: number,
    value: CellValue | undefined
  ) => void | Promise<void>;
  readonly #editorFactory: CellEditorFactory;
  readonly #resolveEditor: ((row: number, column: number) => CellEditor | undefined) | undefined;
  readonly #serializeCellValue: ((row: number, column: number, value: CellValue | undefined) => string) | undefined;
  readonly #onEditStart: ((row: number, column: number) => void) | undefined;
  readonly #onEditCommit: ((row: number, column: number) => void) | undefined;
  readonly #onEditCancel: ((row: number, column: number) => void) | undefined;
  readonly #validateCommit: EditorManagerOptions["validateCommit"];
  #presenter: ValidationPresenter | undefined;

  #active: ActiveEdit | null = null;
  #composing = false;
  /**
   * The rejection dialog is open (or has just closed and focus is being handed back):
   * further commit attempts are ignored. Opening it steals focus from the editor, and
   * removing its focused OK button fires another `focusout`; both look like "clicked
   * elsewhere" to `#handleFocusOut`, which would re-validate and pop the dialog again.
   */
  #dialogOpen = false;

  constructor(options: EditorManagerOptions) {
    this.#container = options.container;
    this.#getScene = options.getScene;
    this.#findMergeAt = options.findMergeAt;
    this.#getCellValue = options.getCellValue;
    this.#commitValue = options.commitValue;
    this.#editorFactory = options.editorFactory ?? createTextCellEditor;
    this.#resolveEditor = options.resolveEditor;
    this.#serializeCellValue = options.serializeCellValue;
    this.#onEditStart = options.onEditStart;
    this.#onEditCommit = options.onEditCommit;
    this.#onEditCancel = options.onEditCancel;
    this.#validateCommit = options.validateCommit;
    this.#presenter = options.validationPresenter;

    this.#container.addEventListener("keydown", this.#handleKeyDown);
    this.#container.addEventListener("compositionstart", this.#handleCompositionStart);
    this.#container.addEventListener("compositionend", this.#handleCompositionEnd);
    this.#container.addEventListener("focusout", this.#handleFocusOut);
  }

  isEditing(): boolean {
    return this.#active !== null;
  }

  /** Reprojects the active cell through one newly published Scene snapshot. */
  updateGeometry(scene: SceneSnapshot = this.#getScene()): void {
    const active = this.#active;
    if (!active) {
      return;
    }
    const activeMerge = this.#findMergeAt?.(active.row, active.column);
    const activeRange: CellRange = activeMerge ?? {
      startRow: active.row,
      endRow: active.row,
      startColumn: active.column,
      endColumn: active.column
    };
    const fragment = scene.getFragmentsForRange(activeRange)[0];
    const rect = fragment ? intersectRect(fragment.rect, fragment.pane.clipRect) : null;
    if (!rect) {
      this.cancelEdit();
      return;
    }
    active.editor.updateGeometry(Object.freeze(rect));
  }

  startEdit(row: number, column: number, initialInput?: string): void {
    if (this.#active) {
      return;
    }

    const scene = this.#getScene();
    const merge = this.#findMergeAt?.(row, column);
    const range: CellRange = merge ?? { startRow: row, endRow: row, startColumn: column, endColumn: column };
    const fragment = scene.getFragmentsForRange(range)[0];
    if (!fragment) {
      return;
    }
    const rect = intersectRect(fragment.rect, fragment.pane.clipRect);
    if (!rect) {
      return;
    }
    const editor = this.#resolveEditor?.(row, column) ?? this.#editorFactory();
    const initialValue = this.#getCellValue(row, column);
    const initialText = this.#serializeCellValue?.(row, column, initialValue);

    editor.mount(this.#container, { row, column, initialValue, initialInput, initialText, rect });
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
    this.#dialogOpen = false;
    this.#presenter?.hide();
    editor.destroy();
    this.#onEditCancel?.(row, column);
  }

  async commitEdit(): Promise<void> {
    if (!this.#active || this.#dialogOpen) {
      return;
    }
    const active = this.#active;
    const { row, column, editor } = active;

    const validation = editor.validate?.();
    if (validation && !validation.valid) {
      // Keep the editor, its value, and focus — the user needs to fix it.
      return;
    }

    const value = editor.getValue();
    const outcome = this.#validateCommit?.(row, column, value) ?? ACCEPT;

    if (outcome.kind === "reject") {
      // 停止：什么都不提交，弹模态对话框说明原因；关掉之后回到编辑器，输入内容还在。
      this.#dialogOpen = true;
      this.#presenter ??= createDomValidationPresenter();
      this.#presenter.show({
        ...(outcome.title !== undefined ? { title: outcome.title } : {}),
        message: outcome.message,
        onClose: () => {
          // **`#dialogOpen` 必须撑到最后才复位**：关对话框会把持有焦点的「确定」按钮摘掉，
          // 浏览器随即发一次 `focusout`（`relatedTarget` 为空），先复位的话
          // `#handleFocusOut` 会当成"点到别处去了"，重新校验、又弹出同一个对话框。
          this.#presenter?.hide();
          if (this.#active?.editor === editor) {
            editor.focus();
          }
          this.#dialogOpen = false;
        }
      });
      return;
    }

    await this.#finishCommit(active, value);
  }

  async #finishCommit(active: ActiveEdit, value: CellValue | undefined): Promise<void> {
    const { row, column, editor } = active;
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
    this.#dialogOpen = false;
    this.#presenter?.dispose();
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
