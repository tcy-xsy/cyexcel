import type { CellRect } from "../layout";
import type { SceneSnapshot } from "../scene";
import type { CellRange } from "../types";
import type { CellEditor, CellEditorFactory } from "./CellEditor";
import { createTextCellEditor } from "./TextCellEditor";
import type { CellValue } from "../sheet/types";
import { ACCEPT, type EditValidationOutcome, type ValidationPresenter } from "./editValidation";
import { createDomValidationPresenter } from "./ValidationBubble";

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
   * decides whether to commit, refuse, ask first, or commit and tell the user
   * (数据验证设计稿 §7.1). Omitted means "always accept". Only this path is
   * gated — paste, fill and `setCell()` mark invalid cells but never refuse.
   */
  readonly validateCommit?: (row: number, column: number, value: CellValue | undefined) => EditValidationOutcome;
  /** How refusals/confirmations are shown. Defaults to a DOM bubble in `container`, created on first use. */
  readonly validationPresenter?: ValidationPresenter;
}

interface ActiveEdit {
  readonly row: number;
  readonly column: number;
  readonly editor: CellEditor;
  /** Latest cell rect, kept so the validation bubble can sit under the cell. */
  rect: CellRect;
}

function intersectRect(rect: CellRect, clipRect: CellRect): CellRect | null {
  const left = Math.max(rect.x, clipRect.x);
  const top = Math.max(rect.y, clipRect.y);
  const right = Math.min(rect.x + rect.width, clipRect.x + clipRect.width);
  const bottom = Math.min(rect.y + rect.height, clipRect.y + clipRect.height);
  if (right <= left || bottom <= top) {
    return null;
  }
  return { x: left, y: top, width: right - left, height: bottom - top };
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
  /** A warning-level "continue anyway?" prompt is open: further commit attempts wait for its answer. */
  #confirming = false;

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
    this.#container.addEventListener("input", this.#handleInput);
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
    active.rect = rect;
    this.#presenter?.reposition(rect);
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

    this.#active = { row, column, editor, rect };
    this.#onEditStart?.(row, column);
  }

  cancelEdit(): void {
    if (!this.#active) {
      return;
    }
    const { row, column, editor } = this.#active;
    this.#active = null;
    this.#composing = false;
    this.#confirming = false;
    this.#presenter?.hide();
    editor.destroy();
    this.#onEditCancel?.(row, column);
  }

  async commitEdit(): Promise<void> {
    if (!this.#active || this.#confirming) {
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
      // 停止：什么都不提交，编辑器、输入值和焦点原样留着，让用户改。
      this.#present(active, outcome);
      editor.focus();
      return;
    }

    if (outcome.kind === "confirm") {
      // 警告：等用户回答。回答之前再来的提交（比如失焦）都被 #confirming 挡掉。
      this.#confirming = true;
      this.#present(active, outcome, {
        onContinue: () => {
          this.#confirming = false;
          this.#presenter?.hide();
          void this.#finishCommit(active, value);
        },
        onCancel: () => {
          this.#confirming = false;
          this.#presenter?.hide();
          editor.focus();
        }
      });
      return;
    }

    await this.#finishCommit(active, value);
    if (outcome.kind === "notify") {
      // 信息：先提交再说。编辑器已经销毁，提示只需要那一格的位置。
      this.#present(active, outcome);
    }
  }

  async #finishCommit(active: ActiveEdit, value: CellValue | undefined): Promise<void> {
    const { row, column, editor } = active;
    // 之前那次被拒绝的提示，这次能提交说明用户已经改对了
    this.#presenter?.hide();
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

  #present(
    active: ActiveEdit,
    outcome: Exclude<EditValidationOutcome, { readonly kind: "accept" }>,
    handlers: { readonly onContinue?: () => void; readonly onCancel?: () => void } = {}
  ): void {
    this.#presenter ??= createDomValidationPresenter(this.#container);
    this.#presenter.show({
      kind: outcome.kind,
      ...(outcome.title !== undefined ? { title: outcome.title } : {}),
      message: outcome.message,
      rect: active.rect,
      ...handlers
    });
  }

  dispose(): void {
    if (this.#active) {
      const { row, column, editor } = this.#active;
      this.#active = null;
      this.#composing = false;
      editor.destroy();
      this.#onEditCancel?.(row, column);
    }
    this.#confirming = false;
    this.#presenter?.dispose();
    this.#container.removeEventListener("keydown", this.#handleKeyDown);
    this.#container.removeEventListener("input", this.#handleInput);
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

  /** 用户又改了内容：之前的拒绝/确认提示作废，下一次提交重新校验。 */
  #handleInput = (): void => {
    if (!this.#active) {
      return;
    }
    this.#confirming = false;
    this.#presenter?.hide();
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
