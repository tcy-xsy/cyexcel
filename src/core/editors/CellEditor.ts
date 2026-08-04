import type { CellRect } from "../layout";

export interface CellEditorContext {
  readonly row: number;
  readonly column: number;
  readonly initialValue: unknown;
  /** The initial typed character when editing was started by typing, if any. */
  readonly initialInput?: string | undefined;
  /** Logical rect from LayoutEngine — editors position themselves from this, not their own math. */
  readonly rect: CellRect;
}

export interface ValidationResult {
  readonly valid: boolean;
  readonly message?: string;
}

/**
 * Matches doc/architecture.md §7's published shape. `mount`/`focus`/
 * `getValue`/`destroy` are called synchronously and exactly once each per
 * edit; `validate` is optional and, when present, is checked before commit.
 */
export interface CellEditor {
  mount(container: HTMLElement, context: CellEditorContext): void;
  focus(): void;
  getValue(): unknown;
  validate?(): ValidationResult;
  destroy(): void;
}

export type CellEditorFactory = () => CellEditor;
