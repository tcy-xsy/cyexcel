export interface CellRange {
  readonly startRow: number;
  readonly startColumn: number;
  readonly endRow: number;
  readonly endColumn: number;
}

export type RowData = Record<string, unknown>;

export interface DataSource {
  getRowCount(): number;
  getColumnCount(): number;
  getCell(row: number, column: number): unknown;
  setCell(row: number, column: number, value: unknown): void;
  getRows?(start: number, end: number): RowData[] | Promise<RowData[]>;
  /**
   * Optional structural mutation capability (Step 12) — a remote/computed
   * DataSource is free to leave all three unimplemented; TableStore checks
   * for their presence and throws a clear capability error instead of
   * assuming they exist. `rows` uses the same "one array of values per
   * row, in column order" convention as `setRange()`/`captureRange()`, not
   * `RowData` — field-name mapping is ArrayDataSource's own concern, never
   * TableStore's or a Command's.
   */
  insertRows?(start: number, rows: readonly (readonly unknown[])[]): void;
  deleteRows?(start: number, count: number): void;
  moveRows?(start: number, count: number, target: number): void;
}

/**
 * Minimal shape shared by every Command; CommandManager (Step 3) will add
 * fields as they become necessary, not replace this one.
 */
export interface CommandContext {
  readonly transactionId: string;
}

export type Patch =
  | { readonly type: "cell-value"; readonly row: number; readonly column: number }
  | { readonly type: "cell-style"; readonly row: number; readonly column: number }
  | { readonly type: "range-value"; readonly range: CellRange }
  | {
      readonly type: "row-structure";
      readonly operation: "insert" | "delete" | "move";
      readonly start: number;
      readonly count: number;
      readonly target?: number;
    }
  | { readonly type: "column-layout"; readonly column: number }
  | { readonly type: "full-reset" };

export interface Command {
  readonly type: string;
  execute(context: CommandContext): Patch | Patch[];
  undo?(context: CommandContext): Patch | Patch[];
}

export type Disposer = () => void;

/**
 * Shared vocabulary for the Hook system. Concrete hook names, execution
 * semantics (series/bail/waterfall/parallel) and contexts are defined in
 * Step 11's HookManager, not here.
 */
export type HookHandler<TContext = unknown> = (context: TContext) => void;

/**
 * Shared vocabulary for the Event system. Concrete event names and payloads
 * are defined per module as each one lands.
 */
export type EventListener<TPayload = unknown> = (payload: TPayload) => void;

/**
 * Minimal subset of the public constructor options. Fields depending on
 * modules that do not exist yet (columns, hooks, plugins, clipboard,
 * contextMenu, history, theme) are added incrementally as those modules
 * land — see doc/api-design.md.
 */
export interface CyExcelOptions {
  readonly dataSource?: DataSource;
  readonly data?: readonly RowData[];
  readonly rowHeight?: number;
  readonly headerHeight?: number;
  readonly readonly?: boolean;
  readonly freeze?: {
    readonly rows?: number;
    readonly columns?: number;
  };
  readonly overscanRows?: number;
  readonly overscanColumns?: number;
}
