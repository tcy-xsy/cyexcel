import type { RuntimeColumn } from "../cells";
import type { CellRange } from "../types";
import { ClipboardWorkerClient, type SerializeWorkerLike } from "./ClipboardWorkerClient";
import { CopyPlanner } from "./CopyPlanner";
import { buildCellGrid } from "./serializers";
import type { CopiedStyles } from "./styleClipboard";
import { decodeTsv } from "./tsv";
import type { ClipboardCopyOptions, ClipboardCopyResult, ClipboardOptions, ClipboardPasteOptions, CopyPlan } from "./types";
import type { CellValue } from "../sheet/types";

const DEFAULT_MAX_CELLS = 100_000;
const DEFAULT_MAX_BYTES = 16 * 1024 * 1024;

export interface ClipboardIO {
  writeText(text: string): Promise<void>;
  readText(): Promise<string>;
  download(blob: Blob, filename: string): void;
}

export interface ClipboardServiceOptions {
  readonly getCellValue: (row: number, column: number) => CellValue | undefined;
  readonly getColumn: (column: number) => RuntimeColumn;
  readonly getRowCount: () => number;
  readonly getColumnCount: () => number;
  readonly parseValue: (row: number, column: number, text: string) => CellValue;
  readonly getSelectionRange: () => CellRange;
  readonly getStoreVersion: () => number;
  /** Expected to be a single Command execution (e.g. one SetRangeCommand), so one paste = one transaction = one History entry. */
  readonly applyRangeCommand: (
    range: CellRange,
    values: readonly (readonly (CellValue | null | undefined)[])[],
    /** 认领成功时带上的样式——值与样式必须落在同一个事务里，否则撤销要按两次。 */
    styles?: CopiedStyles
  ) => void | Promise<void>;
  readonly io: ClipboardIO;
  readonly createWorker?: () => SerializeWorkerLike;
  readonly options?: ClipboardOptions;
  readonly findMergeAt?: (row: number, column: number) => CellRange | undefined;
}

/**
 * Orchestrates plan -> serialize -> write for copy, and read -> parse ->
 * commit for paste. Deliberately depends only on plain callbacks/narrow
 * interfaces (no TableStore/CommandManager/DOM types) so it's fully
 * Node-testable — CyExcelCore wires the real Store/Command/browser
 * Clipboard implementations in.
 */
export class ClipboardService {
  readonly #options: ClipboardServiceOptions;
  readonly #planner: CopyPlanner;
  readonly #worker: ClipboardWorkerClient;
  readonly #maxCells: number;
  readonly #maxBytes: number;
  readonly #oversizedAction: "confirm-or-download" | "download" | "block";

  constructor(options: ClipboardServiceOptions) {
    this.#options = options;
    this.#maxCells = options.options?.maxCells ?? DEFAULT_MAX_CELLS;
    this.#maxBytes = options.options?.maxBytes ?? DEFAULT_MAX_BYTES;
    this.#oversizedAction = options.options?.oversizedAction ?? "confirm-or-download";
    this.#planner = new CopyPlanner({
      getCellValue: options.getCellValue,
      getColumn: options.getColumn,
      maxCells: this.#maxCells,
      maxBytes: this.#maxBytes,
      ...(options.findMergeAt !== undefined ? { findMergeAt: options.findMergeAt } : {})
    });
    this.#worker = new ClipboardWorkerClient(
      options.createWorker !== undefined ? { createWorker: options.createWorker } : {}
    );
  }

  planCopy(range?: CellRange): CopyPlan {
    return this.#planner.plan(range ?? this.#options.getSelectionRange());
  }

  async copy(opts: ClipboardCopyOptions = {}): Promise<ClipboardCopyResult> {
    const range = opts.range ?? this.#options.getSelectionRange();
    const plan = this.#planner.plan(range);

    if (opts.signal?.aborted) {
      return { status: "cancelled", plan };
    }

    const oversized = plan.recommendedMode === "download";
    if (oversized) {
      if (this.#oversizedAction === "block") {
        return { status: "oversized", plan };
      }
      if (this.#oversizedAction === "download" || opts.download) {
        return this.#downloadCopy(range, opts, plan);
      }
      if (!opts.confirmed) {
        return { status: "oversized", plan };
      }
      // confirmed: fall through to the normal copy path below.
    } else if (opts.download) {
      return this.#downloadCopy(range, opts, plan);
    }

    const grid = buildCellGrid(range, opts.mode ?? "display", {
      getCellValue: this.#options.getCellValue,
      getColumn: this.#options.getColumn,
      ...(this.#options.findMergeAt !== undefined ? { findMergeAt: this.#options.findMergeAt } : {})
    });

    if (opts.signal?.aborted) {
      return { status: "cancelled", plan };
    }

    const storeVersion = this.#options.getStoreVersion();
    const text = await this.#worker.serializeTsv(grid, {
      storeVersion,
      currentVersion: this.#options.getStoreVersion,
      ...(opts.signal !== undefined ? { signal: opts.signal } : {})
    });

    // `text` is undefined both when the caller aborted and when the store
    // moved on to a newer version while the worker was busy (see
    // ClipboardWorkerClient) — either way nothing should be written, and
    // from the caller's perspective the outcome is the same: the copy
    // didn't happen.
    if (opts.signal?.aborted || text === undefined) {
      return { status: "cancelled", plan };
    }

    await this.#options.io.writeText(text);
    return { status: "copied", plan };
  }

  async paste(opts: ClipboardPasteOptions = {}): Promise<void> {
    if (opts.signal?.aborted) {
      return;
    }

    const text = opts.text ?? (await this.#options.io.readText());
    if (opts.signal?.aborted) {
      return;
    }

    const grid = decodeTsv(text);
    const anchor = opts.range ?? this.#options.getSelectionRange();
    const pastedRows = grid.length;
    const pastedColumns = grid[0]?.length ?? 0;

    const endRow = Math.min(anchor.startRow + pastedRows - 1, this.#options.getRowCount() - 1);
    const endColumn = Math.min(anchor.startColumn + pastedColumns - 1, this.#options.getColumnCount() - 1);
    if (endRow < anchor.startRow || endColumn < anchor.startColumn) {
      return;
    }

    const values: (CellValue | null | undefined)[][] = [];
    for (let row = anchor.startRow; row <= endRow; row += 1) {
      const rowValues: (CellValue | null | undefined)[] = [];
      for (let column = anchor.startColumn; column <= endColumn; column += 1) {
        const runtimeColumn = this.#options.getColumn(column);
        if (runtimeColumn.clipboard.paste === false) {
          // Keeps this column's existing value so the whole range can
          // still land in one SetRangeCommand (see applyRangeCommand's
          // doc comment) instead of needing per-cell command support.
          rowValues.push(this.#options.getCellValue(row, column));
          continue;
        }
        const cellText = grid[row - anchor.startRow]?.[column - anchor.startColumn] ?? "";
        rowValues.push(this.#options.parseValue(row, column, cellText));
      }
      values.push(rowValues);
    }

    if (opts.signal?.aborted) {
      return;
    }

    await this.#options.applyRangeCommand(
      { startRow: anchor.startRow, startColumn: anchor.startColumn, endRow, endColumn },
      values,
      opts.styles
    );
  }

  dispose(): void {
    this.#worker.dispose();
  }

  async #downloadCopy(
    range: CellRange,
    opts: ClipboardCopyOptions,
    plan: CopyPlan
  ): Promise<ClipboardCopyResult> {
    const grid = buildCellGrid(range, opts.mode ?? "raw", {
      getCellValue: this.#options.getCellValue,
      getColumn: this.#options.getColumn,
      ...(this.#options.findMergeAt !== undefined ? { findMergeAt: this.#options.findMergeAt } : {})
    });

    if (opts.signal?.aborted) {
      return { status: "cancelled", plan };
    }

    const storeVersion = this.#options.getStoreVersion();
    const blob = await this.#worker.serializeTsvBlob(grid, {
      storeVersion,
      currentVersion: this.#options.getStoreVersion,
      ...(opts.signal !== undefined ? { signal: opts.signal } : {})
    });

    if (opts.signal?.aborted || blob === undefined) {
      return { status: "cancelled", plan };
    }

    this.#options.io.download(blob, `cyexcel-export-${Date.now()}.tsv`);
    return { status: "downloaded", plan };
  }
}
