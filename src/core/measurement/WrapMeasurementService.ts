import type { RuntimeColumn } from "../cells";
import type { SemanticChange } from "../data";
import type { RowHeightProvenance } from "../layout";
import { computeWrappedCellHeight, type Canvas2DLike } from "../render";
import { CELL_PADDING } from "../render/theme";
import type { SceneSnapshot } from "../scene";

/**
 * Everything the service needs to know about the authoritative state a
 * measurement is derived from. Passed per call rather than held, because a
 * formatter is user code: it can queue a mutation mid-measurement, and the
 * caller must be able to prove afterwards that the basis it measured against
 * is still the basis it is about to write to.
 */
/**
 * 行高来源的只读视图。
 *
 * 以前这里收的是 `LayoutRoot.rowHeightProvenance` 那张 Map。ADR 0004 阶段 4B T8
 * 之后来源的权威是 `SheetModel` 的 `rowData[row].hp`，没有一张现成的 Map 可交
 * ——而为了传参现造一张，就等于又造出一份能与权威分叉的存储，正是这次迁移要
 * 消灭的东西（设计稿 §4.1）。所以收的是**能力**，不是容器：一次点查加一次
 * 「谁是 auto」的遍历，恰好是这个服务用到的全部。
 */
export interface RowHeightProvenanceView {
  provenanceOf(row: number): RowHeightProvenance;
  /** 来源为 `auto` 的行。只遍历真实存在的条目，不遍历坐标空间。 */
  autoRows(): Iterable<number>;
}

export interface WrapMeasurementBasis {
  readonly scene: SceneSnapshot;
  /** Live model version, asserted against `scene.modelVersion`. */
  readonly modelVersion: number;
  /** Live layout root version, asserted against `scene.layoutVersion`. */
  readonly layoutVersion: number;
  readonly defaultRowHeight: number;
  readonly rowHeightProvenance: RowHeightProvenanceView;
}

/**
 * Upper bound on the "row N has been measured" cache.
 *
 * Without one the cache grows with every row ever scrolled past: a wrap-enabled
 * million-row table retains a million entries, and `forgetRowRange()` — which
 * scans the cache rather than the edited span — gets slower the longer the
 * session runs. Both violate Milestone 0's "memory does not keep growing after
 * five minutes of scrolling".
 *
 * Evicting is always safe: a forgotten row is re-measured the next time it is
 * painted, and measurement is idempotent. The only real constraint is that the
 * bound must stay far above one screenful, or rows would be evicted while still
 * visible and re-measured every frame. A 4K-tall viewport at the 16px minimum
 * row height shows ~135 rows, so this is ~70x the worst case.
 */
const MAX_MEASURED_ROWS = 10_000;

export interface RowHeightMeasurement {
  readonly row: number;
  readonly height: number;
}

export interface WrapMeasurementPlan {
  /** Auto-owned rows whose measured height differs from their current one. */
  readonly updates: readonly RowHeightMeasurement[];
  /** Rows that were measured but need no resize — safe to mark measured now. */
  readonly measuredWithoutResize: readonly number[];
}

export interface WrapMeasurementServiceOptions {
  /**
   * A 2D context used only for `measureText()`. Never appended to any DOM
   * tree, so measurement cannot disturb a real canvas's drawing state.
   */
  readonly measureContext: Canvas2DLike;
  readonly getColumns: () => readonly RuntimeColumn[];
  readonly getCell: (row: number, column: number) => unknown;
  /** Formatters are user code; a throw is reported, never propagated. */
  readonly onError: (error: unknown, eventName: string) => void;
}

/**
 * Owns wrap-height measurement and the "which rows have been measured" cache.
 *
 * Deliberately knows nothing about commands, the mutation queue, rendering or
 * requestAnimationFrame: it answers *what height does this row need* and
 * *which rows still need asking*, and the caller decides how to commit that.
 * Keeping the split here is what lets a shared formatted-text cache live
 * alongside the measurement that produces it rather than inside CyExcelCore.
 */
export class WrapMeasurementService {
  readonly #measureContext: Canvas2DLike;
  readonly #getColumns: () => readonly RuntimeColumn[];
  readonly #getCell: (row: number, column: number) => unknown;
  readonly #onError: (error: unknown, eventName: string) => void;
  readonly #measuredRows = new Set<number>();

  constructor(options: WrapMeasurementServiceOptions) {
    this.#measureContext = options.measureContext;
    this.#getColumns = options.getColumns;
    this.#getCell = options.getCell;
    this.#onError = options.onError;
  }

  /**
   * Whether any column opted into `wrap: true` and uses the default
   * (non-custom-renderer) paint path — the single cheap check every
   * content-driven auto-height call site does first, so a table that never
   * uses `wrap` pays nothing beyond this one `.some()` scan, never a per-row
   * or per-edit measurement.
   */
  hasWrapColumns(): boolean {
    return this.#getColumns().some((column) => column.wrap && column.renderer === undefined);
  }

  /**
   * The tallest height any wrap-enabled column's content needs for `row`, or
   * `defaultRowHeight` if none do. Never smaller than the default.
   * `column.renderer !== undefined` columns are skipped — a custom canvas
   * renderer paints itself and a DOM renderer is owned entirely by
   * `DOMCellManager`; `wrap` only has meaning for the built-in
   * `drawFormattedText()` path.
   */
  computeRequiredRowHeight(row: number, scene: SceneSnapshot, defaultRowHeight: number): number {
    const columns = this.#getColumns();
    let required = defaultRowHeight;
    for (let index = 0; index < columns.length; index += 1) {
      const column = columns[index];
      if (!column || !column.wrap || column.renderer !== undefined) {
        continue;
      }
      const value = this.#getCell(row, index);
      const text = column.format(value, { row, column });
      const columnWidth = scene.getColumnWidth(index);
      if (columnWidth <= CELL_PADDING * 2) {
        continue;
      }
      const maxWidth = columnWidth - CELL_PADDING * 2;
      const height = computeWrappedCellHeight(this.#measureContext, text, maxWidth);
      if (height > required) {
        required = height;
      }
    }
    return required;
  }

  /**
   * Rows about to be painted (visible + overscan, per the Scene's own body
   * panes) that have not been measured yet. Virtual traversal of the published
   * panes — never an expansion of a mutated range, so a million-row paste does
   * not synchronously build a million-entry set.
   */
  visibleUnmeasuredRows(scene: SceneSnapshot): readonly number[] {
    const rows = new Set<number>();
    for (const pane of scene.getPanes("body")) {
      for (let index = 0; index < pane.rows.count; index += 1) {
        const row = pane.rows.start + index;
        if (!this.#measuredRows.has(row)) {
          rows.add(row);
        }
      }
    }
    return [...rows];
  }

  /**
   * Auto-owned rows still sitting at a non-default height when nothing wraps
   * any more — deleting the last wrap column has to shrink them back, and
   * they are not found by `visibleUnmeasuredRows()` because they may well
   * have been measured already.
   */
  rowsNeedingDefaultReset(
    scene: SceneSnapshot,
    rowHeightProvenance: RowHeightProvenanceView,
    defaultRowHeight: number
  ): readonly number[] {
    const rows: number[] = [];
    for (const row of rowHeightProvenance.autoRows()) {
      if (scene.getRowHeight(row) !== defaultRowHeight) {
        rows.push(row);
      }
    }
    return rows;
  }

  /**
   * Which rows a freshly published Scene needs measured, and how many of them
   * came from the visible scan.
   *
   * Assembling this was previously the caller's job: it had to know to skip the
   * visible scan when nothing wraps, *and* to remember that deleting the last
   * wrap column still leaves auto-owned rows that must shrink back — rows the
   * visible scan can never find, because they are already marked measured.
   * Getting only the first half right is a silent bug (rows stay tall forever),
   * so both halves belong behind one call.
   *
   * `visibleUnmeasuredCount` is reported separately because it, not the merged
   * set, decides whether another measurement pass has to be scheduled.
   */
  rowsToMeasureFor(
    scene: SceneSnapshot,
    options: {
      readonly hasColumnStructureChange: boolean;
      readonly rowHeightProvenance: RowHeightProvenanceView;
      readonly defaultRowHeight: number;
    }
  ): { readonly rows: ReadonlySet<number>; readonly visibleUnmeasuredCount: number } {
    const hasWrapColumns = this.hasWrapColumns();
    const visibleUnmeasuredRows = hasWrapColumns ? this.visibleUnmeasuredRows(scene) : [];
    const rows = new Set(visibleUnmeasuredRows);
    if (!hasWrapColumns && options.hasColumnStructureChange) {
      for (const row of this.rowsNeedingDefaultReset(
        scene,
        options.rowHeightProvenance,
        options.defaultRowHeight
      )) {
        rows.add(row);
      }
    }
    return { rows, visibleUnmeasuredCount: visibleUnmeasuredRows.length };
  }

  /**
   * Measures `rows` against `basis` and reports what would have to change.
   * Rows whose height is `manual` are never measured — a hand-resized row
   * keeps its height until the user changes it again.
   *
   * Applies nothing: the caller owns the command, the history policy and the
   * ordering. Rows that need no resize are marked measured here, because that
   * conclusion does not depend on a commit succeeding.
   */
  planUpdates(rows: ReadonlySet<number>, basis: WrapMeasurementBasis): WrapMeasurementPlan {
    if (
      basis.scene.modelVersion !== basis.modelVersion ||
      basis.scene.layoutVersion !== basis.layoutVersion
    ) {
      throw new Error(
        `Wrap measurement Scene basis ${basis.scene.modelVersion}/${basis.scene.layoutVersion} ` +
        `does not match model/layout ${basis.modelVersion}/${basis.layoutVersion}`
      );
    }

    const updates: RowHeightMeasurement[] = [];
    const measuredWithoutResize: number[] = [];
    for (const row of rows) {
      if (basis.rowHeightProvenance.provenanceOf(row) === "manual") {
        measuredWithoutResize.push(row);
        continue;
      }
      try {
        const required = this.computeRequiredRowHeight(row, basis.scene, basis.defaultRowHeight);
        if (required !== basis.scene.getRowHeight(row)) {
          updates.push(Object.freeze({ row, height: required }));
        } else {
          measuredWithoutResize.push(row);
        }
      } catch (error) {
        this.#onError(error, "finalize:wrap-measurement");
      }
    }
    this.markMeasured(measuredWithoutResize);
    return Object.freeze({
      updates: Object.freeze(updates),
      measuredWithoutResize: Object.freeze(measuredWithoutResize)
    });
  }

  /**
   * Applies one commit's worth of semantic changes to the cache.
   *
   * The rules live here rather than at the call site because the cache is this
   * service's own state: a caller that has to remember "a range edit forgets a
   * span, a structural edit forgets everything" is a caller that can get it
   * subtly wrong, and CyExcelCore previously spread exactly those rules across
   * two separate passes over the change list.
   */
  invalidateFor(changes: readonly SemanticChange[]): void {
    for (const change of changes) {
      switch (change.type) {
        // `cell-style` 也在这里：改字号会改变文字占的高度，不失效的话 wrap 列的
        // 行高会停在旧字号上。这不是"每格换行策略"（那要把自动行高的输入从列配置
        // 改成每格样式，是另一个量级的改动），只是让已有的、按列 wrap 触发的测量重跑。
        case "cell-value":
        case "cell-style":
          this.forgetRow(change.row);
          break;
        case "range-value":
        case "range-style":
          this.forgetRowRange(change.range.startRow, change.range.endRow);
          break;
        case "row-structure":
        case "row-reorder":
        case "column-structure":
        case "column-layout":
        case "full-reset":
          // Any axis-level change makes every cached "row N was measured"
          // claim stale — either the row moved or its available width did.
          this.forgetAll();
          return;
      }
    }
  }

  markMeasured(rows: Iterable<number>): void {
    for (const row of rows) {
      this.#measuredRows.add(row);
    }
    this.#evictOldestOverBound();
  }

  /**
   * Drops first-inserted rows until the cache is back within
   * `MAX_MEASURED_ROWS`. Insertion order is the eviction order because a `Set`
   * preserves it and re-adding an existing row does not move it — which is the
   * behaviour wanted here: rows enter the cache when first painted, so the
   * oldest entries are the ones furthest from wherever the user is now.
   */
  #evictOldestOverBound(): void {
    let excess = this.#measuredRows.size - MAX_MEASURED_ROWS;
    if (excess <= 0) {
      return;
    }
    for (const row of this.#measuredRows) {
      this.#measuredRows.delete(row);
      excess -= 1;
      if (excess === 0) {
        return;
      }
    }
  }

  forget(rows: Iterable<number>): void {
    for (const row of rows) {
      this.#measuredRows.delete(row);
    }
  }

  forgetRow(row: number): void {
    this.#measuredRows.delete(row);
  }

  /** Invalidates by row span without materialising the span itself. */
  forgetRowRange(startRow: number, endRow: number): void {
    for (const row of this.#measuredRows) {
      if (row >= startRow && row <= endRow) {
        this.#measuredRows.delete(row);
      }
    }
  }

  /**
   * A row insert/delete/move shifts which row every index refers to, so every
   * cached entry becomes a stale claim about a row that may no longer be
   * there. Clearing wholesale is simpler and safer than remapping.
   */
  forgetAll(): void {
    this.#measuredRows.clear();
  }

  /** Test/diagnostic view of the cache; never handed to production callers. */
  get measuredRowCount(): number {
    return this.#measuredRows.size;
  }
}
