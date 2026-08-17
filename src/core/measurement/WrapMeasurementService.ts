import type { RuntimeColumn } from "../cells";
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
export interface WrapMeasurementBasis {
  readonly scene: SceneSnapshot;
  /** Live model version, asserted against `scene.modelVersion`. */
  readonly modelVersion: number;
  /** Live layout root version, asserted against `scene.layoutVersion`. */
  readonly layoutVersion: number;
  readonly defaultRowHeight: number;
  readonly rowHeightProvenance: ReadonlyMap<number, Exclude<RowHeightProvenance, "default">>;
}

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
    rowHeightProvenance: ReadonlyMap<number, Exclude<RowHeightProvenance, "default">>,
    defaultRowHeight: number
  ): readonly number[] {
    const rows: number[] = [];
    for (const [row, provenance] of rowHeightProvenance) {
      if (provenance === "auto" && scene.getRowHeight(row) !== defaultRowHeight) {
        rows.push(row);
      }
    }
    return rows;
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
      if ((basis.rowHeightProvenance.get(row) ?? "default") === "manual") {
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

  markMeasured(rows: Iterable<number>): void {
    for (const row of rows) {
      this.#measuredRows.add(row);
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
