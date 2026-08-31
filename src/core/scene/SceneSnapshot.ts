import type { CellRange } from "../types";
import type { CellRect, IndexRange, SkeletonSnapshot, ViewportSnapshot } from "../layout";
import { splitByFrozenBoundary } from "../layout";

export interface SceneSurfaceMetrics {
  /** Total fixed height of every column-header row. */
  readonly columnHeaderHeight: number;
  /** Ordered header-band geometry. Omit for one leaf band filling the complete surface. */
  readonly columnHeaderBands?: readonly SceneColumnHeaderBandSpec[];
  readonly rowHeaderWidth: number;
}

export type SceneColumnHeaderBandSpec = Readonly<{
  readonly role: "leaf";
  readonly y: number;
  readonly height: number;
}>;

export interface SceneSnapshotOptions {
  readonly version: number;
  readonly skeleton: SkeletonSnapshot;
  readonly viewport: ViewportSnapshot;
  /** Header bands are measured once and published with every surface in this Scene generation. */
  readonly surfaceMetrics?: SceneSurfaceMetrics;
}

export type SceneSurfaceId = "body" | "column-header" | "row-header" | "corner" | "menu-overlay";
export type SceneLayerId =
  | "grid"
  | "virtual-cell"
  | "selection"
  | "drawing-object"
  | "column-header"
  | "row-header"
  | "dom-cell"
  | "editor"
  | "corner"
  | "menu";
export type SceneBodyPaneId = "frozen-corner" | "frozen-rows" | "frozen-columns" | "body";
export type ScenePaneId = SceneBodyPaneId | "frozen" | "scroll" | "surface";
export type SceneAxisBehavior = "fixed" | "scroll";

export interface SceneTranslation {
  readonly x: number;
  readonly y: number;
}

export interface SceneGenericNodeDescriptor {
  readonly kind: "node";
  readonly id: string;
}

interface SceneColumnHeaderBandDescriptorBase {
  readonly kind: "column-header-band";
  readonly surface: "column-header";
  /** Full-width bounds in column-header surface coordinates. */
  readonly rect: CellRect;
}

export type SceneColumnHeaderLeafBandDescriptor = SceneColumnHeaderBandDescriptorBase & Readonly<{
  readonly id: "leaf";
  readonly role: "leaf";
  readonly level: null;
}>;

export type SceneColumnHeaderBandDescriptor = SceneColumnHeaderLeafBandDescriptor;

export type SceneNodeDescriptor = SceneGenericNodeDescriptor | SceneColumnHeaderBandDescriptor;

export interface SceneLayerDescriptor {
  readonly kind: "layer";
  readonly id: SceneLayerId;
  readonly surface: SceneSurfaceId;
  /** Bounded structural/drawing nodes may persist; grid and header cells stay virtual. */
  readonly nodes: readonly SceneNodeDescriptor[];
}

export interface ScenePaneDescriptor {
  readonly kind: "pane";
  readonly id: ScenePaneId;
  readonly surface: SceneSurfaceId;
  /** Surface-local clipping bounds. Right and bottom edges are open. */
  readonly clipRect: CellRect;
  readonly rows: IndexRange;
  readonly columns: IndexRange;
  readonly frozen: { readonly row: boolean; readonly column: boolean };
  readonly scroll: { readonly x: SceneAxisBehavior; readonly y: SceneAxisBehavior };
  /** Adds to sheet coordinates to obtain coordinates local to this pane's surface. */
  readonly translation: SceneTranslation;
}

export interface SceneSurfaceDescriptor {
  readonly kind: "surface";
  readonly id: SceneSurfaceId;
  /** Surface bounds in Scene-root coordinates; pane clips remain surface-local. */
  readonly rootRect: CellRect;
  readonly layers: readonly SceneLayerDescriptor[];
  /** Pane geometry is canonical at surface scope and shared by every layer. */
  readonly panes: readonly ScenePaneDescriptor[];
}

export interface SceneRootDescriptor {
  readonly kind: "root";
  readonly id: "scene-root";
  readonly surfaces: readonly SceneSurfaceDescriptor[];
}

export interface SceneCellFragment {
  readonly pane: ScenePaneDescriptor;
  readonly surface: "body";
  readonly rows: IndexRange;
  readonly columns: IndexRange;
  readonly frozenRow: boolean;
  readonly frozenColumn: boolean;
  readonly clipRect: CellRect;
  readonly rect: CellRect;
}

export interface SceneColumnHeaderFragment {
  readonly band: SceneColumnHeaderBandDescriptor;
  readonly pane: ScenePaneDescriptor;
  readonly surface: "column-header";
  readonly columns: IndexRange;
  readonly clipRect: CellRect;
  readonly rect: CellRect;
}

export interface SceneRowHeaderFragment {
  readonly pane: ScenePaneDescriptor;
  readonly surface: "row-header";
  readonly rows: IndexRange;
  readonly clipRect: CellRect;
  readonly rect: CellRect;
}

/** Pane-aware geometry for the active selection's visual fill handle. */
export interface SceneSelectionHandleFragment {
  readonly pane: ScenePaneDescriptor;
  readonly surface: "body";
  readonly clipRect: CellRect;
  /** Full handle rect before the canonical pane clip is applied. */
  readonly rect: CellRect;
}

export interface SceneCellAddress {
  readonly row: number;
  readonly column: number;
}

/**
 * A logical scroll position, not a DOM one. Named `logicalTop`/`logicalLeft`
 * rather than `scrollTop`/`scrollLeft` on purpose: these are offsets into the
 * Scene's own coordinate space, which on a segmented axis is not the same
 * number the scroller carries, and the DOM-shaped names invited exactly the
 * confusion the §12 scroll gate exists to prevent.
 */
export interface SceneScrollTarget {
  readonly logicalTop: number;
  readonly logicalLeft: number;
}

export interface SceneLocalPoint {
  readonly x: number;
  readonly y: number;
}

export interface SceneCellHit {
  readonly surface: "body";
  readonly cell: SceneCellAddress;
  readonly pane: ScenePaneDescriptor;
  readonly rect: CellRect;
  readonly clipRect: CellRect;
  readonly localPoint: SceneLocalPoint;
}

export interface SceneColumnHeaderHit {
  readonly surface: "column-header";
  readonly band: SceneColumnHeaderBandDescriptor;
  readonly column: number;
  readonly pane: ScenePaneDescriptor;
  readonly rect: CellRect;
  readonly clipRect: CellRect;
  readonly localPoint: SceneLocalPoint;
}

export interface SceneRowHeaderHit {
  readonly surface: "row-header";
  readonly row: number;
  readonly pane: ScenePaneDescriptor;
  readonly rect: CellRect;
  readonly clipRect: CellRect;
  readonly localPoint: SceneLocalPoint;
}

function assertNonNegativeInteger(name: string, value: number): void {
  if (!Number.isSafeInteger(value) || value < 0) {
    throw new RangeError(`${name} must be a non-negative safe integer: ${value}`);
  }
}

function assertNonNegativeFinite(name: string, value: number): void {
  if (!Number.isFinite(value) || value < 0) {
    throw new RangeError(`${name} must be a non-negative finite number: ${value}`);
  }
}

function freezeRect(rect: CellRect): CellRect {
  return Object.freeze({ ...rect });
}

function freezeRange(range: IndexRange): IndexRange {
  return Object.freeze({ ...range });
}

function rectsIntersect(a: CellRect, b: CellRect): boolean {
  return (
    a.width > 0 &&
    a.height > 0 &&
    b.width > 0 &&
    b.height > 0 &&
    a.x < b.x + b.width &&
    a.y < b.y + b.height &&
    a.x + a.width > b.x &&
    a.y + a.height > b.y
  );
}

function intersectIndexRange(a: IndexRange, b: IndexRange): IndexRange {
  const start = Math.max(a.start, b.start);
  const end = Math.min(a.start + a.count, b.start + b.count);
  return { start, count: Math.max(0, end - start) };
}

function nonEmptyFragments(fragments: readonly SceneCellFragment[]): SceneCellFragment[] {
  return fragments.filter((fragment) => fragment.rows.count > 0 && fragment.columns.count > 0);
}

/**
 * Consumer-facing geometry for one frame. Renderers and hit-test code ask
 * Scene for surface-space fragments instead of reaching into Skeleton and
 * Viewport separately.
 */
export class SceneSnapshot {
  readonly version: number;
  readonly modelVersion: number;
  readonly layoutVersion: number;
  readonly viewportVersion: number;
  readonly width: number;
  readonly height: number;
  readonly overscanRows: number;
  readonly overscanColumns: number;
  readonly surfaceMetrics: SceneSurfaceMetrics;
  readonly root: SceneRootDescriptor;

  readonly #skeleton: SkeletonSnapshot;
  readonly #viewport: ViewportSnapshot;
  readonly #surfaces: readonly SceneSurfaceDescriptor[];
  readonly #bodyPanes: readonly ScenePaneDescriptor[];
  readonly #columnHeaderBands: readonly SceneColumnHeaderBandDescriptor[];

  constructor(options: SceneSnapshotOptions) {
    assertNonNegativeInteger("version", options.version);
    this.version = options.version;
    this.#skeleton = options.skeleton;
    this.#viewport = options.viewport;
    this.modelVersion = options.skeleton.modelVersion;
    this.layoutVersion = options.skeleton.layoutVersion;
    this.viewportVersion = options.viewport.version;
    this.width = options.viewport.width;
    this.height = options.viewport.height;
    this.overscanRows = options.viewport.overscanRows;
    this.overscanColumns = options.viewport.overscanColumns;
    const columnHeaderHeight = options.surfaceMetrics?.columnHeaderHeight ?? 0;
    const columnHeaderBands = options.surfaceMetrics?.columnHeaderBands ?? [
      { role: "leaf", y: 0, height: columnHeaderHeight }
    ];
    const rowHeaderWidth = options.surfaceMetrics?.rowHeaderWidth ?? 0;
    assertNonNegativeFinite("surfaceMetrics.columnHeaderHeight", columnHeaderHeight);
    assertNonNegativeFinite("surfaceMetrics.rowHeaderWidth", rowHeaderWidth);
    const normalizedBands = this.#normalizeColumnHeaderBandSpecs(columnHeaderBands, columnHeaderHeight);
    this.surfaceMetrics = Object.freeze({
      columnHeaderHeight,
      columnHeaderBands: normalizedBands,
      rowHeaderWidth
    });
    this.#columnHeaderBands = this.#createColumnHeaderBands(normalizedBands);
    this.#bodyPanes = this.#createBodyPanes();
    this.#surfaces = this.#createSurfaces(this.#bodyPanes);
    this.root = Object.freeze({
      kind: "root",
      id: "scene-root",
      surfaces: this.#surfaces
    });
    Object.freeze(this);
  }

  get rowCount(): number {
    return this.#skeleton.rowCount;
  }

  get columnCount(): number {
    return this.#skeleton.columnCount;
  }

  get frozenRowCount(): number {
    return this.#skeleton.frozenRowCount;
  }

  get frozenColumnCount(): number {
    return this.#skeleton.frozenColumnCount;
  }

  getTotalWidth(): number {
    return this.#skeleton.getTotalWidth();
  }

  getTotalHeight(): number {
    return this.#skeleton.getTotalHeight();
  }

  getCellRect(row: number, column: number): CellRect {
    return this.#skeleton.getCellRect(row, column);
  }

  getRowHeight(row: number): number {
    return this.#skeleton.getRowHeight(row);
  }

  /** Column-axis geometry without requiring a synthetic body-cell address. */
  getColumnWidth(column: number): number {
    return this.#skeleton.getColumnWidth(column);
  }

  #getSurfaceCellRect(row: number, column: number, pane: ScenePaneDescriptor): CellRect {
    const sheetRect = this.#skeleton.getCellRect(row, column);
    return {
      x: sheetRect.x + pane.translation.x,
      y: sheetRect.y + pane.translation.y,
      width: sheetRect.width,
      height: sheetRect.height
    };
  }

  getVisibleRows(): IndexRange {
    return this.#skeleton.getVisibleRowRange(this.#viewport);
  }

  getVisibleColumns(): IndexRange {
    return this.#skeleton.getVisibleColumnRange(this.#viewport);
  }

  getFrozenRowBoundaryY(): number {
    return this.#skeleton.getRowOffset(this.#skeleton.frozenRowCount);
  }

  getFrozenColumnBoundaryX(): number {
    return this.#skeleton.getColumnLeft(this.#skeleton.frozenColumnCount);
  }

  getPanes(surface: SceneSurfaceId = "body"): readonly ScenePaneDescriptor[] {
    return this.#surface(surface).panes;
  }

  getSurface(surface: SceneSurfaceId): SceneSurfaceDescriptor {
    return this.#surface(surface);
  }

  getCellRectInPane(cell: SceneCellAddress, pane: ScenePaneDescriptor | SceneBodyPaneId): CellRect {
    const descriptor = typeof pane === "string" ? this.#bodyPane(pane) : pane;
    return this.#getSurfaceCellRect(cell.row, cell.column, descriptor);
  }

  getColumnHeaderBands(): readonly SceneColumnHeaderBandDescriptor[] {
    return this.#columnHeaderBands;
  }

  getColumnHeaderLeafBand(): SceneColumnHeaderLeafBandDescriptor {
    return this.#columnHeaderBands.find((band) => band.role === "leaf") as SceneColumnHeaderLeafBandDescriptor;
  }

  /**
   * Returns the traversal ranges for one band, bounded to the columns that
   * can reach the current pane clips. No per-column Scene nodes are retained.
   */
  getVisibleColumnHeaderFragments(
    band: SceneColumnHeaderBandDescriptor
  ): readonly SceneColumnHeaderFragment[] {
    this.#assertOwnedColumnHeaderBand(band);
    const fragments: SceneColumnHeaderFragment[] = [];
    for (const pane of this.#surface("column-header").panes) {
      const columns = this.#visibleColumnsInPane(pane);
      const fragment = this.#createColumnHeaderFragment(band, columns, pane);
      if (fragment) {
        fragments.push(fragment);
      }
    }
    return Object.freeze(fragments);
  }

  /** Creates at most two transient fragments for an inclusive logical span. */
  getColumnHeaderBandFragments(
    band: SceneColumnHeaderBandDescriptor,
    startColumn: number,
    endColumn: number
  ): readonly SceneColumnHeaderFragment[] {
    this.#assertOwnedColumnHeaderBand(band);
    if (
      !Number.isSafeInteger(startColumn) ||
      !Number.isSafeInteger(endColumn) ||
      startColumn < 0 ||
      endColumn < startColumn ||
      startColumn >= this.columnCount
    ) {
      return [];
    }
    const requested = { start: startColumn, count: Math.min(endColumn, this.columnCount - 1) - startColumn + 1 };
    const fragments: SceneColumnHeaderFragment[] = [];
    for (const pane of this.#surface("column-header").panes) {
      const columns = freezeRange(intersectIndexRange(requested, pane.columns));
      const fragment = this.#createColumnHeaderFragment(band, columns, pane);
      if (fragment) {
        fragments.push(fragment);
      }
    }
    return Object.freeze(fragments);
  }

  /** Row-axis counterpart to header fragments; creates at most two transient fragments. */
  getRowHeaderFragments(startRow: number, endRow: number): readonly SceneRowHeaderFragment[] {
    if (
      !Number.isSafeInteger(startRow) ||
      !Number.isSafeInteger(endRow) ||
      startRow < 0 ||
      endRow < startRow ||
      startRow >= this.rowCount
    ) {
      return [];
    }
    const surface = this.#surface("row-header");
    const requested = { start: startRow, count: Math.min(endRow, this.rowCount - 1) - startRow + 1 };
    const fragments: SceneRowHeaderFragment[] = [];
    for (const pane of surface.panes) {
      const rows = freezeRange(intersectIndexRange(requested, pane.rows));
      if (rows.count <= 0) {
        continue;
      }
      const start = this.#skeleton.getRowOffset(rows.start) + pane.translation.y;
      const end = this.#skeleton.getRowOffset(rows.start + rows.count) + pane.translation.y;
      const rect = freezeRect({ x: 0, y: start, width: surface.rootRect.width, height: end - start });
      if (!rectsIntersect(rect, pane.clipRect)) {
        continue;
      }
      fragments.push(Object.freeze({
        pane,
        surface: "row-header",
        rows,
        clipRect: pane.clipRect,
        rect
      }));
    }
    return Object.freeze(fragments);
  }

  getPaneForCell(cell: SceneCellAddress): ScenePaneDescriptor | null {
    if (
      cell.row < 0 ||
      cell.row >= this.rowCount ||
      cell.column < 0 ||
      cell.column >= this.columnCount
    ) {
      return null;
    }
    if (cell.row < this.frozenRowCount && cell.column < this.frozenColumnCount) {
      return this.#bodyPane("frozen-corner");
    }
    if (cell.row < this.frozenRowCount) {
      return this.#bodyPane("frozen-rows");
    }
    if (cell.column < this.frozenColumnCount) {
      return this.#bodyPane("frozen-columns");
    }
    return this.#bodyPane("body");
  }

  getVisibleFragments(): readonly SceneCellFragment[] {
    return nonEmptyFragments(this.#bodyPanes.map((pane) => this.#fragment(pane.rows, pane.columns, pane)));
  }

  getVisibleFragmentsForRange(range: CellRange): readonly SceneCellFragment[] {
    const rowIndexRange: IndexRange = { start: range.startRow, count: range.endRow - range.startRow + 1 };
    const columnIndexRange: IndexRange = {
      start: range.startColumn,
      count: range.endColumn - range.startColumn + 1
    };

    const rowSplit = splitByFrozenBoundary(rowIndexRange, this.#skeleton.frozenRowCount);
    const columnSplit = splitByFrozenBoundary(columnIndexRange, this.#skeleton.frozenColumnCount);

    const scrollableRows = intersectIndexRange(rowSplit.scrollable, this.getVisibleRows());
    const scrollableColumns = intersectIndexRange(columnSplit.scrollable, this.getVisibleColumns());

    return nonEmptyFragments([
      this.#fragment(rowSplit.frozen, columnSplit.frozen, this.#bodyPane("frozen-corner")),
      this.#fragment(rowSplit.frozen, scrollableColumns, this.#bodyPane("frozen-rows")),
      this.#fragment(scrollableRows, columnSplit.frozen, this.#bodyPane("frozen-columns")),
      this.#fragment(scrollableRows, scrollableColumns, this.#bodyPane("body"))
    ]);
  }

  getFragmentsForRange(range: CellRange): readonly SceneCellFragment[] {
    const rowIndexRange: IndexRange = { start: range.startRow, count: range.endRow - range.startRow + 1 };
    const columnIndexRange: IndexRange = {
      start: range.startColumn,
      count: range.endColumn - range.startColumn + 1
    };

    const rowSplit = splitByFrozenBoundary(rowIndexRange, this.#skeleton.frozenRowCount);
    const columnSplit = splitByFrozenBoundary(columnIndexRange, this.#skeleton.frozenColumnCount);

    return nonEmptyFragments([
      this.#fragment(rowSplit.frozen, columnSplit.frozen, this.#bodyPane("frozen-corner")),
      this.#fragment(rowSplit.frozen, columnSplit.scrollable, this.#bodyPane("frozen-rows")),
      this.#fragment(rowSplit.scrollable, columnSplit.frozen, this.#bodyPane("frozen-columns")),
      this.#fragment(rowSplit.scrollable, columnSplit.scrollable, this.#bodyPane("body"))
    ]);
  }

  getCellFragments(cell: SceneCellAddress, surface: "body" = "body"): readonly SceneCellFragment[] {
    if (
      surface !== "body" ||
      cell.row < 0 ||
      cell.row >= this.rowCount ||
      cell.column < 0 ||
      cell.column >= this.columnCount
    ) {
      return [];
    }
    const pane = this.getPaneForCell(cell);
    if (!pane) {
      return [];
    }
    const fragments = [this.#fragment(
      { start: cell.row, count: 1 },
      { start: cell.column, count: 1 },
      pane
    )];
    return fragments.filter(({ rect, clipRect }) => (
      rect.width > 0 &&
      rect.height > 0 &&
      rect.x < clipRect.x + clipRect.width &&
      rect.y < clipRect.y + clipRect.height &&
      rect.x + rect.width > clipRect.x &&
      rect.y + rect.height > clipRect.y
    ));
  }

  getRangeFragments(range: CellRange, surface: "body" = "body"): readonly SceneCellFragment[] {
    return surface === "body" ? this.getFragmentsForRange(range) : [];
  }

  /**
   * Returns the nearest logical scroll offset that fully reveals a cell.
   * Frozen axes remain unchanged because their pane is already fixed.
   */
  getScrollTargetForCell(cell: SceneCellAddress): SceneScrollTarget | null {
    const pane = this.getPaneForCell(cell);
    if (!pane) {
      return null;
    }
    const rect = this.getCellRectInPane(cell, pane);
    let logicalTop = this.#viewport.scrollTop;
    let logicalLeft = this.#viewport.scrollLeft;

    if (pane.scroll.y === "scroll") {
      const clipBottom = pane.clipRect.y + pane.clipRect.height;
      if (rect.y < pane.clipRect.y) {
        logicalTop = Math.max(0, logicalTop + rect.y - pane.clipRect.y);
      } else if (rect.y + rect.height > clipBottom) {
        logicalTop = Math.max(0, logicalTop + rect.y + rect.height - clipBottom);
      }
    }

    if (pane.scroll.x === "scroll") {
      const clipRight = pane.clipRect.x + pane.clipRect.width;
      if (rect.x < pane.clipRect.x) {
        logicalLeft = Math.max(0, logicalLeft + rect.x - pane.clipRect.x);
      } else if (rect.x + rect.width > clipRight) {
        logicalLeft = Math.max(0, logicalLeft + rect.x + rect.width - clipRight);
      }
    }

    return Object.freeze({ logicalTop, logicalLeft });
  }

  getSelectionHandleFragment(range: CellRange, handleSize: number): SceneSelectionHandleFragment | null {
    if (!Number.isFinite(handleSize) || handleSize <= 0) {
      return null;
    }
    const cell = { row: range.endRow, column: range.endColumn };
    const pane = this.getPaneForCell(cell);
    if (!pane) {
      return null;
    }
    const cellRect = this.getCellRectInPane(cell, pane);
    if (cellRect.width <= 0 || cellRect.height <= 0) {
      return null;
    }
    const half = handleSize / 2;
    const rect = Object.freeze({
      x: cellRect.x + cellRect.width - half,
      y: cellRect.y + cellRect.height - half,
      width: handleSize,
      height: handleSize
    });
    const clipRect = pane.clipRect;
    if (
      rect.x >= clipRect.x + clipRect.width ||
      rect.y >= clipRect.y + clipRect.height ||
      rect.x + rect.width <= clipRect.x ||
      rect.y + rect.height <= clipRect.y
    ) {
      return null;
    }
    return Object.freeze({ pane, surface: "body", clipRect, rect });
  }

  /** Returns all body geometry needed by a custom renderer from this Scene generation. */
  hitTestCellGeometry(point: SceneLocalPoint): SceneCellHit | null {
    if (
      !Number.isFinite(point.x) ||
      !Number.isFinite(point.y) ||
      point.x < 0 ||
      point.y < 0 ||
      point.x >= this.width ||
      point.y >= this.height ||
      this.rowCount === 0 ||
      this.columnCount === 0
    ) {
      return null;
    }

    const pane = this.#paneAtSurfacePoint("body", point);
    if (!pane) {
      return null;
    }
    const sheetY = point.y - pane.translation.y;
    const sheetX = point.x - pane.translation.x;

    const row = this.#skeleton.getRowAtY(sheetY);
    if (
      !this.#rangeContains(pane.rows, row) ||
      this.#skeleton.getRowOffset(row) > sheetY
    ) {
      return null;
    }

    const column = this.#skeleton.getColumnAtX(sheetX);
    if (
      !this.#rangeContains(pane.columns, column) ||
      this.#skeleton.getColumnLeft(column) > sheetX
    ) {
      return null;
    }

    const rect = freezeRect(this.#getSurfaceCellRect(row, column, pane));
    if (!this.#pointInClippedRect(point, rect, pane.clipRect)) {
      return null;
    }

    return Object.freeze({
      surface: "body",
      cell: Object.freeze({ row, column }),
      pane,
      rect,
      clipRect: pane.clipRect,
      localPoint: Object.freeze({ x: point.x - rect.x, y: point.y - rect.y })
    });
  }

  /** Resolves an x coordinate local to the column-header surface. */
  resolveColumnHeader(x: number): number | null {
    if (!Number.isFinite(x) || x < 0 || x >= this.width || this.columnCount === 0) {
      return null;
    }
    const pane = this.#paneAtAxisPoint("column-header", "x", x);
    return pane ? this.#resolveColumnInPane(x, pane) : null;
  }

  /** Returns the canonical band and column geometry at a header-surface point. */
  hitTestColumnHeader(point: SceneLocalPoint): SceneColumnHeaderHit | null {
    const pane = this.#paneAtSurfacePoint("column-header", point);
    if (!pane) {
      return null;
    }
    const band = this.#columnHeaderBands.find((candidate) => this.#pointInRect(point, candidate.rect));
    if (!band) {
      return null;
    }
    const column = this.#resolveColumnInPane(point.x, pane);
    if (column === null) {
      return null;
    }
    const fragment = this.getColumnHeaderBandFragments(band, column, column)
      .find((candidate) => candidate.pane === pane);
    if (!fragment || !this.#pointInClippedRect(point, fragment.rect, fragment.clipRect)) {
      return null;
    }
    return Object.freeze({
      surface: "column-header",
      band,
      column,
      pane,
      rect: fragment.rect,
      clipRect: fragment.clipRect,
      localPoint: Object.freeze({
        x: point.x - fragment.rect.x,
        y: point.y - fragment.rect.y
      })
    });
  }

  /** Resolves a y coordinate local to the row-header surface. */
  resolveRowHeader(y: number): number | null {
    if (!Number.isFinite(y) || y < 0 || y >= this.height || this.rowCount === 0) {
      return null;
    }
    const pane = this.#paneAtAxisPoint("row-header", "y", y);
    return pane ? this.#resolveRowInPane(y, pane) : null;
  }

  /** Returns canonical row-header geometry without projecting a body pane. */
  hitTestRowHeader(point: SceneLocalPoint): SceneRowHeaderHit | null {
    const pane = this.#paneAtSurfacePoint("row-header", point);
    if (!pane) {
      return null;
    }
    const row = this.#resolveRowInPane(point.y, pane);
    if (row === null) {
      return null;
    }
    const fragment = this.getRowHeaderFragments(row, row)
      .find((candidate) => candidate.pane === pane);
    if (!fragment || !this.#pointInClippedRect(point, fragment.rect, fragment.clipRect)) {
      return null;
    }
    return Object.freeze({
      surface: "row-header",
      row,
      pane,
      rect: fragment.rect,
      clipRect: fragment.clipRect,
      localPoint: Object.freeze({
        x: point.x - fragment.rect.x,
        y: point.y - fragment.rect.y
      })
    });
  }

  hitTestColumnBoundary(x: number, tolerance: number): number | null {
    if (!Number.isFinite(x) || x < 0 || x >= this.width || !Number.isFinite(tolerance) || tolerance < 0) {
      return null;
    }
    const pane = this.#surface("column-header").panes.find((candidate) =>
      x >= candidate.clipRect.x && x < candidate.clipRect.x + candidate.clipRect.width
    );
    if (!pane) {
      return null;
    }
    return this.#hitTestBoundaryInPane("column", x, tolerance, pane);
  }

  hitTestRowBoundary(y: number, tolerance: number): number | null {
    if (!Number.isFinite(y) || y < 0 || y >= this.height || !Number.isFinite(tolerance) || tolerance < 0) {
      return null;
    }
    const pane = this.#surface("row-header").panes.find((candidate) =>
      y >= candidate.clipRect.y && y < candidate.clipRect.y + candidate.clipRect.height
    );
    if (!pane) {
      return null;
    }
    return this.#hitTestBoundaryInPane("row", y, tolerance, pane);
  }

  resolveRow(y: number): number {
    const frozenHeight = this.getFrozenRowBoundaryY();
    const contentY = y < frozenHeight ? y : this.#viewport.scrollTop + y;
    return this.#skeleton.getRowAtY(Math.max(0, contentY));
  }

  resolveColumn(x: number): number {
    const frozenWidth = this.getFrozenColumnBoundaryX();
    const contentX = x < frozenWidth ? x : this.#viewport.scrollLeft + x;
    return this.#skeleton.getColumnAtX(Math.max(0, contentX));
  }

  getRowBoundaryY(boundaryRow: number): number {
    const absoluteY = this.#skeleton.getRowOffset(boundaryRow);
    return boundaryRow <= this.#skeleton.frozenRowCount ? absoluteY : absoluteY - this.#viewport.scrollTop;
  }

  #hitTestBoundaryInPane(
    axis: "row" | "column",
    point: number,
    tolerance: number,
    pane: ScenePaneDescriptor
  ): number | null {
    const range = axis === "row" ? pane.rows : pane.columns;
    const clipStart = axis === "row" ? pane.clipRect.y : pane.clipRect.x;
    const clipEnd = clipStart + (axis === "row" ? pane.clipRect.height : pane.clipRect.width);
    let closest: { readonly index: number; readonly distance: number } | null = null;

    // At an unscrolled pane split, the leading edge is also the trailing
    // boundary of the preceding (frozen) axis item. Once partial scroll
    // moves the first scrollable item underneath the frozen clip, ownership
    // belongs to that item's trailing boundary instead (covered by the loop).
    const firstIndex = range.start;
    if (firstIndex > 0 && range.count > 0) {
      const precedingSize = axis === "row"
        ? this.#skeleton.getRowHeight(firstIndex - 1)
        : this.#skeleton.getColumnWidth(firstIndex - 1);
      const firstStart = (
        axis === "row"
          ? this.#skeleton.getRowOffset(firstIndex) + pane.translation.y
          : this.#skeleton.getColumnLeft(firstIndex) + pane.translation.x
      );
      if (precedingSize > 0 && firstStart >= clipStart) {
        const distance = Math.abs(point - firstStart);
        if (distance <= tolerance) {
          closest = { index: firstIndex - 1, distance };
        }
      }
    }

    for (let offset = 0; offset < range.count; offset += 1) {
      const index = range.start + offset;
      const start = axis === "row"
        ? this.#skeleton.getRowOffset(index) + pane.translation.y
        : this.#skeleton.getColumnLeft(index) + pane.translation.x;
      const size = axis === "row"
        ? this.#skeleton.getRowHeight(index)
        : this.#skeleton.getColumnWidth(index);
      if (size <= 0) {
        continue;
      }
      const boundary = start + size;
      if (boundary < clipStart || boundary > clipEnd) {
        continue;
      }
      const distance = Math.abs(point - boundary);
      if (distance <= tolerance && (!closest || distance < closest.distance)) {
        closest = { index, distance };
      }
    }
    return closest?.index ?? null;
  }

  #fragment(rows: IndexRange, columns: IndexRange, pane: ScenePaneDescriptor): SceneCellFragment {
    if (rows.count <= 0 || columns.count <= 0) {
      const empty = { x: 0, y: 0, width: 0, height: 0 };
      return Object.freeze({
        pane,
        surface: "body",
        rows,
        columns,
        frozenRow: pane.frozen.row,
        frozenColumn: pane.frozen.column,
        clipRect: pane.clipRect,
        rect: empty
      });
    }

    const range = {
      startRow: rows.start,
      startColumn: columns.start,
      endRow: rows.start + rows.count - 1,
      endColumn: columns.start + columns.count - 1
    };
    const topLeft = this.getCellRectInPane({ row: range.startRow, column: range.startColumn }, pane);
    const bottomRight = this.getCellRectInPane({ row: range.endRow, column: range.endColumn }, pane);
    const rect = {
      x: topLeft.x,
      y: topLeft.y,
      width: bottomRight.x + bottomRight.width - topLeft.x,
      height: bottomRight.y + bottomRight.height - topLeft.y
    };
    return Object.freeze({
      pane,
      surface: "body",
      rows,
      columns,
      frozenRow: pane.frozen.row,
      frozenColumn: pane.frozen.column,
      clipRect: pane.clipRect,
      rect
    });
  }

  #surface(id: SceneSurfaceId): SceneSurfaceDescriptor {
    const surface = this.#surfaces.find((candidate) => candidate.id === id);
    if (!surface) {
      throw new Error(`SceneSnapshot: missing surface ${id}`);
    }
    return surface;
  }

  #bodyPane(id: SceneBodyPaneId): ScenePaneDescriptor {
    const pane = this.#bodyPanes.find((candidate) => candidate.id === id);
    if (!pane) {
      throw new Error(`SceneSnapshot: missing body pane ${id}`);
    }
    return pane;
  }

  #paneAtSurfacePoint(
    surface: "body" | "column-header" | "row-header",
    point: SceneLocalPoint
  ): ScenePaneDescriptor | null {
    if (!Number.isFinite(point.x) || !Number.isFinite(point.y)) {
      return null;
    }
    return this.#surface(surface).panes.find((pane) => (
      point.x >= pane.clipRect.x &&
      point.y >= pane.clipRect.y &&
      point.x < pane.clipRect.x + pane.clipRect.width &&
      point.y < pane.clipRect.y + pane.clipRect.height
    )) ?? null;
  }

  #paneAtAxisPoint(
    surface: "column-header" | "row-header",
    axis: "x" | "y",
    point: number
  ): ScenePaneDescriptor | null {
    return this.#surface(surface).panes.find((pane) => {
      const start = axis === "x" ? pane.clipRect.x : pane.clipRect.y;
      const size = axis === "x" ? pane.clipRect.width : pane.clipRect.height;
      return point >= start && point < start + size;
    }) ?? null;
  }

  #resolveColumnInPane(x: number, pane: ScenePaneDescriptor): number | null {
    const sheetX = x - pane.translation.x;
    const column = this.#skeleton.getColumnAtX(sheetX);
    if (
      !this.#rangeContains(pane.columns, column) ||
      this.#skeleton.getColumnLeft(column) > sheetX
    ) {
      return null;
    }
    const left = this.#skeleton.getColumnLeft(column) + pane.translation.x;
    const width = this.#skeleton.getColumnWidth(column);
    return width > 0 && x >= left && x < left + width ? column : null;
  }

  #resolveRowInPane(y: number, pane: ScenePaneDescriptor): number | null {
    const sheetY = y - pane.translation.y;
    const row = this.#skeleton.getRowAtY(sheetY);
    if (
      !this.#rangeContains(pane.rows, row) ||
      this.#skeleton.getRowOffset(row) > sheetY
    ) {
      return null;
    }
    const top = this.#skeleton.getRowOffset(row) + pane.translation.y;
    const height = this.#skeleton.getRowHeight(row);
    return height > 0 && y >= top && y < top + height ? row : null;
  }

  #rangeContains(range: IndexRange, index: number): boolean {
    return index >= range.start && index < range.start + range.count;
  }

  #pointInClippedRect(point: SceneLocalPoint, rect: CellRect, clipRect: CellRect): boolean {
    const left = Math.max(rect.x, clipRect.x);
    const top = Math.max(rect.y, clipRect.y);
    const right = Math.min(rect.x + rect.width, clipRect.x + clipRect.width);
    const bottom = Math.min(rect.y + rect.height, clipRect.y + clipRect.height);
    return point.x >= left && point.y >= top && point.x < right && point.y < bottom;
  }

  #pointInRect(point: SceneLocalPoint, rect: CellRect): boolean {
    return (
      point.x >= rect.x &&
      point.y >= rect.y &&
      point.x < rect.x + rect.width &&
      point.y < rect.y + rect.height
    );
  }

  #normalizeColumnHeaderBandSpecs(
    specs: readonly SceneColumnHeaderBandSpec[],
    columnHeaderHeight: number
  ): readonly SceneColumnHeaderBandSpec[] {
    if (specs.length !== 1) {
      throw new RangeError("surfaceMetrics.columnHeaderBands must contain exactly one leaf band");
    }
    const spec = specs[0] as SceneColumnHeaderBandSpec;
    assertNonNegativeFinite("surfaceMetrics.columnHeaderBands[0].y", spec.y);
    assertNonNegativeFinite("surfaceMetrics.columnHeaderBands[0].height", spec.height);
    if (spec.y + spec.height > columnHeaderHeight) {
      throw new RangeError("surfaceMetrics.columnHeaderBands[0] exceeds column-header height");
    }
    return Object.freeze([Object.freeze({ role: "leaf", y: spec.y, height: spec.height })]);
  }

  #createColumnHeaderBands(
    specs: readonly SceneColumnHeaderBandSpec[]
  ): readonly SceneColumnHeaderBandDescriptor[] {
    return Object.freeze(specs.map((spec): SceneColumnHeaderBandDescriptor => {
      const rect = freezeRect({ x: 0, y: spec.y, width: this.width, height: spec.height });
      return Object.freeze({
        kind: "column-header-band",
        id: "leaf",
        surface: "column-header",
        role: "leaf",
        level: null,
        rect
      });
    }));
  }

  #assertOwnedColumnHeaderBand(band: SceneColumnHeaderBandDescriptor): void {
    if (!this.#columnHeaderBands.includes(band)) {
      throw new RangeError("column-header band does not belong to this SceneSnapshot");
    }
  }

  #visibleColumnsInPane(pane: ScenePaneDescriptor): IndexRange {
    if (pane.columns.count <= 0 || pane.clipRect.width <= 0 || this.columnCount === 0) {
      return freezeRange({ start: pane.columns.start, count: 0 });
    }
    const sheetLeft = Math.max(0, pane.clipRect.x - pane.translation.x);
    const sheetRight = Math.max(sheetLeft, pane.clipRect.x + pane.clipRect.width - pane.translation.x);
    const first = this.#skeleton.getColumnAtX(sheetLeft);
    const last = this.#skeleton.getColumnAtX(sheetRight);
    const overscan = pane.scroll.x === "scroll" ? this.overscanColumns : 0;
    const clipped = {
      start: Math.max(0, first - overscan),
      count: Math.min(this.columnCount - 1, last + overscan) - Math.max(0, first - overscan) + 1
    };
    return freezeRange(intersectIndexRange(pane.columns, clipped));
  }

  #createColumnHeaderFragment(
    band: SceneColumnHeaderBandDescriptor,
    columns: IndexRange,
    pane: ScenePaneDescriptor
  ): SceneColumnHeaderFragment | null {
    if (columns.count <= 0 || band.rect.width <= 0 || band.rect.height <= 0) {
      return null;
    }
    const start = this.#skeleton.getColumnLeft(columns.start) + pane.translation.x;
    const end = this.#skeleton.getColumnLeft(columns.start + columns.count) + pane.translation.x;
    const rect = freezeRect({
      x: start,
      y: band.rect.y,
      width: end - start,
      height: band.rect.height
    });
    if (!rectsIntersect(rect, pane.clipRect)) {
      return null;
    }
    return Object.freeze({
      band,
      pane,
      surface: "column-header",
      columns,
      clipRect: pane.clipRect,
      rect
    });
  }

  #createBodyPanes(): readonly ScenePaneDescriptor[] {
    const frozenRows = Object.freeze({ start: 0, count: this.#skeleton.frozenRowCount });
    const scrollableRows = Object.freeze(this.getVisibleRows());
    const frozenColumns = Object.freeze({ start: 0, count: this.#skeleton.frozenColumnCount });
    const scrollableColumns = Object.freeze(this.getVisibleColumns());
    const frozenWidth = Math.min(this.width, this.getFrozenColumnBoundaryX());
    const frozenHeight = Math.min(this.height, this.getFrozenRowBoundaryY());

    const createPane = (
      id: SceneBodyPaneId,
      rows: IndexRange,
      columns: IndexRange,
      frozenRow: boolean,
      frozenColumn: boolean
    ): ScenePaneDescriptor => Object.freeze({
      kind: "pane",
      id,
      surface: "body",
      clipRect: Object.freeze({
        x: frozenColumn ? 0 : frozenWidth,
        y: frozenRow ? 0 : frozenHeight,
        width: frozenColumn ? frozenWidth : Math.max(0, this.width - frozenWidth),
        height: frozenRow ? frozenHeight : Math.max(0, this.height - frozenHeight)
      }),
      rows,
      columns,
      frozen: Object.freeze({ row: frozenRow, column: frozenColumn }),
      scroll: Object.freeze({
        x: frozenColumn ? "fixed" : "scroll",
        y: frozenRow ? "fixed" : "scroll"
      }),
      translation: Object.freeze({
        x: frozenColumn ? 0 : -this.#viewport.scrollLeft,
        y: frozenRow ? 0 : -this.#viewport.scrollTop
      })
    });

    return Object.freeze([
      createPane("frozen-corner", frozenRows, frozenColumns, true, true),
      createPane("frozen-rows", frozenRows, scrollableColumns, true, false),
      createPane("frozen-columns", scrollableRows, frozenColumns, false, true),
      createPane("body", scrollableRows, scrollableColumns, false, false)
    ]);
  }

  #createSurfaces(bodyPanes: readonly ScenePaneDescriptor[]): readonly SceneSurfaceDescriptor[] {
    const layer = (
      id: SceneLayerId,
      surface: SceneSurfaceId,
      nodes: readonly SceneNodeDescriptor[] = Object.freeze([])
    ): SceneLayerDescriptor => Object.freeze({
      kind: "layer",
      id,
      surface,
      nodes
    });
    const surface = (
      id: SceneSurfaceId,
      rootRect: CellRect,
      layers: readonly SceneLayerDescriptor[],
      panes: readonly ScenePaneDescriptor[] = Object.freeze([])
    ): SceneSurfaceDescriptor => Object.freeze({
      kind: "surface",
      id,
      rootRect: freezeRect(rootRect),
      layers: Object.freeze(layers),
      panes
    });
    const { columnHeaderHeight, rowHeaderWidth } = this.surfaceMetrics;
    const rootWidth = rowHeaderWidth + this.width;
    const rootHeight = columnHeaderHeight + this.height;
    const emptyRows = freezeRange({ start: 0, count: 0 });
    const emptyColumns = freezeRange({ start: 0, count: 0 });
    const frozenRows = freezeRange({ start: 0, count: this.#skeleton.frozenRowCount });
    const scrollableRows = freezeRange(this.getVisibleRows());
    const frozenColumns = freezeRange({ start: 0, count: this.#skeleton.frozenColumnCount });
    const scrollableColumns = freezeRange(this.getVisibleColumns());
    const frozenWidth = Math.min(this.width, this.getFrozenColumnBoundaryX());
    const frozenHeight = Math.min(this.height, this.getFrozenRowBoundaryY());
    const pane = (
      id: ScenePaneId,
      surfaceId: SceneSurfaceId,
      clipRect: CellRect,
      rows: IndexRange,
      columns: IndexRange,
      frozen: { readonly row: boolean; readonly column: boolean },
      scroll: { readonly x: SceneAxisBehavior; readonly y: SceneAxisBehavior },
      translation: SceneTranslation
    ): ScenePaneDescriptor => Object.freeze({
      kind: "pane",
      id,
      surface: surfaceId,
      clipRect: freezeRect(clipRect),
      rows,
      columns,
      frozen: Object.freeze(frozen),
      scroll: Object.freeze(scroll),
      translation: Object.freeze(translation)
    });
    const columnHeaderPanes = Object.freeze([
      pane(
        "frozen",
        "column-header",
        { x: 0, y: 0, width: frozenWidth, height: columnHeaderHeight },
        emptyRows,
        frozenColumns,
        { row: true, column: true },
        { x: "fixed", y: "fixed" },
        { x: 0, y: 0 }
      ),
      pane(
        "scroll",
        "column-header",
        { x: frozenWidth, y: 0, width: Math.max(0, this.width - frozenWidth), height: columnHeaderHeight },
        emptyRows,
        scrollableColumns,
        { row: true, column: false },
        { x: "scroll", y: "fixed" },
        { x: -this.#viewport.scrollLeft, y: 0 }
      )
    ]);
    const rowHeaderPanes = Object.freeze([
      pane(
        "frozen",
        "row-header",
        { x: 0, y: 0, width: rowHeaderWidth, height: frozenHeight },
        frozenRows,
        emptyColumns,
        { row: true, column: true },
        { x: "fixed", y: "fixed" },
        { x: 0, y: 0 }
      ),
      pane(
        "scroll",
        "row-header",
        { x: 0, y: frozenHeight, width: rowHeaderWidth, height: Math.max(0, this.height - frozenHeight) },
        scrollableRows,
        emptyColumns,
        { row: false, column: true },
        { x: "fixed", y: "scroll" },
        { x: 0, y: -this.#viewport.scrollTop }
      )
    ]);
    const cornerPanes = Object.freeze([
      pane(
        "surface",
        "corner",
        { x: 0, y: 0, width: rowHeaderWidth, height: columnHeaderHeight },
        emptyRows,
        emptyColumns,
        { row: true, column: true },
        { x: "fixed", y: "fixed" },
        { x: 0, y: 0 }
      )
    ]);
    const menuPanes = Object.freeze([
      pane(
        "surface",
        "menu-overlay",
        { x: 0, y: 0, width: rootWidth, height: rootHeight },
        emptyRows,
        emptyColumns,
        { row: true, column: true },
        { x: "fixed", y: "fixed" },
        { x: 0, y: 0 }
      )
    ]);

    return Object.freeze([
      surface("body", { x: rowHeaderWidth, y: columnHeaderHeight, width: this.width, height: this.height }, [
        layer("grid", "body"),
        layer("virtual-cell", "body"),
        layer("selection", "body"),
        layer("drawing-object", "body"),
        layer("dom-cell", "body"),
        layer("editor", "body")
      ], bodyPanes),
      surface(
        "column-header",
        { x: rowHeaderWidth, y: 0, width: this.width, height: columnHeaderHeight },
        [layer("column-header", "column-header", this.#columnHeaderBands)],
        columnHeaderPanes
      ),
      surface(
        "row-header",
        { x: 0, y: columnHeaderHeight, width: rowHeaderWidth, height: this.height },
        [layer("row-header", "row-header")],
        rowHeaderPanes
      ),
      surface(
        "corner",
        { x: 0, y: 0, width: rowHeaderWidth, height: columnHeaderHeight },
        [layer("corner", "corner")],
        cornerPanes
      ),
      surface(
        "menu-overlay",
        { x: 0, y: 0, width: rootWidth, height: rootHeight },
        [layer("menu", "menu-overlay")],
        menuPanes
      )
    ]);
  }
}
