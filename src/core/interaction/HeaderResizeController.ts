import type { SceneSnapshot, SurfacePoint } from "../scene";

export const HEADER_RESIZE_HANDLE_TOLERANCE = 4;
export const ROW_RESIZE_GRIP_WIDTH = 12;
export const DEFAULT_MIN_COLUMN_RESIZE_SIZE = 20;
export const DEFAULT_MIN_ROW_RESIZE_SIZE = 16;

export type HeaderResizePoint =
  | SurfacePoint<"column-header">
  | SurfacePoint<"row-header">;

export type HeaderResizeHandle =
  | Readonly<{
      surface: "column-header";
      axis: "column";
      column: number;
    }>
  | Readonly<{
      surface: "row-header";
      axis: "row";
      row: number;
    }>;

export interface HeaderResizeLimits {
  readonly minSize?: number;
  readonly maxSize?: number;
}

export type HeaderResizePreview =
  | Readonly<{
      axis: "column";
      column: number;
      width: number;
      structureVersion: number;
    }>
  | Readonly<{
      axis: "row";
      row: number;
      height: number;
      structureVersion: number;
    }>;

export type HeaderResizeIntent = HeaderResizePreview;

export interface HeaderResizeControllerOptions {
  readonly handleTolerance?: number;
  readonly rowGripWidth?: number;
  readonly getColumnLimits?: (column: number) => HeaderResizeLimits | undefined;
  readonly getRowLimits?: (row: number) => HeaderResizeLimits | undefined;
}

interface DragState {
  readonly handle: HeaderResizeHandle;
  readonly startAxisPoint: number;
  readonly startSize: number;
  readonly currentSize: number;
  readonly minSize: number;
  readonly maxSize: number;
  readonly structureVersion: number;
}

function assertNonNegativeFinite(name: string, value: number): void {
  if (!Number.isFinite(value) || value < 0) {
    throw new RangeError(`${name} must be a non-negative finite number: ${value}`);
  }
}

function pointInSurface(point: HeaderResizePoint, scene: SceneSnapshot): boolean {
  if (!Number.isFinite(point.x) || !Number.isFinite(point.y)) {
    return false;
  }
  return scene.getSurface(point.surface).panes.some(({ clipRect }) => (
    point.x >= clipRect.x &&
    point.y >= clipRect.y &&
    point.x < clipRect.x + clipRect.width &&
    point.y < clipRect.y + clipRect.height
  ));
}

function freezePreview(
  handle: HeaderResizeHandle,
  size: number,
  structureVersion: number
): HeaderResizePreview {
  return handle.axis === "column"
    ? Object.freeze({
        axis: "column",
        column: handle.column,
        width: size,
        structureVersion
      })
    : Object.freeze({
        axis: "row",
        row: handle.row,
        height: size,
        structureVersion
      });
}

/**
 * Finds resize handles in canonical header-surface coordinates. Column
 * resize is restricted to the leaf-header band; row resize additionally
 * requires the narrow grip at the row header's trailing edge.
 */
export function hitTestHeaderResizeHandle(
  point: HeaderResizePoint,
  scene: SceneSnapshot,
  options: Pick<HeaderResizeControllerOptions, "handleTolerance" | "rowGripWidth"> = {}
): HeaderResizeHandle | null {
  if (!pointInSurface(point, scene)) {
    return null;
  }

  const tolerance = options.handleTolerance ?? HEADER_RESIZE_HANDLE_TOLERANCE;
  assertNonNegativeFinite("handleTolerance", tolerance);

  if (point.surface === "column-header") {
    const hit = scene.hitTestColumnHeader(point);
    if (!hit || hit.band.role !== "leaf") {
      return null;
    }
    const column = scene.hitTestColumnBoundary(point.x, tolerance);
    return column === null
      ? null
      : Object.freeze({ surface: "column-header", axis: "column", column });
  }

  const rowGripWidth = options.rowGripWidth ?? ROW_RESIZE_GRIP_WIDTH;
  assertNonNegativeFinite("rowGripWidth", rowGripWidth);
  const gripStart = Math.max(0, scene.surfaceMetrics.rowHeaderWidth - rowGripWidth);
  if (point.x < gripStart) {
    return null;
  }
  const row = scene.hitTestRowBoundary(point.y, tolerance);
  return row === null
    ? null
    : Object.freeze({ surface: "row-header", axis: "row", row });
}

/**
 * DOM-agnostic header resize state machine. Its caller supplies one fresh
 * immutable Scene and surface-local point per receipt. Move returns a
 * transient preview value; pointerup returns a commit intent. Neither path
 * mutates layout or owns pointer listeners/capture.
 */
export class HeaderResizeController {
  readonly #options: HeaderResizeControllerOptions;
  #drag: DragState | null = null;
  #preview: HeaderResizePreview | null = null;
  #previewPublished = false;

  constructor(options: HeaderResizeControllerOptions = {}) {
    this.#options = options;
  }

  get dragging(): boolean {
    return this.#drag !== null;
  }

  get preview(): HeaderResizePreview | null {
    return this.#preview;
  }

  /**
   * True once a move result has required transient geometry. The owner must
   * read this before `handlePointerUp()` or `cancel()`, both of which clear
   * gesture state. When termination yields no commit intent and this was
   * true, the owner restores its authoritative Scene after releasing the
   * interaction-frame read lease.
   */
  get previewPublished(): boolean {
    return this.#previewPublished;
  }

  hitTest(point: HeaderResizePoint, scene: SceneSnapshot): HeaderResizeHandle | null {
    return hitTestHeaderResizeHandle(point, scene, this.#options);
  }

  handlePointerDown(
    point: HeaderResizePoint,
    scene: SceneSnapshot,
    structureVersion: number
  ): boolean {
    this.cancel();
    const handle = this.hitTest(point, scene);
    if (!handle) {
      return false;
    }

    const limits = handle.axis === "column"
      ? this.#options.getColumnLimits?.(handle.column)
      : this.#options.getRowLimits?.(handle.row);
    const defaultMin = handle.axis === "column"
      ? DEFAULT_MIN_COLUMN_RESIZE_SIZE
      : DEFAULT_MIN_ROW_RESIZE_SIZE;
    const minSize = limits?.minSize ?? defaultMin;
    const maxSize = limits?.maxSize ?? Number.POSITIVE_INFINITY;
    assertNonNegativeFinite("HeaderResizeLimits.minSize", minSize);
    if ((!Number.isFinite(maxSize) && maxSize !== Number.POSITIVE_INFINITY) || maxSize < 0) {
      throw new RangeError(`HeaderResizeLimits.maxSize must be non-negative: ${maxSize}`);
    }
    if (maxSize < minSize) {
      throw new RangeError("HeaderResizeLimits.maxSize cannot be smaller than minSize");
    }

    const startSize = handle.axis === "column"
      ? scene.getColumnWidth(handle.column)
      : scene.getRowHeight(handle.row);
    this.#drag = {
      handle,
      startAxisPoint: handle.axis === "column" ? point.x : point.y,
      startSize,
      currentSize: startSize,
      minSize,
      maxSize,
      structureVersion
    };
    return true;
  }

  handlePointerMove(
    point: HeaderResizePoint,
    scene: SceneSnapshot,
    structureVersion: number
  ): HeaderResizePreview | null {
    const drag = this.#drag;
    if (!drag) {
      return null;
    }
    if (
      structureVersion !== drag.structureVersion ||
      point.surface !== drag.handle.surface ||
      !this.#targetExists(drag.handle, scene)
    ) {
      this.cancel();
      return null;
    }

    const axisPoint = drag.handle.axis === "column" ? point.x : point.y;
    const nextSize = Math.min(
      drag.maxSize,
      Math.max(drag.minSize, drag.startSize + axisPoint - drag.startAxisPoint)
    );
    if (nextSize === drag.currentSize) {
      return null;
    }

    this.#drag = { ...drag, currentSize: nextSize };
    this.#preview = freezePreview(drag.handle, nextSize, drag.structureVersion);
    this.#previewPublished = true;
    return this.#preview;
  }

  handlePointerUp(structureVersion: number): HeaderResizeIntent | null {
    const drag = this.#drag;
    if (!drag) {
      return null;
    }
    if (structureVersion !== drag.structureVersion) {
      this.cancel();
      return null;
    }

    const intent = drag.currentSize === drag.startSize
      ? null
      : freezePreview(drag.handle, drag.currentSize, drag.structureVersion);
    this.#clear();
    return intent;
  }

  cancel(): void {
    this.#clear();
  }

  #targetExists(handle: HeaderResizeHandle, scene: SceneSnapshot): boolean {
    return handle.axis === "column"
      ? handle.column < scene.columnCount
      : handle.row < scene.rowCount;
  }

  #clear(): void {
    this.#drag = null;
    this.#preview = null;
    this.#previewPublished = false;
  }
}
