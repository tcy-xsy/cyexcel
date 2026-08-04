import type { LayoutEngine, ViewportState } from "../layout";
import type { Disposer } from "../types";
import type { SelectionModel, SelectionSnapshot } from "./SelectionModel";

export interface HeaderLayout {
  readonly rowHeaderWidth: number;
  readonly columnHeaderHeight: number;
}

export interface PointerControllerOptions {
  readonly layout: LayoutEngine;
  readonly selection: SelectionModel;
  readonly headerLayout?: HeaderLayout;
  readonly onSelectionChange?: (snapshot: SelectionSnapshot) => void;
}

interface PointerModifiers {
  readonly shift: boolean;
}

const NO_HEADERS: HeaderLayout = { rowHeaderWidth: 0, columnHeaderHeight: 0 };

function frozenColumnWidth(layout: LayoutEngine): number {
  let width = 0;
  for (let column = 0; column < layout.frozenColumnCount; column += 1) {
    width += layout.getColumnWidth(column);
  }
  return width;
}

/**
 * Converts a viewport-relative offset to a row/column index, always
 * clamped to valid bounds — unlike LayoutEngine.getCellAtPoint(), which
 * returns null past the content edge (right for "did I click a real cell",
 * wrong for drag-selection continuing past the table's edge).
 */
function resolveRow(y: number, layout: LayoutEngine, viewport: ViewportState): number {
  const frozenHeight = layout.frozenRowCount * layout.rowHeight;
  const contentY = y < frozenHeight ? y : Math.max(0, viewport.scrollTop) + y;
  return layout.getRowAtY(contentY);
}

function resolveColumn(x: number, layout: LayoutEngine, viewport: ViewportState): number {
  const frozenWidth = frozenColumnWidth(layout);
  const contentX = x < frozenWidth ? x : Math.max(0, viewport.scrollLeft) + x;
  return layout.getColumnAtX(contentX);
}

/**
 * Mouse/pointer input for selection. Knows nothing about RenderScheduler or
 * TableStore — callers observe changes via onSelectionChange and decide
 * what to do (e.g. call renderScheduler.invalidateSelection()).
 */
export class PointerController {
  readonly #layout: LayoutEngine;
  readonly #selection: SelectionModel;
  readonly #headerLayout: HeaderLayout;
  readonly #onSelectionChange: ((snapshot: SelectionSnapshot) => void) | undefined;
  #viewport: ViewportState = { scrollTop: 0, scrollLeft: 0, width: 0, height: 0 };
  #dragging = false;

  constructor(options: PointerControllerOptions) {
    this.#layout = options.layout;
    this.#selection = options.selection;
    this.#headerLayout = options.headerLayout ?? NO_HEADERS;
    this.#onSelectionChange = options.onSelectionChange;
  }

  setViewport(viewport: ViewportState): void {
    this.#viewport = viewport;
  }

  handlePointerDown(
    point: { x: number; y: number },
    modifiers: PointerModifiers = { shift: false }
  ): void {
    const inRowHeader = point.x < this.#headerLayout.rowHeaderWidth;
    const inColumnHeader = point.y < this.#headerLayout.columnHeaderHeight;
    const { row, column } = this.#resolveGridPosition(point);

    if (modifiers.shift) {
      this.#selection.extendTo(row, column);
    } else if (inRowHeader && inColumnHeader) {
      this.#selection.selectAll();
    } else if (inRowHeader) {
      this.#selection.selectRow(row);
    } else if (inColumnHeader) {
      this.#selection.selectColumn(column);
    } else {
      this.#selection.selectCell(row, column);
    }

    this.#dragging = true;
    this.#notify();
  }

  handlePointerMove(point: { x: number; y: number }): void {
    if (!this.#dragging) {
      return;
    }
    const { row, column } = this.#resolveGridPosition(point);
    this.#selection.extendTo(row, column);
    this.#notify();
  }

  handlePointerUp(): void {
    this.#dragging = false;
  }

  attach(element: HTMLElement): Disposer {
    const toPoint = (event: PointerEvent): { x: number; y: number } => {
      const rect = element.getBoundingClientRect();
      return { x: event.clientX - rect.left, y: event.clientY - rect.top };
    };
    const onPointerDown = (event: PointerEvent): void => {
      this.handlePointerDown(toPoint(event), { shift: event.shiftKey });
    };
    const onPointerMove = (event: PointerEvent): void => {
      this.handlePointerMove(toPoint(event));
    };
    const onPointerUp = (): void => {
      this.handlePointerUp();
    };

    element.addEventListener("pointerdown", onPointerDown);
    window.addEventListener("pointermove", onPointerMove);
    window.addEventListener("pointerup", onPointerUp);

    return () => {
      element.removeEventListener("pointerdown", onPointerDown);
      window.removeEventListener("pointermove", onPointerMove);
      window.removeEventListener("pointerup", onPointerUp);
    };
  }

  #resolveGridPosition(point: { x: number; y: number }): { row: number; column: number } {
    const gridX = Math.max(0, point.x - this.#headerLayout.rowHeaderWidth);
    const gridY = Math.max(0, point.y - this.#headerLayout.columnHeaderHeight);
    return {
      row: resolveRow(gridY, this.#layout, this.#viewport),
      column: resolveColumn(gridX, this.#layout, this.#viewport)
    };
  }

  #notify(): void {
    this.#onSelectionChange?.(this.#selection.snapshot);
  }
}
