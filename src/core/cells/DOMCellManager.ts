import type { CellRect, LayoutEngine, ViewportState } from "../layout";
import { toViewportRect } from "../layout";
import type { DomCellContext, DomCellRenderer, DomCellRendererFactory, RuntimeColumn } from "./types";

export interface DOMCellManagerOptions {
  /** Overlay element DOM-rendered cells mount into; must already be connected to the document. */
  readonly container: HTMLElement;
  readonly layout: LayoutEngine;
  readonly getColumn: (column: number) => RuntimeColumn;
  readonly getCellValue: (row: number, column: number) => unknown;
  readonly commitValue: (row: number, column: number, value: unknown) => void;
  /** Injected for Node testability (mirrors RenderScheduler's frameScheduler / EditorManager's editorFactory precedent). Defaults to a real `<div>`. */
  readonly createHost?: () => HTMLElement;
}

interface MountedCell {
  readonly row: number;
  readonly column: number;
  readonly factory: DomCellRendererFactory;
  readonly host: HTMLElement;
  readonly instance: DomCellRenderer;
}

interface PooledInstance {
  readonly host: HTMLElement;
  readonly instance: DomCellRenderer;
}

function cellKey(row: number, column: number): string {
  return `${row}:${column}`;
}

/**
 * Mounts real DOM nodes for cells whose resolved column uses a
 * DomCellRendererFactory, and only those cells — everything else is left to
 * ContentLayerRenderer's canvas painting (see createColumnAwarePainter,
 * which skips drawing `kind: "dom"` cells entirely so the two never
 * overlap). Reuses renderer instances by factory identity as cells scroll
 * in and out of view instead of destroying and recreating them, so the
 * live instance count tracks "currently visible + overscan", never the
 * total row/column count.
 */
export class DOMCellManager {
  readonly #container: HTMLElement;
  readonly #layout: LayoutEngine;
  readonly #getColumn: (column: number) => RuntimeColumn;
  readonly #getCellValue: (row: number, column: number) => unknown;
  readonly #commitValue: (row: number, column: number, value: unknown) => void;
  readonly #createHost: () => HTMLElement;

  readonly #mounted = new Map<string, MountedCell>();
  readonly #pool = new Map<DomCellRendererFactory, PooledInstance[]>();

  constructor(options: DOMCellManagerOptions) {
    this.#container = options.container;
    this.#layout = options.layout;
    this.#getColumn = options.getColumn;
    this.#getCellValue = options.getCellValue;
    this.#commitValue = options.commitValue;
    this.#createHost = options.createHost ?? (() => document.createElement("div"));
  }

  /** Recomputes which cells need a mounted DOM renderer and reconciles against what's currently mounted. */
  sync(viewport: ViewportState): void {
    const needed = this.#neededCells(viewport);

    for (const key of [...this.#mounted.keys()]) {
      if (!needed.has(key)) {
        this.#release(key);
      }
    }

    for (const [key, [row, column]] of needed) {
      this.#ensureMounted(key, row, column, viewport);
    }
  }

  dispose(): void {
    for (const cell of this.#mounted.values()) {
      cell.instance.destroy();
      cell.host.remove();
    }
    this.#mounted.clear();

    for (const pooled of this.#pool.values()) {
      for (const entry of pooled) {
        entry.instance.destroy();
      }
    }
    this.#pool.clear();
  }

  #neededCells(viewport: ViewportState): Map<string, readonly [number, number]> {
    const visibleRows = this.#layout.getVisibleRowRange(viewport);
    const visibleColumns = this.#layout.getVisibleColumnRange(viewport);

    const rows = new Set<number>();
    for (let row = 0; row < this.#layout.frozenRowCount; row += 1) {
      rows.add(row);
    }
    for (let i = 0; i < visibleRows.count; i += 1) {
      rows.add(visibleRows.start + i);
    }

    const columns = new Set<number>();
    for (let column = 0; column < this.#layout.frozenColumnCount; column += 1) {
      columns.add(column);
    }
    for (let i = 0; i < visibleColumns.count; i += 1) {
      columns.add(visibleColumns.start + i);
    }
    const domColumns = [...columns].filter((column) => this.#getColumn(column).renderer?.kind === "dom");

    const needed = new Map<string, readonly [number, number]>();
    for (const row of rows) {
      for (const column of domColumns) {
        needed.set(cellKey(row, column), [row, column]);
      }
    }
    return needed;
  }

  #ensureMounted(key: string, row: number, column: number, viewport: ViewportState): void {
    const runtimeColumn = this.#getColumn(column);
    const factory = runtimeColumn.renderer as DomCellRendererFactory;
    const rect = this.#cellRect(row, column, viewport);
    const context: DomCellContext = {
      row,
      column: runtimeColumn,
      value: this.#getCellValue(row, column),
      rect,
      actions: { commit: (value) => this.#commitValue(row, column, value) }
    };

    const existing = this.#mounted.get(key);
    if (existing) {
      this.#positionHost(existing.host, rect);
      existing.instance.update(context);
      return;
    }

    const pooled = this.#pool.get(factory)?.pop();
    if (pooled) {
      this.#positionHost(pooled.host, rect);
      this.#container.append(pooled.host);
      pooled.instance.update(context);
      this.#mounted.set(key, { row, column, factory, host: pooled.host, instance: pooled.instance });
      return;
    }

    const host = this.#createHost();
    this.#positionHost(host, rect);
    this.#container.append(host);
    const instance = factory.create();
    instance.mount(host, context);
    this.#mounted.set(key, { row, column, factory, host, instance });
  }

  #release(key: string): void {
    const cell = this.#mounted.get(key);
    if (!cell) {
      return;
    }
    this.#mounted.delete(key);
    cell.host.remove();
    const pooled = this.#pool.get(cell.factory) ?? [];
    pooled.push({ host: cell.host, instance: cell.instance });
    this.#pool.set(cell.factory, pooled);
  }

  #positionHost(host: HTMLElement, rect: CellRect): void {
    host.style.position = "absolute";
    host.style.left = `${rect.x}px`;
    host.style.top = `${rect.y}px`;
    host.style.width = `${rect.width}px`;
    host.style.height = `${rect.height}px`;
  }

  #cellRect(row: number, column: number, viewport: ViewportState): CellRect {
    return toViewportRect(this.#layout.getCellRect(row, column), viewport, {
      row: row < this.#layout.frozenRowCount,
      column: column < this.#layout.frozenColumnCount
    });
  }
}
