import type { CellRect, LayoutEngine, ViewportState } from "../layout";
import { toViewportRect } from "../layout";
import type { HeaderManager } from "./HeaderManager";
import type { DomHeaderContext, DomHeaderRenderer, DomHeaderRendererFactory, HeaderState } from "./types";

export interface DomHeaderManagerOptions {
  /** Overlay element DOM-rendered headers mount into; must already be connected to the document. */
  readonly container: HTMLElement;
  readonly layout: LayoutEngine;
  readonly headerManager: HeaderManager;
  readonly columnHeaderHeight: number;
  readonly groupHeaderHeight?: number;
  readonly getHeaderState?: (column: number) => HeaderState;
  /** Injected for Node testability (same precedent as DOMCellManager.createHost). Defaults to a real `<div>`. */
  readonly createHost?: () => HTMLElement;
}

interface MountedHeader {
  readonly column: number;
  readonly factory: DomHeaderRendererFactory;
  readonly host: HTMLElement;
  readonly instance: DomHeaderRenderer;
}

interface PooledInstance {
  readonly host: HTMLElement;
  readonly instance: DomHeaderRenderer;
}

/**
 * One-dimensional analog of DOMCellManager (Step 8) for the header band —
 * mounts real DOM nodes only for leaf columns whose header uses a
 * DomHeaderRendererFactory, only within the frozen ∪ visible column set,
 * pooling instances by factory identity across scroll instead of
 * destroying and recreating them. Deliberately not unified with
 * DOMCellManager into a shared base class — same pooling idea, different
 * dimensionality (columns only, no rows) — see the Step 9 plan's "不做的事".
 */
export class DomHeaderManager {
  readonly #container: HTMLElement;
  readonly #layout: LayoutEngine;
  readonly #headerManager: HeaderManager;
  readonly #columnHeaderHeight: number;
  readonly #groupHeaderHeight: number;
  readonly #getHeaderState: (column: number) => HeaderState;
  readonly #createHost: () => HTMLElement;

  readonly #mounted = new Map<number, MountedHeader>();
  readonly #pool = new Map<DomHeaderRendererFactory, PooledInstance[]>();

  constructor(options: DomHeaderManagerOptions) {
    this.#container = options.container;
    this.#layout = options.layout;
    this.#headerManager = options.headerManager;
    this.#columnHeaderHeight = options.columnHeaderHeight;
    this.#groupHeaderHeight = options.groupHeaderHeight ?? 0;
    this.#getHeaderState = options.getHeaderState ?? (() => ({}));
    this.#createHost = options.createHost ?? (() => document.createElement("div"));
  }

  sync(viewport: ViewportState): void {
    const needed = this.#neededColumns(viewport);

    for (const column of [...this.#mounted.keys()]) {
      if (!needed.has(column)) {
        this.#release(column);
      }
    }
    for (const column of needed) {
      this.#ensureMounted(column, viewport);
    }
  }

  dispose(): void {
    for (const header of this.#mounted.values()) {
      header.instance.destroy();
      header.host.remove();
    }
    this.#mounted.clear();

    for (const pooled of this.#pool.values()) {
      for (const entry of pooled) {
        entry.instance.destroy();
      }
    }
    this.#pool.clear();
  }

  #neededColumns(viewport: ViewportState): Set<number> {
    const columns = new Set<number>();
    for (let column = 0; column < this.#layout.frozenColumnCount; column += 1) {
      columns.add(column);
    }
    const visibleColumns = this.#layout.getVisibleColumnRange(viewport);
    for (let i = 0; i < visibleColumns.count; i += 1) {
      columns.add(visibleColumns.start + i);
    }

    return new Set(
      [...columns].filter((column) => this.#headerManager.getHeader(column).renderer?.kind === "dom")
    );
  }

  #ensureMounted(column: number, viewport: ViewportState): void {
    const header = this.#headerManager.getHeader(column);
    const factory = header.renderer as DomHeaderRendererFactory;
    const rect = this.#leafRect(column, viewport);
    const context: DomHeaderContext = {
      column,
      title: header.title,
      rect,
      state: this.#getHeaderState(column)
    };

    const existing = this.#mounted.get(column);
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
      this.#mounted.set(column, { column, factory, host: pooled.host, instance: pooled.instance });
      return;
    }

    const host = this.#createHost();
    this.#positionHost(host, rect);
    this.#container.append(host);
    const instance = factory.create();
    instance.mount(host, context);
    this.#mounted.set(column, { column, factory, host, instance });
  }

  #release(column: number): void {
    const header = this.#mounted.get(column);
    if (!header) {
      return;
    }
    this.#mounted.delete(column);
    header.host.remove();
    const pooled = this.#pool.get(header.factory) ?? [];
    pooled.push({ host: header.host, instance: header.instance });
    this.#pool.set(header.factory, pooled);
  }

  #positionHost(host: HTMLElement, rect: CellRect): void {
    host.style.position = "absolute";
    host.style.left = `${rect.x}px`;
    host.style.top = `${rect.y}px`;
    host.style.width = `${rect.width}px`;
    host.style.height = `${rect.height}px`;
  }

  #leafRect(column: number, viewport: ViewportState): CellRect {
    const absolute: CellRect = {
      x: this.#layout.getColumnLeft(column),
      y: this.#groupHeaderHeight,
      width: this.#layout.getColumnWidth(column),
      height: this.#columnHeaderHeight
    };
    return toViewportRect(absolute, viewport, {
      row: true,
      column: column < this.#layout.frozenColumnCount
    });
  }
}
