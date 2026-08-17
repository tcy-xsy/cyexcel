import type { CellRect } from "../layout";
import { declareDomPointerOwnership } from "../interaction/domPointerOwnership";
import type { SceneColumnHeaderFragment, SceneSnapshot } from "../scene";
import type { HeaderManager } from "./HeaderManager";
import type { DomHeaderContext, DomHeaderRenderer, DomHeaderRendererFactory, HeaderState } from "./types";

export interface DomHeaderManagerOptions {
  /** Overlay element DOM-rendered headers mount into; must already be connected to the document. */
  readonly container: HTMLElement;
  readonly headerManager: HeaderManager;
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
  readonly #headerManager: HeaderManager;
  readonly #getHeaderState: (column: number) => HeaderState;
  readonly #createHost: () => HTMLElement;

  readonly #mounted = new Map<number, MountedHeader>();
  readonly #pool = new Map<DomHeaderRendererFactory, PooledInstance[]>();

  constructor(options: DomHeaderManagerOptions) {
    this.#container = options.container;
    this.#headerManager = options.headerManager;
    this.#getHeaderState = options.getHeaderState ?? (() => ({}));
    this.#createHost = options.createHost ?? (() => document.createElement("div"));
  }

  sync(scene: SceneSnapshot): void {
    const needed = this.#neededColumns(scene);

    for (const column of [...this.#mounted.keys()]) {
      if (!needed.has(column)) {
        this.#release(column);
      }
    }
    for (const [column, fragment] of needed) {
      this.#ensureMounted(column, fragment);
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

  #neededColumns(scene: SceneSnapshot): Map<number, SceneColumnHeaderFragment> {
    const columns = new Map<number, SceneColumnHeaderFragment>();
    const band = scene.getColumnHeaderLeafBand();
    for (const paneFragment of scene.getVisibleColumnHeaderFragments(band)) {
      for (let offset = 0; offset < paneFragment.columns.count; offset += 1) {
        const column = paneFragment.columns.start + offset;
        if (
          !columns.has(column) &&
          this.#headerManager.getHeader(column).renderer?.kind === "dom"
        ) {
          const fragment = scene
            .getColumnHeaderBandFragments(band, column, column)
            .find((candidate) => candidate.pane === paneFragment.pane);
          if (fragment) {
            columns.set(column, fragment);
          }
        }
      }
    }
    return columns;
  }

  #ensureMounted(column: number, fragment: SceneColumnHeaderFragment): void {
    const header = this.#headerManager.getHeader(column);
    const factory = header.renderer as DomHeaderRendererFactory;
    const rect = fragment.rect;
    const context: DomHeaderContext = {
      column,
      title: header.title,
      rect,
      state: this.#getHeaderState(column)
    };

    const existing = this.#mounted.get(column);
    if (existing) {
      this.#positionHost(existing.host, rect, fragment.clipRect);
      existing.instance.update(context);
      return;
    }

    const pooled = this.#pool.get(factory)?.pop();
    if (pooled) {
      declareDomPointerOwnership(pooled.host, factory.pointerOwnership, "column-header");
      this.#positionHost(pooled.host, rect, fragment.clipRect);
      this.#container.append(pooled.host);
      try {
        pooled.instance.update(context);
      } catch (error) {
        pooled.host.remove();
        try {
          pooled.instance.destroy();
        } catch {
          // Preserve the update failure; the broken pooled instance is discarded.
        }
        throw error;
      }
      this.#mounted.set(column, { column, factory, host: pooled.host, instance: pooled.instance });
      return;
    }

    const host = this.#createHost();
    declareDomPointerOwnership(host, factory.pointerOwnership, "column-header");
    this.#positionHost(host, rect, fragment.clipRect);
    this.#container.append(host);
    let instance: DomHeaderRenderer | undefined;
    try {
      instance = factory.create();
      instance.mount(host, context);
    } catch (error) {
      host.remove();
      try {
        instance?.destroy();
      } catch {
        // Preserve the create/mount failure; the partial instance is discarded.
      }
      throw error;
    }
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

  #positionHost(host: HTMLElement, rect: CellRect, clipRect: CellRect): void {
    host.style.position = "absolute";
    host.style.left = `${rect.x}px`;
    host.style.top = `${rect.y}px`;
    host.style.width = `${rect.width}px`;
    host.style.height = `${rect.height}px`;
    const topInset = Math.max(0, clipRect.y - rect.y);
    const leftInset = Math.max(0, clipRect.x - rect.x);
    const rightInset = Math.max(
      0,
      rect.x + rect.width - (clipRect.x + clipRect.width)
    );
    const bottomInset = Math.max(
      0,
      rect.y + rect.height - (clipRect.y + clipRect.height)
    );
    host.style.clipPath = topInset > 0 || rightInset > 0 || bottomInset > 0 || leftInset > 0
      ? `inset(${topInset}px ${rightInset}px ${bottomInset}px ${leftInset}px)`
      : "";
  }

}
