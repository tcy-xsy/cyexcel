import type { SceneCellFragment, SceneSnapshot } from "../scene";
import { declareDomPointerOwnership } from "../interaction/domPointerOwnership";
import type { DomCellContext, DomCellRenderer, DomCellRendererFactory, RuntimeColumn } from "./types";

export interface DOMCellManagerOptions {
  /** Overlay element DOM-rendered cells mount into; must already be connected to the document. */
  readonly container: HTMLElement;
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
  readonly binding: CommitBinding;
}

interface PooledInstance {
  readonly host: HTMLElement;
  readonly instance: DomCellRenderer;
}

interface NeededCell {
  readonly row: number;
  readonly column: number;
  readonly fragment: SceneCellFragment;
}

interface CommitBinding {
  readonly row: number;
  readonly column: number;
  active: boolean;
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
  readonly #getColumn: (column: number) => RuntimeColumn;
  readonly #getCellValue: (row: number, column: number) => unknown;
  readonly #commitValue: (row: number, column: number, value: unknown) => void;
  readonly #createHost: () => HTMLElement;

  readonly #mounted = new Map<string, MountedCell>();
  readonly #pool = new Map<DomCellRendererFactory, PooledInstance[]>();

  constructor(options: DOMCellManagerOptions) {
    this.#container = options.container;
    this.#getColumn = options.getColumn;
    this.#getCellValue = options.getCellValue;
    this.#commitValue = options.commitValue;
    this.#createHost = options.createHost ?? (() => document.createElement("div"));
  }

  /** Recomputes which cells need a mounted DOM renderer and reconciles against what's currently mounted. */
  sync(scene: SceneSnapshot): void {
    const needed = this.#neededCells(scene);

    for (const key of [...this.#mounted.keys()]) {
      if (!needed.has(key)) {
        this.#release(key);
      }
    }

    for (const [key, cell] of needed) {
      this.#ensureMounted(key, cell);
    }
  }

  dispose(): void {
    for (const cell of this.#mounted.values()) {
      cell.binding.active = false;
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

  #neededCells(scene: SceneSnapshot): Map<string, NeededCell> {
    const needed = new Map<string, NeededCell>();
    for (const fragment of scene.getVisibleFragments()) {
      for (let column = fragment.columns.start; column < fragment.columns.start + fragment.columns.count; column += 1) {
        if (this.#getColumn(column).renderer?.kind !== "dom") {
          continue;
        }
        for (let row = fragment.rows.start; row < fragment.rows.start + fragment.rows.count; row += 1) {
          const cellFragment = scene.getCellFragments({ row, column })[0];
          if (cellFragment) {
            needed.set(cellKey(row, column), { row, column, fragment: cellFragment });
          }
        }
      }
    }
    return needed;
  }

  #ensureMounted(key: string, cell: NeededCell): void {
    const { row, column, fragment } = cell;
    const runtimeColumn = this.#getColumn(column);
    const factory = runtimeColumn.renderer as DomCellRendererFactory;
    const { rect } = fragment;
    const existing = this.#mounted.get(key);
    if (existing) {
      this.#positionHost(existing.host, fragment);
      existing.instance.update(this.#createContext(runtimeColumn, rect, existing.binding));
      return;
    }

    const binding: CommitBinding = { row, column, active: true };
    const context = this.#createContext(runtimeColumn, rect, binding);
    const pooled = this.#pool.get(factory)?.pop();
    if (pooled) {
      declareDomPointerOwnership(pooled.host, factory.pointerOwnership, "body");
      this.#positionHost(pooled.host, fragment);
      this.#container.append(pooled.host);
      try {
        pooled.instance.update(context);
      } catch (error) {
        binding.active = false;
        pooled.host.remove();
        try {
          pooled.instance.destroy();
        } catch {
          // Preserve the update failure; the broken pooled instance is discarded.
        }
        throw error;
      }
      this.#mounted.set(key, {
        row,
        column,
        factory,
        host: pooled.host,
        instance: pooled.instance,
        binding
      });
      return;
    }

    const host = this.#createHost();
    declareDomPointerOwnership(host, factory.pointerOwnership, "body");
    this.#positionHost(host, fragment);
    this.#container.append(host);
    let instance: DomCellRenderer | undefined;
    try {
      instance = factory.create();
      instance.mount(host, context);
    } catch (error) {
      binding.active = false;
      host.remove();
      try {
        instance?.destroy();
      } catch {
        // Preserve the create/mount failure; the partial instance is discarded.
      }
      throw error;
    }
    this.#mounted.set(key, { row, column, factory, host, instance, binding });
  }

  #release(key: string): void {
    const cell = this.#mounted.get(key);
    if (!cell) {
      return;
    }
    this.#mounted.delete(key);
    cell.binding.active = false;
    cell.host.remove();
    const pooled = this.#pool.get(cell.factory) ?? [];
    pooled.push({ host: cell.host, instance: cell.instance });
    this.#pool.set(cell.factory, pooled);
  }

  #createContext(
    runtimeColumn: RuntimeColumn,
    rect: SceneCellFragment["rect"],
    binding: CommitBinding
  ): DomCellContext {
    return {
      row: binding.row,
      column: runtimeColumn,
      value: this.#getCellValue(binding.row, binding.column),
      rect,
      actions: Object.freeze({
        commit: (value: unknown) => {
          if (binding.active) {
            this.#commitValue(binding.row, binding.column, value);
          }
        }
      })
    };
  }

  #positionHost(host: HTMLElement, fragment: SceneCellFragment): void {
    const { rect, pane } = fragment;
    host.style.position = "absolute";
    host.style.left = `${rect.x}px`;
    host.style.top = `${rect.y}px`;
    host.style.width = `${rect.width}px`;
    host.style.height = `${rect.height}px`;
    const topInset = Math.max(0, pane.clipRect.y - rect.y);
    const rightInset = Math.max(
      0,
      rect.x + rect.width - (pane.clipRect.x + pane.clipRect.width)
    );
    const bottomInset = Math.max(
      0,
      rect.y + rect.height - (pane.clipRect.y + pane.clipRect.height)
    );
    const leftInset = Math.max(0, pane.clipRect.x - rect.x);
    host.style.clipPath = topInset > 0 || rightInset > 0 || bottomInset > 0 || leftInset > 0
      ? `inset(${topInset}px ${rightInset}px ${bottomInset}px ${leftInset}px)`
      : "";
  }
}
