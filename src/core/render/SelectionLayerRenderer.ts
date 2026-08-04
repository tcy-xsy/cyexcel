import type { Canvas2DLike } from "./Canvas2DLike";

export interface SelectionLayerRendererOptions {
  readonly ctx: Canvas2DLike;
}

/**
 * Minimal skeleton proving the selection layer can be invalidated and
 * repainted independently of the content layer (RenderScheduler routes
 * invalidateSelection() here without touching ContentLayerRenderer). Real
 * selection-box drawing lands in Step 6 once SelectionManager exists — for
 * now this only clears its own canvas area.
 */
export class SelectionLayerRenderer {
  readonly #ctx: Canvas2DLike;
  #width = 0;
  #height = 0;

  constructor(options: SelectionLayerRendererOptions) {
    this.#ctx = options.ctx;
  }

  setSize(width: number, height: number): void {
    this.#width = width;
    this.#height = height;
  }

  paintSelection(): void {
    this.#ctx.clearRect(0, 0, this.#width, this.#height);
  }
}
