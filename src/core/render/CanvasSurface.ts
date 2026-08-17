import { configureCanvasForDpr, getEffectiveDpr, type ResizableCanvas } from "./dprCanvas";

export interface CanvasTransformContext {
  setTransform(a: number, b: number, c: number, d: number, e: number, f: number): void;
}

export interface CanvasSurfaceSnapshot {
  readonly cssWidth: number;
  readonly cssHeight: number;
  readonly effectiveDpr: number;
  readonly physicalWidth: number;
  readonly physicalHeight: number;
}

/** The only owner of a canvas's effective DPR and backing-store transform. */
export class CanvasSurface {
  readonly #canvas: ResizableCanvas;
  readonly #context: CanvasTransformContext;
  readonly #maxDpr: number | undefined;

  constructor(canvas: ResizableCanvas, context: CanvasTransformContext, maxDpr?: number) {
    this.#canvas = canvas;
    this.#context = context;
    this.#maxDpr = maxDpr;
  }

  resize(cssWidth: number, cssHeight: number, devicePixelRatio: number): CanvasSurfaceSnapshot {
    const effectiveDpr = getEffectiveDpr(devicePixelRatio, this.#maxDpr);
    configureCanvasForDpr(this.#canvas, cssWidth, cssHeight, devicePixelRatio, this.#maxDpr);
    this.#context.setTransform(effectiveDpr, 0, 0, effectiveDpr, 0, 0);

    return Object.freeze({
      cssWidth,
      cssHeight,
      effectiveDpr,
      physicalWidth: this.#canvas.width,
      physicalHeight: this.#canvas.height
    });
  }
}
