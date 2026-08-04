import { toPhysicalPixelSize } from "../layout";

export interface ResizableCanvas {
  width: number;
  height: number;
  readonly style: {
    width: string;
    height: string;
  };
}

/**
 * Sizes a canvas element's backing store for a given device pixel ratio:
 * physical `width`/`height` attributes, logical CSS `style.width`/`height`.
 * Does not call `ctx.scale()` — that needs the 2D context, not the canvas
 * element, and is a one-time surface-setup step the caller does after
 * calling this. ContentLayerRenderer/SelectionLayerRenderer never see a
 * dpr — they always draw in logical coordinates.
 */
export function configureCanvasForDpr(
  canvas: ResizableCanvas,
  logicalWidth: number,
  logicalHeight: number,
  dpr: number,
  maxDpr?: number
): void {
  canvas.width = toPhysicalPixelSize(logicalWidth, dpr, maxDpr);
  canvas.height = toPhysicalPixelSize(logicalHeight, dpr, maxDpr);
  canvas.style.width = `${logicalWidth}px`;
  canvas.style.height = `${logicalHeight}px`;
}
