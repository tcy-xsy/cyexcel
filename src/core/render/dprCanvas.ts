const DEFAULT_MAX_DPR = 3;

export interface ResizableCanvas {
  width: number;
  height: number;
  readonly style: {
    width: string;
    height: string;
  };
}

/**
 * The single read of the `devicePixelRatio` global in the whole core.
 *
 * Everything downstream (CanvasSurface, configureCanvasForDpr, the renderers)
 * takes a dpr *parameter* so it stays injectable in unit tests; this function
 * exists so that the one place which must consult the environment is inside
 * the surface layer that owns backing stores. The lint gate in
 * eslint.config.js enforces that no other module touches the global.
 *
 * Guards `window` so importing the core in a non-DOM runtime stays safe.
 */
export function readDevicePixelRatio(): number {
  if (typeof window === "undefined") {
    return 1;
  }
  return window.devicePixelRatio || 1;
}

/** The exact ratio used for both backing-store size and the context transform. */
export function getEffectiveDpr(devicePixelRatio: number, maxDpr: number = DEFAULT_MAX_DPR): number {
  if (!Number.isFinite(devicePixelRatio) || devicePixelRatio <= 0) {
    return 1;
  }
  if (!Number.isFinite(maxDpr) || maxDpr <= 0) {
    throw new RangeError(`maxDpr must be a positive finite number: ${maxDpr}`);
  }
  return Math.min(devicePixelRatio, maxDpr);
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
  const effectiveDpr = getEffectiveDpr(dpr, maxDpr);
  canvas.width = Math.round(logicalWidth * effectiveDpr);
  canvas.height = Math.round(logicalHeight * effectiveDpr);
  canvas.style.width = `${logicalWidth}px`;
  canvas.style.height = `${logicalHeight}px`;
}
