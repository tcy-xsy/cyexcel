const DEFAULT_MAX_DPR = 3;

/**
 * Converts a logical (CSS pixel) size to a physical (device pixel) size for
 * sizing a Canvas element's backing store. `dpr` is capped at `maxDpr` so a
 * very high-density display doesn't blow up the pixel count — per
 * doc/architecture.md §6.
 */
export function toPhysicalPixelSize(
  logicalSize: number,
  dpr: number,
  maxDpr: number = DEFAULT_MAX_DPR
): number {
  return Math.round(logicalSize * Math.min(dpr, maxDpr));
}
