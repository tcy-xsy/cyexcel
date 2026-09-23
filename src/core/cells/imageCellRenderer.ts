import type { Canvas2DLike } from "../render";
import { ImageCache, sharedImageCache } from "../render/ImageCache";
import type { CanvasCellRenderer, CellRenderContext } from "./types";

export interface ImageCellRendererOptions {
  /** Inset between the cell border and the drawn image, in logical pixels. Defaults to 2. */
  readonly padding?: number;
  /** Defaults to the module-level `sharedImageCache` — override only to isolate caches (e.g. in tests). */
  readonly cache?: ImageCache;
  /**
   * Derives the URL to load/draw from the cell's raw value. Defaults to
   * using the value as-is — override it when a cell's value isn't a bare
   * URL, e.g. a multi-image gallery field, so only one image renders as the
   * cell's thumbnail. A `undefined` return paints nothing (same as an empty
   * cell).
   */
  readonly resolveUrl?: (value: string) => string | undefined;
}

const PLACEHOLDER_FILL = "#e5e7eb";
const PLACEHOLDER_ICON = "#9ca3af";

/** A generic "picture" glyph (frame + sun + mountain) for a cell whose image hasn't loaded yet, or failed to. */
function drawPlaceholder(ctx: Canvas2DLike, x: number, y: number, width: number, height: number): void {
  ctx.fillStyle = PLACEHOLDER_FILL;
  ctx.fillRect(x, y, width, height);
  ctx.strokeStyle = PLACEHOLDER_ICON;
  ctx.lineWidth = 1;
  ctx.strokeRect(x + 0.5, y + 0.5, Math.max(0, width - 1), Math.max(0, height - 1));

  ctx.fillStyle = PLACEHOLDER_ICON;
  ctx.beginPath();
  ctx.arc(x + width * 0.3, y + height * 0.32, Math.max(1, Math.min(width, height) * 0.08), 0, Math.PI * 2);
  ctx.fill();

  ctx.beginPath();
  ctx.moveTo(x + width * 0.12, y + height * 0.82);
  ctx.lineTo(x + width * 0.4, y + height * 0.48);
  ctx.lineTo(x + width * 0.58, y + height * 0.68);
  ctx.lineTo(x + width * 0.78, y + height * 0.38);
  ctx.lineTo(x + width * 0.88, y + height * 0.82);
  ctx.closePath();
  ctx.fill();
}

/** Fits `image` into (width x height), preserving aspect ratio and centering — "contain", never cropping. */
function drawContained(
  ctx: Canvas2DLike,
  image: HTMLImageElement,
  x: number,
  y: number,
  width: number,
  height: number
): void {
  const scale = Math.min(width / image.width, height / image.height);
  const drawWidth = image.width * scale;
  const drawHeight = image.height * scale;
  ctx.drawImage(image, x + (width - drawWidth) / 2, y + (height - drawHeight) / 2, drawWidth, drawHeight);
}

/**
 * A `CanvasCellRenderer` whose cell value is an image URL (http(s) or
 * `data:`) instead of text — analogous to Univer's "cell image"
 * (`FRange.insertCellImageAsync`): the image *is* the cell's content and
 * scales with the cell, unlike a floating drawing anchored over the grid.
 * Assign it to a column's `renderer` to opt that column into image cells;
 * it is never applied automatically (see doc/progress.md's Step 26 entry
 * for why cyexcel doesn't guess "this URL looks like a picture" the way it
 * does guess at hyperlinks — an image cell is always an explicit choice).
 *
 * `paint()` must be synchronous (`ContentLayerRenderer` rejects a thenable
 * return), but decoding an `Image()` is not — so a not-yet-cached URL
 * paints a placeholder immediately and redraws directly onto the same
 * `ctx`/`rect` once decoding finishes. That direct redraw is a deliberate
 * shortcut rather than a scheduler-driven repaint: there is no public "mark
 * this cell dirty" API, so if the cell has scrolled away or been repainted
 * over by the time the image settles, the draw is silently lost until the
 * next real repaint (scroll, edit, resize, ...). Acceptable for a first
 * version; revisit if that proves visible in practice.
 */
export function createImageCellRenderer(options: ImageCellRendererOptions = {}): CanvasCellRenderer<string> {
  const padding = options.padding ?? 2;
  const cache = options.cache ?? sharedImageCache;
  const resolveUrl = options.resolveUrl;

  return {
    kind: "canvas",
    paint(context: CellRenderContext<string>): void {
      const { ctx, rect, value } = context;
      const x = rect.x + padding;
      const y = rect.y + padding;
      const width = rect.width - padding * 2;
      const height = rect.height - padding * 2;
      if (width <= 0 || height <= 0 || !value) {
        return;
      }
      const url = resolveUrl ? resolveUrl(value) : value;
      if (!url) {
        return;
      }

      const entry = cache.request(url, (settled) => {
        if (settled.status === "loaded") {
          drawContained(ctx, settled.image, x, y, width, height);
        } else if (settled.status === "error") {
          drawPlaceholder(ctx, x, y, width, height);
        }
      });

      if (entry.status === "loaded") {
        drawContained(ctx, entry.image, x, y, width, height);
      } else {
        drawPlaceholder(ctx, x, y, width, height);
      }
    }
  };
}
