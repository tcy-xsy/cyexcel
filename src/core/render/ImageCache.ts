export type ImageCacheEntry =
  | { readonly status: "loading" }
  | { readonly status: "loaded"; readonly image: HTMLImageElement }
  | { readonly status: "error" };

/**
 * Loads and caches `HTMLImageElement`s by URL for canvas image-cell
 * renderers (see `cells/imageCellRenderer.ts`). One instance is shared by
 * every image cell by default (`sharedImageCache`) so the same URL used
 * across many rows/columns decodes once, not once per cell.
 */
export class ImageCache {
  readonly #entries = new Map<string, ImageCacheEntry>();
  readonly #pending = new Map<string, Array<(entry: ImageCacheEntry) => void>>();

  get(url: string): ImageCacheEntry | undefined {
    return this.#entries.get(url);
  }

  /**
   * Returns the current entry for `url` synchronously (`"loading"` the
   * first time), and calls `onSettled` once the image has decoded or
   * failed. Every caller in flight for the same URL — not just the first —
   * gets its own `onSettled` call, since scrolling can bring several cells
   * sharing one URL into view before it settles.
   */
  request(url: string, onSettled: (entry: ImageCacheEntry) => void): ImageCacheEntry {
    const existing = this.#entries.get(url);
    if (existing) {
      if (existing.status === "loading") {
        this.#pending.get(url)?.push(onSettled);
      }
      return existing;
    }

    const loading: ImageCacheEntry = { status: "loading" };
    this.#entries.set(url, loading);
    this.#pending.set(url, [onSettled]);

    const settle = (entry: ImageCacheEntry): void => {
      this.#entries.set(url, entry);
      const listeners = this.#pending.get(url) ?? [];
      this.#pending.delete(url);
      for (const listener of listeners) {
        listener(entry);
      }
    };

    const image = new Image();
    image.onload = () => settle({ status: "loaded", image });
    image.onerror = () => settle({ status: "error" });
    image.src = url;
    return loading;
  }
}

export const sharedImageCache = new ImageCache();
