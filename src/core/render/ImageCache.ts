export type ImageCacheEntry =
  | { readonly status: "loading" }
  | { readonly status: "loaded"; readonly image: HTMLImageElement }
  | { readonly status: "error" };

interface Subscription {
  readonly owner: object | undefined;
  readonly callback: (entry: ImageCacheEntry) => void;
}

interface PendingImage {
  readonly image: HTMLImageElement;
  readonly listeners: Map<unknown, Subscription>;
}

/** Shared, bounded LRU for decoded images and in-flight loads. */
export class ImageCache {
  readonly #entries = new Map<string, ImageCacheEntry>();
  readonly #pending = new Map<string, PendingImage>();
  readonly #ownerIds = new WeakMap<object, number>();
  readonly #subscriptionUrls = new Map<string, string>();
  readonly #maxEntries: number;
  #nextOwnerId = 0;

  constructor(maxEntries = 128) {
    if (!Number.isSafeInteger(maxEntries) || maxEntries < 1) {
      throw new RangeError("Image cache capacity must be a positive integer");
    }
    this.#maxEntries = maxEntries;
  }

  get(url: string): ImageCacheEntry | undefined {
    const entry = this.#entries.get(url);
    if (entry) {
      this.#entries.delete(url);
      this.#entries.set(url, entry);
    }
    return entry;
  }

  /**
   * Reuses loads. An owner/key replaces its previous subscription on every
   * repaint; callers without an owner are deduplicated by callback identity.
   */
  request(
    url: string,
    onSettled: (entry: ImageCacheEntry) => void,
    subscription: { readonly owner: object; readonly key: string } | undefined = undefined
  ): ImageCacheEntry {
    let key: string | typeof onSettled = onSettled;
    if (subscription) {
      let ownerId = this.#ownerIds.get(subscription.owner);
      if (ownerId === undefined) {
        ownerId = this.#nextOwnerId++;
        this.#ownerIds.set(subscription.owner, ownerId);
      }
      key = `${ownerId}:${subscription.key}`;
      // A canvas position may have scrolled to a different image URL.
      const previousUrl = this.#subscriptionUrls.get(key);
      if (previousUrl !== undefined && previousUrl !== url) {
        this.#pending.get(previousUrl)?.listeners.delete(key);
      }
      this.#subscriptionUrls.delete(key);
    }
    const listener = { owner: subscription?.owner, callback: onSettled };
    const existing = this.get(url);
    if (existing) {
      const pending = this.#pending.get(url);
      pending?.listeners.set(key, listener);
      if (pending && typeof key === "string") this.#subscriptionUrls.set(key, url);
      return existing;
    }

    while (this.#entries.size >= this.#maxEntries) {
      const oldest = this.#entries.keys().next().value;
      if (oldest !== undefined) this.#remove(oldest);
    }
    const loading: ImageCacheEntry = { status: "loading" };
    const image = new Image();
    const pending: PendingImage = { image, listeners: new Map([[key, listener]]) };
    this.#entries.set(url, loading);
    this.#pending.set(url, pending);
    if (typeof key === "string") this.#subscriptionUrls.set(key, url);

    const settle = (entry: ImageCacheEntry): void => {
      if (this.#pending.get(url) !== pending) return;
      image.onload = null;
      image.onerror = null;
      this.#entries.set(url, entry);
      this.#pending.delete(url);
      for (const [listenerKey, { callback }] of pending.listeners) {
        if (typeof listenerKey === "string") this.#subscriptionUrls.delete(listenerKey);
        callback(entry);
      }
      pending.listeners.clear();
    };
    image.onload = () => settle({ status: "loaded", image });
    image.onerror = () => settle({ status: "error" });
    image.src = url;
    return loading;
  }

  /** Cancels subscriptions held by a disposed canvas without affecting other canvases. */
  release(owner: object): void {
    for (const [url, pending] of this.#pending) {
      for (const [key, listener] of pending.listeners) {
        if (listener.owner === owner) {
          pending.listeners.delete(key);
          if (typeof key === "string") this.#subscriptionUrls.delete(key);
        }
      }
      if (pending.listeners.size === 0) this.#remove(url);
    }
  }

  clear(): void {
    for (const url of this.#entries.keys()) this.#remove(url);
  }

  #remove(url: string): void {
    const pending = this.#pending.get(url);
    this.#pending.delete(url);
    this.#entries.delete(url);
    if (pending) {
      for (const key of pending.listeners.keys()) {
        if (typeof key === "string") this.#subscriptionUrls.delete(key);
      }
      pending.listeners.clear();
      pending.image.onload = null;
      pending.image.onerror = null;
      pending.image.src = "";
    }
  }
}

export const sharedImageCache = new ImageCache();
