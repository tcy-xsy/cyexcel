import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { ImageCache } from "../../../../src/core/render/ImageCache";

class FakeImage {
  width = 10;
  height = 10;
  onload: (() => void) | null = null;
  onerror: (() => void) | null = null;
  #src = "";

  set src(value: string) {
    this.#src = value;
    queueMicrotask(() => {
      if (this.#src.includes("broken")) {
        this.onerror?.();
      } else {
        this.onload?.();
      }
    });
  }

  get src(): string {
    return this.#src;
  }
}

beforeEach(() => {
  vi.stubGlobal("Image", FakeImage);
});

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("ImageCache", () => {
  it("evicts the least recently used URL, including in-flight requests", async () => {
    const cache = new ImageCache(2);
    const evicted = vi.fn();
    cache.request("a", () => {});
    cache.request("b", evicted);
    cache.get("a");
    cache.request("c", () => {});
    await Promise.resolve();

    expect(cache.get("a")?.status).toBe("loaded");
    expect(cache.get("b")).toBeUndefined();
    expect(cache.get("c")?.status).toBe("loaded");
    expect(evicted).not.toHaveBeenCalled();
  });

  it("deduplicates repeated callbacks while retaining independent subscribers", async () => {
    const cache = new ImageCache();
    const callback = vi.fn();
    const other = vi.fn();
    for (let index = 0; index < 100; index += 1) cache.request("a", callback);
    cache.request("a", other);
    await Promise.resolve();

    expect(callback).toHaveBeenCalledTimes(1);
    expect(other).toHaveBeenCalledTimes(1);
  });

  it("replaces callbacks for a canvas slot and cancels callbacks for its previous URL", async () => {
    const cache = new ImageCache();
    const slot = { owner: {}, key: "0,0" };
    const oldUrl = vi.fn();
    const oldPaint = vi.fn();
    const latestPaint = vi.fn();
    cache.request("a", oldUrl, slot);
    cache.request("b", oldPaint, slot);
    cache.request("b", latestPaint, slot);
    await Promise.resolve();

    expect(oldUrl).not.toHaveBeenCalled();
    expect(oldPaint).not.toHaveBeenCalled();
    expect(latestPaint).toHaveBeenCalledTimes(1);
  });

  it("releases a disposed canvas while retaining another canvas subscription", async () => {
    const cache = new ImageCache();
    const owner = {};
    const disposed = vi.fn();
    const alive = vi.fn();
    cache.request("shared", disposed, { owner, key: "0" });
    cache.request("shared", alive, { owner: {}, key: "0" });
    cache.request("exclusive", disposed, { owner, key: "1" });
    cache.release(owner);
    await Promise.resolve();

    expect(disposed).not.toHaveBeenCalled();
    expect(alive).toHaveBeenCalledTimes(1);
    expect(cache.get("exclusive")).toBeUndefined();
  });

  it("clear cancels in-flight callbacks and removes loaded images", async () => {
    const cache = new ImageCache();
    cache.request("loaded", () => {});
    await Promise.resolve();
    const callback = vi.fn();
    cache.request("pending", callback);
    cache.clear();
    await Promise.resolve();

    expect(cache.get("loaded")).toBeUndefined();
    expect(cache.get("pending")).toBeUndefined();
    expect(callback).not.toHaveBeenCalled();
  });

  it("starts a URL as 'loading', then settles to 'loaded'", async () => {
    const cache = new ImageCache();
    const entry = cache.request("https://example.com/a.png", () => {});
    expect(entry.status).toBe("loading");

    await new Promise((resolve) => queueMicrotask(() => resolve(undefined)));

    expect(cache.get("https://example.com/a.png")?.status).toBe("loaded");
  });

  it("settles to 'error' when the image fails to load", async () => {
    const cache = new ImageCache();
    cache.request("https://example.com/broken.png", () => {});

    await new Promise((resolve) => queueMicrotask(() => resolve(undefined)));

    expect(cache.get("https://example.com/broken.png")?.status).toBe("error");
  });

  it("calls onSettled once the image resolves", async () => {
    const cache = new ImageCache();
    const settled: string[] = [];
    cache.request("https://example.com/a.png", (entry) => settled.push(entry.status));

    await new Promise((resolve) => queueMicrotask(() => resolve(undefined)));

    expect(settled).toEqual(["loaded"]);
  });

  it("loads a given URL only once — a second request before settling reuses the same load and still notifies both callers", async () => {
    const cache = new ImageCache();
    const first: string[] = [];
    const second: string[] = [];
    cache.request("https://example.com/a.png", (entry) => first.push(entry.status));
    const secondEntry = cache.request("https://example.com/a.png", (entry) => second.push(entry.status));

    expect(secondEntry.status).toBe("loading");
    await new Promise((resolve) => queueMicrotask(() => resolve(undefined)));

    expect(first).toEqual(["loaded"]);
    expect(second).toEqual(["loaded"]);
  });

  it("returns the cached entry synchronously once loaded, without starting a new load", async () => {
    const cache = new ImageCache();
    cache.request("https://example.com/a.png", () => {});
    await new Promise((resolve) => queueMicrotask(() => resolve(undefined)));

    const notified: string[] = [];
    const entry = cache.request("https://example.com/a.png", (settled) => notified.push(settled.status));

    expect(entry.status).toBe("loaded");
    // Already-settled entries resolve synchronously; onSettled is only for a load in flight.
    expect(notified).toEqual([]);
  });
});
