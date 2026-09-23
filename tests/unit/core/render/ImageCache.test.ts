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
