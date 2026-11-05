import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { createImageCellRenderer } from "../../../../src/core/cells/imageCellRenderer";
import { ImageCache } from "../../../../src/core/render/ImageCache";
import { RecordingCanvasContext } from "../../../helpers/RecordingCanvasContext";

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

const rect = { x: 0, y: 0, width: 20, height: 20 };

describe("createImageCellRenderer", () => {
  it("draws only once after repeated loading repaints of the same canvas slot", async () => {
    const ctx = new RecordingCanvasContext();
    const renderer = createImageCellRenderer({ cache: new ImageCache() });
    for (let index = 0; index < 100; index += 1) {
      renderer.paint({ ctx, rect, row: 0, column: {} as never, value: "a" });
    }
    await Promise.resolve();

    expect(ctx.calls.filter((call) => call.method === "drawImage")).toHaveLength(1);
  });

  it("paints a placeholder immediately for a URL that hasn't loaded yet", () => {
    const ctx = new RecordingCanvasContext();
    const cache = new ImageCache();
    const renderer = createImageCellRenderer({ cache });

    renderer.paint({ ctx, rect, row: 0, column: {} as never, value: "https://example.com/a.png" });

    expect(ctx.calls.some((call) => call.method === "drawImage")).toBe(false);
    // Placeholder draws a fill + an icon (arc/fill), not the image itself.
    expect(ctx.calls.some((call) => call.method === "fillRect")).toBe(true);
  });

  it("draws the decoded image once it settles", async () => {
    const ctx = new RecordingCanvasContext();
    const cache = new ImageCache();
    const renderer = createImageCellRenderer({ cache });

    renderer.paint({ ctx, rect, row: 0, column: {} as never, value: "https://example.com/a.png" });
    await new Promise((resolve) => queueMicrotask(() => resolve(undefined)));

    expect(ctx.calls.some((call) => call.method === "drawImage")).toBe(true);
  });

  it("draws the image synchronously on a cache hit, without a placeholder", async () => {
    const cache = new ImageCache();
    cache.request("https://example.com/a.png", () => {});
    await new Promise((resolve) => queueMicrotask(() => resolve(undefined)));

    const ctx = new RecordingCanvasContext();
    const renderer = createImageCellRenderer({ cache });
    renderer.paint({ ctx, rect, row: 0, column: {} as never, value: "https://example.com/a.png" });

    expect(ctx.calls.some((call) => call.method === "drawImage")).toBe(true);
  });

  it("paints a placeholder when the image fails to load", async () => {
    const ctx = new RecordingCanvasContext();
    const cache = new ImageCache();
    const renderer = createImageCellRenderer({ cache });

    renderer.paint({ ctx, rect, row: 0, column: {} as never, value: "https://example.com/broken.png" });
    await new Promise((resolve) => queueMicrotask(() => resolve(undefined)));

    expect(ctx.calls.some((call) => call.method === "drawImage")).toBe(false);
    expect(ctx.calls.filter((call) => call.method === "fillRect").length).toBeGreaterThan(0);
  });

  it("does nothing for an empty cell", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = createImageCellRenderer();

    renderer.paint({ ctx, rect, row: 0, column: {} as never, value: undefined });

    expect(ctx.calls).toHaveLength(0);
  });

  it("resolveUrl lets a multi-value cell pick one URL to draw", async () => {
    const ctx = new RecordingCanvasContext();
    const cache = new ImageCache();
    const renderer = createImageCellRenderer({
      cache,
      resolveUrl: (value) => value.split("\n")[0]
    });

    renderer.paint({
      ctx,
      rect,
      row: 0,
      column: {} as never,
      value: "https://example.com/a.png\nhttps://example.com/b.png"
    });
    await new Promise((resolve) => queueMicrotask(() => resolve(undefined)));

    expect(ctx.calls.some((call) => call.method === "drawImage")).toBe(true);
    expect(cache.get("https://example.com/a.png")?.status).toBe("loaded");
    expect(cache.get("https://example.com/b.png")).toBeUndefined();
  });

  it("resolveUrl returning undefined paints nothing", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = createImageCellRenderer({ resolveUrl: () => undefined });

    renderer.paint({ ctx, rect, row: 0, column: {} as never, value: "not-a-url" });

    expect(ctx.calls).toHaveLength(0);
  });
});
