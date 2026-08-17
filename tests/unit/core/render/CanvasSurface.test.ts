import { describe, expect, it } from "vitest";

import { CanvasSurface, getEffectiveDpr } from "../../../../src/core/render";
import type { CanvasTransformContext, ResizableCanvas } from "../../../../src/core/render";

function createCanvas(): ResizableCanvas {
  return { width: 0, height: 0, style: { width: "", height: "" } };
}

describe("CanvasSurface", () => {
  it("uses the same capped DPR for backing pixels and the drawing transform", () => {
    const canvas = createCanvas();
    const transforms: number[][] = [];
    const context: CanvasTransformContext = {
      setTransform: (...values) => transforms.push(values)
    };
    const surface = new CanvasSurface(canvas, context, 3);

    const snapshot = surface.resize(100, 50, 4);

    expect(snapshot).toEqual({
      cssWidth: 100,
      cssHeight: 50,
      effectiveDpr: 3,
      physicalWidth: 300,
      physicalHeight: 150
    });
    expect(transforms).toEqual([[3, 0, 0, 3, 0, 0]]);
    expect(canvas.style).toEqual({ width: "100px", height: "50px" });
  });

  it("normalizes invalid browser DPR values to one", () => {
    expect(getEffectiveDpr(0)).toBe(1);
    expect(getEffectiveDpr(Number.NaN)).toBe(1);
  });
});
