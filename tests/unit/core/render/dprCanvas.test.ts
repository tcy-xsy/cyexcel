import { describe, expect, it } from "vitest";

import { configureCanvasForDpr } from "../../../../src/core/render";
import type { ResizableCanvas } from "../../../../src/core/render";

function createFakeCanvas(): ResizableCanvas {
  return { width: 0, height: 0, style: { width: "", height: "" } };
}

describe("configureCanvasForDpr", () => {
  it("sets physical width/height scaled by dpr and logical CSS size", () => {
    const canvas = createFakeCanvas();

    configureCanvasForDpr(canvas, 400, 300, 2);

    expect(canvas.width).toBe(800);
    expect(canvas.height).toBe(600);
    expect(canvas.style.width).toBe("400px");
    expect(canvas.style.height).toBe("300px");
  });

  it("caps the physical size at maxDpr", () => {
    const canvas = createFakeCanvas();

    configureCanvasForDpr(canvas, 400, 300, 10, 3);

    expect(canvas.width).toBe(1200);
    expect(canvas.height).toBe(900);
  });

  it("uses dpr 1 as a passthrough", () => {
    const canvas = createFakeCanvas();

    configureCanvasForDpr(canvas, 400, 300, 1);

    expect(canvas.width).toBe(400);
    expect(canvas.height).toBe(300);
  });
});
