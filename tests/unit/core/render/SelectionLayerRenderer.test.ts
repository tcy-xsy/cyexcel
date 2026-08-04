import { describe, expect, it } from "vitest";

import { SelectionLayerRenderer } from "../../../../src/core";
import { RecordingCanvasContext } from "../../../helpers/RecordingCanvasContext";

describe("SelectionLayerRenderer", () => {
  it("clears its own canvas area at the last size set via setSize", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = new SelectionLayerRenderer({ ctx });

    renderer.setSize(300, 150);
    renderer.paintSelection();

    expect(ctx.calls).toEqual([{ method: "clearRect", args: [0, 0, 300, 150] }]);
  });

  it("defaults to a 0x0 clear when setSize was never called", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = new SelectionLayerRenderer({ ctx });

    renderer.paintSelection();

    expect(ctx.calls).toEqual([{ method: "clearRect", args: [0, 0, 0, 0] }]);
  });
});
