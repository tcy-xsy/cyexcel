import { describe, expect, it } from "vitest";

import { computeWrappedCellHeight, wrapText } from "../../../../src/core/render/textWrap";
import { CELL_LINE_HEIGHT, CELL_PADDING } from "../../../../src/core/render/theme";
import { RecordingCanvasContext } from "../../../helpers/RecordingCanvasContext";

describe("wrapText", () => {
  it("prefers whitespace boundaries for English text", () => {
    const ctx = new RecordingCanvasContext();

    expect(wrapText(ctx, "alpha beta gamma", 70)).toEqual(["alpha beta", "gamma"]);
  });

  it("falls back to character boundaries for CJK text without spaces", () => {
    const ctx = new RecordingCanvasContext();

    expect(wrapText(ctx, "工业温度传感器", 21)).toEqual(["工业温", "度传感", "器"]);
  });

  it("always returns one line for empty or whitespace-only text", () => {
    const ctx = new RecordingCanvasContext();

    expect(wrapText(ctx, "", 100)).toEqual([""]);
    expect(wrapText(ctx, "   ", 100)).toEqual([""]);
  });
});

describe("computeWrappedCellHeight", () => {
  it("adds cell padding around the wrapped line block", () => {
    const ctx = new RecordingCanvasContext();
    const lines = wrapText(ctx, "工业温度传感器", 21);

    expect(computeWrappedCellHeight(ctx, "工业温度传感器", 21)).toBe(
      lines.length * CELL_LINE_HEIGHT + CELL_PADDING * 2
    );
  });
});
