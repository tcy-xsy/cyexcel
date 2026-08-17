import type { Canvas2DLike } from "./Canvas2DLike";
import { CELL_LINE_HEIGHT, CELL_PADDING } from "./theme";

/**
 * Greedy width-first line wrap: breaks preferentially at whitespace, and
 * falls back to breaking mid-word (character by character) when a single
 * word is wider than `maxWidth` on its own — the common case for CJK text,
 * which has no whitespace between "words" at all (cyexcel's own playground
 * data is full of exactly this: "工业温度传感器" has no break opportunity a
 * pure space-based wrap could ever find). Not a full Unicode line-breaking
 * implementation (no locale-aware rules, no hyphenation) — good enough for
 * a spreadsheet cell, not a typesetting engine. `ctx.measureText()` is the
 * only thing this depends on, so it works identically against a real
 * CanvasRenderingContext2D or a test double.
 *
 * Never exceeds `maxWidth` per line, and never infinite-loops even when
 * `maxWidth` is smaller than a single character — the `current === ""`
 * escape hatch in both loops below guarantees at least one word/char is
 * placed per line regardless, so the algorithm always makes forward
 * progress through the input.
 */
export function wrapText(ctx: Canvas2DLike, text: string, maxWidth: number): string[] {
  if (text.length === 0) {
    return [""];
  }

  const safeMaxWidth = Math.max(1, maxWidth);
  const widthOf = (value: string) => ctx.measureText(value).width;

  // Whitespace runs collapse to nothing between wrapped words — matches how
  // HTML/CSS text wrapping already treats whitespace at a line break.
  const words = text.split(/\s+/).filter((word) => word.length > 0);
  if (words.length === 0) {
    return [""]; // text was entirely whitespace
  }

  const lines: string[] = [];
  let currentLine = "";

  const flushLine = () => {
    if (currentLine !== "") {
      lines.push(currentLine);
      currentLine = "";
    }
  };

  for (const word of words) {
    if (widthOf(word) <= safeMaxWidth) {
      const candidate = currentLine === "" ? word : `${currentLine} ${word}`;
      if (currentLine === "" || widthOf(candidate) <= safeMaxWidth) {
        currentLine = candidate;
      } else {
        flushLine();
        currentLine = word;
      }
      continue;
    }

    // The word itself is wider than a whole line — break it character by
    // character. Flushes whatever's on the current line first rather than
    // trying to top it off with a few characters of this word — simpler,
    // and still never exceeds maxWidth per line, just occasionally packs a
    // line a little less tightly than the theoretical optimum.
    flushLine();
    for (const character of word) {
      const candidate = currentLine + character;
      if (currentLine === "" || widthOf(candidate) <= safeMaxWidth) {
        currentLine = candidate;
      } else {
        flushLine();
        currentLine = character;
      }
    }
  }
  flushLine();

  return lines.length > 0 ? lines : [""];
}

/** Vertical space `text` needs when wrapped to `maxWidth` and painted with `CELL_LINE_HEIGHT`-tall lines, plus the same top/bottom padding a single-line cell already reserves via `CELL_PADDING`. */
export function computeWrappedCellHeight(ctx: Canvas2DLike, text: string, maxWidth: number): number {
  return wrapText(ctx, text, maxWidth).length * CELL_LINE_HEIGHT + CELL_PADDING * 2;
}
