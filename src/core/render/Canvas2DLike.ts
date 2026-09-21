/**
 * The subset of CanvasRenderingContext2D that this module's renderers
 * actually use. A real CanvasRenderingContext2D satisfies this structurally
 * with no adapter needed; tests use RecordingCanvasContext
 * (tests/helpers/RecordingCanvasContext.ts) instead of a real canvas.
 */
export interface Canvas2DLike {
  fillRect(x: number, y: number, width: number, height: number): void;
  clearRect(x: number, y: number, width: number, height: number): void;
  strokeRect(x: number, y: number, width: number, height: number): void;
  fillText(text: string, x: number, y: number): void;
  measureText(text: string): { readonly width: number };
  save(): void;
  restore(): void;
  beginPath(): void;
  rect(x: number, y: number, width: number, height: number): void;
  clip(): void;
  /**
   * Arbitrary polygons, for marks whose shape matters at small sizes — the
   * header's sort carets are 7x4px, a size at which a text glyph's shape and
   * weight are at the mercy of whichever font resolves.
   */
  moveTo(x: number, y: number): void;
  lineTo(x: number, y: number): void;
  closePath(): void;
  fill(): void;
  /** 描边当前路径。单元格边框用它画直线（`strokeRect` 画不了单独一条边）。 */
  stroke(): void;
  scale(x: number, y: number): void;
  setTransform(a: number, b: number, c: number, d: number, e: number, f: number): void;
  setLineDash(segments: readonly number[]): void;
  fillStyle: string;
  strokeStyle: string;
  font: string;
  textBaseline: string;
  textAlign: string;
  lineWidth: number;
  globalAlpha: number;
}
