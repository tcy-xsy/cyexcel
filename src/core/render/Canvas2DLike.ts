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
  save(): void;
  restore(): void;
  beginPath(): void;
  rect(x: number, y: number, width: number, height: number): void;
  clip(): void;
  scale(x: number, y: number): void;
  fillStyle: string;
  strokeStyle: string;
  font: string;
  textBaseline: string;
  textAlign: string;
}
