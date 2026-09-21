import type { Canvas2DLike } from "../../src/core/render";

export interface RecordedCall {
  readonly method: string;
  readonly args: readonly unknown[];
}

/**
 * Records every Canvas2DLike call instead of drawing pixels, so paint logic
 * can be asserted on in Node — call counts, coordinates, and style changes —
 * without comparing system-font-rendered pixels.
 */
export class RecordingCanvasContext implements Canvas2DLike {
  readonly calls: RecordedCall[] = [];

  #fillStyle = "";
  #strokeStyle = "";
  #font = "";
  #textBaseline = "";
  #textAlign = "";
  #lineWidth = 1;
  #globalAlpha = 1;

  fillRect(x: number, y: number, width: number, height: number): void {
    this.calls.push({ method: "fillRect", args: [x, y, width, height] });
  }

  clearRect(x: number, y: number, width: number, height: number): void {
    this.calls.push({ method: "clearRect", args: [x, y, width, height] });
  }

  strokeRect(x: number, y: number, width: number, height: number): void {
    this.calls.push({ method: "strokeRect", args: [x, y, width, height] });
  }

  fillText(text: string, x: number, y: number): void {
    this.calls.push({ method: "fillText", args: [text, x, y] });
  }

  /**
   * Not real font metrics — a deterministic per-character width (no
   * font/DOM available in Node) so wrap-line-count assertions are stable
   * and reproducible. Real pixel-accurate wrapping is a browser-test
   * concern (see tests/browser/variable-row-height.spec.ts), not this
   * recorder's job.
   */
  measureText(text: string): { readonly width: number } {
    this.calls.push({ method: "measureText", args: [text] });
    return { width: text.length * 7 };
  }

  save(): void {
    this.calls.push({ method: "save", args: [] });
  }

  restore(): void {
    this.calls.push({ method: "restore", args: [] });
  }

  beginPath(): void {
    this.calls.push({ method: "beginPath", args: [] });
  }

  rect(x: number, y: number, width: number, height: number): void {
    this.calls.push({ method: "rect", args: [x, y, width, height] });
  }

  clip(): void {
    this.calls.push({ method: "clip", args: [] });
  }

  moveTo(x: number, y: number): void {
    this.calls.push({ method: "moveTo", args: [x, y] });
  }

  lineTo(x: number, y: number): void {
    this.calls.push({ method: "lineTo", args: [x, y] });
  }

  closePath(): void {
    this.calls.push({ method: "closePath", args: [] });
  }

  fill(): void {
    this.calls.push({ method: "fill", args: [] });
  }

  stroke(): void {
    this.calls.push({ method: "stroke", args: [] });
  }

  scale(x: number, y: number): void {
    this.calls.push({ method: "scale", args: [x, y] });
  }

  setTransform(a: number, b: number, c: number, d: number, e: number, f: number): void {
    this.calls.push({ method: "setTransform", args: [a, b, c, d, e, f] });
  }

  setLineDash(segments: readonly number[]): void {
    this.calls.push({ method: "setLineDash", args: [segments] });
  }

  get fillStyle(): string {
    return this.#fillStyle;
  }

  set fillStyle(value: string) {
    this.#fillStyle = value;
    this.calls.push({ method: "set:fillStyle", args: [value] });
  }

  get strokeStyle(): string {
    return this.#strokeStyle;
  }

  set strokeStyle(value: string) {
    this.#strokeStyle = value;
    this.calls.push({ method: "set:strokeStyle", args: [value] });
  }

  get font(): string {
    return this.#font;
  }

  set font(value: string) {
    this.#font = value;
    this.calls.push({ method: "set:font", args: [value] });
  }

  get textBaseline(): string {
    return this.#textBaseline;
  }

  set textBaseline(value: string) {
    this.#textBaseline = value;
    this.calls.push({ method: "set:textBaseline", args: [value] });
  }

  get textAlign(): string {
    return this.#textAlign;
  }

  set textAlign(value: string) {
    this.#textAlign = value;
    this.calls.push({ method: "set:textAlign", args: [value] });
  }

  get lineWidth(): number {
    return this.#lineWidth;
  }

  set lineWidth(value: number) {
    this.#lineWidth = value;
    this.calls.push({ method: "set:lineWidth", args: [value] });
  }

  get globalAlpha(): number {
    return this.#globalAlpha;
  }

  set globalAlpha(value: number) {
    this.#globalAlpha = value;
    this.calls.push({ method: "set:globalAlpha", args: [value] });
  }
}
