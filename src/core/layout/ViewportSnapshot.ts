import type { ViewportState } from "./LayoutEngine";

export interface ViewportSnapshotOptions extends ViewportState {
  readonly version: number;
  readonly overscanRows?: number;
  readonly overscanColumns?: number;
}

function assertNonNegativeInteger(name: string, value: number): void {
  if (!Number.isSafeInteger(value) || value < 0) {
    throw new RangeError(`${name} must be a non-negative safe integer: ${value}`);
  }
}

function assertNonNegativeFinite(name: string, value: number): void {
  if (!Number.isFinite(value) || value < 0) {
    throw new RangeError(`${name} must be a non-negative finite number: ${value}`);
  }
}

/** Immutable body viewport state in CSS pixels and logical scroll space. */
export class ViewportSnapshot implements ViewportState {
  readonly version: number;
  readonly scrollTop: number;
  readonly scrollLeft: number;
  readonly width: number;
  readonly height: number;
  readonly overscanRows: number;
  readonly overscanColumns: number;

  constructor(options: ViewportSnapshotOptions) {
    assertNonNegativeInteger("version", options.version);
    assertNonNegativeFinite("scrollTop", options.scrollTop);
    assertNonNegativeFinite("scrollLeft", options.scrollLeft);
    assertNonNegativeFinite("width", options.width);
    assertNonNegativeFinite("height", options.height);
    assertNonNegativeInteger("overscanRows", options.overscanRows ?? 0);
    assertNonNegativeInteger("overscanColumns", options.overscanColumns ?? 0);

    this.version = options.version;
    this.scrollTop = options.scrollTop;
    this.scrollLeft = options.scrollLeft;
    this.width = options.width;
    this.height = options.height;
    this.overscanRows = options.overscanRows ?? 0;
    this.overscanColumns = options.overscanColumns ?? 0;
    Object.freeze(this);
  }
}
