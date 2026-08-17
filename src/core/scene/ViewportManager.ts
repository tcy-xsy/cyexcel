import { ScrollAxisMapper, type ScrollAxisMapperOptions } from "./ScrollAxisMapper";

/** The scroller's own box, as much of an element as this manager needs. */
export interface ScrollerBox {
  scrollTop: number;
  scrollLeft: number;
  readonly clientWidth: number;
  readonly clientHeight: number;
}

/** The sized-to-content element that gives the scroller its range. */
export interface SpacerBox {
  readonly style: { width: string; height: string };
}

/** Pinned bands that share the scroller's box rather than adding to it. */
export interface HeaderBandSizes {
  readonly rowHeaderWidth: number;
  readonly columnHeaderHeight: number;
}

/**
 * The body's logical scroll offsets and its own visible size.
 *
 * Structurally identical to layout's `ViewportState`, but declared here rather
 * than imported: `ViewportState` is the *mutable* viewport shape, and the §12
 * import gate deliberately keeps it out of scene/. Consumers receive a plain
 * readonly fact instead of a handle onto a second coordinate truth.
 */
export interface LogicalViewport {
  readonly scrollTop: number;
  readonly scrollLeft: number;
  readonly width: number;
  readonly height: number;
}

export interface ViewportManagerOptions {
  readonly scroller: ScrollerBox;
  readonly spacer: SpacerBox;
  /** Read per call — the column header band's height changes with group depth. */
  readonly getBandSizes: () => HeaderBandSizes;
  readonly mapperOptions?: ScrollAxisMapperOptions;
}

/**
 * The single owner of the scroller's native scroll offsets and the spacer that
 * gives them their range.
 *
 * Everything above this class works in logical offsets only; the mapping onto
 * what a browser can actually scroll — including segmenting when the content
 * is larger than that — lives here and nowhere else. That is what lets the §12
 * lint gate restrict `scrollTop`/`scrollLeft` to scene/ and layout/ with no
 * exception for CyExcelCore.
 */
export class ViewportManager {
  readonly #scroller: ScrollerBox;
  readonly #spacer: SpacerBox;
  readonly #getBandSizes: () => HeaderBandSizes;
  readonly #vertical: ScrollAxisMapper;
  readonly #horizontal: ScrollAxisMapper;

  constructor(options: ViewportManagerOptions) {
    this.#scroller = options.scroller;
    this.#spacer = options.spacer;
    this.#getBandSizes = options.getBandSizes;
    this.#vertical = new ScrollAxisMapper(options.mapperOptions);
    this.#horizontal = new ScrollAxisMapper(options.mapperOptions);
  }

  /** True once either axis needs more logical range than the browser can scroll. */
  get segmented(): boolean {
    return this.#vertical.segmented || this.#horizontal.segmented;
  }

  /**
   * Records the content's logical size and resizes the spacer to the native
   * extent that allows. The spacer includes each pinned band because the bands
   * share the scroller's box rather than adding to it — without that the last
   * row/column would sit permanently behind unreachable pixels.
   */
  setContentSize(totalWidth: number, totalHeight: number): void {
    const bands = this.#getBandSizes();
    this.#horizontal.setMetrics(totalWidth, this.#bodyWidth(bands));
    this.#vertical.setMetrics(totalHeight, this.#bodyHeight(bands));
    this.#spacer.style.width = `${this.#horizontal.nativeExtent + bands.rowHeaderWidth}px`;
    this.#spacer.style.height = `${this.#vertical.nativeExtent + bands.columnHeaderHeight}px`;
  }

  /** The logical viewport a Scene should be built from. */
  readState(): LogicalViewport {
    const bands = this.#getBandSizes();
    return {
      scrollTop: this.#vertical.toLogical(this.#nativeTop()),
      scrollLeft: this.#horizontal.toLogical(this.#nativeLeft()),
      width: this.#bodyWidth(bands),
      height: this.#bodyHeight(bands)
    };
  }

  /**
   * Reconciles both axes against the offsets the scroller now reports, writing
   * back only when a segment origin actually has to move. Call once per scroll
   * event, before rebuilding the Scene.
   */
  reconcile(): void {
    const vertical = this.#vertical.reconcile(this.#nativeTop());
    const horizontal = this.#horizontal.reconcile(this.#nativeLeft());
    if (vertical) {
      this.#vertical.commit(vertical);
      this.#scroller.scrollTop = vertical.nativeOffset;
    }
    if (horizontal) {
      this.#horizontal.commit(horizontal);
      this.#scroller.scrollLeft = horizontal.nativeOffset;
    }
  }

  /** Moves the scroller to a logical position, picking a segment that can express it. */
  scrollToLogical(logicalTop: number, logicalLeft: number): void {
    const vertical = this.#vertical.toNative(logicalTop);
    const horizontal = this.#horizontal.toNative(logicalLeft);
    this.#vertical.commit(vertical);
    this.#horizontal.commit(horizontal);
    this.#scroller.scrollTop = vertical.nativeOffset;
    this.#scroller.scrollLeft = horizontal.nativeOffset;
  }

  /**
   * One native auto-scroll step, for gestures that drag past an edge. Stays in
   * native pixels on purpose: the caller is nudging the view by a few pixels
   * per frame, not seeking to a position, and routing that through the seek
   * path would rescale it.
   */
  scrollNativeBy(deltaY: number): void {
    this.#scroller.scrollTop = Math.max(0, this.#nativeTop() + deltaY);
  }

  /**
   * Rubber-band overscroll on macOS reports a negative offset. Clamped at this
   * boundary rather than inside the mapper, which stays strict so that a real
   * bug still throws instead of being silently absorbed.
   */
  #nativeTop(): number {
    return Math.max(0, this.#scroller.scrollTop);
  }

  #nativeLeft(): number {
    return Math.max(0, this.#scroller.scrollLeft);
  }

  #bodyWidth(bands: HeaderBandSizes): number {
    return Math.max(0, this.#scroller.clientWidth - bands.rowHeaderWidth);
  }

  #bodyHeight(bands: HeaderBandSizes): number {
    return Math.max(0, this.#scroller.clientHeight - bands.columnHeaderHeight);
  }
}
