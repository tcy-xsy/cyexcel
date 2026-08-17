/**
 * Browsers cap how large an element's scrollable extent may be. The exact
 * ceiling differs per engine (Chromium's is the largest, Firefox's the
 * smallest of the mainstream three), so the default here sits safely below
 * the smallest known limit rather than probing for it: a wrong probe fails
 * silently as content that cannot be scrolled to, which is exactly the bug
 * this mapper exists to prevent.
 *
 * A logical extent at or below this value is mapped 1:1 — `segmentOrigin`
 * stays 0 and every offset passes through unchanged, so ordinary tables
 * behave exactly as they did before segmenting existed.
 */
export const DEFAULT_SAFE_NATIVE_EXTENT = 15_000_000;

export interface ScrollAxisMapperOptions {
  /** Overridable so tests can force segmenting without building a 15M px table. */
  readonly safeNativeExtent?: number;
}

/** A native scroll position paired with the segment it is measured against. */
export interface ScrollAxisPosition {
  readonly segmentOrigin: number;
  readonly nativeOffset: number;
}

function assertFiniteNonNegative(name: string, value: number): void {
  if (!Number.isFinite(value) || value < 0) {
    throw new RangeError(`${name} must be a non-negative finite number: ${value}`);
  }
}

/**
 * Maps one axis between the grid's *logical* offset (which may exceed what a
 * browser can natively scroll) and the *native* offset of the real scroller.
 *
 * The whole contract is one equation:
 *
 *     logicalOffset = segmentOrigin + nativeOffset
 *
 * `segmentOrigin` is the logical position that native offset 0 currently
 * represents. While the user scrolls inside the safe native window the origin
 * stays put; when they approach either end of it the origin is shifted and the
 * native offset is re-centred by the same amount, so the logical position is
 * unchanged and nothing moves on screen.
 *
 * Rounding cannot accumulate: the logical offset is always re-derived from the
 * native offset the scroller actually reports, and `segmentOrigin` is only ever
 * assigned from a freshly clamped computation — never by adding a delta to its
 * previous value.
 *
 * Pure and DOM-free: Scene and Skeleton keep seeing only logical scroll.
 */
export class ScrollAxisMapper {
  readonly #safeNativeExtent: number;
  #logicalExtent = 0;
  #viewportSize = 0;
  #segmentOrigin = 0;
  #version = 0;
  /** Last offset seen by `reconcile()`, used to tell a seek from a continuous scroll. */
  #lastNativeOffset: number | null = null;

  constructor(options: ScrollAxisMapperOptions = {}) {
    const safeNativeExtent = options.safeNativeExtent ?? DEFAULT_SAFE_NATIVE_EXTENT;
    if (!Number.isFinite(safeNativeExtent) || safeNativeExtent <= 0) {
      throw new RangeError(`safeNativeExtent must be a positive finite number: ${safeNativeExtent}`);
    }
    this.#safeNativeExtent = safeNativeExtent;
  }

  /** Bumped whenever the origin moves or the metrics change, so consumers can detect a remap. */
  get version(): number {
    return this.#version;
  }

  get segmentOrigin(): number {
    return this.#segmentOrigin;
  }

  get logicalExtent(): number {
    return this.#logicalExtent;
  }

  /** The size the spacer should be given — never larger than the safe native extent. */
  get nativeExtent(): number {
    return Math.min(this.#logicalExtent, this.#safeNativeExtent);
  }

  /** True once the logical extent no longer fits in one native window. */
  get segmented(): boolean {
    return this.maxSegmentOrigin > 0;
  }

  get maxLogicalScroll(): number {
    return Math.max(0, this.#logicalExtent - this.#viewportSize);
  }

  get maxNativeScroll(): number {
    return Math.max(0, this.nativeExtent - this.#viewportSize);
  }

  get maxSegmentOrigin(): number {
    return Math.max(0, this.maxLogicalScroll - this.maxNativeScroll);
  }

  /**
   * Records new content/viewport metrics. Returns true when the caller must
   * re-apply geometry (spacer size, and possibly a corrected native offset),
   * i.e. when anything observable changed.
   *
   * The origin is re-clamped here: shrinking the content — deleting rows, or a
   * wrap column collapsing back to default heights — can leave the previous
   * origin past the end of the new extent.
   */
  setMetrics(logicalExtent: number, viewportSize: number): boolean {
    assertFiniteNonNegative("logicalExtent", logicalExtent);
    assertFiniteNonNegative("viewportSize", viewportSize);
    if (logicalExtent === this.#logicalExtent && viewportSize === this.#viewportSize) {
      return false;
    }
    this.#logicalExtent = logicalExtent;
    this.#viewportSize = viewportSize;
    const clamped = Math.min(this.#segmentOrigin, this.maxSegmentOrigin);
    this.#segmentOrigin = clamped;
    this.#version += 1;
    return true;
  }

  /** The logical offset the given native offset currently represents. */
  toLogical(nativeOffset: number): number {
    assertFiniteNonNegative("nativeOffset", nativeOffset);
    return Math.min(this.#segmentOrigin + nativeOffset, this.maxLogicalScroll);
  }

  /**
   * Where the scroller must be placed to show `logicalOffset`, choosing an
   * origin that leaves as much native room as possible on both sides so the
   * next few user scrolls do not immediately need another recentre.
   *
   * Does not mutate; call `commit()` with the result to adopt it.
   */
  toNative(logicalOffset: number): ScrollAxisPosition {
    assertFiniteNonNegative("logicalOffset", logicalOffset);
    const logical = Math.min(logicalOffset, this.maxLogicalScroll);
    if (!this.segmented) {
      return { segmentOrigin: 0, nativeOffset: logical };
    }
    const centred = logical - this.maxNativeScroll / 2;
    const segmentOrigin = Math.min(Math.max(centred, 0), this.maxSegmentOrigin);
    return { segmentOrigin, nativeOffset: logical - segmentOrigin };
  }

  /**
   * Called after the user scrolls. Returns the position to write back when the
   * native offset has drifted into either end of the safe window and more
   * logical content exists in that direction, or `null` when the current
   * position is fine and the scroller must be left alone.
   *
   * Returning `null` for "no change" matters: writing the scroller on every
   * scroll event would fight the browser's own smooth-scroll and momentum.
   */
  recenter(nativeOffset: number): ScrollAxisPosition | null {
    assertFiniteNonNegative("nativeOffset", nativeOffset);
    if (!this.segmented) {
      return null;
    }
    const guard = this.#recenterGuard();
    const atLowEdge = nativeOffset < guard && this.#segmentOrigin > 0;
    const atHighEdge =
      nativeOffset > this.maxNativeScroll - guard && this.#segmentOrigin < this.maxSegmentOrigin;
    if (!atLowEdge && !atHighEdge) {
      return null;
    }
    const next = this.toNative(this.toLogical(nativeOffset));
    return next.segmentOrigin === this.#segmentOrigin ? null : next;
  }

  /**
   * Maps a native offset proportionally across the whole logical axis, so the
   * scrollbar's own extremes mean the document's extremes.
   *
   * Used only for discontinuous jumps. A windowed 1:1 mapping cannot serve
   * them: dragging the thumb to the bottom moves the native offset to its
   * maximum while the origin is still 0, which is only `maxNativeScroll` into
   * a longer axis — on this fixture's scale that is 93.75% of the way down,
   * and it gets worse the larger the table.
   */
  proportionalLogical(nativeOffset: number): number {
    assertFiniteNonNegative("nativeOffset", nativeOffset);
    if (this.maxNativeScroll <= 0) {
      return 0;
    }
    const ratio = Math.min(nativeOffset / this.maxNativeScroll, 1);
    return ratio * this.maxLogicalScroll;
  }

  /**
   * Reconciles a freshly observed native offset and returns the position to
   * write back, or `null` to leave the scroller alone.
   *
   * Distinguishes the two ways a native offset changes, because one mapping
   * cannot serve both:
   *
   * - a *continuous* change (wheel, keyboard, drag-select autoscroll) must stay
   *   pixel-accurate, so it keeps the 1:1 windowed mapping and only recentres
   *   near the window's ends;
   * - a *jump* (dragging the scrollbar thumb, or assigning scrollTop) is a seek
   *   to a position in the whole document, so it maps proportionally.
   *
   * The threshold is the viewport itself: no continuous scroll gesture advances
   * by more than one screen between events, and any seek worth distinguishing
   * is larger than that.
   */
  reconcile(nativeOffset: number): ScrollAxisPosition | null {
    assertFiniteNonNegative("nativeOffset", nativeOffset);
    const previous = this.#lastNativeOffset;
    this.#lastNativeOffset = nativeOffset;
    if (!this.segmented) {
      return null;
    }
    const jumped = previous === null || Math.abs(nativeOffset - previous) > this.#jumpThreshold();
    if (!jumped) {
      return this.recenter(nativeOffset);
    }
    const next = this.toNative(this.proportionalLogical(nativeOffset));
    return next.segmentOrigin === this.#segmentOrigin && next.nativeOffset === nativeOffset
      ? null
      : next;
  }

  /** Adopts a position produced by `toNative()`/`recenter()`/`reconcile()`. */
  commit(position: ScrollAxisPosition): void {
    assertFiniteNonNegative("segmentOrigin", position.segmentOrigin);
    if (position.segmentOrigin === this.#segmentOrigin) {
      return;
    }
    this.#segmentOrigin = Math.min(position.segmentOrigin, this.maxSegmentOrigin);
    this.#lastNativeOffset = position.nativeOffset;
    this.#version += 1;
  }

  /**
   * How close to an end of the native window the offset may get before the
   * origin is shifted. A quarter of the window keeps a full half-window of
   * slack on the far side after recentring, so a user scrolling steadily in
   * one direction triggers at most one recentre per half window.
   */
  #recenterGuard(): number {
    return this.maxNativeScroll / 4;
  }

  /**
   * Above this per-event delta an offset change is a seek rather than a
   * continuous scroll. One viewport is the natural line: no wheel, keyboard or
   * autoscroll step advances by a whole screen between scroll events.
   */
  #jumpThreshold(): number {
    return Math.max(this.#viewportSize, 1);
  }
}
