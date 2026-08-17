import type { SceneSurfaceId } from "./SceneSnapshot";

export type SurfaceId = SceneSurfaceId;

export interface ClientPoint {
  readonly space: "client";
  readonly x: number;
  readonly y: number;
}

export interface SurfacePoint<S extends SurfaceId = SurfaceId> {
  readonly space: "surface";
  readonly surface: S;
  readonly x: number;
  readonly y: number;
}

export interface SurfaceRect<S extends SurfaceId = SurfaceId> {
  readonly space: "surface";
  readonly surface: S;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

export interface SurfaceHostRect {
  readonly left: number;
  readonly top: number;
  readonly width: number;
  readonly height: number;
}

/** Minimal DOM shape keeps SurfaceHost geometry independently testable. */
export interface SurfaceHostElement {
  readonly parentElement?: SurfaceHostElement | null;
  getBoundingClientRect(): SurfaceHostRect;
}

export interface SurfaceLocalRegion {
  readonly x: number;
  readonly y: number;
  /** Defaults to the remaining element width after `x`. */
  readonly width?: number;
  /** Defaults to the remaining element height after `y`. */
  readonly height?: number;
}

export interface SurfaceAttachmentOptions {
  /** Client-pixel region relative to the element's measured border box. */
  readonly region?: SurfaceLocalRegion;
  readonly zIndex?: number;
}

export interface SurfaceSnapshot<S extends SurfaceId = SurfaceId> {
  readonly id: S;
  readonly attachmentVersion: number;
  readonly clientOrigin: Readonly<{ x: number; y: number }>;
  readonly cssSize: Readonly<{ width: number; height: number }>;
  readonly clipRect: SurfaceRect<S>;
  readonly zIndex: number;
}

export interface SurfacePick<S extends SurfaceId = SurfaceId> {
  readonly surface: S;
  readonly point: SurfacePoint<S>;
}

export interface SurfaceHostOptions {
  /** Injectable for tests; the browser default delegates to getComputedStyle(). */
  readonly getTransform?: (element: SurfaceHostElement) => string;
}

interface SurfaceAttachment<S extends SurfaceId = SurfaceId> {
  readonly surface: S;
  readonly element: SurfaceHostElement;
  readonly region: SurfaceLocalRegion | undefined;
  readonly zIndex: number;
  readonly attachmentVersion: number;
  readonly order: number;
}

interface OrderedSurfaceSnapshot<S extends SurfaceId = SurfaceId> {
  readonly snapshot: SurfaceSnapshot<S>;
  readonly order: number;
}

function assertFinite(name: string, value: number): void {
  if (!Number.isFinite(value)) {
    throw new RangeError(`${name} must be finite: ${value}`);
  }
}

function assertNonNegativeFinite(name: string, value: number): void {
  if (!Number.isFinite(value) || value < 0) {
    throw new RangeError(`${name} must be a non-negative finite number: ${value}`);
  }
}

function nextVersion(name: string, current: number): number {
  const next = current + 1;
  if (!Number.isSafeInteger(next)) {
    throw new RangeError(`${name} version exhausted`);
  }
  return next;
}

function defaultGetTransform(element: SurfaceHostElement): string {
  if (typeof getComputedStyle !== "function") {
    return "none";
  }
  return getComputedStyle(element as Element).transform;
}

export function clientPoint(x: number, y: number): ClientPoint {
  assertFinite("ClientPoint.x", x);
  assertFinite("ClientPoint.y", y);
  return Object.freeze({ space: "client", x, y });
}

export function surfacePoint<S extends SurfaceId>(surface: S, x: number, y: number): SurfacePoint<S> {
  assertFinite("SurfacePoint.x", x);
  assertFinite("SurfacePoint.y", y);
  return Object.freeze({ space: "surface", surface, x, y });
}

function surfaceRect<S extends SurfaceId>(
  surface: S,
  x: number,
  y: number,
  width: number,
  height: number
): SurfaceRect<S> {
  return Object.freeze({ space: "surface", surface, x, y, width, height });
}

/** Immutable client geometry for every attached surface at one publication. */
export class SurfaceSetSnapshot {
  readonly version: number;
  readonly surfaces: readonly SurfaceSnapshot[];

  readonly #byId: ReadonlyMap<SurfaceId, SurfaceSnapshot>;
  readonly #pickOrder: readonly OrderedSurfaceSnapshot[];

  constructor(version: number, entries: readonly OrderedSurfaceSnapshot[]) {
    if (!Number.isSafeInteger(version) || version < 0) {
      throw new RangeError(`SurfaceSetSnapshot.version must be a non-negative safe integer: ${version}`);
    }
    this.version = version;
    this.surfaces = Object.freeze(entries.map(({ snapshot }) => snapshot));
    this.#byId = new Map(entries.map(({ snapshot }) => [snapshot.id, snapshot]));
    this.#pickOrder = Object.freeze(
      [...entries].sort((left, right) =>
        right.snapshot.zIndex - left.snapshot.zIndex || right.order - left.order
      )
    );
    Object.freeze(this);
  }

  get<S extends SurfaceId>(surface: S): SurfaceSnapshot<S> | undefined {
    return this.#byId.get(surface) as SurfaceSnapshot<S> | undefined;
  }

  clientToSurface<S extends SurfaceId>(point: ClientPoint, surface: S): SurfacePoint<S> | null {
    assertFinite("ClientPoint.x", point.x);
    assertFinite("ClientPoint.y", point.y);
    const snapshot = this.get(surface);
    if (!snapshot) {
      return null;
    }
    return surfacePoint(
      surface,
      point.x - snapshot.clientOrigin.x,
      point.y - snapshot.clientOrigin.y
    );
  }

  surfaceToClient<S extends SurfaceId>(point: SurfacePoint<S>): ClientPoint | null {
    assertFinite("SurfacePoint.x", point.x);
    assertFinite("SurfacePoint.y", point.y);
    const snapshot = this.get(point.surface);
    if (!snapshot) {
      return null;
    }
    return clientPoint(
      snapshot.clientOrigin.x + point.x,
      snapshot.clientOrigin.y + point.y
    );
  }

  /** Picks one owner using half-open bounds and topmost z/order precedence. */
  pick(point: ClientPoint): SurfacePick | null {
    assertFinite("ClientPoint.x", point.x);
    assertFinite("ClientPoint.y", point.y);
    for (const { snapshot } of this.#pickOrder) {
      const { x, y } = snapshot.clientOrigin;
      const { width, height } = snapshot.cssSize;
      if (point.x >= x && point.x < x + width && point.y >= y && point.y < y + height) {
        const local = this.clientToSurface(point, snapshot.id) as SurfacePoint;
        return Object.freeze({ surface: snapshot.id, point: local });
      }
    }
    return null;
  }
}

/**
 * Owns all DOM measurement used to convert browser client coordinates into
 * typed surface coordinates. Attachments are measured atomically by refresh.
 * Non-none CSS transforms are rejected because a DOMRect alone cannot recover
 * a reliable local inverse for arbitrary transformed ancestor chains.
 */
export class SurfaceHost {
  readonly #getTransform: (element: SurfaceHostElement) => string;
  readonly #attachments = new Map<SurfaceId, SurfaceAttachment>();
  #snapshot = new SurfaceSetSnapshot(0, []);
  #attachmentVersion = 0;
  #order = 0;
  #disposed = false;

  constructor(options: SurfaceHostOptions = {}) {
    this.#getTransform = options.getTransform ?? defaultGetTransform;
  }

  get snapshot(): SurfaceSetSnapshot {
    return this.#snapshot;
  }

  attach<S extends SurfaceId>(
    surface: S,
    element: SurfaceHostElement,
    options: SurfaceAttachmentOptions = {}
  ): () => void {
    this.#assertActive();
    if (this.#attachments.has(surface)) {
      throw new Error(`SurfaceHost surface is already attached: ${surface}`);
    }
    assertFinite("SurfaceAttachmentOptions.zIndex", options.zIndex ?? 0);
    this.#attachmentVersion = nextVersion("Surface attachment", this.#attachmentVersion);
    this.#order = nextVersion("Surface attachment order", this.#order);
    const attachment: SurfaceAttachment<S> = Object.freeze({
      surface,
      element,
      region: options.region ? Object.freeze({ ...options.region }) : undefined,
      zIndex: options.zIndex ?? 0,
      attachmentVersion: this.#attachmentVersion,
      order: this.#order
    });
    this.#attachments.set(surface, attachment);

    let detached = false;
    return () => {
      if (detached) {
        return;
      }
      detached = true;
      if (this.#attachments.get(surface) !== attachment) {
        return;
      }
      this.#attachments.delete(surface);
      if (this.#snapshot.get(surface)) {
        this.#publishWithout(surface);
      }
    };
  }

  refresh(): SurfaceSetSnapshot {
    this.#assertActive();
    const entries: OrderedSurfaceSnapshot[] = [];
    for (const attachment of this.#attachments.values()) {
      entries.push(this.#measure(attachment));
    }
    this.#snapshot = new SurfaceSetSnapshot(
      nextVersion("SurfaceSetSnapshot", this.#snapshot.version),
      entries
    );
    return this.#snapshot;
  }

  clientToSurface<S extends SurfaceId>(point: ClientPoint, surface: S): SurfacePoint<S> | null {
    return this.#snapshot.clientToSurface(point, surface);
  }

  surfaceToClient<S extends SurfaceId>(point: SurfacePoint<S>): ClientPoint | null {
    return this.#snapshot.surfaceToClient(point);
  }

  pick(point: ClientPoint): SurfacePick | null {
    return this.#snapshot.pick(point);
  }

  dispose(): void {
    if (this.#disposed) {
      return;
    }
    this.#disposed = true;
    const hadState = this.#attachments.size > 0 || this.#snapshot.surfaces.length > 0;
    this.#attachments.clear();
    if (hadState) {
      this.#snapshot = new SurfaceSetSnapshot(
        nextVersion("SurfaceSetSnapshot", this.#snapshot.version),
        []
      );
    }
  }

  #measure(attachment: SurfaceAttachment): OrderedSurfaceSnapshot {
    this.#assertUntransformed(attachment.element, attachment.surface);
    const measured = attachment.element.getBoundingClientRect();
    assertFinite(`${attachment.surface}.rect.left`, measured.left);
    assertFinite(`${attachment.surface}.rect.top`, measured.top);
    assertNonNegativeFinite(`${attachment.surface}.rect.width`, measured.width);
    assertNonNegativeFinite(`${attachment.surface}.rect.height`, measured.height);

    const regionX = attachment.region?.x ?? 0;
    const regionY = attachment.region?.y ?? 0;
    const regionWidth = attachment.region?.width ?? measured.width - regionX;
    const regionHeight = attachment.region?.height ?? measured.height - regionY;
    assertNonNegativeFinite(`${attachment.surface}.region.x`, regionX);
    assertNonNegativeFinite(`${attachment.surface}.region.y`, regionY);
    assertNonNegativeFinite(`${attachment.surface}.region.width`, regionWidth);
    assertNonNegativeFinite(`${attachment.surface}.region.height`, regionHeight);
    if (regionX + regionWidth > measured.width || regionY + regionHeight > measured.height) {
      throw new RangeError(`SurfaceHost ${attachment.surface} region exceeds its measured element bounds`);
    }

    const clientOrigin = Object.freeze({
      x: measured.left + regionX,
      y: measured.top + regionY
    });
    const cssSize = Object.freeze({ width: regionWidth, height: regionHeight });
    const snapshot: SurfaceSnapshot = Object.freeze({
      id: attachment.surface,
      attachmentVersion: attachment.attachmentVersion,
      clientOrigin,
      cssSize,
      clipRect: surfaceRect(attachment.surface, 0, 0, regionWidth, regionHeight),
      zIndex: attachment.zIndex
    });
    return Object.freeze({ snapshot, order: attachment.order });
  }

  #assertUntransformed(element: SurfaceHostElement, surface: SurfaceId): void {
    let current: SurfaceHostElement | null | undefined = element;
    while (current) {
      const transform = this.#getTransform(current).trim();
      if (transform !== "" && transform !== "none") {
        throw new Error(
          `SurfaceHost does not support CSS transform on ${surface} or its ancestors: ${transform}`
        );
      }
      current = current.parentElement;
    }
  }

  #publishWithout(surface: SurfaceId): void {
    const entries = this.#snapshot.surfaces
      .filter((candidate) => candidate.id !== surface)
      .map((snapshot, order) => Object.freeze({ snapshot, order }));
    this.#snapshot = new SurfaceSetSnapshot(
      nextVersion("SurfaceSetSnapshot", this.#snapshot.version),
      entries
    );
  }

  #assertActive(): void {
    if (this.#disposed) {
      throw new Error("SurfaceHost is disposed");
    }
  }
}
