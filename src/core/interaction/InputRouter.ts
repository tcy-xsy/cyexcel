import type { Disposer } from "../types";
import type { ClientPoint, SurfaceId, SurfacePoint } from "../scene";

/** The only surface capability InputRouter consumes from an InteractionFrame. */
export interface InputRouterSurfaceSet {
  clientToSurface<S extends SurfaceId>(point: ClientPoint, surface: S): SurfacePoint<S> | null;
}

/** Structurally compatible with InteractionFrame without coupling routing to its provider. */
export interface InputRouterFrame {
  readonly surfaces: InputRouterSurfaceSet;
}

export interface InputRouterDispatchInput<TEvent, TPathEntry = unknown> {
  readonly event: TEvent;
  readonly clientPoint: ClientPoint;
  /** Event target first, followed by its ancestors, matching Event.composedPath(). */
  readonly path: readonly TPathEntry[];
  /** An explicit owner bypasses policy resolution but never bypasses surface conversion. */
  readonly surface?: SurfaceId;
}

export interface InputSurfacePolicyContext<
  TFrame extends InputRouterFrame,
  TEvent,
  TPathEntry
> {
  readonly frame: TFrame;
  readonly event: TEvent;
  readonly clientPoint: ClientPoint;
  readonly path: readonly TPathEntry[];
}

export type InputTargetResolution =
  | Readonly<{ readonly kind: "surface"; readonly surface: SurfaceId }>
  | Readonly<{ readonly kind: "native" }>
  | Readonly<{ readonly kind: "none" }>;

const NATIVE_INPUT_TARGET: InputTargetResolution = Object.freeze({ kind: "native" });
const NO_INPUT_TARGET: InputTargetResolution = Object.freeze({ kind: "none" });

export function surfaceInputTarget(surface: SurfaceId): InputTargetResolution {
  return Object.freeze({ kind: "surface", surface });
}

export function nativeInputTarget(): InputTargetResolution {
  return NATIVE_INPUT_TARGET;
}

export function noInputTarget(): InputTargetResolution {
  return NO_INPUT_TARGET;
}

/** Resolves one terminal input owner. Geometry/z-index fallback is intentionally absent. */
export type InputSurfacePolicy<
  TFrame extends InputRouterFrame,
  TEvent,
  TPathEntry = unknown
> = (
  context: InputSurfacePolicyContext<TFrame, TEvent, TPathEntry>
) => InputTargetResolution;

export type EventPathSurfaceLookup<
  TFrame extends InputRouterFrame,
  TEvent,
  TPathEntry
> = (
  entry: TPathEntry,
  index: number,
  context: InputSurfacePolicyContext<TFrame, TEvent, TPathEntry>
) => SurfaceId | null | undefined;

/**
 * Builds a closest-event-path-owner policy. For example, a body canvas earlier
 * in the path wins over a menu-overlay attachment on a shared ancestor.
 */
export function createEventPathSurfacePolicy<
  TFrame extends InputRouterFrame,
  TEvent,
  TPathEntry = unknown
>(
  lookup: EventPathSurfaceLookup<TFrame, TEvent, TPathEntry>
): InputSurfacePolicy<TFrame, TEvent, TPathEntry> {
  return (context) => {
    for (let index = 0; index < context.path.length; index += 1) {
      const entry = context.path[index];
      if (entry === undefined) {
        continue;
      }
      const surface = lookup(entry, index, context);
      if (surface != null) {
        return surfaceInputTarget(surface);
      }
    }
    return noInputTarget();
  };
}

export interface InputCandidateContext<
  TFrame extends InputRouterFrame,
  TEvent,
  TPathEntry = unknown
> {
  /** The exact frame identity supplied to dispatch(). */
  readonly frame: TFrame;
  readonly event: TEvent;
  readonly clientPoint: ClientPoint;
  readonly path: readonly TPathEntry[];
  readonly surface: SurfaceId;
  readonly point: SurfacePoint;
}

export interface InputCandidate<
  TFrame extends InputRouterFrame,
  TEvent,
  TResult = void,
  TPathEntry = unknown
> {
  readonly id: string;
  readonly surface: SurfaceId;
  /** Higher values run first; equal values retain registration order. */
  readonly priority?: number;
  readonly route: (
    context: InputCandidateContext<TFrame, TEvent, TPathEntry>
  ) => InputCandidateResult<TResult>;
}

export type InputCandidateResult<TResult> =
  | Readonly<{ readonly handled: false }>
  | Readonly<{ readonly handled: true; readonly value: TResult }>;

export function declineInput(): InputCandidateResult<never> {
  return Object.freeze({ handled: false });
}

export function handleInput<TResult>(value: TResult): InputCandidateResult<TResult> {
  return Object.freeze({ handled: true, value });
}

export type InputRouteResult<TResult> =
  | Readonly<{
      handled: false;
      reason: "native-owned";
      surface: null;
      point: null;
    }>
  | Readonly<{
      handled: false;
      reason: "no-surface";
      surface: null;
      point: null;
    }>
  | Readonly<{
      handled: false;
      reason: "surface-unavailable";
      surface: SurfaceId;
      point: null;
    }>
  | Readonly<{
      handled: false;
      reason: "not-accepted";
      surface: SurfaceId;
      point: SurfacePoint;
    }>
  | Readonly<{
      handled: true;
      candidateId: string;
      surface: SurfaceId;
      point: SurfacePoint;
      value: TResult;
    }>;

export interface InputRouterOptions<
  TFrame extends InputRouterFrame,
  TEvent,
  TPathEntry = unknown
> {
  /** Optional when every dispatch supplies InputRouterDispatchInput.surface. */
  readonly surfacePolicy?: InputSurfacePolicy<TFrame, TEvent, TPathEntry>;
}

interface CandidateEntry<
  TFrame extends InputRouterFrame,
  TEvent,
  TResult,
  TPathEntry
> {
  readonly sequence: number;
  readonly priority: number;
  readonly candidate: InputCandidate<TFrame, TEvent, TResult, TPathEntry>;
}

/**
 * Pure surface router. The caller owns event capture and frame acquisition;
 * this class only resolves policy, converts through that frame, and dispatches.
 */
export class InputRouter<
  TFrame extends InputRouterFrame,
  TEvent,
  TResult = void,
  TPathEntry = unknown
> {
  readonly #surfacePolicy: InputSurfacePolicy<TFrame, TEvent, TPathEntry> | undefined;
  readonly #candidates = new Map<SurfaceId, CandidateEntry<TFrame, TEvent, TResult, TPathEntry>[]>();
  #sequence = 0;
  #disposed = false;

  constructor(options: InputRouterOptions<TFrame, TEvent, TPathEntry> = {}) {
    this.#surfacePolicy = options.surfacePolicy;
  }

  register(candidate: InputCandidate<TFrame, TEvent, TResult, TPathEntry>): Disposer {
    this.#assertActive();
    if (candidate.id.trim().length === 0) {
      throw new RangeError("InputCandidate.id must not be empty");
    }
    const priority = candidate.priority ?? 0;
    if (!Number.isFinite(priority)) {
      throw new RangeError(`InputCandidate.priority must be finite: ${priority}`);
    }
    this.#sequence += 1;
    if (!Number.isSafeInteger(this.#sequence)) {
      throw new RangeError("InputRouter registration sequence exhausted");
    }
    const entry: CandidateEntry<TFrame, TEvent, TResult, TPathEntry> = {
      sequence: this.#sequence,
      priority,
      candidate
    };
    const candidates = this.#candidates.get(candidate.surface) ?? [];
    candidates.push(entry);
    candidates.sort((left, right) => right.priority - left.priority || left.sequence - right.sequence);
    this.#candidates.set(candidate.surface, candidates);

    let disposed = false;
    return () => {
      if (disposed) {
        return;
      }
      disposed = true;
      const current = this.#candidates.get(candidate.surface);
      if (!current) {
        return;
      }
      const index = current.indexOf(entry);
      if (index >= 0) {
        current.splice(index, 1);
      }
      if (current.length === 0) {
        this.#candidates.delete(candidate.surface);
      }
    };
  }

  dispatch(
    frame: TFrame,
    input: InputRouterDispatchInput<TEvent, TPathEntry>
  ): InputRouteResult<TResult> {
    this.#assertActive();
    const policyContext: InputSurfacePolicyContext<TFrame, TEvent, TPathEntry> = Object.freeze({
      frame,
      event: input.event,
      clientPoint: input.clientPoint,
      path: input.path
    });
    const target = input.surface !== undefined
      ? surfaceInputTarget(input.surface)
      : this.#surfacePolicy?.(policyContext) ?? noInputTarget();
    if (target.kind === "native") {
      return Object.freeze({ handled: false, reason: "native-owned", surface: null, point: null });
    }
    if (target.kind === "none") {
      return Object.freeze({ handled: false, reason: "no-surface", surface: null, point: null });
    }
    const surface = target.surface;

    const point = frame.surfaces.clientToSurface(input.clientPoint, surface);
    if (point === null) {
      return Object.freeze({
        handled: false,
        reason: "surface-unavailable",
        surface,
        point: null
      });
    }
    const context: InputCandidateContext<TFrame, TEvent, TPathEntry> = Object.freeze({
      frame,
      event: input.event,
      clientPoint: input.clientPoint,
      path: input.path,
      surface,
      point
    });

    const candidates = this.#candidates.get(surface)?.slice() ?? [];
    for (const { candidate } of candidates) {
      const candidateResult = candidate.route(context);
      if (!candidateResult.handled) {
        continue;
      }
      return Object.freeze({
        handled: true,
        candidateId: candidate.id,
        surface,
        point,
        value: candidateResult.value
      });
    }
    return Object.freeze({ handled: false, reason: "not-accepted", surface, point });
  }

  dispose(): void {
    if (this.#disposed) {
      return;
    }
    this.#disposed = true;
    this.#candidates.clear();
  }

  #assertActive(): void {
    if (this.#disposed) {
      throw new Error("InputRouter is disposed");
    }
  }
}
