import type { ModelRevision, RevisionGate } from "../model";

export interface InteractionSceneSnapshot {
  readonly modelVersion: number;
}

/** Minimal DeferredMutationQueue surface used by interaction dispatch. */
export interface InteractionIntentQueue {
  enqueueOrRun(intent: () => void): void;
}

export interface InteractionFrame<
  TScene extends InteractionSceneSnapshot,
  TSurfaces,
  TSelection = never
> {
  readonly revision: ModelRevision;
  readonly scene: TScene;
  readonly surfaces: TSurfaces;
  readonly selection: TSelection | undefined;
  readonly structureVersion: number | undefined;
}

export interface InteractionFrameProviderOptions<
  TScene extends InteractionSceneSnapshot,
  TSurfaces,
  TSelection = never
> {
  readonly revisionGate: RevisionGate;
  readonly intentQueue: InteractionIntentQueue;
  readonly getScene: () => TScene;
  readonly getSurfaces: () => TSurfaces;
  readonly getSelection?: () => TSelection;
  readonly getStructureVersion?: () => number;
}

/**
 * Captures one immutable interaction frame under a short model read lease.
 * Gesture sessions retain stable identities or version numbers, never leases.
 */
export class InteractionFrameProvider<
  TScene extends InteractionSceneSnapshot,
  TSurfaces,
  TSelection = never
> {
  readonly #revisionGate: RevisionGate;
  readonly #intentQueue: InteractionIntentQueue;
  readonly #getScene: () => TScene;
  readonly #getSurfaces: () => TSurfaces;
  readonly #getSelection: (() => TSelection) | undefined;
  readonly #getStructureVersion: (() => number) | undefined;

  constructor(options: InteractionFrameProviderOptions<TScene, TSurfaces, TSelection>) {
    this.#revisionGate = options.revisionGate;
    this.#intentQueue = options.intentQueue;
    this.#getScene = options.getScene;
    this.#getSurfaces = options.getSurfaces;
    this.#getSelection = options.getSelection;
    this.#getStructureVersion = options.getStructureVersion;
  }

  dispatch<TResult>(handler: (frame: InteractionFrame<TScene, TSurfaces, TSelection>) => TResult): TResult {
    const lease = this.#revisionGate.acquireReadLease();
    try {
      const scene = this.#getScene();
      if (scene.modelVersion !== lease.revision) {
        throw new Error(
          `Interaction frame revision ${lease.revision} does not match scene modelVersion ${scene.modelVersion}`
        );
      }
      const surfaces = this.#getSurfaces();
      const selection = this.#getSelection?.();
      const structureVersion = this.#getStructureVersion?.();
      const frame: InteractionFrame<TScene, TSurfaces, TSelection> = Object.freeze({
        revision: lease.revision,
        scene,
        surfaces,
        selection,
        structureVersion
      });
      return handler(frame);
    } finally {
      lease.release();
    }
  }

  /** Routes mutation work through the shared queue only after all current read leases release. */
  enqueueIntent(intent: () => void): void {
    this.#intentQueue.enqueueOrRun(intent);
  }
}
