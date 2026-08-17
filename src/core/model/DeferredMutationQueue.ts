import { CyExcelError, CyExcelErrorCode } from "../errors";
import type { RevisionGate } from "./RevisionGate";

interface DeferredMutationIntent {
  readonly run: () => void;
  readonly reject?: (error: unknown) => void;
}

export interface DeferredMutationQueueOptions {
  readonly revisionGate: RevisionGate;
  readonly isGuarded: () => boolean;
  readonly onError: (error: unknown) => void;
  readonly beforeRun?: () => void;
}

/** FIFO gate shared by finalizer reentry and active read leases. */
export class DeferredMutationQueue {
  readonly #revisionGate: RevisionGate;
  readonly #isGuarded: () => boolean;
  readonly #onError: (error: unknown) => void;
  readonly #beforeRun: () => void;
  readonly #intents: DeferredMutationIntent[] = [];
  #draining = false;
  #leaseWakeupArmed = false;
  #disposed = false;

  constructor(options: DeferredMutationQueueOptions) {
    this.#revisionGate = options.revisionGate;
    this.#isGuarded = options.isGuarded;
    this.#onError = options.onError;
    this.#beforeRun = options.beforeRun ?? (() => {});
  }

  get draining(): boolean {
    return this.#draining;
  }

  enqueueOrRun(run: () => void): void {
    this.#assertActive();
    if (this.#mustDefer()) {
      this.#intents.push({ run });
      this.#requestDrain();
      return;
    }
    this.#beforeRun();
    run();
  }

  enqueuePromise(run: () => void): Promise<void> {
    if (this.#disposed) {
      return Promise.reject(this.#disposedError());
    }
    if (!this.#mustDefer()) {
      try {
        this.#beforeRun();
        run();
        return Promise.resolve();
      } catch (error) {
        return Promise.reject(error);
      }
    }
    return new Promise<void>((resolve, reject) => {
      this.#intents.push({
        run: () => {
          run();
          resolve();
        },
        reject
      });
      this.#requestDrain();
    });
  }

  drain(): void {
    if (this.#disposed || this.#draining || this.#isGuarded()) {
      return;
    }
    if (this.#revisionGate.activeReadLeases > 0) {
      this.#armLeaseWakeup();
      return;
    }
    this.#draining = true;
    try {
      while (!this.#disposed && !this.#isGuarded() && this.#revisionGate.activeReadLeases === 0) {
        const intent = this.#intents.shift();
        if (!intent) {
          return;
        }
        this.#runIntent(intent);
      }
    } finally {
      this.#draining = false;
      if (this.#intents.length > 0) {
        this.#requestDrain();
      }
    }
  }

  dispose(): void {
    if (this.#disposed) {
      return;
    }
    this.#disposed = true;
    const error = this.#disposedError();
    for (const intent of this.#intents.splice(0)) {
      intent.reject?.(error);
    }
  }

  #mustDefer(): boolean {
    return this.#draining || this.#isGuarded() || this.#revisionGate.activeReadLeases > 0 || this.#intents.length > 0;
  }

  #requestDrain(): void {
    if (this.#disposed || this.#intents.length === 0 || this.#isGuarded()) {
      return;
    }
    if (this.#revisionGate.activeReadLeases > 0) {
      this.#armLeaseWakeup();
      return;
    }
    this.drain();
  }

  #armLeaseWakeup(): void {
    if (this.#leaseWakeupArmed) {
      return;
    }
    this.#leaseWakeupArmed = true;
    this.#revisionGate.enqueueAfterReadLeases(() => {
      this.#leaseWakeupArmed = false;
      this.drain();
    });
  }

  #runIntent(intent: DeferredMutationIntent): void {
    try {
      this.#beforeRun();
      intent.run();
    } catch (error) {
      if (intent.reject) {
        intent.reject(error);
      } else {
        this.#onError(error);
      }
    }
  }

  #assertActive(): void {
    if (this.#disposed) {
      throw this.#disposedError();
    }
  }

  #disposedError(): CyExcelError {
    return new CyExcelError(CyExcelErrorCode.MutationRuntimeDisposed, "Mutation intent queue is disposed");
  }
}
