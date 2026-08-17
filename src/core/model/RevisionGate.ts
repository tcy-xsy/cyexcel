import { CyExcelError, CyExcelErrorCode } from "../errors";

declare const modelRevisionBrand: unique symbol;

/** A monotonically increasing version of the mutable model. */
export type ModelRevision = number & { readonly [modelRevisionBrand]: "ModelRevision" };

export interface ReadLease {
  readonly revision: ModelRevision;
  readonly released: boolean;
  release(): void;
}

/**
 * Keeps a render or input dispatch on one model revision without copying the
 * model. A future MutationCoordinator commits only after all leases release.
 */
export class RevisionGate {
  #revision: ModelRevision;
  #leaseCount = 0;
  #afterLeaseQueue: Array<() => void> = [];
  #flushingAfterLeaseQueue = false;

  constructor(initialRevision = 0) {
    if (!Number.isSafeInteger(initialRevision) || initialRevision < 0) {
      throw new RangeError(`initialRevision must be a non-negative safe integer: ${initialRevision}`);
    }
    this.#revision = initialRevision as ModelRevision;
  }

  get revision(): ModelRevision {
    return this.#revision;
  }

  get activeReadLeases(): number {
    return this.#leaseCount;
  }

  acquireReadLease(): ReadLease {
    this.#leaseCount += 1;
    const revision = this.#revision;
    let released = false;

    return {
      revision,
      get released(): boolean {
        return released;
      },
      release: () => {
        if (released) {
          return;
        }
        released = true;
        this.#leaseCount -= 1;
        if (this.#leaseCount === 0) {
          this.#flushAfterLeaseQueue();
        }
      }
    };
  }

  assertCanCommit(): void {
    if (this.#leaseCount > 0) {
      throw new CyExcelError(
        CyExcelErrorCode.MutationBlockedByReadLease,
        "Cannot commit a mutation while a render or interaction read lease is active"
      );
    }
  }

  /** Advances the revision after a successful, fully finalized commit. */
  commit(): ModelRevision {
    this.assertCanCommit();
    this.#revision = (this.#revision + 1) as ModelRevision;
    return this.#revision;
  }

  /** Queues a mutation intent until it can be replanned against a fresh revision. */
  enqueueAfterReadLeases(callback: () => void): void {
    this.#afterLeaseQueue.push(callback);
    if (this.#leaseCount === 0) {
      this.#flushAfterLeaseQueue();
    }
  }

  #flushAfterLeaseQueue(): void {
    if (this.#flushingAfterLeaseQueue) {
      return;
    }
    this.#flushingAfterLeaseQueue = true;
    try {
      while (this.#leaseCount === 0) {
        const callback = this.#afterLeaseQueue.shift();
        if (!callback) {
          return;
        }
        callback();
      }
    } finally {
      this.#flushingAfterLeaseQueue = false;
    }
  }
}
