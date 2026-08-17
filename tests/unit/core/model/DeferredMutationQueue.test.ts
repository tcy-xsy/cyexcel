import { describe, expect, it, vi } from "vitest";

import {
  DeferredMutationQueue,
  RevisionGate
} from "../../../../src/core/model";
import { CyExcelError, CyExcelErrorCode } from "../../../../src/core/errors";
import { RenderScheduler } from "../../../../src/core/render";
import type { FrameScheduler, RenderTarget } from "../../../../src/core/render";

function createQueue(options: {
  readonly gate?: RevisionGate;
  readonly isGuarded?: () => boolean;
  readonly onError?: (error: unknown) => void;
} = {}) {
  const gate = options.gate ?? new RevisionGate();
  const calls: string[] = [];
  const queue = new DeferredMutationQueue({
    revisionGate: gate,
    isGuarded: options.isGuarded ?? (() => false),
    onError: options.onError ?? (() => {}),
    beforeRun: () => calls.push("before")
  });
  return { gate, queue, calls };
}

describe("DeferredMutationQueue", () => {
  it("defers a public mutation until the final read lease releases", () => {
    const { gate, queue, calls } = createQueue();
    const lease = gate.acquireReadLease();

    queue.enqueueOrRun(() => {
      calls.push(`run:${gate.activeReadLeases}`);
      gate.commit();
    });

    expect(calls).toEqual([]);
    expect(gate.revision).toBe(0);
    lease.release();
    expect(calls).toEqual(["before", "run:0"]);
    expect(gate.revision).toBe(1);
  });

  it("defers a renderer mutation from an asynchronous frame lease", () => {
    const gate = new RevisionGate();
    let callback: (() => void) | undefined;
    const frameScheduler: FrameScheduler = {
      requestFrame: (next) => {
        callback = next;
        return 1;
      },
      cancelFrame: () => {}
    };
    const calls: string[] = [];
    const queue = new DeferredMutationQueue({
      revisionGate: gate,
      isGuarded: () => false,
      onError: (error) => {
        throw error;
      }
    });
    const target: RenderTarget<{ readonly modelVersion: number }> = {
      paintPresentation() {},
      paintContent() {
        calls.push(`paint:${gate.activeReadLeases}`);
        queue.enqueueOrRun(() => {
          calls.push(`mutation:${gate.activeReadLeases}`);
          gate.commit();
        });
      },
      paintSelection() {}
    };
    const scheduler = new RenderScheduler(target, {
      frameScheduler,
      revisionGate: gate,
      getScene: () => Object.freeze({ modelVersion: gate.revision })
    });

    scheduler.invalidateAll();
    callback?.();

    expect(calls).toEqual(["paint:1", "mutation:0"]);
    expect(gate.revision).toBe(1);
  });

  it("settles a promise intent only after its deferred mutation runs", async () => {
    const { gate, queue, calls } = createQueue();
    const lease = gate.acquireReadLease();
    let settled = false;

    const result = queue.enqueuePromise(() => calls.push("promise-run"));
    void result.then(() => {
      settled = true;
    });
    await Promise.resolve();
    expect(settled).toBe(false);

    lease.release();
    await result;
    expect(calls).toEqual(["before", "promise-run"]);
    expect(settled).toBe(true);
  });

  it("preserves FIFO while finalization and leases unblock in either order", () => {
    const gate = new RevisionGate();
    let guarded = true;
    const { queue, calls } = createQueue({ gate, isGuarded: () => guarded });
    const lease = gate.acquireReadLease();

    queue.enqueueOrRun(() => calls.push("first"));
    queue.enqueueOrRun(() => calls.push("second"));
    lease.release();
    expect(calls).toEqual([]);

    guarded = false;
    queue.drain();
    expect(calls).toEqual(["before", "first", "before", "second"]);

    guarded = true;
    const nestedLease = gate.acquireReadLease();
    queue.enqueueOrRun(() => calls.push("third"));
    guarded = false;
    queue.drain();
    expect(calls).not.toContain("third");
    nestedLease.release();
    expect(calls.at(-2)).toBe("before");
    expect(calls.at(-1)).toBe("third");
  });

  it("queues reentrant intents behind older work while a drain is active", () => {
    const { gate, queue, calls } = createQueue();
    const lease = gate.acquireReadLease();

    queue.enqueueOrRun(() => {
      calls.push("first");
      queue.enqueueOrRun(() => calls.push("third"));
    });
    queue.enqueueOrRun(() => calls.push("second"));
    lease.release();

    expect(calls).toEqual([
      "before",
      "first",
      "before",
      "second",
      "before",
      "third"
    ]);
  });

  it("reports deferred sync failures and continues draining later intents", () => {
    const onError = vi.fn();
    const { gate, queue, calls } = createQueue({ onError });
    const lease = gate.acquireReadLease();
    queue.enqueueOrRun(() => {
      throw new Error("deferred failure");
    });
    queue.enqueueOrRun(() => calls.push("later"));

    lease.release();

    expect(onError).toHaveBeenCalledTimes(1);
    expect(String(onError.mock.calls[0]?.[0])).toContain("deferred failure");
    expect(calls).toEqual(["before", "before", "later"]);
  });

  it("rejects queued promises on dispose and ignores a later lease wakeup", async () => {
    const { gate, queue } = createQueue();
    const lease = gate.acquireReadLease();
    const run = vi.fn();
    const result = queue.enqueuePromise(run);

    queue.dispose();
    await expect(result).rejects.toMatchObject({
      code: CyExcelErrorCode.MutationRuntimeDisposed
    });
    lease.release();
    expect(run).not.toHaveBeenCalled();
    expect(() => queue.enqueueOrRun(run)).toThrow(CyExcelError);
  });
});
