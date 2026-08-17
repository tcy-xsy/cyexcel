import { describe, expect, it } from "vitest";

import { CyExcelError, CyExcelErrorCode } from "../../../../src/core/errors";
import { RevisionGate } from "../../../../src/core/model";

describe("RevisionGate", () => {
  it("pins each lease to its acquisition revision and blocks commits until released", () => {
    const gate = new RevisionGate();
    const lease = gate.acquireReadLease();

    expect(lease.revision).toBe(0);
    expect(() => gate.commit()).toThrow(CyExcelError);
    try {
      gate.commit();
    } catch (error) {
      expect((error as CyExcelError).code).toBe(CyExcelErrorCode.MutationBlockedByReadLease);
    }

    lease.release();
    expect(lease.released).toBe(true);
    expect(gate.commit()).toBe(1);
  });

  it("runs queued intents only after the final nested lease releases", () => {
    const gate = new RevisionGate();
    const first = gate.acquireReadLease();
    const second = gate.acquireReadLease();
    const calls: number[] = [];

    gate.enqueueAfterReadLeases(() => calls.push(gate.activeReadLeases));
    first.release();
    expect(calls).toEqual([]);

    second.release();
    expect(calls).toEqual([0]);
  });

  it("preserves FIFO order when a draining callback enqueues another callback", () => {
    const gate = new RevisionGate();
    const lease = gate.acquireReadLease();
    const calls: string[] = [];

    gate.enqueueAfterReadLeases(() => {
      calls.push("first");
      gate.enqueueAfterReadLeases(() => calls.push("third"));
    });
    gate.enqueueAfterReadLeases(() => calls.push("second"));

    lease.release();
    expect(calls).toEqual(["first", "second", "third"]);
  });

  it("keeps callbacks queued when a callback acquires a new read lease", () => {
    const gate = new RevisionGate();
    const lease = gate.acquireReadLease();
    const calls: string[] = [];
    let nestedLease: ReturnType<RevisionGate["acquireReadLease"]> | undefined;

    gate.enqueueAfterReadLeases(() => {
      calls.push("first");
      nestedLease = gate.acquireReadLease();
    });
    gate.enqueueAfterReadLeases(() => calls.push("second"));

    lease.release();
    expect(calls).toEqual(["first"]);
    nestedLease?.release();
    expect(calls).toEqual(["first", "second"]);
  });
});
