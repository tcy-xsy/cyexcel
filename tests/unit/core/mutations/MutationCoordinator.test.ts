import { describe, expect, it } from "vitest";

import {
  MutationCoordinator,
  type Mutation,
  type PreparedMutationParticipant
} from "../../../../src/core/mutations";
import { CyExcelError, CyExcelErrorCode } from "../../../../src/core/errors";
import type { ModelRevision } from "../../../../src/core/model";
import type { Patch } from "../../../../src/core/types";

interface ParticipantOptions {
  readonly type: string;
  readonly hasEffect?: boolean;
  readonly changes: readonly Patch[];
  readonly prepare?: () => void;
  readonly commit: () => void;
  readonly rollback: () => void;
  readonly dispose?: () => void;
  readonly estimatedBytes?: number;
}

function createMutation(options: ParticipantOptions): Mutation {
  const inverse: Mutation = {
    type: `${options.type}:inverse`,
    changes: options.changes,
    prepare: () => {
      throw new Error("inverse mutation is not prepared in these coordinator tests");
    }
  };

  const mutation: Mutation = {
    type: options.type,
    changes: options.changes,
    prepare: ({ baseRevision }: { readonly baseRevision: ModelRevision }): PreparedMutationParticipant => {
      void baseRevision;
      options.prepare?.();
      return {
        hasEffect: options.hasEffect ?? true,
        forward: mutation,
        inverse,
        estimatedBytes: options.estimatedBytes ?? 0,
        commit: options.commit,
        rollback: options.rollback,
        dispose: options.dispose ?? (() => {})
      };
    }
  };

  return mutation;
}

function expectErrorCode(action: () => unknown, code: CyExcelErrorCode): void {
  try {
    action();
    throw new Error(`Expected CyExcelError with code ${code}`);
  } catch (error) {
    expect(error).toBeInstanceOf(CyExcelError);
    expect((error as CyExcelError).code).toBe(code);
  }
}

describe("MutationCoordinator", () => {
  it("prepares every participant without writing live state", () => {
    const coordinator = new MutationCoordinator();
    const live = { first: 0, second: 0 };
    const preparedTypes: string[] = [];
    const first = createMutation({
      type: "first",
      changes: [{ type: "cell-value", row: 0, column: 0 }],
      prepare: () => preparedTypes.push("first"),
      commit: () => {
        live.first = 1;
      },
      rollback: () => {
        live.first = 0;
      }
    });
    const second = createMutation({
      type: "second",
      changes: [{ type: "cell-value", row: 0, column: 1 }],
      prepare: () => preparedTypes.push("second"),
      commit: () => {
        live.second = 2;
      },
      rollback: () => {
        live.second = 0;
      }
    });

    const prepared = coordinator.prepare("tx-prepare", [first, second]);

    expect(preparedTypes).toEqual(["first", "second"]);
    expect(live).toEqual({ first: 0, second: 0 });
    expect(coordinator.revisionGate.revision).toBe(0);
    expect(coordinator.state).toBe("prepared");

    coordinator.discard(prepared);
  });

  it("commits multiple participants as one revision and preserves ordered changes", () => {
    const coordinator = new MutationCoordinator();
    const writes: string[] = [];
    const firstChange: Patch = { type: "cell-value", row: 2, column: 1 };
    const secondChange: Patch = {
      type: "row-structure",
      operation: "insert",
      start: 4,
      count: 1
    };
    const first = createMutation({
      type: "first",
      changes: [firstChange],
      commit: () => writes.push("first"),
      rollback: () => writes.push("rollback-first"),
      estimatedBytes: 11
    });
    const second = createMutation({
      type: "second",
      changes: [secondChange],
      commit: () => writes.push("second"),
      rollback: () => writes.push("rollback-second"),
      estimatedBytes: 13
    });

    const receipt = coordinator.execute("tx-success", [first, second]);

    expect(writes).toEqual(["first", "second"]);
    expect(receipt.beforeRevision).toBe(0);
    expect(receipt.afterRevision).toBe(1);
    expect(coordinator.revisionGate.revision).toBe(1);
    expect(receipt.changeSet.modelVersion).toBe(1);
    expect(receipt.changeSet.changes).toEqual([firstChange, secondChange]);
    expect(receipt.patches).toEqual([firstChange, secondChange]);
    expect(receipt.historyBytes).toBe(24);
    expect(coordinator.state).toBe("idle");
  });

  it("returns an explicit no-op receipt without committing live state or advancing revision", () => {
    const coordinator = new MutationCoordinator();
    let commitCalls = 0;
    let rollbackCalls = 0;
    let disposeCalls = 0;
    const noOp = createMutation({
      type: "no-op",
      hasEffect: false,
      // A misleading Patch proves that the participant contract is authoritative.
      changes: [{ type: "full-reset" }],
      estimatedBytes: 999,
      commit: () => {
        commitCalls += 1;
      },
      rollback: () => {
        rollbackCalls += 1;
      },
      dispose: () => {
        disposeCalls += 1;
      }
    });

    const prepared = coordinator.prepare("tx-no-op", [noOp]);

    expect(prepared.hasEffect).toBe(false);
    expect(prepared.participants).toEqual([]);
    expect(prepared.changes).toEqual([]);
    expect(prepared.estimatedHistoryBytes).toBe(0);
    expect(disposeCalls).toBe(1);

    const lease = coordinator.revisionGate.acquireReadLease();
    const receipt = coordinator.commit(prepared);
    lease.release();

    expect(receipt.noop).toBe(true);
    expect(receipt.beforeRevision).toBe(0);
    expect(receipt.afterRevision).toBe(0);
    expect(receipt.forward).toEqual([]);
    expect(receipt.inverse).toEqual([]);
    expect(receipt.patches).toEqual([]);
    expect(receipt.changeSet).toEqual({ modelVersion: 0, changes: [] });
    expect(receipt.historyBytes).toBe(0);
    expect(commitCalls).toBe(0);
    expect(rollbackCalls).toBe(0);
    expect(coordinator.revisionGate.revision).toBe(0);
    expect(coordinator.state).toBe("idle");
  });

  it("filters no-op participants while committing a patchless effect in one revision", () => {
    const coordinator = new MutationCoordinator();
    const calls: string[] = [];
    const noOp = createMutation({
      type: "filtered-no-op",
      hasEffect: false,
      changes: [{ type: "full-reset" }],
      estimatedBytes: 1000,
      commit: () => calls.push("unexpected-no-op-commit"),
      rollback: () => calls.push("unexpected-no-op-rollback"),
      dispose: () => calls.push("dispose-no-op")
    });
    const patchlessEffect = createMutation({
      type: "patchless-effect",
      hasEffect: true,
      changes: [],
      estimatedBytes: 17,
      commit: () => calls.push("commit-effect"),
      rollback: () => calls.push("rollback-effect"),
      dispose: () => calls.push("dispose-effect")
    });

    const prepared = coordinator.prepare("tx-mixed-effect", [noOp, patchlessEffect]);

    expect(prepared.hasEffect).toBe(true);
    expect(prepared.participants.map((participant) => participant.forward.type)).toEqual([
      "patchless-effect"
    ]);
    expect(prepared.changes).toEqual([]);
    expect(prepared.estimatedHistoryBytes).toBe(17);

    const receipt = coordinator.commit(prepared);

    expect(receipt.noop).toBe(false);
    expect(receipt.afterRevision).toBe(1);
    expect(receipt.forward.map((mutation) => mutation.type)).toEqual(["patchless-effect"]);
    expect(receipt.inverse.map((mutation) => mutation.type)).toEqual(["patchless-effect:inverse"]);
    expect(receipt.patches).toEqual([]);
    expect(receipt.historyBytes).toBe(17);
    expect(calls).toEqual(["dispose-no-op", "commit-effect", "dispose-effect"]);
    expect(coordinator.revisionGate.revision).toBe(1);
  });

  it("rolls back the throwing participant and earlier participants in reverse order", () => {
    const coordinator = new MutationCoordinator();
    const live = { first: 0, second: 0 };
    const calls: string[] = [];
    const first = createMutation({
      type: "first",
      changes: [{ type: "cell-value", row: 0, column: 0 }],
      commit: () => {
        calls.push("commit-first");
        live.first = 1;
      },
      rollback: () => {
        calls.push("rollback-first");
        live.first = 0;
      }
    });
    const second = createMutation({
      type: "second",
      changes: [{ type: "cell-value", row: 0, column: 1 }],
      commit: () => {
        calls.push("commit-second");
        live.second = 2;
        throw new Error("second participant failed after writing");
      },
      rollback: () => {
        calls.push("rollback-second");
        live.second = 0;
      }
    });

    expect(() => coordinator.execute("tx-rollback", [first, second])).toThrow(
      "second participant failed after writing"
    );

    expect(calls).toEqual(["commit-first", "commit-second", "rollback-second", "rollback-first"]);
    expect(live).toEqual({ first: 0, second: 0 });
    expect(coordinator.revisionGate.revision).toBe(0);
    expect(coordinator.state).toBe("idle");
  });

  it("faults the runtime after rollback failure and rejects later mutations", () => {
    const coordinator = new MutationCoordinator();
    let liveValue = 0;
    let laterPrepareCalls = 0;
    const broken = createMutation({
      type: "broken",
      changes: [{ type: "cell-value", row: 1, column: 1 }],
      commit: () => {
        liveValue = 1;
        throw new Error("commit failed");
      },
      rollback: () => {
        throw new Error("rollback failed");
      }
    });

    expectErrorCode(
      () => coordinator.execute("tx-fault", [broken]),
      CyExcelErrorCode.MutationRuntimeFaulted
    );
    expect(liveValue).toBe(1);
    expect(coordinator.state).toBe("faulted");
    expect(coordinator.revisionGate.revision).toBe(0);

    const later = createMutation({
      type: "later",
      changes: [],
      prepare: () => {
        laterPrepareCalls += 1;
      },
      commit: () => {
        liveValue = 2;
      },
      rollback: () => {
        liveValue = 1;
      }
    });
    expectErrorCode(
      () => coordinator.execute("tx-after-fault", [later]),
      CyExcelErrorCode.MutationRuntimeFaulted
    );
    expect(laterPrepareCalls).toBe(0);
    expect(liveValue).toBe(1);
  });

  it("blocks prepare and commit while a read lease is active without writing", () => {
    const coordinator = new MutationCoordinator();
    let prepareCalls = 0;
    let liveValue = 0;
    const mutation = createMutation({
      type: "leased",
      changes: [{ type: "cell-value", row: 3, column: 3 }],
      prepare: () => {
        prepareCalls += 1;
      },
      commit: () => {
        liveValue = 1;
      },
      rollback: () => {
        liveValue = 0;
      }
    });

    const prepareLease = coordinator.revisionGate.acquireReadLease();
    expectErrorCode(
      () => coordinator.prepare("tx-blocked-prepare", [mutation]),
      CyExcelErrorCode.MutationBlockedByReadLease
    );
    expect(prepareCalls).toBe(0);
    expect(liveValue).toBe(0);
    prepareLease.release();

    const prepared = coordinator.prepare("tx-blocked-commit", [mutation]);
    expect(prepareCalls).toBe(1);
    const commitLease = coordinator.revisionGate.acquireReadLease();
    expectErrorCode(
      () => coordinator.commit(prepared),
      CyExcelErrorCode.MutationBlockedByReadLease
    );
    expect(liveValue).toBe(0);
    expect(coordinator.revisionGate.revision).toBe(0);
    expect(coordinator.state).toBe("prepared");

    commitLease.release();
    coordinator.commit(prepared);
    expect(liveValue).toBe(1);
    expect(coordinator.revisionGate.revision).toBe(1);
  });

  it("disposes an active planning session without returning the coordinator to idle", () => {
    const coordinator = new MutationCoordinator();
    let resourcesDisposed = 0;
    const session = coordinator.begin("tx-dispose-planning", (baseRevision) => ({
      model: {
        baseRevision,
        getRowCount: () => 0,
        getColumnCount: () => 0,
        getCell: () => undefined
      },
      prepareData: () => {
        throw new Error("not used");
      },
      dispose: () => {
        resourcesDisposed += 1;
      }
    }));

    coordinator.dispose();
    session.discard();

    expect(resourcesDisposed).toBe(1);
    expect(coordinator.state).toBe("disposed");
    expectErrorCode(
      () => coordinator.prepare("tx-after-dispose", []),
      CyExcelErrorCode.MutationRuntimeDisposed
    );
  });
});
