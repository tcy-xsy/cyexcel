import { describe, expect, it } from "vitest";
import { MutationCoordinator, StructureIntegrationMutation } from "../../../../src/core/mutations";
import type { AxisChangeDescriptor } from "../../../../src/core/layout";

interface Root {
  readonly count: number;
}

interface Payload {
  readonly amount: number;
  readonly resultCount?: number;
}

function inverse(change: AxisChangeDescriptor): AxisChangeDescriptor {
  return {
    axis: change.axis,
    operation: change.operation === "insert" ? "delete" : "insert",
    start: change.start,
    count: change.count,
    targetSpace: "pre-removal"
  };
}

describe("StructureIntegrationMutation", () => {
  it("prepares off-state and publishes only during commit", () => {
    let root: Root = Object.freeze({ count: 2 });
    let prepareCalls = 0;
    const owner = {
      getRoot: () => root,
      prepare: (base: Root, change: AxisChangeDescriptor, payload: Payload) => {
        prepareCalls += 1;
        const direction = change.operation === "insert" ? 1 : -1;
        const nextRoot = Object.freeze({ count: payload.resultCount ?? base.count + direction * payload.amount });
        return {
          nextRoot,
          forwardPayload: { amount: payload.amount, resultCount: nextRoot.count },
          inverseChange: inverse(change),
          inversePayload: { amount: payload.amount, resultCount: base.count }
        };
      },
      replaceRoot: (next: Root) => {
        root = next;
      }
    };
    const coordinator = new MutationCoordinator();
    const prepared = coordinator.prepare("t-1", [
      new StructureIntegrationMutation(
        owner,
        { axis: "row", operation: "insert", start: 2, count: 3 },
        { amount: 3 }
      )
    ]);

    expect(root.count).toBe(2);
    expect(prepareCalls).toBe(1);
    const receipt = coordinator.commit(prepared);
    expect(root.count).toBe(5);

    coordinator.execute("t-2", receipt.inverse);
    expect(root.count).toBe(2);

    coordinator.execute("t-3", receipt.forward);
    expect(root.count).toBe(5);
  });

  it("rejects a stale owner root before publishing", () => {
    let root: Root = Object.freeze({ count: 2 });
    const owner = {
      getRoot: () => root,
      prepare: (base: Root, change: AxisChangeDescriptor, payload: Payload) => ({
        nextRoot: Object.freeze({ count: base.count + payload.amount }),
        inverseChange: inverse(change),
        inversePayload: payload
      }),
      replaceRoot: (next: Root) => {
        root = next;
      }
    };
    const coordinator = new MutationCoordinator();
    const prepared = coordinator.prepare("t-1", [
      new StructureIntegrationMutation(
        owner,
        { axis: "column", operation: "insert", start: 2, count: 1 },
        { amount: 1 }
      )
    ]);
    root = Object.freeze({ count: 9 });

    expect(() => coordinator.commit(prepared)).toThrow("Structure root changed");
    expect(root.count).toBe(9);
    expect(coordinator.state).toBe("idle");
  });

  it("rolls back when replaceRoot publishes partially and then throws", () => {
    let root: Root = Object.freeze({ count: 2 });
    let shouldThrow = true;
    const owner = {
      getRoot: () => root,
      prepare: (base: Root, change: AxisChangeDescriptor, payload: Payload) => ({
        nextRoot: Object.freeze({ count: base.count + payload.amount }),
        inverseChange: inverse(change),
        inversePayload: payload
      }),
      replaceRoot: (next: Root) => {
        root = next;
        if (shouldThrow) {
          shouldThrow = false;
          throw new Error("partial publish");
        }
      }
    };
    const coordinator = new MutationCoordinator();
    const prepared = coordinator.prepare("t-1", [
      new StructureIntegrationMutation(
        owner,
        { axis: "row", operation: "insert", start: 2, count: 1 },
        { amount: 1 }
      )
    ]);

    expect(() => coordinator.commit(prepared)).toThrow("partial publish");
    expect(root.count).toBe(2);
    expect(coordinator.state).toBe("idle");
  });
});
