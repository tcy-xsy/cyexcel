import { describe, expect, it } from "vitest";

import { CyExcelError, CyExcelErrorCode } from "../../../../src/core/errors";
import { LayoutEngine } from "../../../../src/core/layout";
import {
  MutationCoordinator,
  RowHeightMutation,
  type Mutation
} from "../../../../src/core/mutations";

function createLayout(): LayoutEngine {
  return new LayoutEngine({
    rowCount: 4,
    columnCount: 3,
    rowHeight: 20,
    columnWidths: [50, 60, 70]
  });
}

function createFailingMutation(message: string): Mutation {
  const mutation: Mutation = {
    type: "failing-participant",
    changes: [],
    prepare: () => ({
      hasEffect: true,
      forward: mutation,
      inverse: mutation,
      estimatedBytes: 0,
      commit: () => {
        throw new Error(message);
      },
      rollback: () => {},
      dispose: () => {}
    })
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

describe("RowHeightMutation", () => {
  it("prepares and commits several rows with one Layout root swap", () => {
    const layout = createLayout();
    const coordinator = new MutationCoordinator();
    const initialRoot = layout.root;
    const prepared = coordinator.prepare("row-height-batch", [
      new RowHeightMutation(layout, [
        { row: 2, height: 44, provenance: "manual" },
        { row: 0, height: 28, provenance: "manual" }
      ])
    ]);

    expect(layout.root).toBe(initialRoot);
    expect(layout.version).toBe(0);

    const participant = prepared.participants[0]!;
    const forward = participant.forward as RowHeightMutation;
    const inverse = participant.inverse as RowHeightMutation;
    expect(forward.updates).toEqual([
      { row: 2, height: 44, provenance: "manual" },
      { row: 0, height: 28, provenance: "manual" }
    ]);
    expect(inverse.updates).toEqual([
      { row: 2, height: 20, provenance: "default" },
      { row: 0, height: 20, provenance: "default" }
    ]);
    expect(Object.isFrozen(forward.updates)).toBe(true);
    expect(forward.updates.every((update) => Object.isFrozen(update))).toBe(true);
    expect(Object.isFrozen(inverse.updates)).toBe(true);

    const receipt = coordinator.commit(prepared);

    expect(layout.root).not.toBe(initialRoot);
    expect(layout.getRowHeight(0)).toBe(28);
    expect(layout.getRowHeight(2)).toBe(44);
    expect(layout.version).toBe(1);
    expect(receipt.patches).toEqual([
      {
        type: "row-layout",
        rows: [2, 0],
        range: { startRow: 0, startColumn: 0, endRow: 3, endColumn: 2 }
      }
    ]);
    const rowLayoutPatch = receipt.patches[0];
    expect(rowLayoutPatch?.type).toBe("row-layout");
    if (rowLayoutPatch?.type !== "row-layout") {
      throw new Error("expected row-layout patch");
    }
    expect(Object.isFrozen(rowLayoutPatch)).toBe(true);
    expect(Object.isFrozen(rowLayoutPatch.rows)).toBe(true);
    expect(Object.isFrozen(rowLayoutPatch.range)).toBe(true);
    expect(receipt.changeSet.changes).toHaveLength(1);
    expect(receipt.afterRevision).toBe(1);

    coordinator.execute("undo-row-height-batch", receipt.inverse);
    expect(layout.getRowHeight(0)).toBe(20);
    expect(layout.getRowHeight(2)).toBe(20);
    expect(layout.version).toBe(2);
  });

  it("snapshots input and rejects duplicate or invalid updates", () => {
    const layout = createLayout();
    const updates = [{ row: 1, height: 32, provenance: "manual" as const }];
    const mutation = new RowHeightMutation(layout, updates);

    updates[0]!.height = 90;
    updates.push({ row: 2, height: 48, provenance: "manual" });
    expect(mutation.updates).toEqual([{ row: 1, height: 32, provenance: "manual" }]);
    expect(Object.isFrozen(mutation)).toBe(true);

    expectErrorCode(
      () => new RowHeightMutation(layout, [
        { row: 1, height: 30, provenance: "manual" },
        { row: 1, height: 40, provenance: "manual" }
      ]),
      CyExcelErrorCode.InvalidOptions
    );
    expectErrorCode(
      () => new RowHeightMutation(layout, [{ row: -1, height: 30, provenance: "manual" }]),
      CyExcelErrorCode.InvalidOptions
    );
    expectErrorCode(
      () => new RowHeightMutation(layout, [
        { row: 1, height: Number.NaN, provenance: "manual" }
      ]),
      CyExcelErrorCode.InvalidOptions
    );
  });

  it("composes consecutive batches through the transaction-local Layout root", () => {
    const layout = createLayout();
    const coordinator = new MutationCoordinator();
    const prepared = coordinator.prepare("composed-row-height-batches", [
      new RowHeightMutation(layout, [{ row: 0, height: 30, provenance: "auto" }]),
      new RowHeightMutation(layout, [
        { row: 0, height: 35, provenance: "auto" },
        { row: 3, height: 50, provenance: "manual" }
      ])
    ]);

    const secondInverse = prepared.participants[1]!.inverse as RowHeightMutation;
    expect(secondInverse.updates).toEqual([
      { row: 0, height: 30, provenance: "auto" },
      { row: 3, height: 20, provenance: "default" }
    ]);

    const receipt = coordinator.commit(prepared);
    expect(layout.getRowHeight(0)).toBe(35);
    expect(layout.getRowHeight(3)).toBe(50);
    expect(layout.version).toBe(2);
    expect(receipt.afterRevision).toBe(1);

    coordinator.execute("undo-composed-row-height-batches", receipt.inverse);
    expect(layout.getRowHeight(0)).toBe(20);
    expect(layout.getRowHeight(3)).toBe(20);
    expect(layout.version).toBe(4);
  });

  it("publishes patches and inverse updates only for rows that actually changed", () => {
    const layout = createLayout();
    const coordinator = new MutationCoordinator();
    const prepared = coordinator.prepare("partially-equal-row-height-batch", [
      new RowHeightMutation(layout, [
        { row: 0, height: 20, provenance: "default" },
        { row: 1, height: 36, provenance: "manual" },
        { row: 3, height: 20, provenance: "default" }
      ])
    ]);
    const participant = prepared.participants[0]!;

    expect((participant.forward as RowHeightMutation).updates).toEqual([
      { row: 1, height: 36, provenance: "manual" }
    ]);
    expect((participant.inverse as RowHeightMutation).updates).toEqual([
      { row: 1, height: 20, provenance: "default" }
    ]);

    const receipt = coordinator.commit(prepared);
    expect(receipt.patches).toEqual([{
      type: "row-layout",
      rows: [1],
      range: { startRow: 1, startColumn: 0, endRow: 3, endColumn: 2 }
    }]);
    expect(layout.version).toBe(1);
  });

  it("rolls the whole batch back when a later participant throws", () => {
    const layout = createLayout();
    const coordinator = new MutationCoordinator();
    const initialRoot = layout.root;
    const prepared = coordinator.prepare("rollback-row-height-batch", [
      new RowHeightMutation(layout, [
        { row: 0, height: 28, provenance: "manual" },
        { row: 2, height: 44, provenance: "manual" }
      ]),
      createFailingMutation("later participant failed")
    ]);

    expect(() => coordinator.commit(prepared)).toThrow("later participant failed");
    expect(layout.root).toBe(initialRoot);
    expect(layout.version).toBe(0);
    expect(coordinator.revisionGate.revision).toBe(0);
  });

  it("rejects a stale prepared batch without overwriting the newer root", () => {
    const layout = createLayout();
    const coordinator = new MutationCoordinator();
    const prepared = coordinator.prepare("stale-row-height-batch", [
      new RowHeightMutation(layout, [{ row: 1, height: 48, provenance: "auto" }])
    ]);

    layout.setRowHeight(1, 31);
    const newerRoot = layout.root;

    expectErrorCode(() => coordinator.commit(prepared), CyExcelErrorCode.MutationConflict);
    expect(layout.root).toBe(newerRoot);
    expect(layout.getRowHeight(1)).toBe(31);
    expect(coordinator.revisionGate.revision).toBe(0);
  });

  it("uses a safe no-op participant when every height is already current", () => {
    const layout = createLayout();
    const coordinator = new MutationCoordinator();
    const initialRoot = layout.root;
    const prepared = coordinator.prepare("no-op-row-height-batch", [
      new RowHeightMutation(layout, [
        { row: 0, height: 20, provenance: "default" },
        { row: 2, height: 20, provenance: "default" }
      ])
    ]);

    expect(prepared.hasEffect).toBe(false);
    expect(prepared.participants).toEqual([]);
    expect(prepared.changes).toEqual([]);
    const receipt = coordinator.commit(prepared);

    expect(layout.root).toBe(initialRoot);
    expect(layout.version).toBe(0);
    expect(receipt.noop).toBe(true);
    expect(receipt.forward).toEqual([]);
    expect(receipt.inverse).toEqual([]);
    expect(receipt.patches).toEqual([]);
    expect(receipt.historyBytes).toBe(0);
    expect(receipt.afterRevision).toBe(0);
  });

  it("rejects an out-of-range row during transaction prepare", () => {
    const layout = createLayout();
    const coordinator = new MutationCoordinator();

    expect(() => coordinator.prepare("invalid-row-height-batch", [
      new RowHeightMutation(layout, [{ row: 4, height: 30, provenance: "auto" }])
    ])).toThrow(CyExcelError);
    expect(layout.version).toBe(0);
    expect(coordinator.state).toBe("idle");
  });

  it("treats a protected auto update against a manual row as a root-and-patch no-op", () => {
    const layout = createLayout();
    layout.setRowHeight(1, 33);
    const manualRoot = layout.root;
    const coordinator = new MutationCoordinator();

    const receipt = coordinator.execute("protected-auto-row-height", [
      new RowHeightMutation(
        layout,
        [{ row: 1, height: 60, provenance: "auto" }],
        "preserve-manual"
      )
    ]);

    expect(receipt.patches).toEqual([]);
    expect(receipt.noop).toBe(true);
    expect(receipt.afterRevision).toBe(0);
    expect(coordinator.revisionGate.revision).toBe(0);
    expect(layout.root).toBe(manualRoot);
    expect(layout.getRowHeight(1)).toBe(33);
    expect(layout.getRowHeightProvenance(1)).toBe("manual");
  });

  it("records provenance-only changes and restores the prior owner exactly", () => {
    const layout = createLayout();
    const coordinator = new MutationCoordinator();
    const autoReceipt = coordinator.execute("auto-row-height", [
      new RowHeightMutation(layout, [{ row: 2, height: 40, provenance: "auto" }])
    ]);

    const manualReceipt = coordinator.execute("manual-lock", [
      new RowHeightMutation(layout, [{ row: 2, height: 40, provenance: "manual" }])
    ]);
    expect(manualReceipt.patches).toHaveLength(1);
    expect(layout.getRowHeightProvenance(2)).toBe("manual");

    coordinator.execute("undo-manual-lock", manualReceipt.inverse);
    expect(layout.getRowHeight(2)).toBe(40);
    expect(layout.getRowHeightProvenance(2)).toBe("auto");

    coordinator.execute("undo-auto-row-height", autoReceipt.inverse);
    expect(layout.getRowHeight(2)).toBe(20);
    expect(layout.getRowHeightProvenance(2)).toBe("default");
  });
});
