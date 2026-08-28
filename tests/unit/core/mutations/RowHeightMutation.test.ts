import { describe, expect, it } from "vitest";

import { CyExcelError, CyExcelErrorCode } from "../../../../src/core/errors";
import { LayoutEngine } from "../../../../src/core/layout";
import type { ModelRevision } from "../../../../src/core/model";
import {
  MutationCoordinator,
  RowHeightMutation,
  type Mutation,
  type MutationPlanningResources
} from "../../../../src/core/mutations";
import { createSheetPlanningResources } from "../../../../src/core/sheet";
import type { SheetModel } from "../../../../src/core/sheet/SheetModel";
import { createSheetFromRows } from "../../../helpers/createSheetModel";
import { rowHeightProvenanceOf } from "../../../helpers/rowHeightAuthority";

/**
 * 行高的权威在模型里（ADR 0004 阶段 4B T8），所以这里的夹具是「一张表 + 一个布局」
 * 两件东西，而不是从前的一个布局：一条行高变更要同时写 `rowData` 和 `RowMetrics`，
 * 少了哪一半都 prepare 不出来。
 */
function createFixture(): { layout: LayoutEngine; sheet: SheetModel } {
  const sheet = createSheetFromRows(
    [{ a: "1" }, { a: "2" }, { a: "3" }, { a: "4" }],
    ["a"]
  );
  const layout = new LayoutEngine({
    rowCount: 4,
    columnCount: 3,
    rowHeight: 20,
    columnWidths: [50, 60, 70]
  });
  return { layout, sheet };
}

/** 影子模型的来源，与 `TableStore.beginMutationSession` 用的是同一个工厂。 */
function resourcesFor(sheet: SheetModel): (baseRevision: ModelRevision) => MutationPlanningResources {
  return (baseRevision) => createSheetPlanningResources(sheet, baseRevision);
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
    const { layout, sheet } = createFixture();
    const coordinator = new MutationCoordinator();
    const initialRoot = layout.root;
    const prepared = coordinator.prepare("row-height-batch", [
      new RowHeightMutation(layout, sheet, [
        { row: 2, height: 44, provenance: "manual" },
        { row: 0, height: 28, provenance: "manual" }
      ])
    ], resourcesFor(sheet));

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

    coordinator.execute("undo-row-height-batch", receipt.inverse, resourcesFor(sheet));
    expect(layout.getRowHeight(0)).toBe(20);
    expect(layout.getRowHeight(2)).toBe(20);
    expect(layout.version).toBe(2);
  });

  it("snapshots input and rejects duplicate or invalid updates", () => {
    const { layout, sheet } = createFixture();
    const updates = [{ row: 1, height: 32, provenance: "manual" as const }];
    const mutation = new RowHeightMutation(layout, sheet, updates);

    updates[0]!.height = 90;
    updates.push({ row: 2, height: 48, provenance: "manual" });
    expect(mutation.updates).toEqual([{ row: 1, height: 32, provenance: "manual" }]);
    expect(Object.isFrozen(mutation)).toBe(true);

    expectErrorCode(
      () => new RowHeightMutation(layout, sheet, [
        { row: 1, height: 30, provenance: "manual" },
        { row: 1, height: 40, provenance: "manual" }
      ]),
      CyExcelErrorCode.InvalidOptions
    );
    expectErrorCode(
      () => new RowHeightMutation(layout, sheet, [{ row: -1, height: 30, provenance: "manual" }]),
      CyExcelErrorCode.InvalidOptions
    );
    expectErrorCode(
      () => new RowHeightMutation(layout, sheet, [
        { row: 1, height: Number.NaN, provenance: "manual" }
      ]),
      CyExcelErrorCode.InvalidOptions
    );
  });

  it("composes consecutive batches through the transaction-local Layout root", () => {
    const { layout, sheet } = createFixture();
    const coordinator = new MutationCoordinator();
    const prepared = coordinator.prepare("composed-row-height-batches", [
      new RowHeightMutation(layout, sheet, [{ row: 0, height: 30, provenance: "auto" }]),
      new RowHeightMutation(layout, sheet, [
        { row: 0, height: 35, provenance: "auto" },
        { row: 3, height: 50, provenance: "manual" }
      ])
    ], resourcesFor(sheet));

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

    coordinator.execute("undo-composed-row-height-batches", receipt.inverse, resourcesFor(sheet));
    expect(layout.getRowHeight(0)).toBe(20);
    expect(layout.getRowHeight(3)).toBe(20);
    expect(layout.version).toBe(4);
  });

  it("publishes patches and inverse updates only for rows that actually changed", () => {
    const { layout, sheet } = createFixture();
    const coordinator = new MutationCoordinator();
    const prepared = coordinator.prepare("partially-equal-row-height-batch", [
      new RowHeightMutation(layout, sheet, [
        { row: 0, height: 20, provenance: "default" },
        { row: 1, height: 36, provenance: "manual" },
        { row: 3, height: 20, provenance: "default" }
      ])
    ], resourcesFor(sheet));
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
    const { layout, sheet } = createFixture();
    const coordinator = new MutationCoordinator();
    const initialRoot = layout.root;
    const prepared = coordinator.prepare("rollback-row-height-batch", [
      new RowHeightMutation(layout, sheet, [
        { row: 0, height: 28, provenance: "manual" },
        { row: 2, height: 44, provenance: "manual" }
      ]),
      createFailingMutation("later participant failed")
    ], resourcesFor(sheet));

    expect(() => coordinator.commit(prepared)).toThrow("later participant failed");
    expect(layout.root).toBe(initialRoot);
    expect(layout.version).toBe(0);
    expect(coordinator.revisionGate.revision).toBe(0);
  });

  it("rejects a stale prepared batch without overwriting the newer root", () => {
    const { layout, sheet } = createFixture();
    const coordinator = new MutationCoordinator();
    const prepared = coordinator.prepare("stale-row-height-batch", [
      new RowHeightMutation(layout, sheet, [{ row: 1, height: 48, provenance: "auto" }])
    ], resourcesFor(sheet));

    layout.setRowHeight(1, 31);
    const newerRoot = layout.root;

    expectErrorCode(() => coordinator.commit(prepared), CyExcelErrorCode.MutationConflict);
    expect(layout.root).toBe(newerRoot);
    expect(layout.getRowHeight(1)).toBe(31);
    expect(coordinator.revisionGate.revision).toBe(0);
  });

  it("uses a safe no-op participant when every height is already current", () => {
    const { layout, sheet } = createFixture();
    const coordinator = new MutationCoordinator();
    const initialRoot = layout.root;
    const prepared = coordinator.prepare("no-op-row-height-batch", [
      new RowHeightMutation(layout, sheet, [
        { row: 0, height: 20, provenance: "default" },
        { row: 2, height: 20, provenance: "default" }
      ])
    ], resourcesFor(sheet));

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
    const { layout, sheet } = createFixture();
    const coordinator = new MutationCoordinator();

    expect(() => coordinator.prepare("invalid-row-height-batch", [
      new RowHeightMutation(layout, sheet, [{ row: 4, height: 30, provenance: "auto" }])
    ], resourcesFor(sheet))).toThrow(CyExcelError);
    expect(layout.version).toBe(0);
    expect(coordinator.state).toBe("idle");
  });

  it("treats a protected auto update against a manual row as a root-and-patch no-op", () => {
    const { layout, sheet } = createFixture();
    const coordinator = new MutationCoordinator();
    // 手动那一次也要走变更：`preserve-manual` 问的是**权威**里的来源，而
    // `layout.setRowHeight()` 只动索引，模型那边仍然是 default（设计稿 §4.2）。
    coordinator.execute("manual-row-height", [
      new RowHeightMutation(layout, sheet, [{ row: 1, height: 33, provenance: "manual" }])
    ], resourcesFor(sheet));
    const manualRoot = layout.root;
    const baseRevision = coordinator.revisionGate.revision;

    const receipt = coordinator.execute("protected-auto-row-height", [
      new RowHeightMutation(
        layout,
        sheet,
        [{ row: 1, height: 60, provenance: "auto" }],
        "preserve-manual"
      )
    ], resourcesFor(sheet));

    expect(receipt.patches).toEqual([]);
    expect(receipt.noop).toBe(true);
    expect(receipt.afterRevision).toBe(baseRevision);
    expect(coordinator.revisionGate.revision).toBe(baseRevision);
    expect(layout.root).toBe(manualRoot);
    expect(layout.getRowHeight(1)).toBe(33);
    expect(rowHeightProvenanceOf(sheet, 1)).toBe("manual");
  });

  it("records provenance-only changes and restores the prior owner exactly", () => {
    const { layout, sheet } = createFixture();
    const coordinator = new MutationCoordinator();
    const autoReceipt = coordinator.execute("auto-row-height", [
      new RowHeightMutation(layout, sheet, [{ row: 2, height: 40, provenance: "auto" }])
    ], resourcesFor(sheet));

    const manualReceipt = coordinator.execute("manual-lock", [
      new RowHeightMutation(layout, sheet, [{ row: 2, height: 40, provenance: "manual" }])
    ], resourcesFor(sheet));
    expect(manualReceipt.patches).toHaveLength(1);
    expect(rowHeightProvenanceOf(sheet, 2)).toBe("manual");

    coordinator.execute("undo-manual-lock", manualReceipt.inverse, resourcesFor(sheet));
    expect(layout.getRowHeight(2)).toBe(40);
    expect(rowHeightProvenanceOf(sheet, 2)).toBe("auto");

    coordinator.execute("undo-auto-row-height", autoReceipt.inverse, resourcesFor(sheet));
    expect(layout.getRowHeight(2)).toBe(20);
    expect(rowHeightProvenanceOf(sheet, 2)).toBe("default");
  });
});
