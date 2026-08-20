import { describe, expect, it } from "vitest";

import {
  CellTypeRegistry,
  ColumnManager,
  registerBuiltinCellTypes,
  type ColumnDefinition
} from "../../../../src/core/cells";
import { CyExcelError, CyExcelErrorCode } from "../../../../src/core/errors";
import { HeaderManager } from "../../../../src/core/headers";
import {
  GridStructureOwner,
  createGridStructurePayload
} from "../../../../src/core/integration";
import { SelectionModel } from "../../../../src/core/interaction";
import { LayoutEngine } from "../../../../src/core/layout";
import {
  LayoutRootMutation,
  MutationCoordinator,
  StructureIntegrationMutation,
  type Mutation
} from "../../../../src/core/mutations";

function createLayout(): LayoutEngine {
  return new LayoutEngine({
    rowCount: 3,
    columnCount: 3,
    rowHeight: 20,
    columnWidths: [50, 60, 70]
  });
}

function createGridFixture() {
  const columns: readonly ColumnDefinition[] = Object.freeze([
    Object.freeze({ id: "a", width: 30 }),
    Object.freeze({ id: "b", width: 40 })
  ]);
  const registry = new CellTypeRegistry();
  registerBuiltinCellTypes(registry);
  const columnManager = new ColumnManager({ columns, registry });
  const headerManager = new HeaderManager({ columns });
  const layout = new LayoutEngine({
    rowCount: 3,
    columnCount: columns.length,
    rowHeight: 20,
    columnWidths: [30, 40]
  });
  const selection = new SelectionModel({ rowCount: 3, columnCount: columns.length });
  selection.selectColumn(1);
  let liveColumns = columns;
  const owner = new GridStructureOwner({
    getColumns: () => liveColumns,
    replaceColumns: (next) => {
      liveColumns = next;
    },
    columnManager,
    headerManager,
    layout,
    selection
  });

  return {
    owner,
    columnManager,
    headerManager,
    layout,
    selection,
    getColumns: () => liveColumns
  };
}

function createColumnInsertMutation(owner: GridStructureOwner) {
  return new StructureIntegrationMutation(
    owner,
    { axis: "column", operation: "insert", start: 1, count: 1 },
    createGridStructurePayload({
      columns: [{ id: "x", width: 50 }],
      columnWidths: [50]
    })
  );
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

function expectErrorCode(action: () => unknown, code: string): void {
  try {
    action();
    throw new Error(`Expected CyExcelError with code ${code}`);
  } catch (error) {
    expect(error).toBeInstanceOf(CyExcelError);
    expect((error as CyExcelError).code).toBe(code);
  }
}

describe("LayoutRootMutation", () => {
  it("prepares a replacement root without writing the live layout", () => {
    const layout = createLayout();
    const coordinator = new MutationCoordinator();
    const initialRoot = layout.root;

    const prepared = coordinator.prepare("resize-row", [
      new LayoutRootMutation(layout, {
        type: "set-row-height",
        row: 1,
        height: 36,
        provenance: "manual"
      })
    ]);

    expect(layout.root).toBe(initialRoot);
    expect(layout.getRowHeight(1)).toBe(20);
    expect(layout.version).toBe(0);
    expect(coordinator.revisionGate.revision).toBe(0);
    expect(coordinator.state).toBe("prepared");

    const receipt = coordinator.commit(prepared);

    expect(layout.root).not.toBe(initialRoot);
    expect(layout.getRowHeight(1)).toBe(36);
    expect(layout.version).toBe(1);
    expect(receipt.beforeRevision).toBe(0);
    expect(receipt.afterRevision).toBe(1);
    expect(receipt.patches).toEqual([{
      type: "row-layout",
      rows: [1],
      range: { startRow: 1, startColumn: 0, endRow: 2, endColumn: 2 }
    }]);
    expect(coordinator.revisionGate.revision).toBe(1);
    expect(coordinator.state).toBe("idle");
  });

  it("stages consecutive mutations for the same layout against the preceding staged root", () => {
    const layout = createLayout();
    const coordinator = new MutationCoordinator();
    const initialRoot = layout.root;

    const prepared = coordinator.prepare("resize-layout", [
      new LayoutRootMutation(layout, { type: "set-column-width", column: 1, width: 95 }),
      new LayoutRootMutation(layout, {
        type: "set-row-height",
        row: 2,
        height: 44,
        provenance: "manual"
      })
    ]);

    expect(layout.root).toBe(initialRoot);
    expect(layout.getColumnWidth(1)).toBe(60);
    expect(layout.getRowHeight(2)).toBe(20);
    expect(layout.version).toBe(0);
    expect(coordinator.revisionGate.revision).toBe(0);

    const receipt = coordinator.commit(prepared);

    expect(layout.getColumnWidth(1)).toBe(95);
    expect(layout.getRowHeight(2)).toBe(44);
    expect(layout.version).toBe(2);
    expect(coordinator.revisionGate.revision).toBe(1);
    expect(receipt.patches.map((patch) => patch.type)).toEqual(["column-layout", "row-layout"]);
    expect(receipt.changeSet.modelVersion).toBe(1);
  });

  it("rolls back a committed layout participant when a later participant throws", () => {
    const layout = createLayout();
    const coordinator = new MutationCoordinator();
    const initialRoot = layout.root;
    const prepared = coordinator.prepare("rollback-layout", [
      new LayoutRootMutation(layout, {
        type: "set-row-height",
        row: 0,
        height: 42,
        provenance: "manual"
      }),
      createFailingMutation("later participant failed")
    ]);

    expect(() => coordinator.commit(prepared)).toThrow("later participant failed");

    expect(layout.root).toBe(initialRoot);
    expect(layout.getRowHeight(0)).toBe(20);
    expect(layout.version).toBe(0);
    expect(coordinator.revisionGate.revision).toBe(0);
    expect(coordinator.state).toBe("idle");
  });

  it("composes layout then structure mutations through the shared root registry", () => {
    const fixture = createGridFixture();
    const coordinator = new MutationCoordinator();
    const initialRoot = fixture.owner.getRoot();
    const prepared = coordinator.prepare("layout-then-structure", [
      new LayoutRootMutation(fixture.layout, {
        type: "set-column-width",
        column: 1,
        width: 75
      }),
      createColumnInsertMutation(fixture.owner)
    ]);

    expect(fixture.owner.isCurrent(initialRoot)).toBe(true);
    expect(fixture.getColumns()).toBe(initialRoot.columns);
    expect(fixture.columnManager.getColumns()).toBe(initialRoot.runtimeColumns);
    expect(fixture.headerManager.root).toBe(initialRoot.headers);
    expect(fixture.layout.root).toBe(initialRoot.layout);
    expect(fixture.selection.root).toBe(initialRoot.selection);

    const receipt = coordinator.commit(prepared);

    expect(fixture.getColumns().map((column) => column.id)).toEqual(["a", "x", "b"]);
    expect(fixture.columnManager.getColumns().map((column) => column.id)).toEqual(["a", "x", "b"]);
    expect(fixture.headerManager.getHeaders().map((header) => header.title)).toEqual(["a", "x", "b"]);
    expect(fixture.layout.getColumnWidthsSnapshot()).toEqual([30, 50, 75]);
    expect(fixture.layout.columnCount).toBe(3);
    expect(fixture.layout.version).toBe(2);
    expect(fixture.selection.bounds.columnCount).toBe(3);
    expect(fixture.selection.snapshot.active.column).toBe(2);
    expect(receipt.afterRevision).toBe(1);
  });

  it("composes structure then layout mutations through the shared root registry", () => {
    const fixture = createGridFixture();
    const coordinator = new MutationCoordinator();
    const initialRoot = fixture.owner.getRoot();
    const prepared = coordinator.prepare("structure-then-layout", [
      createColumnInsertMutation(fixture.owner),
      new LayoutRootMutation(fixture.layout, {
        type: "set-column-width",
        column: 2,
        width: 75
      })
    ]);

    expect(fixture.owner.isCurrent(initialRoot)).toBe(true);
    expect(fixture.getColumns()).toBe(initialRoot.columns);
    expect(fixture.columnManager.getColumns()).toBe(initialRoot.runtimeColumns);
    expect(fixture.headerManager.root).toBe(initialRoot.headers);
    expect(fixture.layout.root).toBe(initialRoot.layout);
    expect(fixture.selection.root).toBe(initialRoot.selection);

    const receipt = coordinator.commit(prepared);

    expect(fixture.getColumns().map((column) => column.id)).toEqual(["a", "x", "b"]);
    expect(fixture.columnManager.getColumns().map((column) => column.id)).toEqual(["a", "x", "b"]);
    expect(fixture.headerManager.getHeaders().map((header) => header.title)).toEqual(["a", "x", "b"]);
    expect(fixture.layout.getColumnWidthsSnapshot()).toEqual([30, 50, 75]);
    expect(fixture.layout.columnCount).toBe(3);
    expect(fixture.layout.version).toBe(2);
    expect(fixture.selection.bounds.columnCount).toBe(3);
    expect(fixture.selection.snapshot.active.column).toBe(2);
    expect(receipt.afterRevision).toBe(1);
  });

  it("rolls every real grid owner back after a later participant fails", () => {
    const fixture = createGridFixture();
    const coordinator = new MutationCoordinator();
    const initialRoot = fixture.owner.getRoot();
    const prepared = coordinator.prepare("rollback-mixed-roots", [
      new LayoutRootMutation(fixture.layout, {
        type: "set-column-width",
        column: 1,
        width: 75
      }),
      createColumnInsertMutation(fixture.owner),
      createFailingMutation("mixed transaction failed")
    ]);

    expect(fixture.owner.isCurrent(initialRoot)).toBe(true);
    expect(() => coordinator.commit(prepared)).toThrow("mixed transaction failed");

    expect(fixture.owner.isCurrent(initialRoot)).toBe(true);
    expect(fixture.getColumns()).toBe(initialRoot.columns);
    expect(fixture.columnManager.getColumns()).toBe(initialRoot.runtimeColumns);
    expect(fixture.headerManager.root).toBe(initialRoot.headers);
    expect(fixture.layout.root).toBe(initialRoot.layout);
    expect(fixture.selection.root).toBe(initialRoot.selection);
    expect(fixture.getColumns().map((column) => column.id)).toEqual(["a", "b"]);
    expect(fixture.layout.getColumnWidthsSnapshot()).toEqual([30, 40]);
    expect(fixture.layout.columnCount).toBe(2);
    expect(fixture.layout.version).toBe(0);
    expect(fixture.selection.bounds.columnCount).toBe(2);
    expect(fixture.selection.snapshot.active.column).toBe(1);
    expect(coordinator.revisionGate.revision).toBe(0);
    expect(coordinator.state).toBe("idle");
  });

  it("rejects a stale prepared root without overwriting the newer live root", () => {
    const layout = createLayout();
    const coordinator = new MutationCoordinator();
    const prepared = coordinator.prepare("stale-layout", [
      new LayoutRootMutation(layout, {
        type: "set-row-height",
        row: 1,
        height: 48,
        provenance: "manual"
      })
    ]);

    layout.setRowHeight(1, 31);
    const newerRoot = layout.root;

    expectErrorCode(() => coordinator.commit(prepared), CyExcelErrorCode.MutationConflict);

    expect(layout.root).toBe(newerRoot);
    expect(layout.getRowHeight(1)).toBe(31);
    expect(layout.version).toBe(1);
    expect(coordinator.revisionGate.revision).toBe(0);
    expect(coordinator.state).toBe("idle");
  });

  describe("reorder-rows action", () => {
    // Deliberately asymmetric: [2, 0, 1] is not its own inverse, so a
    // direction mix-up between order and invertRowOrder() cannot pass.
    const order = () => Int32Array.from([2, 0, 1]);

    function createTallRows(): LayoutEngine {
      const layout = createLayout();
      layout.setRowHeight(0, 60);
      layout.setRowHeight(2, 44);
      return layout;
    }

    it("carries each row's height and provenance to its new position", () => {
      const layout = createTallRows();
      const coordinator = new MutationCoordinator();
      const initialRoot = layout.root;

      const prepared = coordinator.prepare("reorder", [
        new LayoutRootMutation(layout, { type: "reorder-rows", order: order() })
      ]);

      expect(layout.root).toBe(initialRoot);
      expect([0, 1, 2].map((row) => layout.getRowHeight(row))).toEqual([60, 20, 44]);

      const receipt = coordinator.commit(prepared);

      // order[newPosition] = oldPosition, so new 0 <- old 2, new 1 <- old 0.
      expect([0, 1, 2].map((row) => layout.getRowHeight(row))).toEqual([44, 60, 20]);
      expect([0, 1, 2].map((row) => layout.getRowHeightProvenance(row)))
        .toEqual(["manual", "manual", "default"]);
      expect(receipt.patches).toEqual([{ type: "row-reorder", order: Int32Array.from([2, 0, 1]) }]);
    });

    it("inverts to the inverse permutation, restoring every height", () => {
      const layout = createTallRows();
      const coordinator = new MutationCoordinator();

      const receipt = coordinator.commit(
        coordinator.prepare("reorder", [
          new LayoutRootMutation(layout, { type: "reorder-rows", order: order() })
        ])
      );
      coordinator.commit(coordinator.prepare("undo-reorder", [...receipt.inverse]));

      expect([0, 1, 2].map((row) => layout.getRowHeight(row))).toEqual([60, 20, 44]);
      expect([0, 1, 2].map((row) => layout.getRowHeightProvenance(row)))
        .toEqual(["manual", "default", "manual"]);
    });

    it("stages a second reorder against the preceding staged root, not the live one", () => {
      const layout = createTallRows();
      const coordinator = new MutationCoordinator();

      const prepared = coordinator.prepare("reorder-twice", [
        new LayoutRootMutation(layout, { type: "reorder-rows", order: order() }),
        new LayoutRootMutation(layout, { type: "reorder-rows", order: order() })
      ]);
      coordinator.commit(prepared);

      // [2,0,1] twice composes to [1,2,0]: new 0 <- old 1, new 1 <- old 2.
      expect([0, 1, 2].map((row) => layout.getRowHeight(row))).toEqual([20, 44, 60]);
    });
  });
});
