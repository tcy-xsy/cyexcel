import { describe, expect, it } from "vitest";
import { CellTypeRegistry, ColumnManager, registerBuiltinCellTypes, type ColumnDefinition } from "../../../../src/core/cells";
import { HeaderManager } from "../../../../src/core/headers";
import { GridStructureOwner } from "../../../../src/core/integration";
import { SelectionModel } from "../../../../src/core/interaction";
import { LayoutEngine } from "../../../../src/core/layout";
import type { MutationPrepareContext, MutationRootDraftRegistry } from "../../../../src/core/mutations";

function createPrepareContext(): MutationPrepareContext {
  const roots = new WeakMap<object, unknown>();
  const registry: MutationRootDraftRegistry = {
    get: <TRoot>(owner: object, fallback: () => TRoot): TRoot => {
      if (!roots.has(owner)) {
        roots.set(owner, fallback());
      }
      return roots.get(owner) as TRoot;
    },
    set: <TRoot>(owner: object, root: TRoot): void => {
      roots.set(owner, root);
    }
  };
  return { roots: registry } as MutationPrepareContext;
}

function createOwner(columns: readonly ColumnDefinition[]) {
  const registry = new CellTypeRegistry();
  registerBuiltinCellTypes(registry);
  const columnManager = new ColumnManager({ columns, registry });
  const headerManager = new HeaderManager({ columns });
  const layout = new LayoutEngine({
    rowCount: 3,
    columnCount: columns.length,
    rowHeight: 20,
    columnWidths: columns.map((column) => column.width ?? 120)
  });
  const selection = new SelectionModel({ rowCount: 3, columnCount: columns.length });
  let liveColumns = columns;
  const owner = new GridStructureOwner({
    getColumns: () => liveColumns,
    replaceColumns: (next) => { liveColumns = next; },
    columnManager,
    headerManager,
    layout,
    selection
  });
  return { owner, columnManager, headerManager, layout, selection, getColumns: () => liveColumns };
}

describe("GridStructureOwner", () => {
  it("prepares all column roots without touching live owners", () => {
    const fixture = createOwner([{ width: 30 }, { width: 40 }]);
    fixture.selection.selectColumn(1);
    const before = fixture.owner.getRoot();
    const prepared = fixture.owner.prepare(
      before,
      { axis: "column", operation: "insert", start: 1, count: 1 },
      { columns: [{ width: 50 }], columnWidths: [50] },
      createPrepareContext()
    );

    expect(fixture.owner.isCurrent(before)).toBe(true);
    expect(prepared.nextRoot.columns.map((column) => column.width)).toEqual([30, 50, 40]);
    expect(prepared.nextRoot.runtimeColumns.map((column) => column.width)).toEqual([30, 50, 40]);
    expect(prepared.nextRoot.headers.headers.map((header) => header.title)).toEqual(["A", "B", "C"]);
    expect(prepared.nextRoot.layout.columnAxis.copyColumnWidths()).toEqual([30, 50, 40]);
    expect(prepared.nextRoot.selection.active.column).toBe(2);
  });

  it("prepares a declarative inverse that restores deleted column state", () => {
    const fixture = createOwner([{ width: 30 }, { width: 40 }]);
    const before = fixture.owner.getRoot();
    const forward = fixture.owner.prepare(
      before,
      { axis: "column", operation: "delete", start: 0, count: 1 },
      {},
      createPrepareContext()
    );
    fixture.owner.replaceRoot(forward.nextRoot);

    const inverse = fixture.owner.prepare(
      forward.nextRoot,
      forward.inverseChange,
      forward.inversePayload,
      createPrepareContext()
    );
    expect(inverse.nextRoot.columns.map((column) => column.width)).toEqual([30, 40]);
    expect(inverse.nextRoot.layout.columnAxis.copyColumnWidths()).toEqual([30, 40]);
    expect(inverse.nextRoot.selection.active.column).toBe(1);
    expect(inverse.nextRoot.layout.version).toBeGreaterThan(forward.nextRoot.layout.version);
    fixture.owner.replaceRoot(inverse.nextRoot);
    expect(fixture.getColumns().map((column) => column.width)).toEqual([30, 40]);
  });

  it("replays history against the current selection instead of restoring a captured selection root", () => {
    const fixture = createOwner([{ width: 30 }, { width: 40 }]);
    fixture.selection.selectColumn(1);
    const forward = fixture.owner.prepare(
      fixture.owner.getRoot(),
      { axis: "column", operation: "insert", start: 0, count: 1 },
      { columns: [{ width: 50 }], columnWidths: [50] },
      createPrepareContext()
    );
    fixture.owner.replaceRoot(forward.nextRoot);

    fixture.selection.selectColumn(0);
    const inverse = fixture.owner.prepare(
      fixture.owner.getRoot(),
      forward.inverseChange,
      forward.inversePayload,
      createPrepareContext()
    );

    expect(inverse.nextRoot.selection.active.column).toBe(0);
    expect(inverse.nextRoot.selection.kind).toBe("column");
  });

  it("keeps later width changes on surviving columns when replaying an insert inverse", () => {
    const fixture = createOwner([{ width: 30 }, { width: 40 }]);
    const forward = fixture.owner.prepare(
      fixture.owner.getRoot(),
      { axis: "column", operation: "insert", start: 1, count: 1 },
      { columns: [{ width: 50 }], columnWidths: [50] },
      createPrepareContext()
    );
    fixture.owner.replaceRoot(forward.nextRoot);
    fixture.layout.setColumnWidth(2, 75);

    const inverse = fixture.owner.prepare(
      fixture.owner.getRoot(),
      forward.inverseChange,
      forward.inversePayload,
      createPrepareContext()
    );

    expect(inverse.nextRoot.layout.columnAxis.copyColumnWidths()).toEqual([30, 75]);
    expect(inverse.nextRoot.layout.version).toBe(fixture.layout.version + 1);
  });

  it("canonicalizes delete history so captured restore data exists only on the inverse insert", () => {
    const fixture = createOwner([{ width: 30 }, { width: 40 }]);
    const forward = fixture.owner.prepare(
      fixture.owner.getRoot(),
      { axis: "column", operation: "delete", start: 0, count: 1 },
      { columns: [{ }], columnWidths: [999] },
      createPrepareContext()
    );

    expect(forward.forwardPayload).toEqual({});
    expect(forward.inversePayload.columns?.map((column) => column.width)).toEqual([30]);
    expect(forward.inversePayload.columnWidths).toEqual([30]);
  });

  it("increments layout version on every semantic undo and redo prepare", () => {
    const fixture = createOwner([{ width: 30 }, { width: 40 }]);
    const forward = fixture.owner.prepare(
      fixture.owner.getRoot(),
      { axis: "row", operation: "insert", start: 1, count: 1 },
      {},
      createPrepareContext()
    );
    fixture.owner.replaceRoot(forward.nextRoot);
    const inverse = fixture.owner.prepare(
      fixture.owner.getRoot(),
      forward.inverseChange,
      forward.inversePayload,
      createPrepareContext()
    );
    fixture.owner.replaceRoot(inverse.nextRoot);
    const redo = fixture.owner.prepare(
      fixture.owner.getRoot(),
      { axis: "row", operation: "insert", start: 1, count: 1 },
      forward.forwardPayload ?? {},
      createPrepareContext()
    );

    expect([
      fixture.layout.root.version - 1,
      fixture.layout.root.version,
      redo.nextRoot.layout.version
    ]).toEqual([1, 2, 3]);
  });

  it("captures deleted row geometry as a persistent fragment for undo", () => {
    const fixture = createOwner([{ width: 30 }]);
    fixture.layout.replaceRoot(fixture.layout.prepareSetRowHeights(new Map([[1, 47]])));
    const before = fixture.owner.getRoot();
    const forward = fixture.owner.prepare(
      before,
      { axis: "row", operation: "delete", start: 1, count: 1 },
      {},
      createPrepareContext()
    );

    expect(forward.inversePayload.rowMetricsFragment?.exportDenseValues()).toEqual(
      new Float64Array([47])
    );
    expect(forward.inversePayload.rowHeights).toBeUndefined();
    // 来源不再由这条载荷搬：`rowData` 是注册侧表，`h`/`hp` 跟着 `restore-rows`
    // 一起回来（ADR 0004 阶段 4B T8）。

    fixture.owner.replaceRoot(forward.nextRoot);
    const inverse = fixture.owner.prepare(
      forward.nextRoot,
      forward.inverseChange,
      forward.inversePayload,
      createPrepareContext()
    );
    expect(inverse.nextRoot.layout.rowMetrics.exportDenseValues()).toEqual(
      new Float64Array([20, 47, 20])
    );
  });
});
