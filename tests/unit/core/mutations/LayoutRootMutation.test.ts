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
import type { ModelRevision } from "../../../../src/core/model";
import {
  LayoutRootMutation,
  MutationCoordinator,
  StructureIntegrationMutation,
  type Mutation,
  type MutationPlanningResources
} from "../../../../src/core/mutations";
import { createSheetPlanningResources } from "../../../../src/core/sheet";
import type { SheetModel } from "../../../../src/core/sheet/SheetModel";
import { createSheetFromRows } from "../../../helpers/createSheetModel";
import { rowHeightProvenanceOf } from "../../../helpers/rowHeightAuthority";

/**
 * 行高的权威在模型里（ADR 0004 阶段 4B T8）、列宽与隐藏的权威也在（T8b），所以夹具是
 * 「一张表 + 一个布局」：`set-row-height` 要同时写 `rowData` 和 `RowMetrics`，
 * `set-column-width` 要同时写 `columnData` 和 `ColumnAxisSnapshot`。
 *
 * 表配三列而不是一列：写第 c 列的元数据会把派生的列数撑到 c+1，而「值写不改坐标空间」
 * 那道闸要求写完还是原来那么宽（spec §3.3）。
 */
function createFixture(): { layout: LayoutEngine; sheet: SheetModel } {
  const sheet = createSheetFromRows([{ a: "1" }, { a: "2" }, { a: "3" }], ["a", "b", "c"]);
  const layout = new LayoutEngine({
    rowCount: 3,
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

function createGridFixture() {
  const columns: readonly ColumnDefinition[] = Object.freeze([
    Object.freeze({ width: 30 }),
    Object.freeze({ width: 40 })
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
    // 列宽的权威也在模型里了（ADR 0004 阶段 4B T8b），所以这张表是真要用的：
    // `set-column-width` 会同时写 `columnData[c].w`。
    //
    // 给三列而不是两列，是因为下面两条组合用例只挂了结构参与者、没挂数据侧那条
    // `insert-columns`：插入后布局有三列，而模型的坐标空间还停在原地，往第 2 列写
    // 元数据会被「值写不改坐标空间」那道闸挡下来（spec §3.3）。生产里两条是一起排的，
    // 这里只测根注册表的组合，就把表先配够宽。
    sheet: createSheetFromRows([{ a: "1" }, { a: "2" }, { a: "3" }], ["a", "b", "c"]),
    selection,
    getColumns: () => liveColumns
  };
}

function createColumnInsertMutation(owner: GridStructureOwner) {
  return new StructureIntegrationMutation(
    owner,
    { axis: "column", operation: "insert", start: 1, count: 1 },
    createGridStructurePayload({
      columns: [{ width: 50 }],
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
    const { layout, sheet } = createFixture();
    const coordinator = new MutationCoordinator();
    const initialRoot = layout.root;

    const prepared = coordinator.prepare("resize-row", [
      new LayoutRootMutation(layout, sheet, {
        type: "set-row-height",
        row: 1,
        height: 36,
        provenance: "manual"
      })
    ], resourcesFor(sheet));

    expect(layout.root).toBe(initialRoot);
    expect(layout.getRowHeight(1)).toBe(20);
    expect(layout.version).toBe(0);
    expect(coordinator.revisionGate.revision).toBe(0);
    expect(coordinator.state).toBe("prepared");

    const receipt = coordinator.commit(prepared);

    expect(layout.root).not.toBe(initialRoot);
    expect(layout.getRowHeight(1)).toBe(36);
    // 索引写了，权威也得写：这一条变更同时暂存了一条 `set-row-meta`（设计稿 §4.3）。
    expect(sheet.getRowMeta(1)?.h).toBe(36);
    expect(rowHeightProvenanceOf(sheet, 1)).toBe("manual");
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
    const { layout, sheet } = createFixture();
    const coordinator = new MutationCoordinator();
    const initialRoot = layout.root;

    const prepared = coordinator.prepare("resize-layout", [
      new LayoutRootMutation(layout, sheet, { type: "set-column-width", column: 1, width: 95 }),
      new LayoutRootMutation(layout, sheet, {
        type: "set-row-height",
        row: 2,
        height: 44,
        provenance: "manual"
      })
    ], resourcesFor(sheet));

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
    const { layout, sheet } = createFixture();
    const coordinator = new MutationCoordinator();
    const initialRoot = layout.root;
    const prepared = coordinator.prepare("rollback-layout", [
      new LayoutRootMutation(layout, sheet, {
        type: "set-row-height",
        row: 0,
        height: 42,
        provenance: "manual"
      }),
      createFailingMutation("later participant failed")
    ], resourcesFor(sheet));

    expect(() => coordinator.commit(prepared)).toThrow("later participant failed");

    expect(layout.root).toBe(initialRoot);
    expect(layout.getRowHeight(0)).toBe(20);
    // 回滚要把两边一起收回去，否则权威与索引就此分家。
    expect(sheet.getRowMeta(0)).toBeUndefined();
    expect(layout.version).toBe(0);
    expect(coordinator.revisionGate.revision).toBe(0);
    expect(coordinator.state).toBe("idle");
  });

  it("composes layout then structure mutations through the shared root registry", () => {
    const fixture = createGridFixture();
    const coordinator = new MutationCoordinator();
    const initialRoot = fixture.owner.getRoot();
    const prepared = coordinator.prepare("layout-then-structure", [
      new LayoutRootMutation(fixture.layout, fixture.sheet, {
        type: "set-column-width",
        column: 1,
        width: 75
      }),
      createColumnInsertMutation(fixture.owner)
    ], resourcesFor(fixture.sheet));

    expect(fixture.owner.isCurrent(initialRoot)).toBe(true);
    expect(fixture.getColumns()).toBe(initialRoot.columns);
    expect(fixture.columnManager.getColumns()).toBe(initialRoot.runtimeColumns);
    expect(fixture.headerManager.root).toBe(initialRoot.headers);
    expect(fixture.layout.root).toBe(initialRoot.layout);
    expect(fixture.selection.root).toBe(initialRoot.selection);

    const receipt = coordinator.commit(prepared);

    expect(fixture.getColumns().map((column) => column.width)).toEqual([30, 50, 40]);
    expect(fixture.columnManager.getColumns().map((column) => column.width)).toEqual([30, 50, 40]);
    expect(fixture.headerManager.getHeaders().map((header) => header.title)).toEqual(["A", "B", "C"]);
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
      new LayoutRootMutation(fixture.layout, fixture.sheet, {
        type: "set-column-width",
        column: 2,
        width: 75
      })
    ], resourcesFor(fixture.sheet));

    expect(fixture.owner.isCurrent(initialRoot)).toBe(true);
    expect(fixture.getColumns()).toBe(initialRoot.columns);
    expect(fixture.columnManager.getColumns()).toBe(initialRoot.runtimeColumns);
    expect(fixture.headerManager.root).toBe(initialRoot.headers);
    expect(fixture.layout.root).toBe(initialRoot.layout);
    expect(fixture.selection.root).toBe(initialRoot.selection);

    const receipt = coordinator.commit(prepared);

    expect(fixture.getColumns().map((column) => column.width)).toEqual([30, 50, 40]);
    expect(fixture.columnManager.getColumns().map((column) => column.width)).toEqual([30, 50, 40]);
    expect(fixture.headerManager.getHeaders().map((header) => header.title)).toEqual(["A", "B", "C"]);
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
      new LayoutRootMutation(fixture.layout, fixture.sheet, {
        type: "set-column-width",
        column: 1,
        width: 75
      }),
      createColumnInsertMutation(fixture.owner),
      createFailingMutation("mixed transaction failed")
    ], resourcesFor(fixture.sheet));

    expect(fixture.owner.isCurrent(initialRoot)).toBe(true);
    expect(() => coordinator.commit(prepared)).toThrow("mixed transaction failed");

    expect(fixture.owner.isCurrent(initialRoot)).toBe(true);
    expect(fixture.getColumns()).toBe(initialRoot.columns);
    expect(fixture.columnManager.getColumns()).toBe(initialRoot.runtimeColumns);
    expect(fixture.headerManager.root).toBe(initialRoot.headers);
    expect(fixture.layout.root).toBe(initialRoot.layout);
    expect(fixture.selection.root).toBe(initialRoot.selection);
    expect(fixture.getColumns().map((column) => column.width)).toEqual([30, 40]);
    expect(fixture.layout.getColumnWidthsSnapshot()).toEqual([30, 40]);
    expect(fixture.layout.columnCount).toBe(2);
    expect(fixture.layout.version).toBe(0);
    expect(fixture.selection.bounds.columnCount).toBe(2);
    expect(fixture.selection.snapshot.active.column).toBe(1);
    expect(coordinator.revisionGate.revision).toBe(0);
    expect(coordinator.state).toBe("idle");
  });

  it("rejects a stale prepared root without overwriting the newer live root", () => {
    const { layout, sheet } = createFixture();
    const coordinator = new MutationCoordinator();
    const prepared = coordinator.prepare("stale-layout", [
      new LayoutRootMutation(layout, sheet, {
        type: "set-row-height",
        row: 1,
        height: 48,
        provenance: "manual"
      })
    ], resourcesFor(sheet));

    layout.setRowHeight(1, 31);
    const newerRoot = layout.root;

    expectErrorCode(() => coordinator.commit(prepared), CyExcelErrorCode.MutationConflict);

    expect(layout.root).toBe(newerRoot);
    expect(layout.getRowHeight(1)).toBe(31);
    // 布局那边被别人改了，这条就整条不落地——权威也不能留下半条。
    expect(sheet.getRowMeta(1)).toBeUndefined();
    expect(layout.version).toBe(1);
    expect(coordinator.revisionGate.revision).toBe(0);
    expect(coordinator.state).toBe("idle");
  });

  describe("reorder-rows action", () => {
    // Deliberately asymmetric: [2, 0, 1] is not its own inverse, so a
    // direction mix-up between order and invertRowOrder() cannot pass.
    const order = () => Int32Array.from([2, 0, 1]);

    /**
     * 这一组只看**索引**这一半。
     *
     * `reorder-rows` 这一支的 `LayoutRootMutation` 只重排 `RowMetrics`；权威那半边的
     * `rowData` 由表自己那条同名变更（走 `IndexedTableRegistry`）搬，两者作为同一个
     * 事务里的两个参与者一起走——所以"重排后来源跟着走"是那一对的性质，归 §4.2 的
     * 属性测试管，这里不重复。这也是为什么下面直接用 `layout.setRowHeight()` 起高度：
     * 本组不读权威。
     */
    function createTallRows(): LayoutEngine {
      const { layout } = createFixture();
      layout.setRowHeight(0, 60);
      layout.setRowHeight(2, 44);
      return layout;
    }

    function reorderFixture(): { layout: LayoutEngine; sheet: SheetModel } {
      return { layout: createTallRows(), sheet: createSheetFromRows([{ a: "1" }], ["a"]) };
    }

    it("carries each row's height to its new position", () => {
      const { layout, sheet } = reorderFixture();
      const coordinator = new MutationCoordinator();
      const initialRoot = layout.root;

      const prepared = coordinator.prepare("reorder", [
        new LayoutRootMutation(layout, sheet, { type: "reorder-rows", order: order() })
      ]);

      expect(layout.root).toBe(initialRoot);
      expect([0, 1, 2].map((row) => layout.getRowHeight(row))).toEqual([60, 20, 44]);

      const receipt = coordinator.commit(prepared);

      // order[newPosition] = oldPosition, so new 0 <- old 2, new 1 <- old 0.
      expect([0, 1, 2].map((row) => layout.getRowHeight(row))).toEqual([44, 60, 20]);
      expect(receipt.patches).toEqual([{ type: "row-reorder", order: Int32Array.from([2, 0, 1]) }]);
    });

    it("inverts to the inverse permutation, restoring every height", () => {
      const { layout, sheet } = reorderFixture();
      const coordinator = new MutationCoordinator();

      const receipt = coordinator.commit(
        coordinator.prepare("reorder", [
          new LayoutRootMutation(layout, sheet, { type: "reorder-rows", order: order() })
        ])
      );
      coordinator.commit(coordinator.prepare("undo-reorder", [...receipt.inverse]));

      expect([0, 1, 2].map((row) => layout.getRowHeight(row))).toEqual([60, 20, 44]);
    });

    it("describes the inverse permutation in the inverse mutation's own patch", () => {
      const { layout, sheet } = reorderFixture();
      const coordinator = new MutationCoordinator();

      const receipt = coordinator.commit(
        coordinator.prepare("reorder", [
          new LayoutRootMutation(layout, sheet, { type: "reorder-rows", order: order() })
        ])
      );

      // Read straight off the *unprepared* inverse, the way a consumer that
      // wants to know what an undo will do (a History size estimate, a
      // dirty-region preview) reasonably might. prepare() recomputes
      // `changes`, so nothing today depends on this — but a reorder's patch
      // carries the permutation itself, and the inverse used to be handed
      // the forward one: [2, 0, 1] instead of its inverse [1, 2, 0].
      expect([...receipt.inverse].map((mutation) => mutation.changes)).toEqual([
        [{ type: "row-reorder", order: Int32Array.from([1, 2, 0]) }]
      ]);
    });

    it("stages a second reorder against the preceding staged root, not the live one", () => {
      const { layout, sheet } = reorderFixture();
      const coordinator = new MutationCoordinator();

      const prepared = coordinator.prepare("reorder-twice", [
        new LayoutRootMutation(layout, sheet, { type: "reorder-rows", order: order() }),
        new LayoutRootMutation(layout, sheet, { type: "reorder-rows", order: order() })
      ]);
      coordinator.commit(prepared);

      // [2,0,1] twice composes to [1,2,0]: new 0 <- old 1, new 1 <- old 2.
      expect([0, 1, 2].map((row) => layout.getRowHeight(row))).toEqual([20, 44, 60]);
    });
  });
});
