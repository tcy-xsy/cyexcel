import fc from "fast-check";
import { describe, expect, it } from "vitest";

import {
  CellTypeRegistry,
  ColumnManager,
  registerBuiltinCellTypes,
  type ColumnDefinition
} from "../../../../src/core/cells";
import {
  DeleteColumnsCommand,
  DeleteRowsCommand,
  InsertColumnsCommand,
  InsertRowsCommand,
  MoveColumnsCommand,
  MoveRowsCommand,
  ReorderRowsCommand,
  ResizeRowCommand,
  SetRowHeightsCommand
} from "../../../../src/core/commands";
import { TableStore } from "../../../../src/core/data";
import { HeaderManager } from "../../../../src/core/headers";
import { GridStructureOwner, createGridCommandManager } from "../../../../src/core/integration";
import { SelectionModel } from "../../../../src/core/interaction";
import { LayoutEngine } from "../../../../src/core/layout";
import type { CommandManager } from "../../../../src/core/commands";
import { createSheetFromRows } from "../../../helpers/createSheetModel";
import { rowHeightProvenanceOf } from "../../../helpers/rowHeightAuthority";

const DEFAULT_ROW_HEIGHT = 20;
const INITIAL_ROWS = 6;
const INITIAL_COLUMNS = 3;

interface Fixture {
  readonly store: TableStore;
  readonly layout: LayoutEngine;
  readonly manager: CommandManager;
}

function createFixture(): Fixture {
  let columns: readonly ColumnDefinition[] = Object.freeze(
    Array.from({ length: INITIAL_COLUMNS }, () => Object.freeze({ width: 50 }))
  );
  const store = new TableStore(createSheetFromRows(
    Array.from({ length: INITIAL_ROWS }, (_unused, row) => ({ a: row, b: row, c: row })),
    ["a", "b", "c"]
  ));
  const registry = new CellTypeRegistry();
  registerBuiltinCellTypes(registry);
  const columnManager = new ColumnManager({ columns, registry });
  const headerManager = new HeaderManager({ columns });
  const layout = new LayoutEngine({
    rowCount: INITIAL_ROWS,
    columnCount: INITIAL_COLUMNS,
    rowHeight: DEFAULT_ROW_HEIGHT,
    columnWidths: [50, 50, 50]
  });
  const selection = new SelectionModel({ rowCount: INITIAL_ROWS, columnCount: INITIAL_COLUMNS });
  const structureOwner = new GridStructureOwner({
    getColumns: () => columns,
    replaceColumns: (next) => {
      columns = next;
    },
    columnManager,
    headerManager,
    layout,
    selection
  });
  const manager = createGridCommandManager({ store, layout, structureOwner });
  return { store, layout, manager };
}

/**
 * 设计稿 §4.2 的那条不变量，逐行核对。
 *
 * 权威是 `rowData[row].h` / `.hp`，`RowMetrics` 只是它的稠密索引；缺省项取的是
 * **配置**里的 `layout.rowHeight`，不是 `SheetModel` 内部那个默认值——"这一行没有
 * 显式高度"要落到布局的缺省行高上。
 */
function expectRowHeightAuthorityMatchesIndex(fixture: Fixture, step: string): void {
  const { store, layout } = fixture;
  expect(`${step}: rowCount ${layout.rowCount}`).toBe(`${step}: rowCount ${store.getRowCount()}`);

  const indexed: string[] = [];
  const authoritative: string[] = [];
  for (let row = 0; row < layout.rowCount; row += 1) {
    const meta = store.sheet.getRowMeta(row);
    indexed.push(`${row}:${layout.getRowHeight(row)}/${rowHeightProvenanceOf(store.sheet, row)}`);
    authoritative.push(`${row}:${meta?.h ?? layout.rowHeight}/${meta?.hp ?? "default"}`);
  }
  expect(`${step} ${indexed.join(" ")}`).toBe(`${step} ${authoritative.join(" ")}`);
}

type Operation =
  | { readonly kind: "resize-row"; readonly row: number; readonly height: number }
  | {
    readonly kind: "set-row-heights";
    readonly rows: readonly number[];
    readonly height: number;
    readonly provenance: "auto" | "manual";
  }
  | { readonly kind: "insert-rows"; readonly start: number; readonly count: number }
  | { readonly kind: "delete-rows"; readonly start: number; readonly count: number }
  | { readonly kind: "move-rows"; readonly start: number; readonly target: number }
  | { readonly kind: "reorder-rows"; readonly rotation: number }
  | { readonly kind: "insert-columns"; readonly start: number; readonly count: number }
  | { readonly kind: "delete-columns"; readonly start: number; readonly count: number }
  | { readonly kind: "move-columns"; readonly start: number; readonly target: number }
  | { readonly kind: "undo" }
  | { readonly kind: "redo" };

const operationArbitrary: fc.Arbitrary<Operation> = fc.oneof(
  fc.record({
    kind: fc.constant("resize-row" as const),
    row: fc.nat({ max: 20 }),
    height: fc.integer({ min: 12, max: 90 })
  }),
  fc.record({
    kind: fc.constant("set-row-heights" as const),
    rows: fc.uniqueArray(fc.nat({ max: 20 }), { minLength: 1, maxLength: 3 }),
    height: fc.integer({ min: 12, max: 90 }),
    provenance: fc.constantFrom("auto" as const, "manual" as const)
  }),
  fc.record({
    kind: fc.constant("insert-rows" as const),
    start: fc.nat({ max: 20 }),
    count: fc.integer({ min: 1, max: 2 })
  }),
  fc.record({
    kind: fc.constant("delete-rows" as const),
    start: fc.nat({ max: 20 }),
    count: fc.integer({ min: 1, max: 2 })
  }),
  fc.record({
    kind: fc.constant("move-rows" as const),
    start: fc.nat({ max: 20 }),
    target: fc.nat({ max: 20 })
  }),
  fc.record({ kind: fc.constant("reorder-rows" as const), rotation: fc.nat({ max: 20 }) }),
  fc.record({
    kind: fc.constant("insert-columns" as const),
    start: fc.nat({ max: 20 }),
    count: fc.integer({ min: 1, max: 2 })
  }),
  fc.record({
    kind: fc.constant("delete-columns" as const),
    start: fc.nat({ max: 20 }),
    count: fc.integer({ min: 1, max: 2 })
  }),
  fc.record({
    kind: fc.constant("move-columns" as const),
    start: fc.nat({ max: 20 }),
    target: fc.nat({ max: 20 })
  }),
  fc.record({ kind: fc.constant("undo" as const) }),
  fc.record({ kind: fc.constant("redo" as const) })
);

/**
 * 生成的下标一律**夹到**当前合法范围，而不是让越界的那些去撞校验。
 *
 * 撞校验也能过——命令是原子的，抛出时状态没动——但那样绝大多数样本会退化成
 * "什么也没做"，真正在跑的交叉序列就没剩几条了。夹一下，每个样本都真的动了一次。
 */
function applyOperation(fixture: Fixture, operation: Operation): string | undefined {
  const { store, manager } = fixture;
  const rowCount = store.getRowCount();
  const columnCount = store.getColumnCount();

  switch (operation.kind) {
    case "resize-row": {
      const row = operation.row % rowCount;
      manager.execute(new ResizeRowCommand(row, operation.height));
      return `resize-row ${row}=${operation.height}`;
    }
    case "set-row-heights": {
      const rows = [...new Set(operation.rows.map((row) => row % rowCount))];
      manager.execute(new SetRowHeightsCommand(
        rows.map((row) => ({ row, height: operation.height })),
        operation.provenance
      ));
      return `set-row-heights ${rows.join(",")}=${operation.height}/${operation.provenance}`;
    }
    case "insert-rows": {
      const start = operation.start % (rowCount + 1);
      const rows = Array.from({ length: operation.count }, () =>
        Array.from({ length: columnCount }, () => "new"));
      manager.execute(new InsertRowsCommand(start, rows));
      return `insert-rows ${start}+${operation.count}`;
    }
    case "delete-rows": {
      // 一行都不剩的表不是这条不变量要谈的东西，留一行。
      const count = Math.min(operation.count, rowCount - 1);
      if (count <= 0) {
        return undefined;
      }
      const start = operation.start % (rowCount - count + 1);
      manager.execute(new DeleteRowsCommand(start, count));
      return `delete-rows ${start}+${count}`;
    }
    case "move-rows": {
      if (rowCount < 2) {
        return undefined;
      }
      const start = operation.start % rowCount;
      const target = operation.target % rowCount;
      if (start === target) {
        return undefined;
      }
      manager.execute(new MoveRowsCommand(start, 1, target));
      return `move-rows ${start}->${target}`;
    }
    case "reorder-rows": {
      const rotation = operation.rotation % rowCount;
      if (rotation === 0) {
        return undefined;
      }
      const order = Int32Array.from(
        { length: rowCount },
        (_unused, row) => (row + rotation) % rowCount
      );
      manager.execute(new ReorderRowsCommand(order));
      return `reorder-rows +${rotation}`;
    }
    case "insert-columns": {
      const start = operation.start % (columnCount + 1);
      manager.execute(new InsertColumnsCommand(
        start,
        Array.from({ length: operation.count }, () => ({ width: 40 }))
      ));
      return `insert-columns ${start}+${operation.count}`;
    }
    case "delete-columns": {
      const count = Math.min(operation.count, columnCount - 1);
      if (count <= 0) {
        return undefined;
      }
      const start = operation.start % (columnCount - count + 1);
      manager.execute(new DeleteColumnsCommand(start, count));
      return `delete-columns ${start}+${count}`;
    }
    case "move-columns": {
      if (columnCount < 2) {
        return undefined;
      }
      const start = operation.start % columnCount;
      const target = operation.target % columnCount;
      if (start === target) {
        return undefined;
      }
      manager.execute(new MoveColumnsCommand(start, 1, target));
      return `move-columns ${start}->${target}`;
    }
    case "undo": {
      if (!manager.canUndo()) {
        return undefined;
      }
      manager.undo();
      return "undo";
    }
    case "redo": {
      if (!manager.canRedo()) {
        return undefined;
      }
      manager.redo();
      return "redo";
    }
  }
}

/**
 * ADR 0004 阶段 4B T8 / 设计稿 §4.2 的硬门槛。
 *
 * 行高有**一个权威**（`rowData`）和**一个增量维护的索引**（`RowMetrics`）。不做全量
 * 重建的代价，就是每一个结构操作都必须把两边搬得一样多；这条属性测试是唯一能挡住
 * 两者分叉的东西，所以它覆盖 7 个结构操作（行的插 / 删 / 移 / 重排、列的插 / 删 /
 * 移）与行高设置的随机交叉序列，外加撤销 / 重做——逆变更那条路正是分叉最爱藏身的
 * 地方。
 */
describe("row-height authority vs. RowMetrics index (design §4.2)", () => {
  it("keeps every row's height and provenance identical on both sides", () => {
    fc.assert(
      fc.property(fc.array(operationArbitrary, { maxLength: 24 }), (operations) => {
        const fixture = createFixture();
        expectRowHeightAuthorityMatchesIndex(fixture, "initial");

        const trail: string[] = [];
        for (const operation of operations) {
          const applied = applyOperation(fixture, operation);
          if (applied === undefined) {
            continue;
          }
          trail.push(applied);
          expectRowHeightAuthorityMatchesIndex(fixture, `after [${trail.join(" | ")}]`);
        }
      }),
      { numRuns: 300 }
    );
  });
});
