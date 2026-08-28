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
  HideColumnCommand,
  InsertColumnsCommand,
  InsertRowsCommand,
  MoveColumnsCommand,
  MoveRowsCommand,
  ResizeColumnCommand,
  ShowColumnCommand
} from "../../../../src/core/commands";
import { TableStore } from "../../../../src/core/data";
import { HeaderManager } from "../../../../src/core/headers";
import { GridStructureOwner, createGridCommandManager } from "../../../../src/core/integration";
import { SelectionModel } from "../../../../src/core/interaction";
import type { LayoutEngine } from "../../../../src/core/layout";
import type { CommandManager } from "../../../../src/core/commands";
import {
  createSeededLayout,
  geometricColumnWidthOf
} from "../../../helpers/columnWidthAuthority";
import { createSheetFromRows } from "../../../helpers/createSheetModel";

const DEFAULT_ROW_HEIGHT = 20;
const INITIAL_ROWS = 6;
const INITIAL_COLUMNS = 3;

interface Fixture {
  readonly store: TableStore;
  readonly layout: LayoutEngine;
  readonly manager: CommandManager;
}

function createFixture(): Fixture {
  let columns: readonly ColumnDefinition[] = Object.freeze([
    Object.freeze({ width: 50 }),
    Object.freeze({ width: 60 }),
    Object.freeze({ width: 70 })
  ]);
  const sheet = createSheetFromRows(
    Array.from({ length: INITIAL_ROWS }, (_unused, row) => ({ a: row, b: row, c: row })),
    ["a", "b", "c"]
  );
  const store = new TableStore(sheet);
  const registry = new CellTypeRegistry();
  registerBuiltinCellTypes(registry);
  const columnManager = new ColumnManager({ columns, registry });
  const headerManager = new HeaderManager({ columns });
  // 三列宽度**各不相同**：全给 50 的话，一条把宽度搬错了列的实现照样全绿。
  const layout = createSeededLayout(sheet, {
    rowCount: INITIAL_ROWS,
    columnCount: INITIAL_COLUMNS,
    rowHeight: DEFAULT_ROW_HEIGHT,
    columnWidths: [50, 60, 70]
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
 * 设计稿 §4.2 的列轴那条不变量，逐列核对。
 *
 * 权威是 `columnData[column]` 的 `w`/`hd`，`ColumnAxisSnapshot` 只是由它派生的稠密
 * 索引；「隐藏」在索引里就是**几何宽度 0**，T8b 之后布局层已经不另存一份隐藏表了。
 * 缺省项取的是**配置**里的 `layout.defaultColumnWidth`，不是 `SheetModel` 的内部默认值。
 */
function expectColumnWidthAuthorityMatchesIndex(fixture: Fixture, step: string): void {
  const { store, layout } = fixture;
  expect(`${step}: columnCount ${layout.columnCount}`)
    .toBe(`${step}: columnCount ${store.getColumnCount()}`);

  const indexed: string[] = [];
  const authoritative: string[] = [];
  for (let column = 0; column < layout.columnCount; column += 1) {
    indexed.push(`${column}:${layout.getColumnWidth(column)}`);
    authoritative.push(
      `${column}:${geometricColumnWidthOf(store.sheet, column, layout.defaultColumnWidth)}`
    );
  }
  expect(`${step} ${indexed.join(" ")}`).toBe(`${step} ${authoritative.join(" ")}`);

  // 偏移量是宽度的前缀和，本来该自动成立——但它是增量维护的另一条链，
  // 只核宽度的话「宽度对了、偏移没跟着重算」这种分叉会漏网。
  let offset = 0;
  for (let column = 0; column < layout.columnCount; column += 1) {
    expect(`${step}: offset ${column} = ${layout.getColumnLeft(column)}`)
      .toBe(`${step}: offset ${column} = ${offset}`);
    offset += layout.getColumnWidth(column);
  }
}

type Operation =
  | { readonly kind: "resize-column"; readonly column: number; readonly width: number }
  | { readonly kind: "hide-column"; readonly column: number }
  | { readonly kind: "show-column"; readonly column: number }
  | { readonly kind: "insert-columns"; readonly start: number; readonly count: number; readonly width: number | undefined }
  | { readonly kind: "delete-columns"; readonly start: number; readonly count: number }
  | { readonly kind: "move-columns"; readonly start: number; readonly target: number }
  | { readonly kind: "insert-rows"; readonly start: number; readonly count: number }
  | { readonly kind: "delete-rows"; readonly start: number; readonly count: number }
  | { readonly kind: "move-rows"; readonly start: number; readonly target: number }
  | { readonly kind: "undo" }
  | { readonly kind: "redo" };

const operationArbitrary: fc.Arbitrary<Operation> = fc.oneof(
  fc.record({
    kind: fc.constant("resize-column" as const),
    column: fc.nat({ max: 20 }),
    width: fc.integer({ min: 15, max: 200 })
  }),
  fc.record({ kind: fc.constant("hide-column" as const), column: fc.nat({ max: 20 }) }),
  fc.record({ kind: fc.constant("show-column" as const), column: fc.nat({ max: 20 }) }),
  fc.record({
    kind: fc.constant("insert-columns" as const),
    start: fc.nat({ max: 20 }),
    count: fc.integer({ min: 1, max: 2 }),
    // `undefined` 这一支专门打「插入时没给宽度」：那时权威里根本不该有条目，
    // 索引要落到缺省宽上。
    width: fc.option(fc.integer({ min: 15, max: 200 }), { nil: undefined })
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
  fc.record({ kind: fc.constant("undo" as const) }),
  fc.record({ kind: fc.constant("redo" as const) })
);

/**
 * 生成的下标一律**夹到**当前合法范围，理由与行轴那份同样：让越界的去撞校验，
 * 绝大多数样本会退化成「什么也没做」。
 */
function applyOperation(fixture: Fixture, operation: Operation): string | undefined {
  const { store, manager } = fixture;
  const rowCount = store.getRowCount();
  const columnCount = store.getColumnCount();

  switch (operation.kind) {
    case "resize-column": {
      const column = operation.column % columnCount;
      manager.execute(new ResizeColumnCommand(column, operation.width));
      return `resize-column ${column}=${operation.width}`;
    }
    case "hide-column": {
      const column = operation.column % columnCount;
      manager.execute(new HideColumnCommand(column));
      return `hide-column ${column}`;
    }
    case "show-column": {
      const column = operation.column % columnCount;
      manager.execute(new ShowColumnCommand(column));
      return `show-column ${column}`;
    }
    case "insert-columns": {
      const start = operation.start % (columnCount + 1);
      const inserted = Array.from(
        { length: operation.count },
        (): ColumnDefinition => (operation.width === undefined ? {} : { width: operation.width })
      );
      manager.execute(new InsertColumnsCommand(start, inserted));
      return `insert-columns ${start}+${operation.count}@${operation.width ?? "default"}`;
    }
    case "delete-columns": {
      // 一列都不剩的表不是这条不变量要谈的东西，留一列。
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
    case "insert-rows": {
      const start = operation.start % (rowCount + 1);
      const rows = Array.from({ length: operation.count }, () =>
        Array.from({ length: columnCount }, () => "new"));
      manager.execute(new InsertRowsCommand(start, rows));
      return `insert-rows ${start}+${operation.count}`;
    }
    case "delete-rows": {
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
 * ADR 0004 阶段 4B T8b / 设计稿 §4.2 的硬门槛，列轴这一半。
 *
 * 列宽与「这列是不是隐藏的」有**一个权威**（`columnData`）和**一个增量维护的索引**
 * （`ColumnAxisSnapshot`）。T8b 之前布局层另存了一张 `#hiddenColumnWidths`，那是第二个
 * 存储、分叉的温床；删掉它之后隐藏就只是几何宽度 0，而「恢复成多宽」记在权威的 `w` 里。
 * 这条属性测试覆盖列的改宽 / 隐藏 / 取消隐藏 / 插 / 删 / 移、行的插 / 删 / 移，以及撤销
 * 与重做的随机交叉序列——逆变更那条路正是分叉最爱藏身的地方。
 */
describe("column-width authority vs. ColumnAxisSnapshot index (design §4.2)", () => {
  it("keeps every column's geometric width identical on both sides", () => {
    fc.assert(
      fc.property(fc.array(operationArbitrary, { maxLength: 24 }), (operations) => {
        const fixture = createFixture();
        expectColumnWidthAuthorityMatchesIndex(fixture, "initial");

        const trail: string[] = [];
        for (const operation of operations) {
          const applied = applyOperation(fixture, operation);
          if (applied === undefined) {
            continue;
          }
          trail.push(applied);
          expectColumnWidthAuthorityMatchesIndex(fixture, `after [${trail.join(" | ")}]`);
        }
      }),
      { numRuns: 300 }
    );
  });
});
