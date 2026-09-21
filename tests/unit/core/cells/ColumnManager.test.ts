import { describe, expect, it } from "vitest";

import {
  CellTypeRegistry,
  ColumnManager,
  registerBuiltinCellTypes
} from "../../../../src/core/cells";
import type { ResolveColumnContext, RuntimeColumn } from "../../../../src/core/cells";
import { HookManager, PRODUCTION_HOOK_REGISTRY } from "../../../../src/core/hooks";

function registryWithMoney(): CellTypeRegistry {
  const registry = new CellTypeRegistry();
  registerBuiltinCellTypes(registry);
  registry.register("money", {
    extends: "number",
    format: (value) => `¥${String(value)}`
  });
  // 一个自称"没有总序"的类型。内建类型里已经没有这样的了——`array`/`json` 随
  // `CellValue` 收紧一起删掉了——但 `sortable` 的解析链仍然要能表达这件事，所以
  // 由测试自己注册一个，而不是借一个恰好如此的内建类型。
  registry.register("opaque", {
    extends: "text",
    sortable: false
  });
  return registry;
}

describe("ColumnManager — structural defaults and column overrides", () => {
  // 这两条原来还各钉了一半 `field`（默认取 `id`、可被覆盖）。列已经没有 `id`
  // 也没有 `field` 了（ADR 0004 §3.8.2），剩下的结构性默认值只有宽度这一个。
  it("defaults width to 120 when the column doesn't specify it", () => {
    const manager = new ColumnManager({ columns: [{ }], registry: registryWithMoney() });

    expect(manager.getColumn(0).width).toBe(120);
  });

  it("lets the column definition override the structural defaults", () => {
    const manager = new ColumnManager({
      columns: [{ width: 200 }],
      registry: registryWithMoney()
    });

    expect(manager.getColumn(0).width).toBe(200);
  });

  it("defaults wrap to false and preserves an explicit true value", () => {
    const manager = new ColumnManager({
      columns: [{ }, { wrap: true }],
      registry: registryWithMoney()
    });

    expect(manager.getColumn(0).wrap).toBe(false);
    expect(manager.getColumn(1).wrap).toBe(true);
  });

  it("defaults clipboard.copy and clipboard.paste to true when the column doesn't specify them", () => {
    const manager = new ColumnManager({ columns: [{ }], registry: registryWithMoney() });

    expect(manager.getColumn(0).clipboard).toEqual({ copy: true, paste: true });
  });

  it("lets the column definition override clipboard.copy and clipboard.paste independently", () => {
    const manager = new ColumnManager({
      columns: [{ clipboard: { copy: false } }, { clipboard: { paste: false } }],
      registry: registryWithMoney()
    });

    expect(manager.getColumn(0).clipboard).toEqual({ copy: false, paste: true });
    expect(manager.getColumn(1).clipboard).toEqual({ copy: true, paste: false });
  });
});

describe("ColumnManager — type chain resolution", () => {
  it("resolves format/etc from the registered type, honoring its extends chain", () => {
    const manager = new ColumnManager({
      columns: [{ type: "money" }],
      registry: registryWithMoney()
    });

    const column = manager.getColumn(0);

    expect(column.typeName).toBe("money");
    expect(column.format(10, { row: 0, column })).toBe("¥10");
  });

  it("defaults to the \"text\" type when the column doesn't specify one", () => {
    const registry = new CellTypeRegistry();
    registry.register("text", { format: (value) => `text:${String(value)}` });
    const manager = new ColumnManager({ columns: [{ }], registry });

    expect(manager.getColumn(0).typeName).toBe("text");
    expect(manager.getColumn(0).format("x", { row: 0, column: manager.getColumn(0) })).toBe("text:x");
  });
});

describe("ColumnManager — select options", () => {
  it("takes the type's options when the column doesn't override them", () => {
    const registry = registryWithMoney();
    registry.register("status", {
      extends: "select",
      options: [{ value: "draft", label: "草稿" }]
    });
    const manager = new ColumnManager({ columns: [{ type: "status" }], registry });

    expect(manager.getColumn(0).options).toEqual([{ value: "draft", label: "草稿" }]);
  });

  it("lets the column definition override the type's options", () => {
    const registry = registryWithMoney();
    registry.register("status", {
      extends: "select",
      options: [{ value: "draft", label: "草稿" }]
    });
    const manager = new ColumnManager({
      columns: [{ type: "status", options: [{ value: "done", label: "完成" }] }],
      registry
    });

    expect(manager.getColumn(0).options).toEqual([{ value: "done", label: "完成" }]);
  });
});

describe("ColumnManager — createEditor", () => {
  it("column-level createEditor overrides the type's createEditor, same precedent as renderer", () => {
    const typeEditor = {} as never;
    const columnEditor = {} as never;
    const registry = registryWithMoney();
    registry.register("checkbox", { extends: "boolean", createEditor: () => typeEditor });
    const manager = new ColumnManager({
      columns: [{ type: "checkbox", createEditor: () => columnEditor }],
      registry
    });

    expect(manager.getColumn(0).createEditor?.({ row: 0, column: manager.getColumn(0) })).toBe(columnEditor);
  });

  it("a column can opt into a specific editor without registering a named type for it", () => {
    const columnEditor = {} as never;
    const manager = new ColumnManager({
      columns: [{ createEditor: () => columnEditor }],
      registry: registryWithMoney()
    });

    expect(manager.getColumn(0).createEditor?.({ row: 0, column: manager.getColumn(0) })).toBe(columnEditor);
  });
});

describe("ColumnManager — resolveColumn hook", () => {
  it("runs last and can override anything the type chain or column definition produced", () => {
    const hooks = new HookManager({ registry: PRODUCTION_HOOK_REGISTRY });
    hooks.tap("resolveColumn", { name: "override-column" }, (column: RuntimeColumn) => ({
      ...column,
      width: 999,
      format: () => "overridden"
    }));
    const manager = new ColumnManager({
      columns: [{ type: "money", width: 90 }],
      registry: registryWithMoney(),
      hooks
    });

    const column = manager.getColumn(0);

    expect(column.width).toBe(999);
    expect(column.format(1, { row: 0, column })).toBe("overridden");
  });
});

describe("ColumnManager — shared HookManager", () => {
  it("runs a registered resolveColumn waterfall", () => {
    const hooks = new HookManager({ registry: PRODUCTION_HOOK_REGISTRY });
    hooks.tap("resolveColumn", { name: "force-width" }, (column: RuntimeColumn) => ({
      ...column,
      width: 999
    }));

    const manager = new ColumnManager({
      columns: [{ type: "money", width: 90 }],
      registry: registryWithMoney(),
      hooks
    });

    expect(manager.getColumn(0).width).toBe(999);
  });

  it("chains multiple taps in waterfall order, each seeing the previous tap's output", () => {
    const hooks = new HookManager({ registry: PRODUCTION_HOOK_REGISTRY });
    hooks.tap("resolveColumn", { name: "double-width" }, (column: RuntimeColumn) => ({
      ...column,
      width: column.width * 2
    }));
    hooks.tap("resolveColumn", { name: "add-ten" }, (column: RuntimeColumn) => ({
      ...column,
      width: column.width + 10
    }));

    const manager = new ColumnManager({
      columns: [{ width: 90 }],
      registry: registryWithMoney(),
      hooks
    });

    expect(manager.getColumn(0).width).toBe(190); // (90 * 2) + 10
  });

  it("hands each tap the declared ColumnDefinition and its index", () => {
    const hooks = new HookManager({ registry: PRODUCTION_HOOK_REGISTRY });
    // 断言的是「拿到的确实是**这一列自己**声明的那份定义」。原来用 `id` 区分两
    // 列，列已经没有 id 了（ADR 0004 §3.8.2），改用两个不同的宽度当标记——重点
    // 从来不是那个值叫什么，而是它和 `index` 对得上。
    const seen: Array<{ width: number | undefined; index: number }> = [];
    hooks.tap<ResolveColumnContext, RuntimeColumn>(
      "resolveColumn",
      { name: "observe" },
      (column: RuntimeColumn, context: ResolveColumnContext) => {
        seen.push({ width: context.definition.width, index: context.index });
        return column;
      }
    );

    new ColumnManager({
      columns: [{ width: 10 }, { width: 20 }],
      registry: registryWithMoney(),
      hooks
    });

    expect(seen).toEqual([
      { width: 10, index: 0 },
      { width: 20, index: 1 }
    ]);
  });
});

describe("ColumnManager — resolves once and caches", () => {
  it("only calls resolveColumn once per column at construction, not on every getColumn() call", () => {
    let calls = 0;
    const hooks = new HookManager({ registry: PRODUCTION_HOOK_REGISTRY });
    hooks.tap("resolveColumn", { name: "count-resolutions" }, (column: RuntimeColumn) => {
      calls += 1;
      return column;
    });
    const manager = new ColumnManager({
      columns: [{ }, { }],
      registry: registryWithMoney(),
      hooks
    });

    manager.getColumn(0);
    manager.getColumn(0);
    manager.getColumn(1);
    manager.getColumn(1);

    expect(calls).toBe(2);
  });

  it("getColumn() returns the exact same object on repeated calls", () => {
    const manager = new ColumnManager({ columns: [{ }], registry: registryWithMoney() });

    expect(manager.getColumn(0)).toBe(manager.getColumn(0));
  });

  it("getColumns() returns all resolved columns in order", () => {
    // 「有序」以前靠 id 序列断言。列没有 id 了（ADR 0004 §3.8.2），换成三个互不
    // 相同的宽度——同样能把顺序钉死，而且这次钉的是真存在的属性。
    const manager = new ColumnManager({
      columns: [{ width: 10 }, { width: 20 }, { width: 30 }],
      registry: registryWithMoney()
    });

    expect(manager.getColumns().map((column) => column.width)).toEqual([10, 20, 30]);
  });

  it("columnCount reflects the number of resolved columns", () => {
    const manager = new ColumnManager({
      columns: [{ }, { }],
      registry: registryWithMoney()
    });

    expect(manager.columnCount).toBe(2);
  });
});

describe("ColumnManager — out-of-range access", () => {
  it("getColumn() throws for an index outside the resolved columns", () => {
    const manager = new ColumnManager({ columns: [{ }], registry: registryWithMoney() });

    expect(() => manager.getColumn(1)).toThrow(RangeError);
  });
});

describe("ColumnManager root replacement", () => {
  it("publishes a resolved root from a new column list", () => {
    const manager = new ColumnManager({
      columns: [{ width: 1 }, { width: 2 }],
      registry: registryWithMoney()
    });
    expect(manager.columnCount).toBe(2);

    manager.replaceRoot(manager.prepareRoot([{ width: 10 }, { width: 20 }, { width: 30 }]));

    expect(manager.columnCount).toBe(3);
    expect(manager.getColumns().map((column) => column.width)).toEqual([10, 20, 30]);
  });

  it("re-runs the resolveColumn hook for every column in the new list", () => {
    const hooks = new HookManager({ registry: PRODUCTION_HOOK_REGISTRY });
    hooks.tap("resolveColumn", { name: "force-width" }, (column: RuntimeColumn) => ({
      ...column,
      width: 999
    }));
    const manager = new ColumnManager({ columns: [{ }], registry: registryWithMoney(), hooks });

    manager.replaceRoot(manager.prepareRoot([{ }, { }]));

    expect(manager.getColumn(1).width).toBe(999);
  });

  it("getColumn()/getColumns() reflect the replacement root immediately", () => {
    const manager = new ColumnManager({
      columns: [{ width: 100 }],
      registry: registryWithMoney()
    });

    manager.replaceRoot(manager.prepareRoot([{ width: 500 }]));

    expect(manager.getColumn(0).width).toBe(500);
  });
});

describe("ColumnManager prepared roots", () => {
  it("prepareRoot resolves a complete immutable root without changing live columns", () => {
    const manager = new ColumnManager({
      columns: [{ width: 100 }],
      registry: registryWithMoney()
    });

    const root = manager.prepareRoot([{ width: 180 }, { width: 240 }]);

    // 宽度在这里同时充当「哪一份 root」的标记：活的还是那一列 100，备好的是新的两列。
    expect(manager.getColumns().map((column) => column.width)).toEqual([100]);
    expect(root.map((column) => column.width)).toEqual([180, 240]);
    expect(Object.isFrozen(root)).toBe(true);
    expect(root.every(Object.isFrozen)).toBe(true);
    expect(root.every((column) => Object.isFrozen(column.clipboard))).toBe(true);
  });

  it("contains an isolated resolveColumn hook failure inside the prepared root", () => {
    const hooks = new HookManager({ registry: PRODUCTION_HOOK_REGISTRY });
    // 三列的区分标记从 id 换成了声明宽度（ADR 0004 §3.8.2）：钩子只在中间那一列
    // 上抛，于是它退回自己声明的 2，前后两列照常被改写成 999。
    const resolved: number[] = [];
    hooks.tap("resolveColumn", { name: "fail-second" }, (column: RuntimeColumn) => {
      resolved.push(column.width);
      if (column.width === 2) {
        throw new Error("column resolution failed");
      }
      return { ...column, width: 999 };
    });
    const manager = new ColumnManager({
      columns: [{ width: 100 }],
      registry: registryWithMoney(),
      hooks
    });
    resolved.length = 0;

    const root = manager.prepareRoot([{ width: 1 }, { width: 2 }, { width: 3 }]);

    expect(resolved).toEqual([1, 2, 3]);
    expect(root.map((column) => column.width)).toEqual([999, 2, 999]);
    expect(manager.columnCount).toBe(1);
    expect(manager.getColumn(0).width).toBe(999);
  });

  it("does not leak partially resolved columns when later resolution throws", () => {
    const manager = new ColumnManager({
      columns: [{ width: 100 }],
      registry: registryWithMoney()
    });
    const liveRoot = manager.getColumns();

    expect(() =>
      manager.prepareRoot([
        { type: "money" },
        { type: "missing-cell-type" },
        { }
      ])
    ).toThrow('Unknown cell type "missing-cell-type"');

    expect(manager.getColumns()).toBe(liveRoot);
    expect(manager.columnCount).toBe(1);
    expect(manager.getColumn(0).width).toBe(100);
  });

  it("replaceRoot swaps only the root and never re-runs resolveColumn hooks", () => {
    let hookCalls = 0;
    const hooks = new HookManager({ registry: PRODUCTION_HOOK_REGISTRY });
    hooks.tap("resolveColumn", { name: "count-root-resolutions" }, (column: RuntimeColumn) => {
      hookCalls += 1;
      return { ...column, width: column.width + hookCalls };
    });
    const manager = new ColumnManager({
      columns: [{ width: 100 }],
      registry: registryWithMoney(),
      hooks
    });
    const root = manager.prepareRoot([{ width: 200 }, { width: 300 }]);
    const callsAfterPrepare = hookCalls;

    manager.replaceRoot(root);

    expect(hookCalls).toBe(callsAfterPrepare);
    expect(manager.getColumns()).toBe(root);
    // 每列的宽度里带着「钩子跑到第几次」——replaceRoot 之后它们还是 prepareRoot
    // 那一轮的 2、3 次，没有被重跑一遍加成 4、5 次。
    expect(manager.getColumns().map((column) => column.width)).toEqual([202, 303]);
  });

});

describe("ColumnManager — sortable resolution", () => {
  it("treats an ordinary column as sortable without being told", () => {
    const manager = new ColumnManager({ columns: [{ }], registry: registryWithMoney() });

    expect(manager.getColumn(0).sortable).toBe(true);
  });

  it("takes sortable from the cell type, so a type with no total order opts itself out", () => {
    // A type that declares itself unordered — sorting it would fall back to
    // comparing string forms, which is a fake sort.
    const manager = new ColumnManager({
      columns: [{ type: "opaque" }, { type: "text" }],
      registry: registryWithMoney()
    });

    expect(manager.getColumn(0).sortable).toBe(false);
    // Premise: the default really is true, so the false above came from the type.
    expect(manager.getColumn(1).sortable).toBe(true);
  });

  it("lets a column override the type's own answer in both directions", () => {
    const manager = new ColumnManager({
      columns: [
        { type: "opaque", sortable: true },
        { sortable: false }
      ],
      registry: registryWithMoney()
    });

    expect(manager.getColumn(0).sortable).toBe(true);
    expect(manager.getColumn(1).sortable).toBe(false);
  });

  it("inherits sortable through `extends`", () => {
    const registry = registryWithMoney();
    registry.register("tagList", { extends: "opaque" });
    const manager = new ColumnManager({ columns: [{ type: "tagList" }], registry });

    expect(manager.getColumn(0).sortable).toBe(false);
  });
});
