import { describe, expect, it } from "vitest";

import {
  CellTypeRegistry,
  ColumnManager,
  HookManager,
  PRODUCTION_HOOK_REGISTRY,
  registerBuiltinCellTypes
} from "../../../../src/core";
import type { RuntimeColumn } from "../../../../src/core";

function registryWithMoney(): CellTypeRegistry {
  const registry = new CellTypeRegistry();
  registerBuiltinCellTypes(registry);
  registry.register("money", {
    extends: "number",
    format: (value) => `¥${String(value)}`
  });
  return registry;
}

describe("ColumnManager — structural defaults and column overrides", () => {
  it("defaults field to the column id and width to 120 when the column doesn't specify them", () => {
    const manager = new ColumnManager({ columns: [{ id: "sku" }], registry: registryWithMoney() });

    const column = manager.getColumn(0);

    expect(column.field).toBe("sku");
    expect(column.width).toBe(120);
  });

  it("lets the column definition override the structural defaults", () => {
    const manager = new ColumnManager({
      columns: [{ id: "sku", field: "productSku", width: 200 }],
      registry: registryWithMoney()
    });

    const column = manager.getColumn(0);

    expect(column.field).toBe("productSku");
    expect(column.width).toBe(200);
  });

  it("defaults clipboard.copy and clipboard.paste to true when the column doesn't specify them", () => {
    const manager = new ColumnManager({ columns: [{ id: "sku" }], registry: registryWithMoney() });

    expect(manager.getColumn(0).clipboard).toEqual({ copy: true, paste: true });
  });

  it("lets the column definition override clipboard.copy and clipboard.paste independently", () => {
    const manager = new ColumnManager({
      columns: [{ id: "secret", clipboard: { copy: false } }, { id: "locked", clipboard: { paste: false } }],
      registry: registryWithMoney()
    });

    expect(manager.getColumn(0).clipboard).toEqual({ copy: false, paste: true });
    expect(manager.getColumn(1).clipboard).toEqual({ copy: true, paste: false });
  });
});

describe("ColumnManager — type chain resolution", () => {
  it("resolves format/etc from the registered type, honoring its extends chain", () => {
    const manager = new ColumnManager({
      columns: [{ id: "price", type: "money" }],
      registry: registryWithMoney()
    });

    const column = manager.getColumn(0);

    expect(column.typeName).toBe("money");
    expect(column.format(10, { row: 0, column })).toBe("¥10");
  });

  it("defaults to the \"text\" type when the column doesn't specify one", () => {
    const registry = new CellTypeRegistry();
    registry.register("text", { format: (value) => `text:${String(value)}` });
    const manager = new ColumnManager({ columns: [{ id: "name" }], registry });

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
    const manager = new ColumnManager({ columns: [{ id: "s", type: "status" }], registry });

    expect(manager.getColumn(0).options).toEqual([{ value: "draft", label: "草稿" }]);
  });

  it("lets the column definition override the type's options", () => {
    const registry = registryWithMoney();
    registry.register("status", {
      extends: "select",
      options: [{ value: "draft", label: "草稿" }]
    });
    const manager = new ColumnManager({
      columns: [{ id: "s", type: "status", options: [{ value: "done", label: "完成" }] }],
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
      columns: [{ id: "c", type: "checkbox", createEditor: () => columnEditor }],
      registry
    });

    expect(manager.getColumn(0).createEditor?.({ row: 0, column: manager.getColumn(0) })).toBe(columnEditor);
  });

  it("a column can opt into a specific editor without registering a named type for it", () => {
    const columnEditor = {} as never;
    const manager = new ColumnManager({
      columns: [{ id: "c", createEditor: () => columnEditor }],
      registry: registryWithMoney()
    });

    expect(manager.getColumn(0).createEditor?.({ row: 0, column: manager.getColumn(0) })).toBe(columnEditor);
  });
});

describe("ColumnManager — resolveColumn hook", () => {
  it("runs last and can override anything the type chain or column definition produced", () => {
    const manager = new ColumnManager({
      columns: [{ id: "price", type: "money", width: 90 }],
      registry: registryWithMoney(),
      resolveColumn: (column) => ({ ...column, width: 999, format: () => "overridden" })
    });

    const column = manager.getColumn(0);

    expect(column.width).toBe(999);
    expect(column.format(1, { row: 0, column })).toBe("overridden");
  });
});

describe("ColumnManager — shared HookManager", () => {
  it("hooks.tap('resolveColumn', ...) is equivalent to the resolveColumn option", () => {
    const hooks = new HookManager({ registry: PRODUCTION_HOOK_REGISTRY });
    hooks.tap("resolveColumn", { name: "force-width" }, (column: RuntimeColumn) => ({
      ...column,
      width: 999
    }));

    const manager = new ColumnManager({
      columns: [{ id: "price", type: "money", width: 90 }],
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
      columns: [{ id: "price", width: 90 }],
      registry: registryWithMoney(),
      hooks
    });

    expect(manager.getColumn(0).width).toBe(190); // (90 * 2) + 10
  });

  it("hands each tap the declared ColumnDefinition and its index", () => {
    const hooks = new HookManager({ registry: PRODUCTION_HOOK_REGISTRY });
    const seen: Array<{ id: string; index: number }> = [];
    hooks.tap<{ definition: { id: string }; index: number }, RuntimeColumn>(
      "resolveColumn",
      { name: "observe" },
      (column: RuntimeColumn, context: { definition: { id: string }; index: number }) => {
        seen.push({ id: context.definition.id, index: context.index });
        return column;
      }
    );

    new ColumnManager({
      columns: [{ id: "a" }, { id: "b" }],
      registry: registryWithMoney(),
      hooks
    });

    expect(seen).toEqual([
      { id: "a", index: 0 },
      { id: "b", index: 1 }
    ]);
  });
});

describe("ColumnManager — resolves once and caches", () => {
  it("only calls resolveColumn once per column at construction, not on every getColumn() call", () => {
    let calls = 0;
    const manager = new ColumnManager({
      columns: [{ id: "a" }, { id: "b" }],
      registry: registryWithMoney(),
      resolveColumn: (column) => {
        calls += 1;
        return column;
      }
    });

    manager.getColumn(0);
    manager.getColumn(0);
    manager.getColumn(1);
    manager.getColumn(1);

    expect(calls).toBe(2);
  });

  it("getColumn() returns the exact same object on repeated calls", () => {
    const manager = new ColumnManager({ columns: [{ id: "a" }], registry: registryWithMoney() });

    expect(manager.getColumn(0)).toBe(manager.getColumn(0));
  });

  it("getColumns() returns all resolved columns in order", () => {
    const manager = new ColumnManager({
      columns: [{ id: "a" }, { id: "b" }, { id: "c" }],
      registry: registryWithMoney()
    });

    expect(manager.getColumns().map((column) => column.id)).toEqual(["a", "b", "c"]);
  });

  it("columnCount reflects the number of resolved columns", () => {
    const manager = new ColumnManager({
      columns: [{ id: "a" }, { id: "b" }],
      registry: registryWithMoney()
    });

    expect(manager.columnCount).toBe(2);
  });
});

describe("ColumnManager — out-of-range access", () => {
  it("getColumn() throws for an index outside the resolved columns", () => {
    const manager = new ColumnManager({ columns: [{ id: "a" }], registry: registryWithMoney() });

    expect(() => manager.getColumn(1)).toThrow(RangeError);
  });
});
