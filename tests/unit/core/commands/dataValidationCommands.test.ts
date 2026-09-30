import { describe, expect, it } from "vitest";

import {
  ClearDataValidationCommand,
  DeleteColumnsCommand,
  DeleteRowsCommand,
  InsertRowsCommand,
  MoveColumnsCommand,
  MoveRowsCommand,
  ReorderRowsCommand,
  SetDataValidationCommand
} from "../../../../src/core/commands";
import { TableStore } from "../../../../src/core/data";
import { CyExcelError, CyExcelErrorCode } from "../../../../src/core/errors";
import { createTableCommandManager } from "../../../../src/core/integration";
import { SheetModel } from "../../../../src/core/sheet/SheetModel";
import type { DataValidationRule } from "../../../../src/core/sheet/validations";
import type { CellRange } from "../../../../src/core/types";
import { createSheetFromRows } from "../../../helpers/createSheetModel";

function range(startRow: number, startColumn: number, endRow: number, endColumn: number): CellRange {
  return { startRow, startColumn, endRow, endColumn };
}

function rule(id: string, ranges: readonly CellRange[]): DataValidationRule {
  return { id, ranges, criteria: { type: "whole", operator: "between", value1: 0, value2: 10 } };
}

/** 6 行 x 4 列。 */
function createStore(): TableStore {
  const fields = ["c0", "c1", "c2", "c3"];
  const rows = Array.from({ length: 6 }, (_unused, r) =>
    Object.fromEntries(fields.map((field, c) => [field, `R${r}C${c}`]))
  );
  return new TableStore(createSheetFromRows(rows, fields));
}

function ranges(store: TableStore, id: string): readonly CellRange[] | undefined {
  return store.sheet.validations().find((r) => r.id === id)?.ranges;
}

describe("SetDataValidationCommand", () => {
  it("stores the rule and emits a validation-layout patch", () => {
    const store = createStore();
    const manager = createTableCommandManager(store);

    const patches = manager.execute(new SetDataValidationCommand(rule("a", [range(0, 0, 3, 0)])));

    expect(patches).toContainEqual({ type: "validation-layout" });
    expect(store.sheet.validations()).toHaveLength(1);
    expect(ranges(store, "a")).toEqual([range(0, 0, 3, 0)]);
  });

  it("carves an older rule and undo restores it exactly", () => {
    const store = createStore();
    const manager = createTableCommandManager(store);
    manager.execute(new SetDataValidationCommand(rule("old", [range(0, 0, 3, 0)])));
    const before = store.sheet.validations();

    manager.execute(new SetDataValidationCommand(rule("new", [range(1, 0, 2, 0)])));
    expect(store.sheet.validations().map((r) => r.id)).toEqual(["old", "new"]);
    expect(ranges(store, "old")).toEqual([range(0, 0, 0, 0), range(3, 0, 3, 0)]);

    manager.undo();
    expect(store.sheet.validations()).toEqual(before);
    manager.redo();
    expect(store.sheet.validations().map((r) => r.id)).toEqual(["old", "new"]);
  });

  it("rejects a malformed rule with InvalidDataValidation and leaves state and history alone", () => {
    const store = createStore();
    const manager = createTableCommandManager(store);

    const bad = { id: "x", ranges: [range(0, 0, 0, 0)], criteria: { type: "list", items: [] } } as unknown as DataValidationRule;
    expect(() => manager.execute(new SetDataValidationCommand(bad))).toThrow(CyExcelError);
    try {
      manager.execute(new SetDataValidationCommand(bad));
    } catch (error) {
      expect((error as CyExcelError).code).toBe(CyExcelErrorCode.InvalidDataValidation);
    }
    expect(store.sheet.validations()).toEqual([]);
    expect(manager.historySize).toBe(0);
  });

  it("freezes the command's params", () => {
    const command = new SetDataValidationCommand(rule("a", [range(0, 0, 0, 0)]));
    expect(Object.isFrozen(command.params)).toBe(true);
    expect(Object.isFrozen(command.params.rule.ranges[0])).toBe(true);
  });

  it("does not touch cell values", () => {
    const store = createStore();
    const manager = createTableCommandManager(store);
    manager.execute(new SetDataValidationCommand(rule("a", [range(0, 0, 5, 3)])));
    expect(store.getCell(2, 2)).toBe("R2C2");
  });
});

describe("ClearDataValidationCommand", () => {
  it("removes covered parts, drops emptied rules, and undoes", () => {
    const store = createStore();
    const manager = createTableCommandManager(store);
    manager.execute(new SetDataValidationCommand(rule("a", [range(0, 0, 3, 0)])));
    manager.execute(new SetDataValidationCommand(rule("b", [range(5, 0, 5, 0)])));
    const before = store.sheet.validations();

    manager.execute(new ClearDataValidationCommand([range(0, 0, 1, 0), range(5, 0, 5, 0)]));
    expect(store.sheet.validations().map((r) => r.id)).toEqual(["a"]);
    expect(ranges(store, "a")).toEqual([range(2, 0, 3, 0)]);

    manager.undo();
    expect(store.sheet.validations()).toEqual(before);
  });

  it("rejects an invalid range", () => {
    const manager = createTableCommandManager(createStore());
    expect(() => manager.execute(new ClearDataValidationCommand([range(3, 0, 1, 0)]))).toThrow(CyExcelError);
  });
});

describe("structural operations carry rules along, and undo brings them back", () => {
  function setup(): { store: TableStore; manager: ReturnType<typeof createTableCommandManager> } {
    const store = createStore();
    const manager = createTableCommandManager(store);
    manager.execute(new SetDataValidationCommand(rule("a", [range(2, 1, 4, 2)])));
    return { store, manager };
  }

  it("insert rows shifts or expands the rule; undo restores it", () => {
    const { store, manager } = setup();
    const before = store.sheet.validations();

    manager.execute(new InsertRowsCommand(3, [[], []]));
    expect(ranges(store, "a")).toEqual([range(2, 1, 6, 2)]); // 插入点在区域内部：撑大

    manager.undo();
    expect(store.sheet.validations()).toEqual(before);
  });

  it("delete rows shrinks the rule, and undo restores the pre-delete rule", () => {
    const { store, manager } = setup();
    const before = store.sheet.validations();

    manager.execute(new DeleteRowsCommand(3, 1));
    expect(ranges(store, "a")).toEqual([range(2, 1, 3, 2)]);

    manager.undo();
    expect(store.sheet.validations()).toEqual(before);
    manager.redo();
    expect(ranges(store, "a")).toEqual([range(2, 1, 3, 2)]);
  });

  it("deleting every row a rule covers removes it, and undo brings it back (validationsBefore)", () => {
    const { store, manager } = setup();
    const before = store.sheet.validations();

    manager.execute(new DeleteRowsCommand(2, 3));
    expect(store.sheet.validations()).toEqual([]);

    manager.undo();
    expect(store.sheet.validations()).toEqual(before);
  });

  it("delete columns removes the rule when fully covered; undo restores it", () => {
    const { store, manager } = setup();
    const before = store.sheet.validations();

    manager.execute(new DeleteColumnsCommand(1, 2));
    expect(store.sheet.validations()).toEqual([]);

    manager.undo();
    expect(store.sheet.validations()).toEqual(before);
  });

  it("move rows is never rejected: the rule splits and follows its rows", () => {
    const { store, manager } = setup();

    // 把 row 3 搬到最前面：区域 rows 2..4 里的 row 3 落到 row 0
    manager.execute(new MoveRowsCommand(3, 1, 0));
    const covered = (r: number, c: number) => store.sheet.validations().some((v) => v.ranges.some((x) => r >= x.startRow && r <= x.endRow && c >= x.startColumn && c <= x.endColumn));
    expect(covered(0, 1)).toBe(true); // 原 row 3
    expect(covered(1, 1)).toBe(false); // 原 row 0
    expect(covered(3, 1)).toBe(true); // 原 row 2 顺延到 3
    expect(covered(4, 1)).toBe(true); // 原 row 4 不动

    manager.undo();
    // 撤销之后覆盖的格子集合相同（区域划分不保证逐字相同）
    for (let r = 0; r < 6; r += 1) {
      expect(covered(r, 1)).toBe(r >= 2 && r <= 4);
    }
  });

  it("move columns moves the rule with the columns", () => {
    const { store, manager } = setup();
    manager.execute(new MoveColumnsCommand(1, 2, 4));
    // 摘出 cols 1..2 插回 4 之前：块落在 cols 2..3
    expect(ranges(store, "a")).toEqual([range(2, 2, 4, 3)]);
    manager.undo();
    expect(ranges(store, "a")).toEqual([range(2, 1, 4, 2)]);
  });

  it("sorting (reorder rows) leaves the rule where it is", () => {
    const { store, manager } = setup();
    const before = store.sheet.validations();
    manager.execute(new ReorderRowsCommand(Int32Array.from([5, 4, 3, 2, 1, 0])));
    expect(store.sheet.validations()).toEqual(before);
  });

  it("a rule set and a structural change in one transaction roll back together", () => {
    const { store, manager } = setup();
    const before = store.sheet.validations();
    expect(() =>
      manager.transaction(() => {
        manager.execute(new SetDataValidationCommand(rule("b", [range(0, 0, 0, 0)])));
        manager.execute(new DeleteRowsCommand(100, 1)); // 越界，整个事务放弃
      })
    ).toThrow();
    expect(store.sheet.validations()).toEqual(before);
  });
});

describe("snapshots", () => {
  it("round-trips rules through toSnapshot / fromSnapshot", () => {
    const store = createStore();
    const manager = createTableCommandManager(store);
    manager.execute(new SetDataValidationCommand(rule("a", [range(0, 0, 3, 0), range(5, 1, 5, 1)])));
    manager.execute(
      new SetDataValidationCommand({
        id: "l",
        ranges: [range(0, 2, 5, 2)],
        criteria: { type: "list", items: ["x", "y"] },
        errorStyle: "warning",
        errorMessage: "选一个"
      })
    );

    const snapshot = store.sheet.toSnapshot();
    expect(snapshot.dataValidations).toHaveLength(2);

    const restored = SheetModel.fromSnapshot(JSON.parse(JSON.stringify(snapshot)));
    expect(restored.validations()).toEqual(store.sheet.validations());
  });

  it("omits the field when there are no rules, so old snapshots stay byte-identical", () => {
    const snapshot = createStore().sheet.toSnapshot();
    expect("dataValidations" in snapshot).toBe(false);
  });

  it("treats a snapshot without the field as having no rules", () => {
    const snapshot = createStore().sheet.toSnapshot();
    expect(SheetModel.fromSnapshot(snapshot).validations()).toEqual([]);
  });

  it("rejects a snapshot whose rules overlap", () => {
    const snapshot = {
      ...createStore().sheet.toSnapshot(),
      dataValidations: [rule("a", [range(0, 0, 2, 2)]), rule("b", [range(1, 1, 3, 3)])]
    };
    expect(() => SheetModel.fromSnapshot(snapshot)).toThrow(CyExcelError);
  });

  it("clone() gives the shadow model the same rules without sharing later edits", () => {
    const store = createStore();
    createTableCommandManager(store).execute(new SetDataValidationCommand(rule("a", [range(0, 0, 0, 0)])));
    const copy = store.sheet.clone();
    copy.setValidations([]);
    expect(store.sheet.validations()).toHaveLength(1);
    expect(copy.validations()).toEqual([]);
  });
});
