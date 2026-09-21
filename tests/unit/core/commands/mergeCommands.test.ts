import { describe, expect, it } from "vitest";

import { MergeCellsCommand, UnmergeCellsCommand } from "../../../../src/core/commands";
import { TableStore } from "../../../../src/core/data";
import { CyExcelError, CyExcelErrorCode } from "../../../../src/core/errors";
import { createTableCommandManager } from "../../../../src/core/integration";
import type { CellRange } from "../../../../src/core/types";
import { createSheetFromRows } from "../../../helpers/createSheetModel";

function range(startRow: number, startColumn: number, endRow: number, endColumn: number): CellRange {
  return { startRow, startColumn, endRow, endColumn };
}

/** 5 行 x 4 列，每格的值是它自己的坐标标签，方便断言"哪些格子被清空、哪些没动"。 */
function createStore(): { store: TableStore } {
  const fields = ["c0", "c1", "c2", "c3"];
  const rows = Array.from({ length: 5 }, (_unused, r) =>
    Object.fromEntries(fields.map((field, c) => [field, `R${r}C${c}`]))
  );
  return { store: new TableStore(createSheetFromRows(rows, fields)) };
}

describe("MergeCellsCommand", () => {
  it("records the merge and emits a merge-layout patch", () => {
    const { store } = createStore();
    const manager = createTableCommandManager(store);

    const patches = manager.execute(new MergeCellsCommand(range(0, 0, 1, 1)));

    expect(patches).toContainEqual({ type: "merge-layout" });
    expect(store.sheet.merges()).toEqual([range(0, 0, 1, 1)]);
  });

  it("clears every non-anchor cell in the range but preserves the anchor's value", () => {
    const { store } = createStore();
    const manager = createTableCommandManager(store);

    manager.execute(new MergeCellsCommand(range(0, 0, 1, 1)));

    // 清空 = 没有记录 = 读出来是 undefined。`CellValue` 里没有 `null` 了
    // （对齐 xlsx/Univer），空格子只有这一种表示。
    expect(store.getCell(0, 0)).toBe("R0C0"); // anchor untouched
    expect(store.getCell(0, 1)).toBeUndefined();
    expect(store.getCell(1, 0)).toBeUndefined();
    expect(store.getCell(1, 1)).toBeUndefined();
    // Outside the merge range, nothing changed.
    expect(store.getCell(0, 2)).toBe("R0C2");
    expect(store.getCell(2, 0)).toBe("R2C0");
  });

  it("does not stage any clear write for a degenerate 1x1 range (anchor is the only cell)", () => {
    const { store } = createStore();
    const manager = createTableCommandManager(store);

    const patches = manager.execute(new MergeCellsCommand(range(0, 0, 0, 0)));

    expect(patches).toEqual([{ type: "merge-layout" }]);
    expect(store.getCell(0, 0)).toBe("R0C0");
  });

  it("rejects a range that partially overlaps an existing merge, leaving state untouched", () => {
    const { store } = createStore();
    const manager = createTableCommandManager(store);
    manager.execute(new MergeCellsCommand(range(0, 0, 1, 1)));

    expect(() => manager.execute(new MergeCellsCommand(range(1, 1, 2, 2)))).toThrow(CyExcelError);
    try {
      manager.execute(new MergeCellsCommand(range(1, 1, 2, 2)));
    } catch (error) {
      expect((error as CyExcelError).code).toBe(CyExcelErrorCode.MergeOverlapsExistingMerge);
    }

    // Nothing from the rejected attempt leaked through.
    expect(store.sheet.merges()).toEqual([range(0, 0, 1, 1)]);
    expect(store.getCell(2, 2)).toBe("R2C2");
    expect(manager.historySize).toBe(1);
  });

  it("rejects an inverted range (endRow before startRow) with InvalidRange (final-review M1)", () => {
    const { store } = createStore();
    const manager = createTableCommandManager(store);

    expect(() => manager.execute(new MergeCellsCommand(range(2, 0, 1, 0)))).toThrow(CyExcelError);
    try {
      manager.execute(new MergeCellsCommand(range(2, 0, 1, 0)));
    } catch (error) {
      expect((error as CyExcelError).code).toBe(CyExcelErrorCode.InvalidRange);
    }
    expect(store.sheet.merges()).toEqual([]);
  });

  it("absorbs an existing merge that the new range fully contains", () => {
    const { store } = createStore();
    const manager = createTableCommandManager(store);
    manager.execute(new MergeCellsCommand(range(0, 0, 0, 1)));

    manager.execute(new MergeCellsCommand(range(0, 0, 2, 2)));

    expect(store.sheet.merges()).toEqual([range(0, 0, 2, 2)]);
  });

  it("leaves an unrelated, non-intersecting merge untouched", () => {
    const { store } = createStore();
    const manager = createTableCommandManager(store);
    manager.execute(new MergeCellsCommand(range(3, 0, 3, 1)));

    manager.execute(new MergeCellsCommand(range(0, 0, 1, 1)));

    expect(store.sheet.merges()).toEqual(
      expect.arrayContaining([range(3, 0, 3, 1), range(0, 0, 1, 1)])
    );
    expect(store.sheet.merges()).toHaveLength(2);
  });

  it("undo restores both the merge table and the values it cleared", () => {
    const { store } = createStore();
    const manager = createTableCommandManager(store);
    manager.execute(new MergeCellsCommand(range(0, 0, 1, 1)));

    manager.undo();

    expect(store.sheet.merges()).toEqual([]);
    expect(store.getCell(0, 1)).toBe("R0C1");
    expect(store.getCell(1, 0)).toBe("R1C0");
    expect(store.getCell(1, 1)).toBe("R1C1");
  });

  it("redo re-applies the merge and re-clears the same cells", () => {
    const { store } = createStore();
    const manager = createTableCommandManager(store);
    manager.execute(new MergeCellsCommand(range(0, 0, 1, 1)));
    manager.undo();

    manager.redo();

    expect(store.sheet.merges()).toEqual([range(0, 0, 1, 1)]);
    expect(store.getCell(0, 1)).toBeUndefined();
  });
});

describe("UnmergeCellsCommand", () => {
  function createStoreWithTwoMerges(): { store: TableStore; manager: ReturnType<typeof createTableCommandManager> } {
    const { store } = createStore();
    const manager = createTableCommandManager(store);
    manager.execute(new MergeCellsCommand(range(0, 0, 1, 1)));
    manager.execute(new MergeCellsCommand(range(3, 0, 3, 2)));
    return { store, manager };
  }

  it("removes every merge intersecting the given range", () => {
    const { store, manager } = createStoreWithTwoMerges();

    const patches = manager.execute(new UnmergeCellsCommand(range(1, 1, 1, 1)));

    expect(patches).toEqual([{ type: "merge-layout" }]);
    expect(store.sheet.merges()).toEqual([range(3, 0, 3, 2)]);
  });

  it("removes multiple merges at once when the range intersects all of them", () => {
    const { store, manager } = createStoreWithTwoMerges();

    manager.execute(new UnmergeCellsCommand(range(0, 0, 3, 2)));

    expect(store.sheet.merges()).toEqual([]);
  });

  it("leaves the merge table untouched when the range intersects nothing", () => {
    const { store, manager } = createStoreWithTwoMerges();

    manager.execute(new UnmergeCellsCommand(range(4, 0, 4, 3)));

    expect(store.sheet.merges()).toEqual(
      expect.arrayContaining([range(0, 0, 1, 1), range(3, 0, 3, 2)])
    );
    expect(store.sheet.merges()).toHaveLength(2);
  });

  it("does not touch cell values — only the merge table changes", () => {
    const { store, manager } = createStoreWithTwoMerges();

    manager.execute(new UnmergeCellsCommand(range(0, 0, 1, 1)));

    expect(store.getCell(0, 0)).toBe("R0C0");
    expect(store.getCell(0, 1)).toBeUndefined(); // cleared earlier by the merge itself, unmerge doesn't restore it
  });

  it("rejects an inverted range (endColumn before startColumn) with InvalidRange (final-review M1)", () => {
    const { store, manager } = createStoreWithTwoMerges();

    expect(() => manager.execute(new UnmergeCellsCommand(range(0, 1, 0, 0)))).toThrow(CyExcelError);
    try {
      manager.execute(new UnmergeCellsCommand(range(0, 1, 0, 0)));
    } catch (error) {
      expect((error as CyExcelError).code).toBe(CyExcelErrorCode.InvalidRange);
    }
    // Both existing merges are untouched by the rejected attempt.
    expect(store.sheet.merges()).toEqual(
      expect.arrayContaining([range(0, 0, 1, 1), range(3, 0, 3, 2)])
    );
    expect(store.sheet.merges()).toHaveLength(2);
  });

  it("undo restores the merge table to include the removed merge again", () => {
    const { store, manager } = createStoreWithTwoMerges();
    manager.execute(new UnmergeCellsCommand(range(1, 1, 1, 1)));

    manager.undo();

    expect(store.sheet.merges()).toEqual(
      expect.arrayContaining([range(0, 0, 1, 1), range(3, 0, 3, 2)])
    );
    expect(store.sheet.merges()).toHaveLength(2);
  });
});
