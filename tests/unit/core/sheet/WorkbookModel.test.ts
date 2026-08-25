import { describe, expect, it } from "vitest";
import { WorkbookModel } from "../../../../src/core/sheet/WorkbookModel";
import { SNAPSHOT_VERSION } from "../../../../src/core/sheet/types";

describe("WorkbookModel — 结构", () => {
  it("空工作簿有一张 sheet，版本号是当前 schema 版本", () => {
    const snapshot = WorkbookModel.createEmpty().toSnapshot();
    expect(snapshot.version).toBe(SNAPSHOT_VERSION);
    expect(snapshot.sheetOrder).toHaveLength(1);
    expect(Object.keys(snapshot.sheets)).toEqual([...snapshot.sheetOrder]);
  });

  it("快照里不写 rowCount / columnCount（spec §11 不变量 5）", () => {
    const sheet = WorkbookModel.createEmpty().toSnapshot();
    const first = sheet.sheets[sheet.sheetOrder[0] as string] as unknown as Record<string, unknown>;
    expect("rowCount" in first).toBe(false);
    expect("columnCount" in first).toBe(false);
  });

  it("多于一张 sheet 的快照被拒绝（不变量 6）", () => {
    const base = WorkbookModel.createEmpty().toSnapshot();
    const id = base.sheetOrder[0] as string;
    const two = {
      ...base,
      sheetOrder: [id, "sheet-2"],
      sheets: { ...base.sheets, "sheet-2": base.sheets[id] as never }
    };
    expect(() => WorkbookModel.fromSnapshot(two)).toThrow();
  });
});

describe("WorkbookModel — resources 往返", () => {
  it("indexedBy: none 的资源原样往返", () => {
    const workbook = WorkbookModel.createEmpty();
    workbook.addResource({ name: "blob", indexedBy: "none", data: { anything: [1, 2] } });
    const restored = WorkbookModel.fromSnapshot(workbook.toSnapshot());
    expect(restored.resource("blob")).toStrictEqual({
      name: "blob",
      indexedBy: "none",
      data: { anything: [1, 2] }
    });
  });

  it("按行索引的资源往返，且回来之后仍然登记在 sheet 上", () => {
    const workbook = WorkbookModel.createEmpty();
    workbook.addResource({ name: "rowKeys", indexedBy: "row", data: { 0: "k0", 1: "k1" } });
    const restored = WorkbookModel.fromSnapshot(workbook.toSnapshot());
    expect(restored.sheet().tableNames()).toContain("rowKeys");
    restored.sheet().insertRows(0, 1);
    expect(restored.resource("rowKeys")?.data).toStrictEqual({ 1: "k0", 2: "k1" });
  });

  it("经过 JSON 往返", () => {
    const workbook = WorkbookModel.createEmpty();
    workbook.sheet().setCell(1, 1, "v");
    workbook.addResource({ name: "rowKeys", indexedBy: "row", data: { 1: "k1" } });
    const json = JSON.parse(JSON.stringify(workbook.toSnapshot())) as ReturnType<
      WorkbookModel["toSnapshot"]
    >;
    const restored = WorkbookModel.fromSnapshot(json);
    expect(restored.sheet().getCell(1, 1)).toBe("v");
    expect(restored.resource("rowKeys")?.data).toStrictEqual({ 1: "k1" });
  });

  it("重名资源被拒绝", () => {
    const workbook = WorkbookModel.createEmpty();
    workbook.addResource({ name: "dup", indexedBy: "none", data: 1 });
    expect(() => workbook.addResource({ name: "dup", indexedBy: "none", data: 2 })).toThrow(/dup/);
  });

  it("资源名不能撞上核心自己的表名", () => {
    const workbook = WorkbookModel.createEmpty();
    expect(() => workbook.addResource({ name: "rowData", indexedBy: "row", data: {} })).toThrow(
      /rowData/
    );
  });

  it("resources 数组按加入顺序往返，不按 indexedBy 分组", () => {
    const workbook = WorkbookModel.createEmpty();
    workbook.addResource({ name: "first", indexedBy: "none", data: 1 });
    workbook.addResource({ name: "second", indexedBy: "row", data: { 0: "a" } });
    workbook.addResource({ name: "third", indexedBy: "none", data: 3 });
    const snapshot = workbook.toSnapshot();
    expect(snapshot.resources.map((r) => r.name)).toEqual(["first", "second", "third"]);
    // 往返之后顺序必须还是这个——否则 fromSnapshot(s).toSnapshot() 不等于 s
    const restored = WorkbookModel.fromSnapshot(snapshot).toSnapshot();
    expect(restored.resources.map((r) => r.name)).toEqual(["first", "second", "third"]);
  });

  it("跨 indexedBy 的重名也被拒绝", () => {
    const workbook = WorkbookModel.createEmpty();
    workbook.addResource({ name: "dup", indexedBy: "none", data: 1 });
    expect(() => workbook.addResource({ name: "dup", indexedBy: "row", data: {} })).toThrow(/dup/);
  });
});

describe("spec §7.2 的重键契约 —— 两个消费者，一套机制", () => {
  it("两张各自登记的行表在同一次结构操作里一起被正确重键", () => {
    const workbook = WorkbookModel.createEmpty();
    workbook.sheet().setCell(0, 0, "a");
    workbook.sheet().setCell(1, 0, "b");
    workbook.sheet().setRowMeta(1, { h: 40 });
    // 消费者一：业务主键（spec §7.1）
    workbook.addResource({ name: "rowKeys", indexedBy: "row", data: { 0: "pk-0", 1: "pk-1" } });
    // 消费者二：一组行下标，形如 ADR 0003 §3.2 的 filteredRows
    workbook.addResource({ name: "filtered", indexedBy: "row", data: { 1: true } });

    workbook.sheet().insertRows(0, 2);

    expect(workbook.sheet().getCell(2, 0)).toBe("a");
    expect(workbook.sheet().getRowMeta(3)).toStrictEqual({ h: 40 });
    expect(workbook.resource("rowKeys")?.data).toStrictEqual({ 2: "pk-0", 3: "pk-1" });
    expect(workbook.resource("filtered")?.data).toStrictEqual({ 3: true });
  });

  it("四个行操作都带着资源走，一个都不漏", () => {
    const build = () => {
      const workbook = WorkbookModel.createEmpty();
      workbook.sheet().setCell(0, 0, "a");
      workbook.sheet().setCell(1, 0, "b");
      workbook.sheet().setCell(2, 0, "c");
      workbook.addResource({
        name: "rowKeys",
        indexedBy: "row",
        data: { 0: "k0", 1: "k1", 2: "k2" }
      });
      return workbook;
    };

    const inserted = build();
    inserted.sheet().insertRows(1, 1);
    expect(inserted.resource("rowKeys")?.data).toStrictEqual({ 0: "k0", 2: "k1", 3: "k2" });

    const deleted = build();
    deleted.sheet().deleteRows(0, 1);
    expect(deleted.resource("rowKeys")?.data).toStrictEqual({ 0: "k1", 1: "k2" });

    const moved = build();
    moved.sheet().moveRows(0, 1, 2);
    expect(moved.resource("rowKeys")?.data).toStrictEqual({ 0: "k1", 1: "k0", 2: "k2" });

    const reordered = build();
    reordered.sheet().reorderRows(Int32Array.from([2, 1, 0]));
    expect(reordered.resource("rowKeys")?.data).toStrictEqual({ 0: "k2", 1: "k1", 2: "k0" });
  });

  it("三个列操作都带着按列索引的资源走", () => {
    const workbook = WorkbookModel.createEmpty();
    workbook.sheet().setCell(0, 0, "a");
    workbook.addResource({ name: "colTag", indexedBy: "column", data: { 0: "t0", 1: "t1" } });

    workbook.sheet().insertColumns(0, 1);
    expect(workbook.resource("colTag")?.data).toStrictEqual({ 1: "t0", 2: "t1" });

    workbook.sheet().deleteColumns(0, 1);
    expect(workbook.resource("colTag")?.data).toStrictEqual({ 0: "t0", 1: "t1" });

    workbook.sheet().moveColumns(0, 1, 2);
    expect(workbook.resource("colTag")?.data).toStrictEqual({ 0: "t1", 1: "t0" });
  });

  it("indexedBy: none 的资源不被结构操作碰", () => {
    const workbook = WorkbookModel.createEmpty();
    workbook.addResource({ name: "blob", indexedBy: "none", data: { 0: "untouched" } });
    workbook.sheet().insertRows(0, 5);
    expect(workbook.resource("blob")?.data).toStrictEqual({ 0: "untouched" });
  });
});
