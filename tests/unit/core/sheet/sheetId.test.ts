import { describe, expect, it } from "vitest";
import { WorkbookModel } from "../../../../src/core/sheet/WorkbookModel";
import {
  assertResourceSheetId,
  assertValidResourceEntry
} from "../../../../src/core/sheet/resources";

describe("ResourceEntry.sheetId — 引用完整性", () => {
  it("缺省是合法的：单表下没有歧义（spec §11 不变量 6）", () => {
    expect(() =>
      assertResourceSheetId({ name: "k", indexedBy: "row", data: {} }, "r", ["sheet-1"])
    ).not.toThrow();
  });

  it("给出且与已知 sheet 相符是合法的", () => {
    expect(() =>
      assertResourceSheetId(
        { name: "k", indexedBy: "row", sheetId: "sheet-1", data: {} },
        "r",
        ["sheet-1"]
      )
    ).not.toThrow();
  });

  it("指向不存在的 sheet 必须拒绝——下标失去所指就是静默错位", () => {
    expect(() =>
      assertResourceSheetId(
        { name: "k", indexedBy: "row", sheetId: "nope", data: {} },
        "r",
        ["sheet-1"]
      )
    ).toThrow(/nope/);
  });
});

describe("ResourceEntry.sheetId — 往返", () => {
  it("给出时原样往返", () => {
    const workbook = WorkbookModel.createEmpty();
    const id = workbook.sheet().id();
    workbook.addResource({ name: "k", indexedBy: "row", sheetId: id, data: { 0: "a" } });
    const restored = WorkbookModel.fromSnapshot(workbook.toSnapshot());
    expect(restored.resource("k")?.sheetId).toBe(id);
  });

  it("缺省时**不**被补成显式值——缺省本身是要往返的信息", () => {
    // 补默认值会让"没指定"和"指定了当前这张"在快照里无法区分，
    // 而 M3 放开多 Sheet 后这两者含义不同。
    const workbook = WorkbookModel.createEmpty();
    workbook.addResource({ name: "k", indexedBy: "row", data: { 0: "a" } });
    const entry = workbook.toSnapshot().resources[0] as unknown as Record<string, unknown>;
    expect("sheetId" in entry).toBe(false);
  });
});

describe("ResourceEntry.sheetId — indexedBy: none", () => {
  it("不透明资源不接受 sheetId——它本来就是 workbook 级的", () => {
    // 这条是**形状**检查，不需要 sheet 列表，所以归 assertValidResourceEntry
    expect(() =>
      assertValidResourceEntry({ name: "k", indexedBy: "none", sheetId: "sheet-1", data: 1 }, "r")
    ).toThrow(/none/);
  });
});
