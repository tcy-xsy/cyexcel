import { describe, expect, it } from "vitest";
import { IndexedTable } from "../../../../src/core/sheet/IndexedTable";
import {
  assertValidResourceEntry,
  resourceToTable,
  tableToResource
} from "../../../../src/core/sheet/resources";
import { assertValidWorkbookSnapshot } from "../../../../src/core/sheet/serialize";
import { SNAPSHOT_VERSION, type ResourceEntry } from "../../../../src/core/sheet/types";
import { SheetModel } from "../../../../src/core/sheet/SheetModel";

describe("assertValidResourceEntry", () => {
  it("接受一个按行索引的稀疏资源", () => {
    expect(() =>
      assertValidResourceEntry(
        { name: "com.acme.rowKeys", indexedBy: "row", data: { 0: "a", 3: "b" } },
        "resources[0]"
      )
    ).not.toThrow();
  });

  it("接受 indexedBy: none 的不透明 blob，不管 data 长什么样", () => {
    expect(() =>
      assertValidResourceEntry({ name: "x", indexedBy: "none", data: "任意字符串" }, "resources[0]")
    ).not.toThrow();
  });

  it("拒绝空名字——名字是 resources 数组里的身份", () => {
    expect(() =>
      assertValidResourceEntry({ name: "", indexedBy: "none", data: {} }, "resources[0]")
    ).toThrow(/name/);
  });

  it("拒绝不认识的 indexedBy", () => {
    expect(() =>
      assertValidResourceEntry(
        { name: "x", indexedBy: "sheet" as "none", data: {} },
        "resources[0]"
      )
    ).toThrow(/indexedBy/);
  });

  it("按行索引的资源，data 必须是下标映射而不是数组", () => {
    expect(() =>
      assertValidResourceEntry({ name: "x", indexedBy: "row", data: ["a", "b"] }, "resources[0]")
    ).toThrow(/data/);
  });

  it("按行索引的资源，键必须是规范的非负整数字符串", () => {
    expect(() =>
      assertValidResourceEntry({ name: "x", indexedBy: "row", data: { "-1": "a" } }, "resources[0]")
    ).toThrow(/key/);
    expect(() =>
      assertValidResourceEntry({ name: "x", indexedBy: "row", data: { "01": "a" } }, "resources[0]")
    ).toThrow(/key/);
  });

  it("按行索引的资源，data 不能是 null", () => {
    expect(() =>
      assertValidResourceEntry({ name: "x", indexedBy: "row", data: null }, "resources[0]")
    ).toThrow(/data/);
  });
});

describe("resources ⇄ IndexedTable", () => {
  it("往返保持名字、轴与内容", () => {
    const entry = { name: "k", indexedBy: "row" as const, data: { 2: "v" } };
    const table = resourceToTable(entry);
    expect(table.name).toBe("k");
    expect(table.indexedBy).toBe("row");
    expect(table.read()).toStrictEqual({ 2: "v" });
    expect(tableToResource(table)).toStrictEqual(entry);
  });

  it("转出来的表与原 data 不共享对象", () => {
    const data = { 0: "v" };
    const table = resourceToTable({ name: "k", indexedBy: "row", data });
    table.set(1, "w");
    expect(data).toStrictEqual({ 0: "v" });
  });

  it("indexedBy: none 的资源不能转成表——核心不碰它", () => {
    expect(() => resourceToTable({ name: "k", indexedBy: "none", data: {} })).toThrow(/none/);
  });
});

function workbookWith(resources: ResourceEntry[]) {
  const sheet = SheetModel.createEmpty().toSnapshot();
  return {
    version: SNAPSHOT_VERSION,
    sheetOrder: [sheet.id],
    sheets: { [sheet.id]: sheet },
    styles: {},
    resources
  };
}

describe("assertValidWorkbookSnapshot — resources", () => {
  it("拒绝重名的资源", () => {
    expect(() =>
      assertValidWorkbookSnapshot(
        workbookWith([
          { name: "dup", indexedBy: "none", data: 1 },
          { name: "dup", indexedBy: "none", data: 2 }
        ])
      )
    ).toThrow(/dup/);
  });

  it("拒绝按行索引但 data 键非法的资源", () => {
    expect(() =>
      assertValidWorkbookSnapshot(workbookWith([{ name: "x", indexedBy: "row", data: { a: 1 } }]))
    ).toThrow(/key/);
  });
});
