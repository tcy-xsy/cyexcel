import { describe, expect, it } from "vitest";
import type { SheetMutation } from "../../../../src/core/sheet/mutation";
import { changesForSheetMutation, type SheetAxisSpace } from "../../../../src/core/sheet/patches";
import type { Patch } from "../../../../src/core/types";

/** 行列都不对称，这样"补丁里的数取自哪根轴"错了就能看出来。 */
const SPACE: SheetAxisSpace = { rowCount: 10, columnCount: 4 };

function onlyPatch(mutation: SheetMutation, space: SheetAxisSpace = SPACE): Patch {
  const patches = changesForSheetMutation(mutation, space);
  expect(patches).toHaveLength(1);
  return patches[0] as Patch;
}

describe("changesForSheetMutation — 十三个成员一条不落", () => {
  const table: readonly { mutation: SheetMutation; patchType: Patch["type"] }[] = [
    { mutation: { type: "set-cell", row: 0, column: 0, record: { v: 1 } }, patchType: "cell-value" },
    {
      mutation: {
        type: "set-range",
        range: { startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 },
        records: [[{ v: 1 }]]
      },
      patchType: "range-value"
    },
    { mutation: { type: "insert-rows", start: 0, rows: [{}] }, patchType: "row-structure" },
    { mutation: { type: "delete-rows", start: 0, count: 1 }, patchType: "row-structure" },
    {
      mutation: { type: "restore-rows", start: 0, rows: [{}], meta: [undefined], tables: {} },
      patchType: "row-structure"
    },
    { mutation: { type: "move-rows", start: 0, count: 1, target: 3 }, patchType: "row-structure" },
    { mutation: { type: "reorder-rows", order: Int32Array.from([1, 0]) }, patchType: "row-reorder" },
    { mutation: { type: "insert-columns", start: 0, count: 1 }, patchType: "column-structure" },
    { mutation: { type: "delete-columns", start: 0, count: 1 }, patchType: "column-structure" },
    {
      mutation: { type: "restore-columns", start: 0, count: 1, cells: [{}], meta: [undefined], tables: {} },
      patchType: "column-structure"
    },
    { mutation: { type: "move-columns", start: 0, count: 1, target: 3 }, patchType: "column-structure" },
    { mutation: { type: "set-row-meta", row: 0, meta: { h: 20 } }, patchType: "row-layout" },
    { mutation: { type: "set-column-meta", column: 0, meta: { w: 80 } }, patchType: "column-layout" }
  ];

  it("每个成员都恰好产出一条补丁，类型按 §5 的映射表", () => {
    // 反空洞：上面那张表少抄一个成员，这一条先红。TypeScript 只保证 `switch` 穷尽，
    // 保证不了这里真的把十三个都过了一遍。
    expect(new Set(table.map((entry) => entry.mutation.type)).size).toBe(13);
    for (const entry of table) {
      expect(onlyPatch(entry.mutation).type).toBe(entry.patchType);
    }
  });
});

/**
 * `restore-*` 映射成 `insert` 是有实质的：它们**确实**把行/列插回去。下游的
 * `remapIndex()` 若不照插入处理，撤销一次删除之后选区与视口会停在旧下标上，而模型
 * 已经整体下移——不报错的那种错位。
 */
describe("changesForSheetMutation — `restore-*` 在下游眼里就是一次插入", () => {
  it("`restore-rows` 与同形状的 `insert-rows` 产出逐字段相同的补丁", () => {
    const rows = [{}, {}, {}];
    const restore = onlyPatch({ type: "restore-rows", start: 2, rows, meta: [undefined, undefined, undefined], tables: {} });
    const insert = onlyPatch({ type: "insert-rows", start: 2, rows });
    expect(restore).toStrictEqual(insert);
    expect(restore).toStrictEqual({ type: "row-structure", operation: "insert", start: 2, count: 3 });
  });

  it("`restore-columns` 与同形状的 `insert-columns` 产出逐字段相同的补丁", () => {
    const restore = onlyPatch({
      type: "restore-columns",
      start: 1,
      count: 2,
      cells: [{}, {}],
      meta: [undefined, undefined],
      tables: {}
    });
    const insert = onlyPatch({ type: "insert-columns", start: 1, count: 2 });
    expect(restore).toStrictEqual(insert);
    expect(restore).toStrictEqual({ type: "column-structure", operation: "insert", start: 1, count: 2 });
  });
});

describe("changesForSheetMutation — 数取自载荷本身", () => {
  it("`insert-rows` 的 count 是 `rows.length`，不是别处的某个数", () => {
    expect(onlyPatch({ type: "insert-rows", start: 4, rows: [{}, {}, {}] })).toStrictEqual({
      type: "row-structure",
      operation: "insert",
      start: 4,
      count: 3
    });
  });

  it("`move-rows` / `move-columns` 带上 target——少带一个下游就不知道搬去哪了", () => {
    expect(onlyPatch({ type: "move-rows", start: 1, count: 2, target: 7 })).toStrictEqual({
      type: "row-structure",
      operation: "move",
      start: 1,
      count: 2,
      target: 7
    });
    expect(onlyPatch({ type: "move-columns", start: 0, count: 1, target: 3 })).toStrictEqual({
      type: "column-structure",
      operation: "move",
      start: 0,
      count: 1,
      target: 3
    });
  });

  it("`set-cell` / `set-range` 原样带着坐标与范围", () => {
    expect(onlyPatch({ type: "set-cell", row: 3, column: 5, record: undefined })).toStrictEqual({
      type: "cell-value",
      row: 3,
      column: 5
    });
    const range = { startRow: 1, startColumn: 2, endRow: 4, endColumn: 6 };
    expect(onlyPatch({ type: "set-range", range, records: [] })).toStrictEqual({
      type: "range-value",
      range
    });
  });
});

/**
 * `types.ts:134` 已经写下的理由：`row-structure` 的每一个消费者都按 `start`/`count`
 * 理解下标位移，而一个排列没有这个形状。折进去会编译通过，然后静默地把下标重映射错。
 */
describe("changesForSheetMutation — `reorder-rows` 带着 order 走", () => {
  it("产出的是 `row-reorder`，且 order 就是传进来的那一个", () => {
    const order = Int32Array.from([2, 0, 1]);
    const patch = onlyPatch({ type: "reorder-rows", order });
    expect(patch.type).toBe("row-reorder");
    if (patch.type !== "row-reorder") {
      throw new Error("unreachable");
    }
    // 这里不拷贝：拷贝发生在 `createChangeSet` 的 `snapshotChange`（`Object.freeze`
    // 对非空 TypedArray 会抛，那边只能拷）。这一层交出去的是同一个数组。
    expect(patch.order).toBe(order);
  });
});

describe("changesForSheetMutation — `set-row-meta` 的 range 从这一行铺到最后一行", () => {
  it("改一行的高度会推走它下面每一行，所以 range 的下界是坐标空间的末行", () => {
    expect(onlyPatch({ type: "set-row-meta", row: 3, meta: { h: 44 } })).toStrictEqual({
      type: "row-layout",
      rows: [3],
      range: { startRow: 3, startColumn: 0, endRow: 9, endColumn: 3 }
    });
  });

  it("range 真的取自传进来的 space，不是被忽略掉的参数", () => {
    // 反空洞：如果实现把 `space` 丢掉、写死一个下界，两次调用会给出同一条 range。
    const narrow = onlyPatch({ type: "set-row-meta", row: 0, meta: undefined }, { rowCount: 2, columnCount: 2 });
    const wide = onlyPatch({ type: "set-row-meta", row: 0, meta: undefined }, { rowCount: 500, columnCount: 9 });
    expect(narrow).not.toStrictEqual(wide);
    expect(wide).toStrictEqual({
      type: "row-layout",
      rows: [0],
      range: { startRow: 0, startColumn: 0, endRow: 499, endColumn: 8 }
    });
  });

  it("空列轴上 endColumn 不会变成 -1", () => {
    // `Math.max(0, columnCount - 1)` 这道下界保护与 `RowHeightMutation` /
    // `LayoutRootMutation` 逐字一致，去掉它会产出一条 endColumn 为 -1 的范围。
    expect(onlyPatch({ type: "set-row-meta", row: 0, meta: { h: 20 } }, { rowCount: 1, columnCount: 0 })).toStrictEqual({
      type: "row-layout",
      rows: [0],
      range: { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 }
    });
  });
});

describe("changesForSheetMutation — `set-column-meta` → `column-layout`", () => {
  it("只带列号，不替这个变体补一个 range", () => {
    // `column-layout` 这个变体本来就没有 `range` 字段（`types.ts:121`），列宽变化的
    // 失效范围由下游按整条列轴处理。`toStrictEqual` 在这里就是"没有多出别的字段"。
    expect(onlyPatch({ type: "set-column-meta", column: 2, meta: { hd: true } })).toStrictEqual({
      type: "column-layout",
      column: 2
    });
  });
});
