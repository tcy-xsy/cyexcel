import { describe, expect, it } from "vitest";

import { CyExcelError, CyExcelErrorCode } from "../../../../src/core/errors";
import {
  assertMergesDoNotOverlap,
  buildMergeClearGrid,
  expandRangeForMerges,
  findMergeAt,
  mergesIntersectingRange,
  rangeContains,
  rangesIntersect,
  relocateMergesForDelete,
  relocateMergesForInsert,
  relocateMergesForMove,
  resolveMergeCells
} from "../../../../src/core/sheet/merges";
import type { CellRange } from "../../../../src/core/types";

function range(startRow: number, startColumn: number, endRow: number, endColumn: number): CellRange {
  return { startRow, startColumn, endRow, endColumn };
}

describe("findMergeAt", () => {
  const merges = [range(1, 1, 2, 2), range(5, 0, 5, 3)];

  it("returns the merge whose range contains the point", () => {
    expect(findMergeAt(merges, 2, 1)).toEqual(range(1, 1, 2, 2));
    expect(findMergeAt(merges, 5, 3)).toEqual(range(5, 0, 5, 3));
  });

  it("returns undefined when no merge contains the point", () => {
    expect(findMergeAt(merges, 0, 0)).toBeUndefined();
    expect(findMergeAt(merges, 3, 1)).toBeUndefined();
  });
});

describe("rangesIntersect", () => {
  it("is true when the ranges overlap", () => {
    expect(rangesIntersect(range(0, 0, 2, 2), range(1, 1, 3, 3))).toBe(true);
  });

  it("is true when one range fully contains the other", () => {
    expect(rangesIntersect(range(0, 0, 5, 5), range(1, 1, 2, 2))).toBe(true);
  });

  it("is false for ranges that only touch on an edge without sharing a cell", () => {
    // row 3 谁都不占：a 止于 row 2，b 起于 row 3，行轴上没有公共格
    expect(rangesIntersect(range(0, 0, 2, 2), range(3, 0, 4, 2))).toBe(false);
  });

  it("is false for ranges disjoint on the column axis", () => {
    expect(rangesIntersect(range(0, 0, 2, 2), range(0, 3, 2, 4))).toBe(false);
  });
});

describe("rangeContains", () => {
  it("is true when outer strictly encloses inner", () => {
    expect(rangeContains(range(0, 0, 5, 5), range(1, 1, 2, 2))).toBe(true);
  });

  it("is true (reflexive) when the two ranges are equal", () => {
    expect(rangeContains(range(1, 1, 2, 2), range(1, 1, 2, 2))).toBe(true);
  });

  it("is false when the ranges only partially overlap", () => {
    expect(rangeContains(range(0, 0, 2, 2), range(1, 1, 3, 3))).toBe(false);
  });

  it("is false when outer is smaller than inner", () => {
    expect(rangeContains(range(1, 1, 2, 2), range(0, 0, 5, 5))).toBe(false);
  });
});

describe("mergesIntersectingRange", () => {
  it("returns an empty array for an empty merges list", () => {
    expect(mergesIntersectingRange([], range(0, 0, 5, 5))).toEqual([]);
  });

  it("returns an empty array when no merge intersects the range", () => {
    const merges = [range(0, 0, 1, 1), range(10, 10, 11, 11)];
    expect(mergesIntersectingRange(merges, range(3, 3, 4, 4))).toEqual([]);
  });

  it("returns only the merges that partially overlap the range", () => {
    const overlapping = range(1, 1, 2, 2);
    const disjoint = range(10, 10, 11, 11);
    expect(mergesIntersectingRange([overlapping, disjoint], range(0, 0, 1, 1))).toEqual([overlapping]);
  });

  it("returns a merge that fully contains the range", () => {
    const merge = range(0, 0, 5, 5);
    expect(mergesIntersectingRange([merge], range(2, 2, 3, 3))).toEqual([merge]);
  });
});

describe("expandRangeForMerges", () => {
  it("returns the range unchanged when no merge intersects it", () => {
    const original = range(3, 3, 4, 4);
    expect(expandRangeForMerges([range(0, 0, 1, 1)], original)).toEqual(original);
  });

  it("grows the range to enclose a single intersecting merge", () => {
    // range [2,2]-[3,3] touches a merge at [3,3]-[5,5] on its bottom-right corner.
    const merge = range(3, 3, 5, 5);
    expect(expandRangeForMerges([merge], range(2, 2, 3, 3))).toEqual(range(2, 2, 5, 5));
  });

  it("does not shrink or otherwise alter a range that already fully contains the merge", () => {
    const merge = range(1, 1, 2, 2);
    const original = range(0, 0, 5, 5);
    expect(expandRangeForMerges([merge], original)).toEqual(original);
  });

  it("chains: growing to enclose one merge can newly intersect a second merge, which also gets enclosed", () => {
    // range [0,0]-[1,1] touches mergeA at [1,1]-[2,2]; growing to [0,0]-[2,2]
    // then newly touches mergeB at [2,2]-[4,4], which must also get folded in.
    const mergeA = range(1, 1, 2, 2);
    const mergeB = range(2, 2, 4, 4);
    expect(expandRangeForMerges([mergeA, mergeB], range(0, 0, 1, 1))).toEqual(range(0, 0, 4, 4));
  });
});

describe("assertMergesDoNotOverlap", () => {
  it("does not throw for a disjoint candidate", () => {
    expect(() => assertMergesDoNotOverlap([range(0, 0, 1, 1)], range(3, 3, 4, 4))).not.toThrow();
  });

  it("does not throw when the candidate fully contains an existing merge", () => {
    expect(() =>
      assertMergesDoNotOverlap([range(1, 1, 2, 2)], range(0, 0, 5, 5))
    ).not.toThrow();
  });

  it("does not throw when an existing merge fully contains the candidate", () => {
    expect(() =>
      assertMergesDoNotOverlap([range(0, 0, 5, 5)], range(1, 1, 2, 2))
    ).not.toThrow();
  });

  it("throws MergeOverlapsExistingMerge for a partial overlap", () => {
    try {
      assertMergesDoNotOverlap([range(0, 0, 2, 2)], range(1, 1, 3, 3));
      expect.unreachable("should have thrown");
    } catch (error) {
      expect(error).toBeInstanceOf(CyExcelError);
      expect((error as CyExcelError).code).toBe(CyExcelErrorCode.MergeOverlapsExistingMerge);
    }
  });
});

describe("resolveMergeCells", () => {
  it("rejects a candidate that partially overlaps an existing merge", () => {
    expect(() => resolveMergeCells([range(0, 0, 2, 2)], range(1, 1, 3, 3))).toThrow(CyExcelError);
  });

  it("absorbs an existing merge the new range fully contains", () => {
    const existing = [range(1, 1, 2, 2)];
    const result = resolveMergeCells(existing, range(0, 0, 5, 5));
    expect(result).toEqual([range(0, 0, 5, 5)]);
  });

  it("keeps unrelated merges untouched and appends the new range", () => {
    const untouched = range(8, 8, 9, 9);
    const result = resolveMergeCells([untouched], range(0, 0, 1, 1));
    expect(result).toEqual([untouched, range(0, 0, 1, 1)]);
  });

  it("treats re-merging an identical range as replacing the equal entry", () => {
    const existing = range(1, 1, 2, 2);
    const result = resolveMergeCells([existing], range(1, 1, 2, 2));
    expect(result).toEqual([range(1, 1, 2, 2)]);
  });
});

describe("buildMergeClearGrid", () => {
  it("is undefined throughout — every cell in the range gets cleared", () => {
    // 锚点不在这张网格的职责里：`stageMergeClear` 本来就把锚点排除在写入范围
    // 之外，而 `CellValue` 收紧之后 `null` 也不再是"清空"的写法。
    expect(buildMergeClearGrid(range(0, 0, 1, 2))).toEqual([
      [undefined, undefined, undefined],
      [undefined, undefined, undefined]
    ]);
  });

  it("is a single undefined cell for a 1x1 range (no merge, degenerate input)", () => {
    expect(buildMergeClearGrid(range(3, 3, 3, 3))).toEqual([[undefined]]);
  });
});

describe("relocateMergesForInsert", () => {
  describe("row axis", () => {
    const merge = range(2, 1, 4, 1); // rows 2-4, columns 未受影响

    it("shifts the whole merge when the insertion point is at or before its start", () => {
      expect(relocateMergesForInsert([merge], "row", 2, 2)).toEqual([range(4, 1, 6, 1)]);
    });

    it("shifts the whole merge when the insertion point is strictly before its start", () => {
      expect(relocateMergesForInsert([merge], "row", 1, 2)).toEqual([range(4, 1, 6, 1)]);
    });

    it("leaves the merge untouched when the insertion point is after its end", () => {
      expect(relocateMergesForInsert([merge], "row", 6, 2)).toEqual([merge]);
    });

    it("expands the merge when the insertion point falls strictly inside it", () => {
      expect(relocateMergesForInsert([merge], "row", 3, 2)).toEqual([range(2, 1, 6, 1)]);
    });
  });

  describe("column axis", () => {
    const merge = range(1, 2, 1, 4); // columns 2-4, 行不受影响

    it("shifts the whole merge when the insertion point is at its start", () => {
      expect(relocateMergesForInsert([merge], "column", 2, 3)).toEqual([range(1, 5, 1, 7)]);
    });

    it("leaves the merge untouched when the insertion point is after its end", () => {
      expect(relocateMergesForInsert([merge], "column", 5, 3)).toEqual([merge]);
    });

    it("expands the merge when the insertion point falls strictly inside it", () => {
      expect(relocateMergesForInsert([merge], "column", 3, 3)).toEqual([range(1, 2, 1, 7)]);
    });
  });
});

describe("relocateMergesForDelete", () => {
  describe("row axis", () => {
    const merge = range(3, 0, 6, 0); // rows 3-6

    it("shifts the merge up when the deleted band is entirely before it", () => {
      expect(relocateMergesForDelete([merge], "row", 0, 2)).toEqual([range(1, 0, 4, 0)]);
    });

    it("leaves the merge untouched when the deleted band is entirely after it", () => {
      expect(relocateMergesForDelete([merge], "row", 8, 2)).toEqual([merge]);
    });

    it("removes the merge entirely when the deleted band fully covers it", () => {
      expect(relocateMergesForDelete([merge], "row", 2, 6)).toEqual([]);
    });

    it("shrinks and closes the gap when the deleted band cuts the upper edge (删上沿)", () => {
      // 删除 rows [2,4]：merge 的 3,4 被删掉，剩下的 5,6 整体上移 3 位 => [2,3]
      expect(relocateMergesForDelete([merge], "row", 2, 3)).toEqual([range(2, 0, 3, 0)]);
    });

    it("shrinks and closes the gap when the deleted band cuts the lower edge (删下沿)", () => {
      // 删除 rows [5,7]：merge 的 5,6 被删掉，剩下的 3,4 位置不变 => [3,4]
      expect(relocateMergesForDelete([merge], "row", 5, 3)).toEqual([range(3, 0, 4, 0)]);
    });

    it("shrinks and closes the gap when the deleted band cuts the middle (删中段)", () => {
      // 删除 row [4,4]：merge 的 4 被删掉，5,6 各上移 1 位 => 3,4,5 => [3,5]
      expect(relocateMergesForDelete([merge], "row", 4, 1)).toEqual([range(3, 0, 5, 0)]);
    });
  });

  describe("column axis", () => {
    const merge = range(0, 3, 0, 6); // columns 3-6

    it("shifts the merge left when the deleted band is entirely before it", () => {
      expect(relocateMergesForDelete([merge], "column", 0, 2)).toEqual([range(0, 1, 0, 4)]);
    });

    it("leaves the merge untouched when the deleted band is entirely after it", () => {
      expect(relocateMergesForDelete([merge], "column", 8, 2)).toEqual([merge]);
    });

    it("removes the merge entirely when the deleted band fully covers it", () => {
      expect(relocateMergesForDelete([merge], "column", 2, 6)).toEqual([]);
    });

    it("shrinks and closes the gap on a partial overlap", () => {
      // 删除 columns [2,4]：merge 的 3,4 被删掉，剩下的 5,6 整体左移 3 位 => [2,3]
      expect(relocateMergesForDelete([merge], "column", 2, 3)).toEqual([range(0, 2, 0, 3)]);
    });
  });
});

describe("relocateMergesForMove", () => {
  describe("row axis", () => {
    it("carries a merge that sits entirely inside the moved block", () => {
      // block = rows [2,3]，target=6 => insertAt = 6-2 = 4，delta = 4-2 = 2
      const merge = range(2, 0, 3, 0);
      const result = relocateMergesForMove([merge], "row", 2, 2, 6);
      expect(result).toEqual({ ok: true, merges: [range(4, 0, 5, 0)] });
    });

    it("rejects a merge that straddles the move boundary", () => {
      // block = rows [2,3]；merge rows [1,2] 与块相交但没有整体落在块内
      const merge = range(1, 0, 2, 0);
      const result = relocateMergesForMove([merge], "row", 2, 2, 6);
      expect(result).toEqual({ ok: false, reason: "straddles-move-boundary" });
    });

    it("composes delete-then-insert for a merge fully outside the block", () => {
      // block = rows [2,3]（先摘除），target=10 => insertAt = 10-2 = 8（再插回）。
      // merge rows [5,6] 在块之后：摘除后左移 2 位 => [3,4]；插回点(8) 在 [3,4] 之后
      // 不受影响 => 最终 [3,4]（块从merge前面搬到了merge后面，merge整体让出2行）。
      const merge = range(5, 0, 6, 0);
      const result = relocateMergesForMove([merge], "row", 2, 2, 10);
      expect(result).toEqual({ ok: true, merges: [range(3, 0, 4, 0)] });
    });

    it("rejects a merge fully outside the block whose reinsertion point would land strictly inside it (final-review Finding 1)", () => {
      // merge rows [2,6]；block = rows [8,9]（start=8, count=2），target=4 =>
      // insertAt = 4（target(4) 不大于 start(8)，取直取支）。
      // outsideBlock：mEnd(6) < blockStart(8) 成立。
      // afterDelete：删除区间 [8,9] 整体在 merge 之后 => 不受影响，仍是 [2,6]。
      // insertAt(4) 严格落在 [2,6] 内部（2 < 4 <= 6）——这正是
      // `relocateIntervalForInsert` 的撑大分支，撑大后块会落在合并区内部，逆向
      // 搬回去必然撞上"跨越移动块边界"（504-组合 brute-force 里的代表实例）。
      // 修复前：{ ok: true, merges: [range(2, 0, 8, 0)] }（错误地撑大）。
      const merge = range(2, 0, 6, 0);
      const result = relocateMergesForMove([merge], "row", 8, 2, 4);
      expect(result).toEqual({ ok: false, reason: "straddles-move-boundary" });
    });

    it("still allows a merge fully outside the block when the reinsertion point lands at its start (shift, not expand)", () => {
      // 插入点恰好等于摘除后区间的起点——`relocateIntervalForInsert` 的"平移"支
      // （start <= mStart），不是撑大，仍然允许。
      // merge rows [4,6]；block = rows [8,9]，target=4 => insertAt=4。
      // afterDelete：删除区间 [8,9] 在 merge 之后，不受影响，仍是 [4,6]。
      // insertAt(4) <= afterDeleteStart(4)：平移支，merge 整体 +count(2) => [6,8]。
      const merge = range(4, 0, 6, 0);
      const result = relocateMergesForMove([merge], "row", 8, 2, 4);
      expect(result).toEqual({ ok: true, merges: [range(6, 0, 8, 0)] });
    });
  });

  describe("column axis", () => {
    it("carries a merge that sits entirely inside the moved block", () => {
      const merge = range(0, 2, 0, 3);
      const result = relocateMergesForMove([merge], "column", 2, 2, 6);
      expect(result).toEqual({ ok: true, merges: [range(0, 4, 0, 5)] });
    });

    it("rejects a merge that straddles the move boundary", () => {
      const merge = range(0, 1, 0, 2);
      const result = relocateMergesForMove([merge], "column", 2, 2, 6);
      expect(result).toEqual({ ok: false, reason: "straddles-move-boundary" });
    });

    it("composes delete-then-insert for a merge fully outside the block", () => {
      const merge = range(0, 5, 0, 6);
      const result = relocateMergesForMove([merge], "column", 2, 2, 10);
      expect(result).toEqual({ ok: true, merges: [range(0, 3, 0, 4)] });
    });

    it("rejects a merge fully outside the block whose reinsertion point would land strictly inside it (final-review Finding 1)", () => {
      // 与行轴那条同构：merge columns [2,6]，block columns [8,9]，target=4。
      const merge = range(0, 2, 0, 6);
      const result = relocateMergesForMove([merge], "column", 8, 2, 4);
      expect(result).toEqual({ ok: false, reason: "straddles-move-boundary" });
    });
  });
});
