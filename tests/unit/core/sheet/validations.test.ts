import { describe, expect, it } from "vitest";

import { CyExcelError, CyExcelErrorCode } from "../../../../src/core/errors";
import {
  applyRule,
  assertValidRule,
  assertValidRules,
  buildRule,
  clearRanges,
  collectInvalidCells,
  describeCellDropdown,
  describeDropdownItems,
  defaultValidationMessage,
  findRuleAt,
  intersectRanges,
  relocateValidationsForDelete,
  relocateValidationsForInsert,
  relocateValidationsForMove,
  rulesIntersecting,
  subtractRange,
  toDisjointRanges,
  validateValue,
  type DataValidationRule
} from "../../../../src/core/sheet/validations";
import type { CellRange } from "../../../../src/core/types";

function range(startRow: number, startColumn: number, endRow: number, endColumn: number): CellRange {
  return { startRow, startColumn, endRow, endColumn };
}

function whole(id: string, ranges: readonly CellRange[], value1 = 0, value2 = 10): DataValidationRule {
  return { id, ranges, criteria: { type: "whole", operator: "between", value1, value2 } };
}

function freezeRules(rules: DataValidationRule[]): readonly DataValidationRule[] {
  return Object.freeze(rules.map(rule => Object.freeze({ ...rule,
    ranges: Object.freeze(rule.ranges.map(r => Object.freeze(r))) })));
}

/** 把一组规则展开成 "行,列 -> 规则 id" 的映射，便于逐格比较覆盖范围。 */
function coverage(rules: readonly DataValidationRule[]): Map<string, string> {
  const cells = new Map<string, string>();
  for (const rule of rules) {
    for (const r of rule.ranges) {
      for (let row = r.startRow; row <= r.endRow; row += 1) {
        for (let column = r.startColumn; column <= r.endColumn; column += 1) {
          const key = `${row},${column}`;
          if (cells.has(key)) {
            throw new Error(`cell ${key} is covered twice`);
          }
          cells.set(key, rule.id);
        }
      }
    }
  }
  return cells;
}

describe("subtractRange", () => {
  it("returns the range itself when the cut does not intersect", () => {
    expect(subtractRange(range(0, 0, 2, 2), range(5, 5, 6, 6))).toEqual([range(0, 0, 2, 2)]);
  });

  it("returns nothing when the cut covers the range", () => {
    expect(subtractRange(range(1, 1, 2, 2), range(0, 0, 5, 5))).toEqual([]);
  });

  it("leaves four disjoint pieces when the cut is strictly inside", () => {
    const pieces = subtractRange(range(0, 0, 4, 4), range(2, 2, 2, 2));
    expect(pieces).toHaveLength(4);
    const cells = coverage([{ id: "x", ranges: pieces, criteria: { type: "list", items: ["a"] } }]);
    expect(cells.size).toBe(24);
    expect(cells.has("2,2")).toBe(false);
  });
});

describe("toDisjointRanges", () => {
  it("keeps the earlier range whole and trims the later one", () => {
    const result = toDisjointRanges([range(0, 0, 1, 1), range(1, 1, 2, 2)]);
    expect(result[0]).toEqual(range(0, 0, 1, 1));
    const cells = coverage([{ id: "x", ranges: result, criteria: { type: "list", items: ["a"] } }]);
    expect(cells.size).toBe(7); // 4 + (9 - 2 ... 3x3 minus the shared cell) => 4 + 3
  });
});

describe("findRuleAt / rulesIntersecting", () => {
  const rules = [whole("a", [range(0, 0, 1, 1)]), whole("b", [range(5, 5, 6, 6), range(9, 0, 9, 0)])];

  it("finds the rule containing the cell, across multi-range rules", () => {
    expect(findRuleAt(rules, 1, 1)?.id).toBe("a");
    expect(findRuleAt(rules, 9, 0)?.id).toBe("b");
    expect(findRuleAt(rules, 3, 3)).toBeUndefined();
  });

  it("lists only rules that intersect the window", () => {
    expect(rulesIntersecting(rules, range(4, 4, 7, 7)).map((r) => r.id)).toEqual(["b"]);
    expect(rulesIntersecting(rules, range(20, 20, 30, 30))).toEqual([]);
  });

  it("indexes large rule sets without scanning offscreen ranges on each lookup", () => {
    let reads = 0;
    const many = freezeRules(Array.from({ length: 10_000 }, (_, column) => whole(String(column), [{
      get startRow() { reads += 1; return 0; },
      endRow: 100,
      startColumn: column,
      endColumn: column
    }])));
    expect(findRuleAt(many, 50, 9999)?.id).toBe("9999");
    reads = 0;
    expect(findRuleAt(many, 75, 8000)?.id).toBe("8000");
    expect(reads).toBeLessThan(50);
    expect(findRuleAt(many, 101, 8000)).toBeUndefined();
    const changed = freezeRules([...many, whole("new", [range(101, 8000, 102, 8000)])]);
    expect(findRuleAt(changed, 101, 8000)?.id).toBe("new");
  });

  it("preserves first-match precedence and multi-range lookup in the index", () => {
    const indexed = freezeRules([
      whole("first", [range(4, 4, 6, 6), range(90, 90, 99, 99)]),
      whole("overlap", [range(0, 0, 10, 10)]),
      ...Array.from({ length: 20 }, (_, i) => whole(String(i), [range(i + 20, 0, i + 20, 2)]))
    ]);
    for (let row = 0; row < 105; row += 1) {
      for (let column = 0; column < 105; column += 1) {
        const expected = indexed.find(rule => rule.ranges.some(r =>
          row >= r.startRow && row <= r.endRow && column >= r.startColumn && column <= r.endColumn));
        expect(findRuleAt(indexed, row, column)).toBe(expected);
      }
    }
  });

  it("observes changes to mutable public rule arrays and nested ranges", () => {
    const target = range(0, 0, 0, 0);
    const mutable = Array.from({ length: 9 }, (_, i) => whole(String(i), [i === 0 ? target : range(i, 0, i, 0)]));
    expect(findRuleAt(mutable, 0, 0)?.id).toBe("0");
    Object.assign(target, { startRow: 20, endRow: 20 });
    expect(findRuleAt(mutable, 0, 0)).toBeUndefined();
    expect(findRuleAt(mutable, 20, 0)?.id).toBe("0");
    mutable[0] = whole("replacement", [range(20, 0, 20, 0)]);
    expect(findRuleAt(mutable, 20, 0)?.id).toBe("replacement");
  });

  it("observes list item replacements without stale membership or messages", () => {
    const items = ["a"];
    const rule: DataValidationRule = { id: "list", ranges: [range(0, 0, 0, 0)], criteria: { type: "list", items } };
    expect(validateValue(rule, "a").valid).toBe(true);
    expect(defaultValidationMessage(rule.criteria)).toContain("a");
    items[0] = "b";
    expect(validateValue(rule, "a").valid).toBe(false);
    expect(validateValue(rule, "b").valid).toBe(true);
    expect(defaultValidationMessage(rule.criteria)).toBe("请从列表中选择：b");
  });
});

describe("applyRule", () => {
  it("carves the overlap out of an older rule so a cell belongs to at most one rule", () => {
    const older = whole("old", [range(0, 0, 3, 0)]);
    const next = applyRule([older], whole("new", [range(1, 0, 2, 0)]));

    expect(next.map((r) => r.id)).toEqual(["old", "new"]);
    const cells = coverage(next);
    expect(cells.get("0,0")).toBe("old");
    expect(cells.get("1,0")).toBe("new");
    expect(cells.get("2,0")).toBe("new");
    expect(cells.get("3,0")).toBe("old");
  });

  it("drops an older rule whose whole area is covered", () => {
    const next = applyRule([whole("old", [range(1, 1, 2, 2)])], whole("new", [range(0, 0, 5, 5)]));
    expect(next.map((r) => r.id)).toEqual(["new"]);
  });

  it("replaces a rule with the same id and moves it to the end", () => {
    const a = whole("a", [range(0, 0, 0, 0)]);
    const b = whole("b", [range(5, 5, 5, 5)]);
    const next = applyRule([a, b], whole("a", [range(9, 9, 9, 9)]));
    expect(next.map((r) => r.id)).toEqual(["b", "a"]);
    expect(next[1]!.ranges).toEqual([range(9, 9, 9, 9)]);
  });

  it("does not touch unrelated rules (same object identity)", () => {
    const untouched = whole("u", [range(8, 8, 9, 9)]);
    const next = applyRule([untouched], whole("n", [range(0, 0, 1, 1)]));
    expect(next[0]).toBe(untouched);
  });
});

describe("clearRanges", () => {
  it("removes only the covered part and deletes rules that end up empty", () => {
    const next = clearRanges(
      [whole("a", [range(0, 0, 3, 0)]), whole("b", [range(5, 0, 5, 0)])],
      [range(0, 0, 1, 0), range(5, 0, 5, 0)]
    );
    expect(next.map((r) => r.id)).toEqual(["a"]);
    expect(next[0]!.ranges).toEqual([range(2, 0, 3, 0)]);
  });
});

describe("relocateValidationsForInsert / Delete", () => {
  const rules = [whole("a", [range(2, 0, 4, 1)])];

  it("shifts when the insert point is at or before the start, expands when strictly inside", () => {
    expect(relocateValidationsForInsert(rules, "row", 2, 2)[0]!.ranges).toEqual([range(4, 0, 6, 1)]);
    expect(relocateValidationsForInsert(rules, "row", 3, 2)[0]!.ranges).toEqual([range(2, 0, 6, 1)]);
    expect(relocateValidationsForInsert(rules, "row", 5, 2)[0]!.ranges).toEqual([range(2, 0, 4, 1)]);
  });

  it("works on the column axis", () => {
    expect(relocateValidationsForInsert(rules, "column", 1, 3)[0]!.ranges).toEqual([range(2, 0, 4, 4)]);
  });

  it("shrinks on partial delete and removes the rule when fully covered", () => {
    expect(relocateValidationsForDelete(rules, "row", 3, 1)[0]!.ranges).toEqual([range(2, 0, 3, 1)]);
    expect(relocateValidationsForDelete(rules, "row", 0, 1)[0]!.ranges).toEqual([range(1, 0, 3, 1)]);
    expect(relocateValidationsForDelete(rules, "row", 2, 3)).toEqual([]);
    expect(relocateValidationsForDelete(rules, "column", 0, 2)).toEqual([]);
  });

  it("drops only the emptied range of a multi-range rule", () => {
    const multi = [whole("m", [range(0, 0, 0, 0), range(5, 0, 5, 0)])];
    const next = relocateValidationsForDelete(multi, "row", 0, 1);
    expect(next[0]!.ranges).toEqual([range(4, 0, 4, 0)]);
  });
});

describe("relocateValidationsForMove", () => {
  /** 独立于实现的参照：一个下标经过"摘出再插回"落在哪。 */
  function reference(index: number, start: number, count: number, target: number): number {
    const insertAt = target > start ? target - count : target;
    if (index >= start && index < start + count) {
      return insertAt + (index - start);
    }
    const afterDelete = index < start ? index : index - count;
    return afterDelete >= insertAt ? afterDelete + count : afterDelete;
  }

  it("sends every covered cell to where its row goes, for every legal move", () => {
    const rows = 8;
    const rules = [whole("a", [range(2, 0, 5, 0)]), whole("b", [range(7, 1, 7, 1), range(0, 1, 0, 1)])];
    const before = coverage(rules);

    for (let start = 0; start < rows; start += 1) {
      for (let count = 1; start + count <= rows; count += 1) {
        for (let target = 0; target <= rows; target += 1) {
          if (target >= start && target <= start + count) {
            continue; // 落在自己身上，不是一次真正的移动
          }
          const after = coverage(relocateValidationsForMove(rules, "row", start, count, target));
          const expected = new Map<string, string>();
          for (const [key, id] of before) {
            const [row, column] = key.split(",").map(Number) as [number, number];
            expected.set(`${reference(row, start, count, target)},${column}`, id);
          }
          expect(after, `move(${start}, ${count}, ${target})`).toEqual(expected);
        }
      }
    }
  });

  it("splits a range when the moved block is pulled out of its middle", () => {
    // 区域 rows 2..5，把 rows 3..4 搬到最前面：区域被拆成 上下两段
    const next = relocateValidationsForMove([whole("a", [range(2, 0, 5, 0)])], "row", 3, 2, 0);
    const cells = coverage(next);
    expect([...cells.keys()].sort()).toEqual(["0,0", "1,0", "4,0", "5,0"]);
    expect(next[0]!.ranges.length).toBeGreaterThan(1);
  });

  it("moving a block and then moving it back restores the covered cell set", () => {
    const rules = [whole("a", [range(1, 0, 6, 1)])];
    const moved = relocateValidationsForMove(rules, "row", 2, 2, 6);
    // 摘出 2..3 之后插回 6 之前：块落在 4..5
    const back = relocateValidationsForMove(moved, "row", 4, 2, 2);
    expect(coverage(back)).toEqual(coverage(rules));
  });

  it("works on the column axis", () => {
    const next = relocateValidationsForMove([whole("a", [range(0, 0, 0, 1)])], "column", 0, 1, 4);
    expect([...coverage(next).keys()].sort()).toEqual(["0,0", "0,3"]);
  });
});

describe("assertValidRule / assertValidRules", () => {
  function code(fn: () => void): string | undefined {
    try {
      fn();
    } catch (error) {
      return (error as CyExcelError).code;
    }
    return undefined;
  }

  it("accepts every criteria shape", () => {
    expect(() => assertValidRule(whole("a", [range(0, 0, 0, 0)]))).not.toThrow();
    expect(() =>
      assertValidRule({ id: "d", ranges: [range(0, 0, 0, 0)], criteria: { type: "date", operator: "greaterThan", value1: 45000 } })
    ).not.toThrow();
    expect(() =>
      assertValidRule({ id: "l", ranges: [range(0, 0, 0, 0)], criteria: { type: "list", items: ["a", "b"] } })
    ).not.toThrow();
    expect(() =>
      assertValidRule({ id: "c", ranges: [range(0, 0, 0, 0)], criteria: { type: "custom", name: "unique" } })
    ).not.toThrow();
  });

  it.each([
    ["empty id", { ...whole("", [range(0, 0, 0, 0)]) }],
    ["no ranges", whole("a", [])],
    ["inverted range", whole("a", [range(3, 0, 1, 0)])],
    ["negative range", whole("a", [range(-1, 0, 0, 0)])],
    ["overlapping ranges inside one rule", whole("a", [range(0, 0, 2, 2), range(1, 1, 3, 3)])],
    ["between without value2", { id: "a", ranges: [range(0, 0, 0, 0)], criteria: { type: "whole", operator: "between", value1: 1 } }],
    ["non-integer whole", whole("a", [range(0, 0, 0, 0)], 1.5, 3)],
    ["negative textLength", { id: "a", ranges: [range(0, 0, 0, 0)], criteria: { type: "textLength", operator: "equal", value1: -1 } }],
    ["unknown operator", { id: "a", ranges: [range(0, 0, 0, 0)], criteria: { type: "whole", operator: "nope", value1: 1 } }],
    ["empty list", { id: "a", ranges: [range(0, 0, 0, 0)], criteria: { type: "list", items: [] } }],
    ["custom without name", { id: "a", ranges: [range(0, 0, 0, 0)], criteria: { type: "custom", name: "" } }],
    ["unknown type", { id: "a", ranges: [range(0, 0, 0, 0)], criteria: { type: "mystery" } }],
    ["unknown errorStyle", { ...whole("a", [range(0, 0, 0, 0)]), errorStyle: "loud" }]
  ])("rejects %s with InvalidDataValidation", (_label, rule) => {
    expect(code(() => assertValidRule(rule as unknown as DataValidationRule))).toBe(CyExcelErrorCode.InvalidDataValidation);
  });

  it("rejects duplicate ids and overlapping rules in a list", () => {
    const a = whole("a", [range(0, 0, 1, 1)]);
    expect(code(() => assertValidRules([a, whole("a", [range(5, 5, 5, 5)])]))).toBe(CyExcelErrorCode.InvalidDataValidation);
    expect(code(() => assertValidRules([a, whole("b", [range(1, 1, 2, 2)])]))).toBe(CyExcelErrorCode.InvalidDataValidation);
    expect(() => assertValidRules([a, whole("b", [range(2, 2, 3, 3)])])).not.toThrow();
  });
});

describe("buildRule", () => {
  const ids = () => "generated";

  it("fills in an id and makes overlapping input ranges disjoint", () => {
    const rule = buildRule(
      { criteria: { type: "whole", operator: "equal", value1: 1 } },
      [range(0, 0, 1, 1), range(1, 1, 2, 2)],
      ids
    );
    expect(rule.id).toBe("generated");
    expect(() => assertValidRule(rule)).not.toThrow();
  });

  it("turns ISO date strings into serial numbers", () => {
    const rule = buildRule(
      { criteria: { type: "date", operator: "between", value1: "2026-01-01", value2: "2026-01-31" } },
      [range(0, 0, 0, 0)],
      ids
    );
    expect(rule.criteria).toEqual({ type: "date", operator: "between", value1: 46023, value2: 46053 });
  });

  it("rejects a date string it cannot parse", () => {
    expect(() =>
      buildRule({ criteria: { type: "date", operator: "equal", value1: "not a date" } }, [range(0, 0, 0, 0)], ids)
    ).toThrow(CyExcelError);
  });
});

describe("validateValue", () => {
  const cell = [range(0, 0, 0, 0)];

  it("treats blank as valid unless allowBlank is false", () => {
    const rule = whole("a", cell);
    expect(validateValue(rule, undefined).valid).toBe(true);
    expect(validateValue(rule, "").valid).toBe(true);
    expect(validateValue({ ...rule, allowBlank: false }, undefined).valid).toBe(false);
  });

  it("whole: integers only, no implicit string-to-number", () => {
    const rule = whole("a", cell, 1, 10);
    expect(validateValue(rule, 5).valid).toBe(true);
    expect(validateValue(rule, 5.5).valid).toBe(false);
    expect(validateValue(rule, 11).valid).toBe(false);
    expect(validateValue(rule, "5").valid).toBe(false);
    expect(validateValue(rule, true).valid).toBe(false);
  });

  it("decimal: any finite number", () => {
    const rule: DataValidationRule = { id: "d", ranges: cell, criteria: { type: "decimal", operator: "lessThan", value1: 1 } };
    expect(validateValue(rule, 0.5).valid).toBe(true);
    expect(validateValue(rule, 1).valid).toBe(false);
    expect(validateValue(rule, Number.NaN).valid).toBe(false);
  });

  it.each([
    ["between", 3, 1, 5, true],
    ["between", 6, 1, 5, false],
    ["between", 3, 5, 1, true], // value1 > value2 is not a config error
    ["notBetween", 6, 1, 5, true],
    ["notBetween", 3, 1, 5, false],
    ["equal", 4, 4, undefined, true],
    ["notEqual", 4, 4, undefined, false],
    ["greaterThan", 5, 4, undefined, true],
    ["greaterThan", 4, 4, undefined, false],
    ["greaterThanOrEqual", 4, 4, undefined, true],
    ["lessThan", 3, 4, undefined, true],
    ["lessThanOrEqual", 5, 4, undefined, false]
  ] as const)("operator %s on %s vs (%s, %s) is %s", (operator, n, value1, value2, expected) => {
    const rule: DataValidationRule = {
      id: "o",
      ranges: cell,
      criteria: { type: "decimal", operator, value1, ...(value2 !== undefined ? { value2 } : {}) }
    };
    expect(validateValue(rule, n).valid).toBe(expected);
  });

  it("textLength counts characters of the string form", () => {
    const rule: DataValidationRule = { id: "t", ranges: cell, criteria: { type: "textLength", operator: "lessThanOrEqual", value1: 3 } };
    expect(validateValue(rule, "abc").valid).toBe(true);
    expect(validateValue(rule, "abcd").valid).toBe(false);
    expect(validateValue(rule, 123).valid).toBe(true);
    expect(validateValue(rule, 1234).valid).toBe(false);
  });

  it("date compares serial numbers and parses ISO strings", () => {
    const rule: DataValidationRule = { id: "d", ranges: cell, criteria: { type: "date", operator: "greaterThanOrEqual", value1: 46023 } };
    expect(validateValue(rule, 46023).valid).toBe(true);
    expect(validateValue(rule, 46022).valid).toBe(false);
    expect(validateValue(rule, "2026-01-15").valid).toBe(true);
    expect(validateValue(rule, "2025-12-31").valid).toBe(false);
    expect(validateValue(rule, "banana").valid).toBe(false);
    expect(validateValue(rule, true).valid).toBe(false);
  });

  it("time compares only the fractional part of the serial", () => {
    const rule: DataValidationRule = { id: "t", ranges: cell, criteria: { type: "time", operator: "between", value1: 0.375, value2: 0.75 } };
    expect(validateValue(rule, 0.5).valid).toBe(true); // 12:00
    expect(validateValue(rule, 46023.5).valid).toBe(true); // 与是哪一天无关
    expect(validateValue(rule, 0.1).valid).toBe(false);
  });

  it("list checks the string form against items", () => {
    const rule: DataValidationRule = { id: "l", ranges: cell, criteria: { type: "list", items: ["红", "绿", "1"] } };
    expect(validateValue(rule, "红").valid).toBe(true);
    expect(validateValue(rule, 1).valid).toBe(true);
    expect(validateValue(rule, "蓝").valid).toBe(false);
  });

  it("carries errorStyle, title and custom message on failure, with sensible defaults", () => {
    const plain = validateValue(whole("a", cell, 1, 10), 99);
    expect(plain).toMatchObject({ valid: false, ruleId: "a", errorStyle: "stop" });
    expect(plain.message).toContain("整数");

    const custom = validateValue(
      { ...whole("a", cell), errorStyle: "warning", errorTitle: "注意", errorMessage: "太大了" },
      99
    );
    expect(custom).toEqual({ valid: false, ruleId: "a", errorStyle: "warning", title: "注意", message: "太大了" });
  });

  describe("custom", () => {
    const rule: DataValidationRule = { id: "c", ranges: cell, criteria: { type: "custom", name: "even", params: { min: 2 } } };

    it("treats an unregistered name as valid", () => {
      expect(validateValue(rule, 3).valid).toBe(true);
      expect(validateValue(rule, 3, { getCustomValidator: () => undefined }).valid).toBe(true);
    });

    it("passes value, params and context to the function", () => {
      let seen: unknown[] = [];
      const fn = (value: unknown, params: unknown, context: { row: number; column: number; getCell(r: number, c: number): unknown }) => {
        seen = [value, params, context.row, context.column, context.getCell(9, 9)];
        return true;
      };
      validateValue(rule, 4, {
        getCustomValidator: () => fn,
        row: 2,
        column: 3,
        getCell: () => "neighbour"
      });
      expect(seen).toEqual([4, { min: 2 }, 2, 3, "neighbour"]);
    });

    it("treats false as invalid with the default text and a string as the message", () => {
      const options = (result: boolean | string) => ({ getCustomValidator: () => () => result });
      expect(validateValue(rule, 3, options(true)).valid).toBe(true);
      expect(validateValue(rule, 3, options(false))).toMatchObject({ valid: false, message: "输入的值不符合规则" });
      expect(validateValue(rule, 3, options("必须是偶数"))).toMatchObject({ valid: false, message: "必须是偶数" });
    });
  });
});

describe("intersectRanges / collectInvalidCells", () => {
  it("intersects rectangles", () => {
    expect(intersectRanges(range(0, 0, 5, 5), range(3, 3, 9, 9))).toEqual(range(3, 3, 5, 5));
    expect(intersectRanges(range(0, 0, 1, 1), range(3, 3, 4, 4))).toBeUndefined();
  });

  it("lists invalid cells in row-major order, clipped to the scope", () => {
    const rules = [whole("a", [range(0, 0, 2, 1)], 0, 9)];
    const values: Record<string, number> = { "0,0": 5, "0,1": 50, "1,0": -1, "2,1": 7 };
    const getValue = (r: number, c: number) => values[`${r},${c}`];

    const all = collectInvalidCells(rules, range(0, 0, 9, 9), getValue, () => ({}));
    expect(all).toEqual([
      { row: 0, column: 1 },
      { row: 1, column: 0 }
    ]);

    const clipped = collectInvalidCells(rules, range(1, 0, 1, 0), getValue, () => ({}));
    expect(clipped).toEqual([{ row: 1, column: 0 }]);
  });

  it("counts blank cells only when allowBlank is false", () => {
    const strict = [{ ...whole("a", [range(0, 0, 1, 0)]), allowBlank: false }];
    const found = collectInvalidCells(strict, range(0, 0, 1, 0), () => undefined, () => ({}));
    expect(found).toHaveLength(2);
    const lax = collectInvalidCells([whole("a", [range(0, 0, 1, 0)])], range(0, 0, 1, 0), () => undefined, () => ({}));
    expect(lax).toEqual([]);
  });
});

describe("describeDropdownItems", () => {
  const cell = [range(0, 0, 0, 0)];

  it("lists the items of a list rule, unless showDropDown is false", () => {
    const rule: DataValidationRule = { id: "l", ranges: cell, criteria: { type: "list", items: ["a", "b"] } };
    expect(describeDropdownItems(rule)).toEqual(["a", "b"]);
    expect(describeDropdownItems({ ...rule, showDropDown: false })).toBeUndefined();
  });

  it("is undefined for rules that are not lists", () => {
    expect(describeDropdownItems(whole("w", cell))).toBeUndefined();
  });
});

describe("list renderMode and colors", () => {
  const cell = [range(0, 0, 0, 0)];
  const list = (extra: Partial<DataValidationRule> = {}, colors?: Record<string, string>): DataValidationRule => ({
    id: "l",
    ranges: cell,
    criteria: { type: "list", items: ["红", "绿"], ...(colors !== undefined ? { colors } : {}) },
    ...extra
  });

  it("accepts every render mode and rejects unknown ones", () => {
    for (const renderMode of ["text", "arrow", "custom"] as const) {
      expect(() => assertValidRule(list({ renderMode }))).not.toThrow();
    }
    expect(() => assertValidRule(list({ renderMode: "pill" as never }))).toThrow(CyExcelError);
  });

  it("accepts hex colours for known items", () => {
    expect(() => assertValidRule(list({}, { 红: "#d93025", 绿: "#188038cc", }))).not.toThrow();
    expect(() => assertValidRule(list({}, { 红: "#f00" }))).not.toThrow();
  });

  it.each([
    ["a colour for a non-item", { 蓝: "#0000ff" }],
    ["a named colour", { 红: "red" }],
    ["an rgb() string", { 红: "rgb(1,2,3)" }],
    ["a malformed hex", { 红: "#ff00" }],
    ["a non-string", { 红: 12 as never }]
  ])("rejects %s", (_label, colors) => {
    expect(() => assertValidRule(list({}, colors as Record<string, string>))).toThrow(CyExcelError);
  });

  it("rejects colors that is not a plain object", () => {
    const bad = { ...list(), criteria: { type: "list", items: ["a"], colors: ["#fff"] } } as unknown as DataValidationRule;
    expect(() => assertValidRule(bad)).toThrow(CyExcelError);
  });

  describe("describeCellDropdown", () => {
    it("defaults to arrow mode and carries the current value's colour", () => {
      const rule = list({}, { 红: "#d93025" });
      expect(describeCellDropdown(rule, "红")).toEqual({ mode: "arrow", color: "#d93025" });
      expect(describeCellDropdown(rule, "绿")).toEqual({ mode: "arrow" });
      expect(describeCellDropdown(rule, undefined)).toEqual({ mode: "arrow" });
    });

    it("reports the configured mode", () => {
      expect(describeCellDropdown(list({ renderMode: "custom" }), "红")).toEqual({ mode: "custom" });
      expect(describeCellDropdown(list({ renderMode: "text" }), "红")).toEqual({ mode: "text" });
    });

    it("is undefined for non-list rules and when showDropDown is false", () => {
      expect(describeCellDropdown(whole("w", cell), 1)).toBeUndefined();
      expect(describeCellDropdown(list({ showDropDown: false }), "红")).toBeUndefined();
    });

    it("matches numbers by their string form", () => {
      const rule: DataValidationRule = { id: "n", ranges: cell, criteria: { type: "list", items: ["1"], colors: { "1": "#fff" } } };
      expect(describeCellDropdown(rule, 1)).toEqual({ mode: "arrow", color: "#fff" });
    });
  });
});
