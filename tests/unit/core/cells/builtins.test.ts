import { describe, expect, it } from "vitest";

import { CellTypeRegistry, registerBuiltinCellTypes } from "../../../../src/core/cells";
import type { CellValue } from "../../../../src/core/sheet/types";

function registry(): CellTypeRegistry {
  const registry = new CellTypeRegistry();
  registerBuiltinCellTypes(registry);
  return registry;
}

const ctx = { row: 0, column: {} as never };

describe("builtin cell types — text", () => {
  it("formats values as their string form, empty for an unoccupied cell", () => {
    const type = registry().resolve("text");
    expect(type.format("hello", ctx)).toBe("hello");
    expect(type.format(undefined, ctx)).toBe("");
  });
});

describe("builtin cell types — number", () => {
  it("formats finite numbers, empty string for non-finite", () => {
    const type = registry().resolve("number");
    expect(type.format(42, ctx)).toBe("42");
    expect(type.format(Number.NaN, ctx)).toBe("");
  });

  it("parses text back into a number", () => {
    const type = registry().resolve("number");
    expect(type.parse?.("42", ctx)).toBe(42);
  });

  it("treats empty text as an empty cell instead of zero", () => {
    const type = registry().resolve("number");
    expect(type.parse?.("", ctx)).toBeUndefined();
    expect(type.parse?.("   ", ctx)).toBeUndefined();
  });

  it("compares numerically", () => {
    const type = registry().resolve("number");
    expect(type.compare?.(1, 2)).toBeLessThan(0);
    expect(type.compare?.(2, 1)).toBeGreaterThan(0);
  });
});

describe("builtin cell types — boolean", () => {
  it("formats as true/false", () => {
    const type = registry().resolve("boolean");
    expect(type.format(true, ctx)).toBe("true");
    expect(type.format(false, ctx)).toBe("false");
  });

  it("parses text back into a boolean", () => {
    const type = registry().resolve("boolean");
    expect(type.parse?.("true", ctx)).toBe(true);
    expect(type.parse?.("false", ctx)).toBe(false);
  });

  it("serializes for editing the same way it formats for display", () => {
    const type = registry().resolve("boolean");
    expect(type.serialize?.(true, ctx)).toBe("true");
    expect(type.serialize?.(false, ctx)).toBe("false");
  });

  it("round-trips through serialize -> parse without changing the value", () => {
    const type = registry().resolve("boolean");
    for (const value of [true, false]) {
      expect(type.parse?.(type.serialize?.(value, ctx) ?? "", ctx)).toBe(value);
    }
  });

  it("orders by truthiness, which is not the order the raw values' strings give", () => {
    const type = registry().resolve("boolean");
    // A "boolean" column may hold any truthy/falsy value — format() and
    // parse() both coerce — so sorting has to agree with what is drawn.
    const values: CellValue[] = [true, 0, false, 1];

    const byString = [...values].sort((left, right) =>
      String(left) < String(right) ? -1 : String(left) > String(right) ? 1 : 0
    );
    // Premise: string order really does disagree ("0" < "1" < "false").
    expect(byString).toEqual([0, 1, false, true]);

    expect([...values].sort((left, right) => type.compare?.(left, right) ?? 0))
      .toEqual([0, false, true, 1]);
    expect(type.compare?.(false, true)).toBeLessThan(0);
    expect(type.compare?.(true, true)).toBe(0);
  });
});

describe("builtin cell types — date", () => {
  // 日期不是一种值类型：格子里存的是 Excel 序列号（数字），显示成什么样由
  // numFmt 决定，这个类型只负责序列号 <-> 日期串。46037 = 2026-01-15。
  const JAN_15_2026 = 46037;

  it("formats a serial number as an ISO date (YYYY-MM-DD)", () => {
    const type = registry().resolve("date");
    expect(type.format(JAN_15_2026, ctx)).toBe("2026-01-15");
  });

  it("falls back to the raw text when the value isn't a serial number", () => {
    const type = registry().resolve("date");
    expect(type.format("not a date", ctx)).toBe("not a date");
    expect(type.format(undefined, ctx)).toBe("");
  });

  it("parses an ISO date string into a serial number", () => {
    const type = registry().resolve("date");
    expect(type.parse?.("2026-01-15", ctx)).toBe(JAN_15_2026);
  });

  it("keeps unparseable text as text instead of inventing a serial", () => {
    const type = registry().resolve("date");
    expect(type.parse?.("not a date", ctx)).toBe("not a date");
  });

  it("serializes for editing the same way it formats for display", () => {
    const type = registry().resolve("date");
    expect(type.serialize?.(JAN_15_2026, ctx)).toBe("2026-01-15");
  });

  it("round-trips through serialize -> parse without changing the value", () => {
    const type = registry().resolve("date");
    expect(type.parse?.(type.serialize?.(JAN_15_2026, ctx) ?? "", ctx)).toBe(JAN_15_2026);
  });

  it("round-trips a whole-day serial without picking up a time-of-day fraction", () => {
    // Regression: going through `new Date("2026-01-15")` parses as UTC
    // midnight, which in a UTC+8 zone comes back as 46037.333. A serial is a
    // wall clock — it must survive the round trip in every timezone.
    const type = registry().resolve("date");
    const parsed = type.parse?.("2026-01-15", ctx);
    expect(Number.isInteger(parsed)).toBe(true);
  });

  it("needs no compare(): serial numbers are already in chronological order", () => {
    const type = registry().resolve("date");
    // Premise — the three serials really are consecutive days, so plain
    // numeric order is chronological order and a comparator would be dead code.
    const jan01 = type.parse?.("2020-01-01", ctx) as number;
    const jan06 = type.parse?.("2020-01-06", ctx) as number;
    const jan11 = type.parse?.("2020-01-11", ctx) as number;
    expect([jan11, jan01, jan06].sort((left, right) => left - right)).toEqual([jan01, jan06, jan11]);
    expect(type.compare).toBeUndefined();
  });
});

describe("builtin cell types — select", () => {
  it("formats the current value as text when no options are configured", () => {
    const type = registry().resolve("select");
    expect(type.format("draft", ctx)).toBe("draft");
  });

  it("shows the matching option's label when options are configured", () => {
    const type = registry().resolve("select");
    const context = { row: 0, column: { options: [{ value: "draft", label: "草稿" }] } as never };
    expect(type.format("draft", context)).toBe("草稿");
  });

  it("falls back to the raw value when it isn't in the configured options", () => {
    const type = registry().resolve("select");
    const context = { row: 0, column: { options: [{ value: "draft", label: "草稿" }] } as never };
    expect(type.format("archived", context)).toBe("archived");
  });
});
