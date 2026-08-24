import { describe, expect, it } from "vitest";

import { CellTypeRegistry, registerBuiltinCellTypes } from "../../../../src/core/cells";

function registry(): CellTypeRegistry {
  const registry = new CellTypeRegistry();
  registerBuiltinCellTypes(registry);
  return registry;
}

const ctx = { row: 0, column: {} as never };

describe("builtin cell types — text", () => {
  it("formats values as their string form, empty for null/undefined", () => {
    const type = registry().resolve("text");
    expect(type.format("hello", ctx)).toBe("hello");
    expect(type.format(undefined, ctx)).toBe("");
    expect(type.format(null, ctx)).toBe("");
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
    const values: unknown[] = [true, 0, false, 1];

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
  it("formats a Date as an ISO date (YYYY-MM-DD)", () => {
    const type = registry().resolve("date");
    expect(type.format(new Date(Date.UTC(2026, 0, 15)), ctx)).toBe("2026-01-15");
  });

  it("returns an empty string for an invalid date", () => {
    const type = registry().resolve("date");
    expect(type.format(new Date(Number.NaN), ctx)).toBe("");
  });

  it("parses an ISO date string back into a Date", () => {
    const type = registry().resolve("date");
    const parsed = type.parse?.("2026-01-15", ctx) as Date;
    expect(parsed.toISOString().slice(0, 10)).toBe("2026-01-15");
  });

  it("serializes for editing the same way it formats for display", () => {
    const type = registry().resolve("date");
    expect(type.serialize?.(new Date(Date.UTC(2026, 0, 15)), ctx)).toBe("2026-01-15");
  });

  it("round-trips through serialize -> parse without changing the value", () => {
    const type = registry().resolve("date");
    const original = new Date(Date.UTC(2026, 0, 15));
    const parsed = type.parse?.(type.serialize?.(original, ctx) ?? "", ctx) as Date;
    expect(parsed.toISOString().slice(0, 10)).toBe("2026-01-15");
  });

  it("orders chronologically, which is not the order the raw Date strings give", () => {
    const type = registry().resolve("date");
    // String(date) starts with the weekday name, so the default string
    // comparator orders these Mon < Sat < Wed. Noon UTC keeps the weekday
    // the same in every plausible local timezone.
    const wedJan01 = new Date(Date.UTC(2020, 0, 1, 12));
    const monJan06 = new Date(Date.UTC(2020, 0, 6, 12));
    const satJan11 = new Date(Date.UTC(2020, 0, 11, 12));
    const chronological = [wedJan01, monJan06, satJan11];

    const byString = [...chronological].sort((left, right) =>
      String(left) < String(right) ? -1 : String(left) > String(right) ? 1 : 0
    );
    // Premise: the two orders genuinely disagree, so this test can only pass
    // because of compare(), never by coincidence.
    expect(byString).toEqual([monJan06, satJan11, wedJan01]);

    expect([satJan11, wedJan01, monJan06].sort((left, right) => type.compare?.(left, right) ?? 0))
      .toEqual(chronological);
  });

  it("groups values that are not valid dates ahead of the real ones", () => {
    const type = registry().resolve("date");
    const invalid = new Date(Number.NaN);
    const real = new Date(Date.UTC(2020, 0, 1));

    expect(type.compare?.(invalid, real)).toBeLessThan(0);
    expect(type.compare?.(real, invalid)).toBeGreaterThan(0);
    expect(type.compare?.(invalid, new Date("nope"))).toBe(0);
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

describe("builtin cell types — array", () => {
  it("previews the first few items and the total count", () => {
    const type = registry().resolve("array");
    expect(type.format(["a", "b", "c", "d"], ctx)).toBe("[4] a, b, c, ...");
    expect(type.format(["a"], ctx)).toBe("[1] a");
    expect(type.format(undefined, ctx)).toBe("");
  });

  it("parses JSON array syntax back into an array", () => {
    const type = registry().resolve("array");
    expect(type.parse?.('["a","b",3]', ctx)).toEqual(["a", "b", 3]);
  });

  it("serializes using JSON array syntax, not the comma-joined display preview", () => {
    const type = registry().resolve("array");
    expect(type.serialize?.(["a", "b", 3], ctx)).toBe('["a","b",3]');
  });

  it("round-trips arbitrary nested values through serialize -> parse", () => {
    const type = registry().resolve("array");
    const original = ["a", 1, { nested: true }];
    expect(type.parse?.(type.serialize?.(original, ctx) ?? "", ctx)).toEqual(original);
  });

  it("degrades to an empty array instead of throwing on malformed input", () => {
    const type = registry().resolve("array");
    expect(type.parse?.("not json", ctx)).toEqual([]);
    expect(type.parse?.('{"not":"an array"}', ctx)).toEqual([]);
  });
});

describe("builtin cell types — json", () => {
  it("previews a serialized, length-capped form", () => {
    const type = registry().resolve("json");
    expect(type.format({ a: 1 }, ctx)).toBe('{"a":1}');
  });

  it("returns an empty string when the value can't be serialized to JSON (e.g. undefined)", () => {
    const type = registry().resolve("json");
    expect(type.format(undefined, ctx)).toBe("");
  });

  it("truncates long serialized values", () => {
    const type = registry().resolve("json");
    const big = { text: "x".repeat(200) };
    const formatted = type.format(big, ctx);
    expect(formatted.length).toBeLessThan(90);
    expect(formatted.endsWith("...")).toBe(true);
  });

  it("parses JSON text back into the original value", () => {
    const type = registry().resolve("json");
    expect(type.parse?.('{"a":1}', ctx)).toEqual({ a: 1 });
  });

  it("serializes the same way it previews for display when short", () => {
    const type = registry().resolve("json");
    expect(type.serialize?.({ a: 1 }, ctx)).toBe('{"a":1}');
  });

  it("round-trips through serialize -> parse without changing the value", () => {
    const type = registry().resolve("json");
    const original = { a: 1, nested: { b: [1, 2, 3] } };
    expect(type.parse?.(type.serialize?.(original, ctx) ?? "", ctx)).toEqual(original);
  });

  it("degrades to undefined instead of throwing on malformed input", () => {
    const type = registry().resolve("json");
    expect(type.parse?.("not json", ctx)).toBeUndefined();
  });
});
