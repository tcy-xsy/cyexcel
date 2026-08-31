import { describe, expect, it } from "vitest";

import type { CanvasHeaderRenderer } from "../../../../src/core/cells";
import { CyExcelError } from "../../../../src/core/errors";
import { HeaderManager } from "../../../../src/core/headers";

describe("HeaderManager — leaf header resolution", () => {
  it("resolves a plain string header as its title", () => {
    const manager = new HeaderManager({ columns: [{ header: "Name" }] });

    expect(manager.getHeader(0)).toEqual({ column: 0, title: "Name", renderer: undefined });
  });

  it("resolves a HeaderDefinition's title and renderer", () => {
    const renderer: CanvasHeaderRenderer = { kind: "canvas", paint: () => {} };
    const manager = new HeaderManager({
      columns: [{ header: { title: "Name", renderer } }]
    });

    const header = manager.getHeader(0);
    expect(header.title).toBe("Name");
    expect(header.renderer).toBe(renderer);
  });

  it("falls back to the spreadsheet letter when no header is declared", () => {
    const manager = new HeaderManager({ columns: [{ }] });

    expect(manager.getHeader(0)).toEqual({ column: 0, title: "A", renderer: undefined });
  });

  it("getHeaders() returns every leaf header in column order", () => {
    const manager = new HeaderManager({
      columns: [{ header: "第一列" }, { header: "第二列" }, { }]
    });

    expect(manager.getHeaders().map((header) => header.title)).toEqual(["第一列", "第二列", "C"]);
  });

  it("getHeader() throws for an out-of-range column", () => {
    const manager = new HeaderManager({ columns: [{ }] });

    expect(() => manager.getHeader(5)).toThrow(CyExcelError);
  });
});

describe("HeaderManager — columnLetterFallback (spreadsheet-style A/B/C.../AA/AB... labels)", () => {
  // 列没有 `id` 之后，这个开关的两个值就是仅剩的两种答案（ADR 0004 §3.8.1/
  // §3.8.2）：开是 A/B/C（默认从 false 翻成了 true），关是**没有标题**——不是
  // 回退到别的名字，而是这一列不显示标题。
  it("defaults to true — an undeclared header shows its spreadsheet letter", () => {
    const manager = new HeaderManager({ columns: [{ }] });

    expect(manager.getHeader(0).title).toBe("A");
  });

  it("turning it off leaves an undeclared header with no title at all, not some other fallback", () => {
    const manager = new HeaderManager({
      columns: [{ }, { header: "Price" }],
      columnLetterFallback: false
    });

    expect(manager.getHeader(0).title).toBe("");
    expect(manager.getHeader(1).title).toBe("Price");
  });

  it("falls back to the spreadsheet letter when explicitly enabled", () => {
    const manager = new HeaderManager({
      columns: [{ }, { }],
      columnLetterFallback: true
    });

    expect(manager.getHeader(0).title).toBe("A");
    expect(manager.getHeader(1).title).toBe("B");
  });

  it("wraps past Z into AA, AB, ... like Excel/Google Sheets/Univer column references", () => {
    const columns = Array.from({ length: 30 }, () => ({ }));
    const manager = new HeaderManager({ columns, columnLetterFallback: true });

    expect(manager.getHeader(25).title).toBe("Z");
    expect(manager.getHeader(26).title).toBe("AA");
    expect(manager.getHeader(27).title).toBe("AB");
  });

  it("an explicit header still overrides the letter for that one column — mutually exclusive per column, matching Univer's customizeColumnHeader model", () => {
    const manager = new HeaderManager({
      columns: [{ }, { header: "Price" }, { }],
      columnLetterFallback: true
    });

    expect(manager.getHeader(0).title).toBe("A");
    expect(manager.getHeader(1).title).toBe("Price");
    expect(manager.getHeader(2).title).toBe("C");
  });

  it("prepareRoot() keeps using the letter fallback the manager was constructed with", () => {
    const manager = new HeaderManager({ columns: [{ }], columnLetterFallback: true });
    expect(manager.getHeader(0).title).toBe("A");

    manager.replaceRoot(manager.prepareRoot([{ }, { }]));

    expect(manager.getHeader(0).title).toBe("A");
    expect(manager.getHeader(1).title).toBe("B");
  });
});

describe("HeaderManager — resolves once and caches", () => {
  it("getHeader() returns the same object on repeated calls", () => {
    const manager = new HeaderManager({ columns: [{ }] });

    expect(manager.getHeader(0)).toBe(manager.getHeader(0));
  });
});

describe("HeaderManager.rebuild", () => {
  it("prepares a validated root without changing the live root", () => {
    const manager = new HeaderManager({ columns: [{ header: "A" }] });
    const oldHeaders = manager.getHeaders();
    const root = manager.prepareRoot([
      { header: "X" },
      { header: "Y" }
    ]);

    expect(manager.getHeaders()).toBe(oldHeaders);
    expect(manager.getHeaders().map((header) => header.title)).toEqual(["A"]);

    manager.replaceRoot(root);
    expect(manager.getHeaders().map((header) => header.title)).toEqual(["X", "Y"]);
  });

  it("re-resolves headers from a new column list", () => {
    const manager = new HeaderManager({ columns: [{ header: "A" }] });
    expect(manager.getHeaders().map((header) => header.title)).toEqual(["A"]);

    manager.replaceRoot(manager.prepareRoot([
      { header: "A" },
      { header: "B" }
    ]));

    expect(manager.getHeaders().map((header) => header.title)).toEqual(["A", "B"]);
  });

  it("getHeader() reflects the replacement root immediately", () => {
    const manager = new HeaderManager({ columns: [{ header: "A" }] });

    manager.replaceRoot(manager.prepareRoot([{ header: "X" }]));

    expect(manager.getHeader(0).title).toBe("X");
  });
});
