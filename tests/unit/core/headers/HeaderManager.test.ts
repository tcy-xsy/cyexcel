import { describe, expect, it } from "vitest";

import type { CanvasHeaderRenderer } from "../../../../src/core/cells";
import { CyExcelError } from "../../../../src/core/errors";
import { HeaderManager } from "../../../../src/core/headers";

describe("HeaderManager — leaf header resolution", () => {
  it("resolves a plain string header as its title", () => {
    const manager = new HeaderManager({ columns: [{ id: "name", header: "Name" }] });

    expect(manager.getHeader(0)).toEqual({ column: 0, title: "Name", renderer: undefined });
  });

  it("resolves a HeaderDefinition's title and renderer", () => {
    const renderer: CanvasHeaderRenderer = { kind: "canvas", paint: () => {} };
    const manager = new HeaderManager({
      columns: [{ id: "name", header: { title: "Name", renderer } }]
    });

    const header = manager.getHeader(0);
    expect(header.title).toBe("Name");
    expect(header.renderer).toBe(renderer);
  });

  it("falls back to the column id when no header is declared", () => {
    const manager = new HeaderManager({ columns: [{ id: "sku" }] });

    expect(manager.getHeader(0)).toEqual({ column: 0, title: "sku", renderer: undefined });
  });

  it("getHeaders() returns every leaf header in column order", () => {
    const manager = new HeaderManager({
      columns: [{ id: "a", header: "A" }, { id: "b", header: "B" }, { id: "c" }]
    });

    expect(manager.getHeaders().map((header) => header.title)).toEqual(["A", "B", "c"]);
  });

  it("getHeader() throws for an out-of-range column", () => {
    const manager = new HeaderManager({ columns: [{ id: "a" }] });

    expect(() => manager.getHeader(5)).toThrow(CyExcelError);
  });
});

describe("HeaderManager — columnLetterFallback (spreadsheet-style A/B/C.../AA/AB... labels)", () => {
  it("defaults to false — unset behavior (fall back to column id) is unchanged", () => {
    const manager = new HeaderManager({ columns: [{ id: "sku" }] });

    expect(manager.getHeader(0).title).toBe("sku");
  });

  it("falls back to the spreadsheet letter instead of the column id when enabled", () => {
    const manager = new HeaderManager({
      columns: [{ id: "sku" }, { id: "name" }],
      columnLetterFallback: true
    });

    expect(manager.getHeader(0).title).toBe("A");
    expect(manager.getHeader(1).title).toBe("B");
  });

  it("wraps past Z into AA, AB, ... like Excel/Google Sheets/Univer column references", () => {
    const columns = Array.from({ length: 30 }, (_unused, index) => ({ id: `c${index}` }));
    const manager = new HeaderManager({ columns, columnLetterFallback: true });

    expect(manager.getHeader(25).title).toBe("Z");
    expect(manager.getHeader(26).title).toBe("AA");
    expect(manager.getHeader(27).title).toBe("AB");
  });

  it("an explicit header still overrides the letter for that one column — mutually exclusive per column, matching Univer's customizeColumnHeader model", () => {
    const manager = new HeaderManager({
      columns: [{ id: "a" }, { id: "b", header: "Price" }, { id: "c" }],
      columnLetterFallback: true
    });

    expect(manager.getHeader(0).title).toBe("A");
    expect(manager.getHeader(1).title).toBe("Price");
    expect(manager.getHeader(2).title).toBe("C");
  });

  it("prepareRoot() keeps using the letter fallback the manager was constructed with", () => {
    const manager = new HeaderManager({ columns: [{ id: "a" }], columnLetterFallback: true });
    expect(manager.getHeader(0).title).toBe("A");

    manager.replaceRoot(manager.prepareRoot([{ id: "a" }, { id: "b" }]));

    expect(manager.getHeader(0).title).toBe("A");
    expect(manager.getHeader(1).title).toBe("B");
  });
});

describe("HeaderManager — two-level groups (先写测试 #3: colspan)", () => {
  it("computes a group's colspan from its leaf column range", () => {
    const manager = new HeaderManager({
      columns: [{ id: "a" }, { id: "b" }, { id: "c" }, { id: "d" }],
      groups: [{ id: "g1", title: "Basics", columns: [0, 1, 2] }]
    });

    const [group] = manager.getGroups();
    expect(group).toEqual({ id: "g1", title: "Basics", startColumn: 0, endColumn: 2 });
    expect((group?.endColumn ?? 0) - (group?.startColumn ?? 0) + 1).toBe(3); // colspan
  });

  it("supports multiple independent groups", () => {
    const manager = new HeaderManager({
      columns: [{ id: "a" }, { id: "b" }, { id: "c" }, { id: "d" }],
      groups: [
        { id: "g1", title: "Basics", columns: [0, 1] },
        { id: "g2", title: "Sales", columns: [2, 3] }
      ]
    });

    expect(manager.getGroups().map((group) => group.id)).toEqual(["g1", "g2"]);
  });

  it("rejects a group whose leaf columns aren't contiguous", () => {
    expect(
      () =>
        new HeaderManager({
          columns: [{ id: "a" }, { id: "b" }, { id: "c" }],
          groups: [{ id: "g1", title: "Bad", columns: [0, 2] }]
        })
    ).toThrow(CyExcelError);
  });

  it("rejects a group referencing an out-of-range column", () => {
    expect(
      () =>
        new HeaderManager({
          columns: [{ id: "a" }],
          groups: [{ id: "g1", title: "Bad", columns: [0, 1] }]
        })
    ).toThrow(CyExcelError);
  });

  it("rejects a group with an empty columns list", () => {
    expect(
      () =>
        new HeaderManager({
          columns: [{ id: "a" }],
          groups: [{ id: "g1", title: "Bad", columns: [] }]
        })
    ).toThrow(CyExcelError);
  });
});

describe("HeaderManager — resolves once and caches", () => {
  it("getHeader()/getGroups() return the same objects on repeated calls", () => {
    const manager = new HeaderManager({
      columns: [{ id: "a" }],
      groups: [{ id: "g1", title: "G", columns: [0] }]
    });

    expect(manager.getHeader(0)).toBe(manager.getHeader(0));
    expect(manager.getGroups()).toBe(manager.getGroups());
  });
});

describe("HeaderManager.rebuild", () => {
  it("prepares a validated root without changing the live root", () => {
    const manager = new HeaderManager({ columns: [{ id: "a", header: "A" }] });
    const oldHeaders = manager.getHeaders();
    const root = manager.prepareRoot([
      { id: "x", header: "X" },
      { id: "y", header: "Y" }
    ]);

    expect(manager.getHeaders()).toBe(oldHeaders);
    expect(manager.getHeaders().map((header) => header.title)).toEqual(["A"]);

    manager.replaceRoot(root);
    expect(manager.getHeaders().map((header) => header.title)).toEqual(["X", "Y"]);
  });

  it("keeps the old root stable when root preparation fails", () => {
    const manager = new HeaderManager({
      columns: [{ id: "a" }, { id: "b" }],
      groups: [{ id: "good", title: "Good", columns: [0, 1] }]
    });
    const oldHeaders = manager.getHeaders();
    const oldGroups = manager.getGroups();

    expect(() =>
      manager.prepareRoot(
        [{ id: "a" }],
        [{ id: "bad", title: "Bad", columns: [0, 1] }]
      )
    ).toThrow(CyExcelError);

    expect(manager.getHeaders()).toBe(oldHeaders);
    expect(manager.getGroups()).toBe(oldGroups);
    expect(manager.getHeader(0).title).toBe("a");
    expect(manager.getGroups()[0]?.id).toBe("good");
  });

  it("re-resolves headers from a new column list", () => {
    const manager = new HeaderManager({ columns: [{ id: "a", header: "A" }] });
    expect(manager.getHeaders().map((header) => header.title)).toEqual(["A"]);

    manager.replaceRoot(manager.prepareRoot([
      { id: "a", header: "A" },
      { id: "b", header: "B" }
    ]));

    expect(manager.getHeaders().map((header) => header.title)).toEqual(["A", "B"]);
  });

  it("replaces groups with whatever (possibly empty) groups are supplied to prepareRoot()", () => {
    const manager = new HeaderManager({
      columns: [{ id: "a" }, { id: "b" }],
      groups: [{ id: "g", title: "G", columns: [0, 1] }]
    });
    expect(manager.getGroups()).toHaveLength(1);

    manager.replaceRoot(manager.prepareRoot([{ id: "a" }, { id: "b" }, { id: "c" }]));

    expect(manager.getGroups()).toHaveLength(0);
  });

  it("getHeader() reflects the replacement root immediately", () => {
    const manager = new HeaderManager({ columns: [{ id: "a", header: "A" }] });

    manager.replaceRoot(manager.prepareRoot([{ id: "x", header: "X" }]));

    expect(manager.getHeader(0).title).toBe("X");
  });
});

describe("HeaderManager — arbitrary-depth nested groups (Step 19)", () => {
  function threeLevelManager() {
    // Leaf columns 0..3. Level 0 (leaf-adjacent): "cat-a" [0,1], "cat-b" [2,3].
    // Level 1: "dept" wraps both categories -> [0,3].
    // Level 2: "root" wraps dept -> [0,3].
    return new HeaderManager({
      columns: [{ id: "a" }, { id: "b" }, { id: "c" }, { id: "d" }],
      groups: [
        {
          id: "root",
          title: "Root",
          children: [
            {
              id: "dept",
              title: "Dept",
              children: [
                { id: "cat-a", title: "Cat A", columns: [0, 1] },
                { id: "cat-b", title: "Cat B", columns: [2, 3] }
              ]
            }
          ]
        }
      ]
    });
  }

  it("resolves level 0 as the leaf-adjacent groups", () => {
    const manager = threeLevelManager();

    expect(manager.getGroups(0)).toEqual([
      { id: "cat-a", title: "Cat A", startColumn: 0, endColumn: 1 },
      { id: "cat-b", title: "Cat B", startColumn: 2, endColumn: 3 }
    ]);
  });

  it("resolves level 1 as the union of its children's column ranges", () => {
    const manager = threeLevelManager();

    expect(manager.getGroups(1)).toEqual([{ id: "dept", title: "Dept", startColumn: 0, endColumn: 3 }]);
  });

  it("resolves level 2 (the outermost group) the same way", () => {
    const manager = threeLevelManager();

    expect(manager.getGroups(2)).toEqual([{ id: "root", title: "Root", startColumn: 0, endColumn: 3 }]);
  });

  it("groupDepth reflects the deepest branch", () => {
    expect(threeLevelManager().groupDepth).toBe(3);
    expect(new HeaderManager({ columns: [{ id: "a" }] }).groupDepth).toBe(0);
    expect(
      new HeaderManager({
        columns: [{ id: "a" }],
        groups: [{ id: "g", title: "G", columns: [0] }]
      }).groupDepth
    ).toBe(1); // Step 9's single-level case is depth 1.
  });

  it("getGroups() with no argument defaults to level 0 — Step 9 behavior unchanged", () => {
    const manager = new HeaderManager({
      columns: [{ id: "a" }, { id: "b" }],
      groups: [{ id: "g1", title: "Basics", columns: [0, 1] }]
    });

    expect(manager.getGroups()).toEqual(manager.getGroups(0));
  });

  it("getGroups() for a level beyond groupDepth returns an empty array, not an error", () => {
    const manager = threeLevelManager();

    expect(manager.getGroups(3)).toEqual([]);
    expect(manager.getGroups(99)).toEqual([]);
  });

  it("allows branches of different depth — a shallower branch has no group at a level its siblings reach", () => {
    // Columns 0-1 are only grouped one level deep ("shallow"); columns 2-3
    // are grouped two levels deep via "wrapper" -> "deep".
    const manager = new HeaderManager({
      columns: [{ id: "a" }, { id: "b" }, { id: "c" }, { id: "d" }],
      groups: [
        { id: "shallow", title: "Shallow", columns: [0, 1] },
        {
          id: "wrapper",
          title: "Wrapper",
          children: [{ id: "deep", title: "Deep", columns: [2, 3] }]
        }
      ]
    });

    expect(manager.groupDepth).toBe(2);
    expect(manager.getGroups(0)).toEqual([
      { id: "shallow", title: "Shallow", startColumn: 0, endColumn: 1 },
      { id: "deep", title: "Deep", startColumn: 2, endColumn: 3 }
    ]);
    // "shallow" has nothing above it at level 1 — only "wrapper" is there.
    expect(manager.getGroups(1)).toEqual([{ id: "wrapper", title: "Wrapper", startColumn: 2, endColumn: 3 }]);
  });

  it("rejects a node that specifies both columns and children", () => {
    expect(
      () =>
        new HeaderManager({
          columns: [{ id: "a" }, { id: "b" }],
          groups: [{ id: "bad", title: "Bad", columns: [0, 1], children: [{ id: "x", title: "X", columns: [0] }] }]
        })
    ).toThrow(CyExcelError);
  });

  it("rejects a node that specifies neither columns nor children", () => {
    expect(
      () =>
        new HeaderManager({
          columns: [{ id: "a" }],
          groups: [{ id: "bad", title: "Bad" }]
        })
    ).toThrow(CyExcelError);
  });

  it("rejects an empty children array", () => {
    expect(
      () =>
        new HeaderManager({
          columns: [{ id: "a" }],
          groups: [{ id: "bad", title: "Bad", children: [] }]
        })
    ).toThrow(CyExcelError);
  });

  it("rejects children whose column ranges leave a gap", () => {
    expect(
      () =>
        new HeaderManager({
          columns: [{ id: "a" }, { id: "b" }, { id: "c" }, { id: "d" }],
          groups: [
            {
              id: "parent",
              title: "Parent",
              children: [
                { id: "c1", title: "C1", columns: [0] },
                { id: "c2", title: "C2", columns: [2, 3] } // gap at column 1
              ]
            }
          ]
        })
    ).toThrow(CyExcelError);
  });

  it("rejects children whose column ranges overlap", () => {
    expect(
      () =>
        new HeaderManager({
          columns: [{ id: "a" }, { id: "b" }, { id: "c" }],
          groups: [
            {
              id: "parent",
              title: "Parent",
              children: [
                { id: "c1", title: "C1", columns: [0, 1] },
                { id: "c2", title: "C2", columns: [1, 2] } // overlaps at column 1
              ]
            }
          ]
        })
    ).toThrow(CyExcelError);
  });

  it("prepareRoot() re-resolves a nested tree from scratch", () => {
    const manager = threeLevelManager();
    expect(manager.groupDepth).toBe(3);

    manager.replaceRoot(manager.prepareRoot(
      [{ id: "a" }, { id: "b" }],
      [{ id: "flat", title: "Flat", columns: [0, 1] }]
    ));

    expect(manager.groupDepth).toBe(1);
    expect(manager.getGroups(0)).toEqual([{ id: "flat", title: "Flat", startColumn: 0, endColumn: 1 }]);
  });
});
