import { describe, expect, it } from "vitest";

import { CyExcelError, HeaderManager } from "../../../../src/core";
import type { CanvasHeaderRenderer } from "../../../../src/core";

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

describe("HeaderManager.getGroupFragments — 先写测试 #4: split at the frozen boundary, same logical id", () => {
  it("returns one fragment when the group doesn't straddle the frozen boundary", () => {
    const manager = new HeaderManager({
      columns: [{ id: "a" }, { id: "b" }, { id: "c" }],
      groups: [{ id: "g1", title: "G", columns: [1, 2] }]
    });

    const fragments = manager.getGroupFragments(1); // frozenColumnCount 1 — group starts after it
    expect(fragments).toEqual([
      { group: manager.getGroups()[0], startColumn: 1, endColumn: 2, frozenColumn: false }
    ]);
  });

  it("splits a group straddling the boundary into two fragments sharing the same group object/id", () => {
    const manager = new HeaderManager({
      columns: [{ id: "a" }, { id: "b" }, { id: "c" }],
      groups: [{ id: "g1", title: "G", columns: [0, 1, 2] }]
    });

    const fragments = manager.getGroupFragments(1); // frozenColumnCount 1 splits [0,2] into [0,0] + [1,2]

    expect(fragments).toHaveLength(2);
    expect(fragments[0]).toEqual({
      group: manager.getGroups()[0],
      startColumn: 0,
      endColumn: 0,
      frozenColumn: true
    });
    expect(fragments[1]).toEqual({
      group: manager.getGroups()[0],
      startColumn: 1,
      endColumn: 2,
      frozenColumn: false
    });
    // Same logical group — both fragments reference the identical object.
    expect(fragments[0]?.group).toBe(fragments[1]?.group);
  });

  it("returns no fragments for freeze counts that don't apply (frozenColumnCount 0)", () => {
    const manager = new HeaderManager({
      columns: [{ id: "a" }, { id: "b" }],
      groups: [{ id: "g1", title: "G", columns: [0, 1] }]
    });

    const fragments = manager.getGroupFragments(0);

    expect(fragments).toEqual([
      { group: manager.getGroups()[0], startColumn: 0, endColumn: 1, frozenColumn: false }
    ]);
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
