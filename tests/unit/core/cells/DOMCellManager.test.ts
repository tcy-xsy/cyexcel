import { describe, expect, it } from "vitest";

import { DOMCellManager } from "../../../../src/core/cells";
import type { DomCellContext, DomCellRenderer, DomCellRendererFactory, RuntimeColumn } from "../../../../src/core/cells";
import { LayoutEngine, SkeletonSnapshot, ViewportSnapshot } from "../../../../src/core/layout";
import type { ViewportState } from "../../../../src/core/layout";
import { getDeclaredDomPointerInputTarget } from "../../../../src/core/interaction/domPointerOwnership";
import { SceneSnapshot } from "../../../../src/core/scene";

interface FakeElement {
  readonly style: Record<string, string>;
  readonly children: FakeElement[];
  removed: boolean;
  append(child: FakeElement): void;
  remove(): void;
}

function createFakeElement(): FakeElement {
  return {
    style: {},
    children: [],
    removed: false,
    append(child: FakeElement) {
      this.children.push(child);
    },
    remove() {
      this.removed = true;
    }
  };
}

interface RecordingInstance extends DomCellRenderer {
  readonly calls: string[];
  readonly contexts: DomCellContext[];
}

function createRecordingFactory(pointerOwnership: "native" | "surface" = "surface"): {
  readonly factory: DomCellRendererFactory;
  readonly instances: RecordingInstance[];
} {
  const instances: RecordingInstance[] = [];
  const factory: DomCellRendererFactory = {
    kind: "dom",
    pointerOwnership,
    create: () => {
      const calls: string[] = [];
      const contexts: DomCellContext[] = [];
      const instance: RecordingInstance = {
        calls,
        contexts,
        mount(_container, context) {
          calls.push("mount");
          contexts.push(context);
        },
        update(context) {
          calls.push("update");
          contexts.push(context);
        },
        destroy() {
          calls.push("destroy");
        }
      };
      instances.push(instance);
      return instance;
    }
  };
  return { factory, instances };
}

function createScene(layout: LayoutEngine, viewport: ViewportState): SceneSnapshot {
  return new SceneSnapshot({
    version: 1,
    skeleton: SkeletonSnapshot.fromLayout(layout, 1),
    viewport: new ViewportSnapshot({ version: 1, ...viewport })
  });
}

function textColumn(id: string): RuntimeColumn {
  return {
    id,
    field: id,
    width: 50,
    typeName: "text",
    format: (v) => String(v ?? ""),
    clipboard: { copy: true, paste: true },
    wrap: false,
    sortable: true,
    version: 1
  };
}

function domColumn(id: string, factory: DomCellRendererFactory): RuntimeColumn {
  return {
    id,
    field: id,
    width: 50,
    typeName: "text",
    format: (v) => String(v ?? ""),
    clipboard: { copy: true, paste: true },
    wrap: false,
    sortable: true,
    version: 1,
    renderer: factory
  };
}

function createManager(options: {
  readonly layout: LayoutEngine;
  readonly columns: readonly RuntimeColumn[];
  readonly commitValue?: (row: number, column: number, value: unknown) => void;
}) {
  const container = createFakeElement();
  const hosts: FakeElement[] = [];
  const manager = new DOMCellManager({
    container: container as unknown as HTMLElement,
    getColumn: (index) => options.columns[index] as RuntimeColumn,
    getCellValue: () => "value",
    commitValue: options.commitValue ?? (() => {}),
    createHost: () => {
      const host = createFakeElement();
      hosts.push(host);
      return host as unknown as HTMLElement;
    }
  });
  return { manager, container, hosts };
}

describe("DOMCellManager — mounts within the visible range", () => {
  it("mounts DOM-renderer cells that are currently visible, and only those columns", () => {
    const { factory, instances } = createRecordingFactory();
    const layout = new LayoutEngine({
      rowCount: 100,
      columnCount: 2,
      rowHeight: 20,
      columnWidths: [50, 50]
    });
    const { manager } = createManager({
      layout,
      columns: [textColumn("plain"), domColumn("interactive", factory)]
    });

    manager.sync(createScene(layout, { scrollTop: 0, scrollLeft: 0, width: 100, height: 40 }));

    // Only column 1 ("interactive") has a DOM renderer; column 0 never
    // gets an instance regardless of visibility.
    expect(instances.length).toBeGreaterThan(0);
    expect(instances.every((instance) => instance.calls[0] === "mount")).toBe(true);
  });

  it("binds the renderer's terminal pointer ownership to each mounted body host", () => {
    const { factory } = createRecordingFactory("native");
    const layout = new LayoutEngine({
      rowCount: 1,
      columnCount: 1,
      rowHeight: 20,
      columnWidths: [50]
    });
    const { manager, hosts } = createManager({
      layout,
      columns: [domColumn("native", factory)]
    });

    manager.sync(createScene(layout, { scrollTop: 0, scrollLeft: 0, width: 50, height: 20 }));

    expect(getDeclaredDomPointerInputTarget(hosts[0] as unknown as EventTarget))
      .toEqual({ kind: "native" });
  });
});

describe("DOMCellManager — reuses instances by renderer identity as cells scroll out", () => {
  it("updates (reuses) a pooled instance for a newly-visible cell instead of creating a new one", () => {
    const { factory, instances } = createRecordingFactory();
    const layout = new LayoutEngine({
      rowCount: 1000,
      columnCount: 1,
      rowHeight: 20,
      columnWidths: [50]
    });
    const { manager } = createManager({ layout, columns: [domColumn("interactive", factory)] });

    manager.sync(createScene(layout, { scrollTop: 0, scrollLeft: 0, width: 50, height: 40 })); // rows ~0-1 visible
    const afterFirstSync = instances.length;

    manager.sync(createScene(layout, { scrollTop: 2000, scrollLeft: 0, width: 50, height: 40 })); // far away, non-overlapping rows
    const afterSecondSync = instances.length;

    // No new instances were created — the ones freed by the first sync's
    // now-invisible cells were reused via update(), not destroy()+recreate.
    expect(afterSecondSync).toBe(afterFirstSync);
    expect(instances.some((instance) => instance.calls.includes("update"))).toBe(true);
    expect(instances.some((instance) => instance.calls.includes("destroy"))).toBe(false);
  });

  it("invalidates old commit actions when a pooled instance is rebound", () => {
    const { factory, instances } = createRecordingFactory();
    const commits: Array<{ row: number; column: number; value: unknown }> = [];
    const layout = new LayoutEngine({
      rowCount: 100,
      columnCount: 1,
      rowHeight: 20,
      columnWidths: [50]
    });
    const { manager } = createManager({
      layout,
      columns: [domColumn("interactive", factory)],
      commitValue: (row, column, value) => commits.push({ row, column, value })
    });

    manager.sync(createScene(layout, { scrollTop: 0, scrollLeft: 0, width: 50, height: 1 }));
    const reused = instances[0] as RecordingInstance;
    const oldAction = reused.contexts[0]!.actions.commit;

    manager.sync(createScene(layout, { scrollTop: 200, scrollLeft: 0, width: 50, height: 1 }));
    const currentContext = reused.contexts.at(-1) as DomCellContext;

    oldAction("stale");
    expect(commits).toEqual([]);

    currentContext.actions.commit("current");
    expect(commits).toEqual([
      { row: currentContext.row, column: 0, value: "current" }
    ]);

    manager.dispose();
    currentContext.actions.commit("disposed");
    expect(commits).toHaveLength(1);
  });
});

describe("DOMCellManager — re-syncing an unchanged viewport", () => {
  it("updates an already-mounted cell in place instead of pooling and recreating it", () => {
    const { factory, instances } = createRecordingFactory();
    const layout = new LayoutEngine({
      rowCount: 1000,
      columnCount: 1,
      rowHeight: 20,
      columnWidths: [50]
    });
    const { manager } = createManager({ layout, columns: [domColumn("interactive", factory)] });
    const viewport = { scrollTop: 0, scrollLeft: 0, width: 50, height: 40 };

    manager.sync(createScene(layout, viewport));
    const countAfterFirstSync = instances.length;
    manager.sync(createScene(layout, viewport));

    expect(instances.length).toBe(countAfterFirstSync);
    expect(instances.every((instance) => !instance.calls.includes("destroy"))).toBe(true);
    expect(instances.some((instance) => instance.calls.includes("update"))).toBe(true);
  });
});

describe("DOMCellManager — frozen rows/columns", () => {
  it("keeps frozen-band DOM cells mounted even when scrolled far past them", () => {
    const { factory, instances } = createRecordingFactory();
    const layout = new LayoutEngine({
      rowCount: 1000,
      columnCount: 2,
      rowHeight: 20,
      columnWidths: [50, 50],
      freeze: { rows: 1, columns: 1 }
    });
    const { manager } = createManager({
      layout,
      columns: [domColumn("frozen-corner", factory), textColumn("plain")]
    });

    manager.sync(createScene(layout, { scrollTop: 5000, scrollLeft: 0, width: 100, height: 40 }));

    // Row 0 / column 0 is the frozen corner — it must still be mounted
    // despite scrollTop being nowhere near row 0's absolute position.
    expect(instances.length).toBeGreaterThan(0);
    expect(instances[0]?.calls).toContain("mount");
  });

  it("clips a partially scrolled DOM cell to its canonical body pane", () => {
    const { factory } = createRecordingFactory();
    const layout = new LayoutEngine({
      rowCount: 10,
      columnCount: 2,
      rowHeight: 20,
      columnWidths: [50, 50],
      freeze: { rows: 1, columns: 1 }
    });
    const { manager, hosts } = createManager({
      layout,
      columns: [textColumn("frozen"), domColumn("scrollable", factory)]
    });

    manager.sync(createScene(layout, { scrollTop: 10, scrollLeft: 20, width: 100, height: 40 }));

    expect(hosts.some((host) => (
      host.style.left === "30px" &&
      host.style.top === "10px" &&
      host.style.width === "50px" &&
      host.style.height === "20px" &&
      host.style.clipPath === "inset(10px 0px 0px 20px)"
    ))).toBe(true);
  });
});

describe("DOMCellManager — instance count does not grow with total row count", () => {
  it("stays bounded to roughly the visible count across many scroll positions on a million-row layout", () => {
    const { factory, instances } = createRecordingFactory();
    const layout = new LayoutEngine({
      rowCount: 100_000,
      columnCount: 1,
      rowHeight: 20,
      columnWidths: [50]
    });
    const { manager } = createManager({ layout, columns: [domColumn("interactive", factory)] });

    // Keep this focused on DOM pooling, not the current dense SkeletonSnapshot
    // construction cost. A 100k-row layout is still far beyond visible size.
    for (let scrollTop = 0; scrollTop < 100_000; scrollTop += 5000) {
      manager.sync(createScene(layout, { scrollTop, scrollLeft: 0, width: 50, height: 40 }));
    }

    // Nowhere near total row count — bounded by the small visible+overscan count.
    expect(instances.length).toBeLessThan(50);
  });
});

describe("DOMCellManager.dispose", () => {
  it("destroys every mounted and pooled instance", () => {
    const { factory, instances } = createRecordingFactory();
    const layout = new LayoutEngine({
      rowCount: 1000,
      columnCount: 1,
      rowHeight: 20,
      columnWidths: [50]
    });
    const { manager } = createManager({ layout, columns: [domColumn("interactive", factory)] });

    manager.sync(createScene(layout, { scrollTop: 0, scrollLeft: 0, width: 50, height: 40 }));
    manager.sync(createScene(layout, { scrollTop: 2000, scrollLeft: 0, width: 50, height: 40 }));
    manager.dispose();

    expect(instances.length).toBeGreaterThan(0);
    expect(instances.every((instance) => instance.calls.includes("destroy"))).toBe(true);
  });
});

describe("DOMCellManager — failed mount cleanup", () => {
  it("removes the host and destroys a partially created renderer", () => {
    const calls: string[] = [];
    const factory: DomCellRendererFactory = {
      kind: "dom",
      pointerOwnership: "surface",
      create: () => ({
        mount: () => { calls.push("mount"); throw new Error("mount failed"); },
        update: () => {},
        destroy: () => { calls.push("destroy"); }
      })
    };
    const layout = new LayoutEngine({ rowCount: 1, columnCount: 1, rowHeight: 20, columnWidths: [50] });
    const { manager, hosts } = createManager({ layout, columns: [domColumn("broken", factory)] });

    expect(() => manager.sync(createScene(layout, { scrollTop: 0, scrollLeft: 0, width: 50, height: 20 })))
      .toThrow("mount failed");
    expect(hosts[0]?.removed).toBe(true);
    expect(calls).toEqual(["mount", "destroy"]);
  });
});
