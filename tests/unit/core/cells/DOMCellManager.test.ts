import { describe, expect, it } from "vitest";

import { DOMCellManager, LayoutEngine } from "../../../../src/core";
import type { DomCellContext, DomCellRenderer, DomCellRendererFactory, RuntimeColumn } from "../../../../src/core";

interface FakeElement {
  readonly style: Record<string, string>;
  readonly children: FakeElement[];
  append(child: FakeElement): void;
  remove(): void;
}

function createFakeElement(): FakeElement {
  return {
    style: {},
    children: [],
    append(child: FakeElement) {
      this.children.push(child);
    },
    remove() {
      // A real element detaches itself from whatever parent holds it; the
      // fake only needs to be inspectable, not to actually track a parent
      // back-reference for these tests.
    }
  };
}

interface RecordingInstance extends DomCellRenderer {
  readonly calls: string[];
  readonly contexts: DomCellContext[];
}

function createRecordingFactory(): {
  readonly factory: DomCellRendererFactory;
  readonly instances: RecordingInstance[];
} {
  const instances: RecordingInstance[] = [];
  const factory: DomCellRendererFactory = {
    kind: "dom",
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

function textColumn(id: string): RuntimeColumn {
  return {
    id,
    field: id,
    width: 50,
    typeName: "text",
    format: (v) => String(v ?? ""),
    clipboard: { copy: true, paste: true },
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
    version: 1,
    renderer: factory
  };
}

function createManager(options: {
  readonly layout: LayoutEngine;
  readonly columns: readonly RuntimeColumn[];
}) {
  const container = createFakeElement();
  const hosts: FakeElement[] = [];
  const manager = new DOMCellManager({
    container: container as unknown as HTMLElement,
    layout: options.layout,
    getColumn: (index) => options.columns[index] as RuntimeColumn,
    getCellValue: () => "value",
    commitValue: () => {},
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

    manager.sync({ scrollTop: 0, scrollLeft: 0, width: 100, height: 40 });

    // Only column 1 ("interactive") has a DOM renderer; column 0 never
    // gets an instance regardless of visibility.
    expect(instances.length).toBeGreaterThan(0);
    expect(instances.every((instance) => instance.calls[0] === "mount")).toBe(true);
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

    manager.sync({ scrollTop: 0, scrollLeft: 0, width: 50, height: 40 }); // rows ~0-1 visible
    const afterFirstSync = instances.length;

    manager.sync({ scrollTop: 2000, scrollLeft: 0, width: 50, height: 40 }); // far away, non-overlapping rows
    const afterSecondSync = instances.length;

    // No new instances were created — the ones freed by the first sync's
    // now-invisible cells were reused via update(), not destroy()+recreate.
    expect(afterSecondSync).toBe(afterFirstSync);
    expect(instances.some((instance) => instance.calls.includes("update"))).toBe(true);
    expect(instances.some((instance) => instance.calls.includes("destroy"))).toBe(false);
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

    manager.sync(viewport);
    const countAfterFirstSync = instances.length;
    manager.sync(viewport);

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

    manager.sync({ scrollTop: 5000, scrollLeft: 0, width: 100, height: 40 });

    // Row 0 / column 0 is the frozen corner — it must still be mounted
    // despite scrollTop being nowhere near row 0's absolute position.
    expect(instances.length).toBeGreaterThan(0);
    expect(instances[0]?.calls).toContain("mount");
  });
});

describe("DOMCellManager — instance count does not grow with total row count", () => {
  it("stays bounded to roughly the visible count across many scroll positions on a million-row layout", () => {
    const { factory, instances } = createRecordingFactory();
    const layout = new LayoutEngine({
      rowCount: 1_000_000,
      columnCount: 1,
      rowHeight: 20,
      columnWidths: [50]
    });
    const { manager } = createManager({ layout, columns: [domColumn("interactive", factory)] });

    for (let scrollTop = 0; scrollTop < 1_000_000; scrollTop += 5000) {
      manager.sync({ scrollTop, scrollLeft: 0, width: 50, height: 40 });
    }

    // Nowhere near 1,000,000 — bounded by the small visible+overscan count.
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

    manager.sync({ scrollTop: 0, scrollLeft: 0, width: 50, height: 40 });
    manager.sync({ scrollTop: 2000, scrollLeft: 0, width: 50, height: 40 });
    manager.dispose();

    expect(instances.length).toBeGreaterThan(0);
    expect(instances.every((instance) => instance.calls.includes("destroy"))).toBe(true);
  });
});
