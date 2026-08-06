import { describe, expect, it } from "vitest";

import { DomHeaderManager, HeaderManager, LayoutEngine } from "../../../../src/core";
import type { DomHeaderContext, DomHeaderRenderer, DomHeaderRendererFactory } from "../../../../src/core";

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

interface RecordingInstance extends DomHeaderRenderer {
  readonly calls: string[];
  readonly contexts: DomHeaderContext[];
}

function createRecordingFactory(): {
  readonly factory: DomHeaderRendererFactory;
  readonly instances: RecordingInstance[];
} {
  const instances: RecordingInstance[] = [];
  const factory: DomHeaderRendererFactory = {
    kind: "dom",
    create: () => {
      const calls: string[] = [];
      const contexts: DomHeaderContext[] = [];
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

function createLayout(columnCount: number, freeze?: { columns?: number }): LayoutEngine {
  return new LayoutEngine({
    rowCount: 10,
    columnCount,
    rowHeight: 20,
    columnWidths: Array.from({ length: columnCount }, () => 50),
    ...(freeze ? { freeze } : {})
  });
}

function createManager(options: { readonly layout: LayoutEngine; readonly headerManager: HeaderManager }) {
  const container = createFakeElement();
  const hosts: FakeElement[] = [];
  const manager = new DomHeaderManager({
    container: container as unknown as HTMLElement,
    layout: options.layout,
    headerManager: options.headerManager,
    columnHeaderHeight: 24,
    createHost: () => {
      const host = createFakeElement();
      hosts.push(host);
      return host as unknown as HTMLElement;
    }
  });
  return { manager, container, hosts };
}

describe("DomHeaderManager — mounts within the visible/frozen column set (先写测试 #2)", () => {
  it("mounts DOM-renderer headers that are currently visible, and only those columns", () => {
    const { factory, instances } = createRecordingFactory();
    const headerManager = new HeaderManager({
      columns: [
        { id: "plain", header: "Plain" },
        { id: "interactive", header: { title: "Interactive", renderer: factory } }
      ]
    });
    const { manager } = createManager({ layout: createLayout(2), headerManager });

    manager.sync({ scrollTop: 0, scrollLeft: 0, width: 100, height: 40 });

    expect(instances.length).toBeGreaterThan(0);
    expect(instances.every((instance) => instance.calls[0] === "mount")).toBe(true);
  });

  it("keeps frozen-column DOM headers mounted even when scrolled far past them", () => {
    const { factory, instances } = createRecordingFactory();
    const headerManager = new HeaderManager({
      columns: [
        { id: "frozen", header: { title: "Frozen", renderer: factory } },
        ...Array.from({ length: 20 }, (_unused, index) => ({ id: `c${index}`, header: `C${index}` }))
      ]
    });
    const { manager } = createManager({ layout: createLayout(21, { columns: 1 }), headerManager });

    manager.sync({ scrollTop: 0, scrollLeft: 800, width: 100, height: 40 });

    expect(instances.length).toBeGreaterThan(0);
    expect(instances[0]?.calls).toContain("mount");
  });
});

describe("DomHeaderManager — scrolling out (先写测试 #6: update or destroy)", () => {
  it("reuses an instance via update() for a newly-visible column instead of creating a new one", () => {
    const { factory, instances } = createRecordingFactory();
    const headerManager = new HeaderManager({
      columns: Array.from({ length: 20 }, (_unused, index) => ({
        id: `c${index}`,
        header: { title: `C${index}`, renderer: factory }
      }))
    });
    const { manager } = createManager({ layout: createLayout(20), headerManager });

    manager.sync({ scrollTop: 0, scrollLeft: 0, width: 100, height: 40 });
    const afterFirstSync = instances.length;

    manager.sync({ scrollTop: 0, scrollLeft: 800, width: 100, height: 40 });
    const afterSecondSync = instances.length;

    expect(afterSecondSync).toBe(afterFirstSync);
    expect(instances.some((instance) => instance.calls.includes("update"))).toBe(true);
    expect(instances.some((instance) => instance.calls.includes("destroy"))).toBe(false);
  });

  it("destroys every instance on dispose()", () => {
    const { factory, instances } = createRecordingFactory();
    const headerManager = new HeaderManager({
      columns: [{ id: "a", header: { title: "A", renderer: factory } }]
    });
    const { manager } = createManager({ layout: createLayout(1), headerManager });

    manager.sync({ scrollTop: 0, scrollLeft: 0, width: 100, height: 40 });
    manager.dispose();

    expect(instances.length).toBeGreaterThan(0);
    expect(instances.every((instance) => instance.calls.includes("destroy"))).toBe(true);
  });

  it("removes a mounted header's node when it's hidden by no longer being needed anywhere", () => {
    const { factory, instances } = createRecordingFactory();
    const headerManager = new HeaderManager({
      columns: [{ id: "a", header: { title: "A", renderer: factory } }]
    });
    const { manager, hosts } = createManager({ layout: createLayout(1), headerManager });

    manager.sync({ scrollTop: 0, scrollLeft: 0, width: 100, height: 40 });
    manager.sync({ scrollTop: 0, scrollLeft: 0, width: 0, height: 40 }); // nothing visible now

    // The host was detached (pooled) once its column stopped being needed.
    expect(hosts[0]?.removed).toBe(true);
    expect(instances[0]?.calls).not.toContain("destroy"); // pooled, not destroyed
  });
});
