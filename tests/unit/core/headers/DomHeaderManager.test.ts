import { afterEach, describe, expect, it, vi } from "vitest";

import type { DomHeaderContext, DomHeaderRenderer, DomHeaderRendererFactory } from "../../../../src/core/cells";
import { DomHeaderManager, HeaderManager } from "../../../../src/core/headers";
import { getDeclaredDomPointerInputTarget } from "../../../../src/core/interaction/domPointerOwnership";
import { LayoutEngine, SkeletonSnapshot, ViewportSnapshot } from "../../../../src/core/layout";
import type { ViewportState } from "../../../../src/core/layout";
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

interface RecordingInstance extends DomHeaderRenderer {
  readonly calls: string[];
  readonly contexts: DomHeaderContext[];
}

function createRecordingFactory(pointerOwnership: "native" | "surface" = "surface"): {
  readonly factory: DomHeaderRendererFactory;
  readonly instances: RecordingInstance[];
} {
  const instances: RecordingInstance[] = [];
  const factory: DomHeaderRendererFactory = {
    kind: "dom",
    pointerOwnership,
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

function createScene(
  layout: LayoutEngine,
  viewport: ViewportState,
  columnHeaderHeight = 24,
  leafY = 0,
  leafHeight = columnHeaderHeight - leafY
): SceneSnapshot {
  return new SceneSnapshot({
    version: 1,
    skeleton: SkeletonSnapshot.fromLayout(layout, 1),
    viewport: new ViewportSnapshot({ version: 1, ...viewport }),
    surfaceMetrics: {
      columnHeaderHeight,
      columnHeaderBands: [{ role: "leaf", y: leafY, height: leafHeight }],
      rowHeaderWidth: 0
    }
  });
}

function createManager(options: {
  readonly layout: LayoutEngine;
  readonly headerManager: HeaderManager;
}) {
  const container = createFakeElement();
  const hosts: FakeElement[] = [];
  const manager = new DomHeaderManager({
    container: container as unknown as HTMLElement,
    headerManager: options.headerManager,
    createHost: () => {
      const host = createFakeElement();
      hosts.push(host);
      return host as unknown as HTMLElement;
    }
  });
  return { manager, container, hosts };
}

afterEach(() => {
  vi.restoreAllMocks();
});

describe("DomHeaderManager — mounts within the visible/frozen column set (先写测试 #2)", () => {
  it("mounts DOM-renderer headers that are currently visible, and only those columns", () => {
    const { factory, instances } = createRecordingFactory();
    const headerManager = new HeaderManager({
      columns: [
        { header: "Plain" },
        { header: { title: "Interactive", renderer: factory } }
      ]
    });
    const layout = createLayout(2);
    const { manager } = createManager({ layout, headerManager });

    manager.sync(createScene(layout, { scrollTop: 0, scrollLeft: 0, width: 100, height: 40 }));

    expect(instances.length).toBeGreaterThan(0);
    expect(instances.every((instance) => instance.calls[0] === "mount")).toBe(true);
  });

  it("binds surface ownership to the mounted column-header host", () => {
    const { factory } = createRecordingFactory("surface");
    const headerManager = new HeaderManager({
      columns: [{ header: { title: "Surface", renderer: factory } }]
    });
    const layout = createLayout(1);
    const { manager, hosts } = createManager({ layout, headerManager });

    manager.sync(createScene(layout, { scrollTop: 0, scrollLeft: 0, width: 50, height: 20 }));

    expect(getDeclaredDomPointerInputTarget(hosts[0] as unknown as EventTarget))
      .toEqual({ kind: "surface", surface: "column-header" });
  });

  it("keeps frozen-column DOM headers mounted even when scrolled far past them", () => {
    const { factory, instances } = createRecordingFactory();
    const headerManager = new HeaderManager({
      columns: [
        { header: { title: "Frozen", renderer: factory } },
        ...Array.from({ length: 20 }, (_unused, index) => ({ id: `c${index}`, header: `C${index}` }))
      ]
    });
    const layout = createLayout(21, { columns: 1 });
    const { manager } = createManager({ layout, headerManager });

    manager.sync(createScene(layout, { scrollTop: 0, scrollLeft: 800, width: 100, height: 40 }));

    expect(instances.length).toBeGreaterThan(0);
    expect(instances[0]?.calls).toContain("mount");
  });

  it("clips a partially scrolled DOM header to its canonical scrollable pane", () => {
    const { factory } = createRecordingFactory();
    const headerManager = new HeaderManager({
      columns: [
        { header: "Frozen" },
        { header: { title: "Scroll", renderer: factory } }
      ]
    });
    const layout = createLayout(2, { columns: 1 });
    const { manager, hosts } = createManager({ layout, headerManager });

    manager.sync(createScene(layout, { scrollTop: 0, scrollLeft: 20, width: 100, height: 40 }));

    expect(hosts[0]?.style.left).toBe("30px");
    expect(hosts[0]?.style.width).toBe("50px");
    expect(hosts[0]?.style.clipPath).toBe("inset(0px 0px 0px 20px)");
  });

  it("uses canonical column-header fragments without projecting body pane cells", () => {
    const { factory } = createRecordingFactory();
    const headerManager = new HeaderManager({
      columns: [{ header: { title: "Interactive", renderer: factory } }]
    });
    const layout = createLayout(1);
    const { manager } = createManager({ layout, headerManager });
    const getVisibleFragments = vi.spyOn(SceneSnapshot.prototype, "getVisibleColumnHeaderFragments");
    const getBandFragments = vi.spyOn(SceneSnapshot.prototype, "getColumnHeaderBandFragments");
    const getCellRectInPane = vi.spyOn(SceneSnapshot.prototype, "getCellRectInPane");
    const scene = createScene(layout, {
      scrollTop: 0,
      scrollLeft: 0,
      width: 100,
      height: 40
    });

    manager.sync(scene);

    expect(getVisibleFragments).toHaveBeenCalledWith(scene.getColumnHeaderLeafBand());
    expect(getBandFragments).toHaveBeenCalled();
    expect(getCellRectInPane).not.toHaveBeenCalled();
    const [fragment] = scene.getColumnHeaderBandFragments(
      scene.getColumnHeaderLeafBand(),
      0,
      0
    );
    expect(fragment?.surface).toBe("column-header");
    expect(fragment?.pane.surface).toBe("column-header");
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
    const layout = createLayout(20);
    const { manager } = createManager({ layout, headerManager });

    manager.sync(createScene(layout, { scrollTop: 0, scrollLeft: 0, width: 100, height: 40 }));
    const afterFirstSync = instances.length;

    manager.sync(createScene(layout, { scrollTop: 0, scrollLeft: 800, width: 100, height: 40 }));
    const afterSecondSync = instances.length;

    expect(afterSecondSync).toBe(afterFirstSync);
    expect(instances.some((instance) => instance.calls.includes("update"))).toBe(true);
    expect(instances.some((instance) => instance.calls.includes("destroy"))).toBe(false);
  });

  it("destroys every instance on dispose()", () => {
    const { factory, instances } = createRecordingFactory();
    const headerManager = new HeaderManager({
      columns: [{ header: { title: "A", renderer: factory } }]
    });
    const layout = createLayout(1);
    const { manager } = createManager({ layout, headerManager });

    manager.sync(createScene(layout, { scrollTop: 0, scrollLeft: 0, width: 100, height: 40 }));
    manager.dispose();

    expect(instances.length).toBeGreaterThan(0);
    expect(instances.every((instance) => instance.calls.includes("destroy"))).toBe(true);
  });

  it("removes a mounted header's node when it's hidden by no longer being needed anywhere", () => {
    const { factory, instances } = createRecordingFactory();
    const headerManager = new HeaderManager({
      columns: [{ header: { title: "A", renderer: factory } }]
    });
    const layout = createLayout(1);
    const { manager, hosts } = createManager({ layout, headerManager });

    manager.sync(createScene(layout, { scrollTop: 0, scrollLeft: 0, width: 100, height: 40 }));
    manager.sync(createScene(layout, { scrollTop: 0, scrollLeft: 0, width: 0, height: 40 })); // nothing visible now

    // The host was detached (pooled) once its column stopped being needed.
    expect(hosts[0]?.removed).toBe(true);
    expect(instances[0]?.calls).not.toContain("destroy"); // pooled, not destroyed
  });
});

describe("DomHeaderManager — failed mount cleanup", () => {
  it("removes the host and destroys a partially created renderer", () => {
    const calls: string[] = [];
    const factory: DomHeaderRendererFactory = {
      kind: "dom",
      pointerOwnership: "surface",
      create: () => ({
        mount: () => { calls.push("mount"); throw new Error("mount failed"); },
        update: () => {},
        destroy: () => { calls.push("destroy"); }
      })
    };
    const headerManager = new HeaderManager({
      columns: [{ header: { title: "Broken", renderer: factory } }]
    });
    const layout = createLayout(1);
    const { manager, hosts } = createManager({ layout, headerManager });

    expect(() => manager.sync(createScene(layout, { scrollTop: 0, scrollLeft: 0, width: 50, height: 20 })))
      .toThrow("mount failed");
    expect(hosts[0]?.removed).toBe(true);
    expect(calls).toEqual(["mount", "destroy"]);
  });
});
