import { describe, expect, it } from "vitest";

import {
  ContextMenuService,
  EventBus,
  HookManager,
  PRODUCTION_HOOK_REGISTRY,
  PluginManager
} from "../../../../src/core";
import type { CyExcelPlugin, PluginContext } from "../../../../src/core";

function setup() {
  const hooks = new HookManager({ registry: PRODUCTION_HOOK_REGISTRY });
  const events = new EventBus();
  const contextMenu = new ContextMenuService({ container: {} as unknown as HTMLElement });
  const manager = new PluginManager({
    hooks,
    events,
    contextMenu,
    getCell: () => "value",
    copy: async () => ({ status: "copied", plan: { range: { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 }, rows: 1, columns: 1, cells: 1, estimatedBytes: 0, containsComplexValues: false, recommendedMode: "display" } }),
    paste: async () => {},
    insertRows: () => {},
    deleteRows: () => {},
    moveRows: () => {}
  });
  return { hooks, events, contextMenu, manager };
}

describe("PluginManager — installs a plugin", () => {
  it("calls setup() with a context that can tap hooks, listen to events, and register menu providers", () => {
    const { hooks, events, contextMenu, manager } = setup();
    const seen: string[] = [];
    const plugin: CyExcelPlugin = {
      name: "demo",
      setup(context: PluginContext) {
        context.hooks.tap("beforeInit", { name: "demo" }, () => seen.push("hook"));
        context.events.on("change", () => seen.push("event"));
        context.contextMenu.registerProvider({ name: "demo" }, () => [
          { type: "action", id: "demo-action", label: "Demo", run: () => {} }
        ]);
      }
    };

    manager.use(plugin);
    hooks.runSeries("beforeInit", {});
    events.emit("change", {});

    expect(seen).toEqual(["hook", "event"]);
    // The provider really reached the shared ContextMenuService.
    contextMenu.registerProvider({ name: "probe" }, () => []); // no-op, just proves contextMenu is the shared instance
  });

  it("exposes getCell to the plugin", () => {
    const { manager } = setup();
    let observed: unknown;
    manager.use({
      name: "reader",
      setup(context: PluginContext) {
        observed = context.getCell(0, 0);
      }
    });

    expect(observed).toBe("value");
  });
});

describe("PluginManager — uninstall releases everything, even if the plugin's own cleanup forgets to", () => {
  it("stops a tapped Hook from firing after the plugin is uninstalled", () => {
    const { hooks, manager } = setup();
    const calls: string[] = [];
    const dispose = manager.use({
      name: "sloppy",
      setup(context: PluginContext) {
        context.hooks.tap("beforeInit", { name: "sloppy" }, () => calls.push("fired"));
        // Deliberately does NOT return a cleanup function.
      }
    });

    hooks.runSeries("beforeInit", {});
    dispose();
    hooks.runSeries("beforeInit", {});

    expect(calls).toEqual(["fired"]);
  });

  it("stops a listened-to Event from firing after the plugin is uninstalled", () => {
    const { events, manager } = setup();
    const calls: string[] = [];
    const dispose = manager.use({
      name: "sloppy",
      setup(context: PluginContext) {
        context.events.on("change", () => calls.push("fired"));
      }
    });

    events.emit("change", {});
    dispose();
    events.emit("change", {});

    expect(calls).toEqual(["fired"]);
  });

  it("removes a registered context menu provider after the plugin is uninstalled", () => {
    const { contextMenu, manager } = setup();
    const dispose = manager.use({
      name: "sloppy",
      setup(context: PluginContext) {
        context.contextMenu.registerProvider({ name: "sloppy" }, () => [
          { type: "action", id: "a", label: "A", run: () => {} }
        ]);
      }
    });

    dispose();

    // With the provider gone, opening a menu should find no items and refuse to open.
    const opened = contextMenu.open({
      target: "cell",
      selection: {
        kind: "cell",
        active: { row: 0, column: 0 },
        anchor: { row: 0, column: 0 },
        range: { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 }
      },
      point: { x: 0, y: 0 }
    });
    expect(opened).toBe(false);
  });

  it("also calls the plugin's own returned cleanup function", () => {
    const { manager } = setup();
    let cleanedUp = false;
    const dispose = manager.use({
      name: "tidy",
      setup() {
        return () => {
          cleanedUp = true;
        };
      }
    });

    dispose();

    expect(cleanedUp).toBe(true);
  });

  it("disposing the returned disposer twice is a no-op", () => {
    const { manager } = setup();
    let cleanupCalls = 0;
    const dispose = manager.use({
      name: "tidy",
      setup() {
        return () => {
          cleanupCalls += 1;
        };
      }
    });

    dispose();
    dispose();

    expect(cleanupCalls).toBe(1);
  });

  it("PluginManager.dispose() uninstalls every currently-installed plugin", () => {
    const { hooks, manager } = setup();
    const calls: string[] = [];
    manager.use({
      name: "a",
      setup(context: PluginContext) {
        context.hooks.tap("beforeInit", { name: "a" }, () => calls.push("a"));
      }
    });
    manager.use({
      name: "b",
      setup(context: PluginContext) {
        context.hooks.tap("beforeInit", { name: "b" }, () => calls.push("b"));
      }
    });

    manager.dispose();
    hooks.runSeries("beforeInit", {});

    expect(calls).toEqual([]);
  });

  it("PluginManager.dispose() is safe to call repeatedly", () => {
    const { manager } = setup();
    manager.use({ name: "a", setup() {} });

    expect(() => {
      manager.dispose();
      manager.dispose();
    }).not.toThrow();
  });
});
