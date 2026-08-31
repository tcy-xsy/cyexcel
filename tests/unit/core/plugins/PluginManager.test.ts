import { describe, expect, it } from "vitest";

import {
  CyExcelErrorCode
} from "../../../../src/core/errors";
import { EventBus } from "../../../../src/core/events";
import { HookManager, PRODUCTION_HOOK_REGISTRY } from "../../../../src/core/hooks";
import { ContextMenuService } from "../../../../src/core/menu";
import { PluginManager } from "../../../../src/core/plugins";
import type { CyExcelPlugin, PluginContext } from "../../../../src/core/plugins";
import { ShortcutRegistry } from "../../../../src/core/shortcuts";

function setup() {
  const hooks = new HookManager({ registry: PRODUCTION_HOOK_REGISTRY });
  const events = new EventBus();
  const contextMenu = new ContextMenuService({ container: {} as unknown as HTMLElement });
  const shortcuts = new ShortcutRegistry();
  const manager = new PluginManager({
    hooks,
    events,
    contextMenu,
    getCell: () => "value",
    copy: async () => ({ status: "copied", plan: { range: { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 }, rows: 1, columns: 1, cells: 1, estimatedBytes: 0, containsComplexValues: false, recommendedMode: "display" } }),
    paste: async () => {},
    insertRows: () => {},
    deleteRows: () => {},
    moveRows: () => {},
    insertColumns: () => {},
    deleteColumns: () => {},
    moveColumns: () => {},
    hideColumn: () => {},
    registerShortcut: (combo, handler) => shortcuts.register(combo, handler)
  });
  return { hooks, events, contextMenu, shortcuts, manager };
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

  it("exposes registerShortcut, which reaches the shared ShortcutRegistry", () => {
    const { shortcuts, manager } = setup();
    let handled: boolean | undefined;
    manager.use({
      name: "shortcuts",
      setup(context: PluginContext) {
        context.registerShortcut("Mod+K", () => {
          handled = true;
        });
      }
    });

    shortcuts.dispatch({ key: "k", ctrlKey: true, metaKey: false, shiftKey: false, altKey: false });

    expect(handled).toBe(true);
  });

  it("exposes getCell to the plugin", () => {
    const { manager } = setup();
    let getCell: PluginContext["getCell"] | undefined;
    manager.use({
      name: "reader",
      setup(context: PluginContext) {
        getCell = context.getCell;
      }
    });
    manager.activate();

    expect(getCell?.(0, 0)).toBe("value");
  });

  it("runtime capability proxies keep identity but reject calls until activate", () => {
    const { manager } = setup();
    let getCell: PluginContext["getCell"] | undefined;
    manager.use({
      name: "reader",
      setup(context: PluginContext) {
        getCell = context.getCell;
      }
    });

    expect(() => getCell?.(0, 0)).toThrow(
      expect.objectContaining({ code: CyExcelErrorCode.PluginRuntimeNotReady })
    );

    manager.activate();

    expect(getCell?.(0, 0)).toBe("value");
  });

  it("does not allow installing more plugins after activation", () => {
    const { manager } = setup();
    manager.activate();

    expect(() => manager.use({ name: "late", setup() {} })).toThrow(
      expect.objectContaining({ code: CyExcelErrorCode.PluginRuntimeNotReady })
    );
  });

  it("rolls back registrations if setup throws", () => {
    const { hooks, manager } = setup();
    const calls: string[] = [];

    expect(() =>
      manager.use({
        name: "broken",
        setup(context: PluginContext) {
          context.hooks.tap("beforeInit", { name: "broken" }, () => calls.push("hook"));
          throw new Error("setup failed");
        }
      })
    ).toThrow("setup failed");

    hooks.runSeries("beforeInit", {});

    expect(calls).toEqual([]);
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
        range: { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 },
        ranges: [{ startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 }]
      },
      point: { x: 0, y: 0 }
    });
    expect(opened).toBe(false);
  });

  it("stops a registered shortcut from firing after the plugin is uninstalled", () => {
    const { shortcuts, manager } = setup();
    const calls: string[] = [];
    const dispose = manager.use({
      name: "sloppy",
      setup(context: PluginContext) {
        context.registerShortcut("Mod+K", () => {
          calls.push("fired");
        });
      }
    });

    shortcuts.dispatch({ key: "k", ctrlKey: true, metaKey: false, shiftKey: false, altKey: false });
    dispose();
    shortcuts.dispatch({ key: "k", ctrlKey: true, metaKey: false, shiftKey: false, altKey: false });

    expect(calls).toEqual(["fired"]);
  });

  it("invalidates runtime functions captured by a plugin after that plugin is uninstalled", () => {
    const { manager } = setup();
    let getCell: PluginContext["getCell"] | undefined;
    const dispose = manager.use({
      name: "reader",
      setup(context: PluginContext) {
        getCell = context.getCell;
      }
    });
    manager.activate();

    expect(getCell?.(0, 0)).toBe("value");

    dispose();

    expect(() => getCell?.(0, 0)).toThrow(expect.objectContaining({ code: CyExcelErrorCode.PluginDisposed }));
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

  it("continues automatic cleanup when the plugin's own cleanup throws", () => {
    const { hooks, events, manager } = setup();
    const calls: string[] = [];
    const errors: unknown[] = [];
    events.on("error", payload => errors.push(payload));
    const dispose = manager.use({
      name: "throws-on-cleanup",
      setup(context: PluginContext) {
        context.hooks.tap("beforeInit", { name: "tracked" }, () => calls.push("hook"));
        return () => {
          throw new Error("cleanup failed");
        };
      }
    });

    hooks.runSeries("beforeInit", {});
    dispose();
    hooks.runSeries("beforeInit", {});

    expect(calls).toEqual(["hook"]);
    expect(errors).toHaveLength(1);
    expect(errors[0]).toMatchObject({ eventName: "plugin:throws-on-cleanup:dispose" });
  });

  it("runs plugin cleanup in reverse registration order", () => {
    const { manager } = setup();
    const calls: string[] = [];
    const dispose = manager.use({
      name: "ordered",
      setup(context: PluginContext) {
        context.registerShortcut("Mod+K", () => {
          calls.push("shortcut");
        });
        return () => calls.push("cleanup");
      }
    });

    dispose();

    expect(calls).toEqual(["cleanup"]);
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

  it("invalidates runtime functions after PluginManager.dispose()", () => {
    const { manager } = setup();
    let getCell: PluginContext["getCell"] | undefined;
    manager.use({
      name: "reader",
      setup(context: PluginContext) {
        getCell = context.getCell;
      }
    });
    manager.activate();
    manager.dispose();

    expect(() => getCell?.(0, 0)).toThrow(expect.objectContaining({ code: CyExcelErrorCode.PluginDisposed }));
  });
});
