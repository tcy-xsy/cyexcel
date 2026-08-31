import { describe, expect, it, vi } from "vitest";

import { HookManager, PRODUCTION_HOOK_REGISTRY } from "../../../../src/core/hooks";
import { ContextMenuService } from "../../../../src/core/menu";
import type {
  ContextMenuContext,
  ContextMenuItem,
  ContextMenuRenderContext,
  ContextMenuRenderer,
  ContextMenuRendererFactory
} from "../../../../src/core/menu";

function fakeContext(overrides: Partial<ContextMenuContext> = {}): ContextMenuContext {
  return {
    target: "cell",
    row: 0,
    column: 0,
    selection: {
      kind: "cell",
      active: { row: 0, column: 0 },
      anchor: { row: 0, column: 0 },
      range: { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 },
      ranges: [{ startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 }]
    },
    point: { x: 10, y: 10 },
    ...overrides
  };
}

interface FakeRenderer extends ContextMenuRenderer {
  readonly mountCalls: ContextMenuRenderContext[];
  readonly updateCalls: ContextMenuRenderContext[];
  destroyCalls: number;
  focusCalls: number;
}

function createFakeRendererFactory(): { factory: ContextMenuRendererFactory; renderers: FakeRenderer[] } {
  const renderers: FakeRenderer[] = [];
  const factory: ContextMenuRendererFactory = {
    create() {
      const renderer: FakeRenderer = {
        mountCalls: [],
        updateCalls: [],
        destroyCalls: 0,
        focusCalls: 0,
        mount(_container, context) {
          renderer.mountCalls.push(context);
        },
        update(context) {
          renderer.updateCalls.push(context);
        },
        focus() {
          renderer.focusCalls += 1;
        },
        destroy() {
          renderer.destroyCalls += 1;
        }
      };
      renderers.push(renderer);
      return renderer;
    }
  };
  return { factory, renderers };
}

function fakeContainer(): HTMLElement {
  return {} as unknown as HTMLElement;
}

describe("ContextMenuService — provider merging", () => {
  it("merges items from every applicable provider, in priority then registration order", () => {
    const { factory, renderers } = createFakeRendererFactory();
    const service = new ContextMenuService({ container: fakeContainer(), rendererFactory: factory });
    service.registerProvider({ name: "low", priority: 0 }, () => [
      { type: "action", id: "low-item", label: "Low", run: () => {} }
    ]);
    service.registerProvider({ name: "high", priority: 100 }, () => [
      { type: "action", id: "high-item", label: "High", run: () => {} }
    ]);

    service.open(fakeContext());

    const items = renderers[0]?.mountCalls[0]?.items ?? [];
    expect(items.map((i) => i.id)).toEqual(["high-item", "low-item"]);
  });

  it("only invokes a provider registered for a specific target when that target matches", () => {
    const { factory, renderers } = createFakeRendererFactory();
    const service = new ContextMenuService({ container: fakeContainer(), rendererFactory: factory });
    service.registerProvider({ name: "cell-only", target: "cell" }, () => [
      { type: "action", id: "cell-item", label: "Cell", run: () => {} }
    ]);
    service.registerProvider({ name: "row-header-only", target: "row-header" }, () => [
      { type: "action", id: "row-item", label: "Row", run: () => {} }
    ]);

    service.open(fakeContext({ target: "cell" }));

    const items = renderers[0]?.mountCalls[0]?.items ?? [];
    expect(items.map((i) => i.id)).toEqual(["cell-item"]);
  });

  it("groups items via a submenu returned by a provider", () => {
    const { factory, renderers } = createFakeRendererFactory();
    const service = new ContextMenuService({ container: fakeContainer(), rendererFactory: factory });
    service.registerProvider({ name: "grouped" }, () => [
      {
        type: "submenu",
        id: "group",
        label: "Group",
        children: [{ type: "action", id: "child", label: "Child", run: () => {} }]
      }
    ]);

    service.open(fakeContext());

    const items = renderers[0]?.mountCalls[0]?.items ?? [];
    expect(items).toHaveLength(1);
    expect(items[0]?.type).toBe("submenu");
  });

  it("does not open when no provider contributes any items", () => {
    const { factory, renderers } = createFakeRendererFactory();
    const service = new ContextMenuService({ container: fakeContainer(), rendererFactory: factory });

    const opened = service.open(fakeContext());

    expect(opened).toBe(false);
    expect(renderers).toHaveLength(0);
  });
});

describe("ContextMenuService — extendContextMenu Hook (delete/replace/disable)", () => {
  function serviceWithHooks() {
    const hooks = new HookManager({ registry: PRODUCTION_HOOK_REGISTRY });
    const { factory, renderers } = createFakeRendererFactory();
    const service = new ContextMenuService({ container: fakeContainer(), hooks, rendererFactory: factory });
    service.registerProvider({ name: "base" }, () => [
      { type: "action", id: "a", label: "A", run: () => {} },
      { type: "action", id: "b", label: "B", run: () => {} }
    ]);
    return { hooks, service, renderers };
  }

  it("deletes an item by id", () => {
    const { hooks, service, renderers } = serviceWithHooks();
    hooks.tap(
      "extendContextMenu",
      { name: "delete-b" },
      (items: readonly ContextMenuItem[]) => items.filter((i) => i.id !== "b")
    );

    service.open(fakeContext());

    expect(renderers[0]?.mountCalls[0]?.items.map((i) => i.id)).toEqual(["a"]);
  });

  it("replaces an item by id", () => {
    const { hooks, service, renderers } = serviceWithHooks();
    hooks.tap("extendContextMenu", { name: "replace-b" }, (items: readonly ContextMenuItem[]) =>
      items.map((i) => (i.id === "b" ? { type: "action" as const, id: "b", label: "Replaced", run: () => {} } : i))
    );

    service.open(fakeContext());

    const items = renderers[0]?.mountCalls[0]?.items ?? [];
    const replaced = items.find((i) => i.id === "b");
    expect(replaced && "label" in replaced ? replaced.label : undefined).toBe("Replaced");
  });

  it("disables an item by id", () => {
    const { hooks, service, renderers } = serviceWithHooks();
    hooks.tap("extendContextMenu", { name: "disable-b" }, (items: readonly ContextMenuItem[]) =>
      items.map((i) => (i.id === "b" ? { ...i, disabled: true } : i))
    );

    service.open(fakeContext());

    const items = renderers[0]?.mountCalls[0]?.items ?? [];
    const disabled = items.find((i) => i.id === "b");
    expect(disabled && "disabled" in disabled ? disabled.disabled : undefined).toBe(true);
  });
});

describe("ContextMenuService — renderer destroy lifecycle", () => {
  it("destroys the renderer on close()", () => {
    const { factory, renderers } = createFakeRendererFactory();
    const service = new ContextMenuService({ container: fakeContainer(), rendererFactory: factory });
    service.registerProvider({ name: "p" }, () => [{ type: "action", id: "a", label: "A", run: () => {} }]);

    service.open(fakeContext());
    service.close();

    expect(renderers[0]?.destroyCalls).toBe(1);
  });

  it("destroys the previous renderer once when open() replaces an already-open menu", () => {
    const { factory, renderers } = createFakeRendererFactory();
    const service = new ContextMenuService({ container: fakeContainer(), rendererFactory: factory });
    service.registerProvider({ name: "p" }, () => [{ type: "action", id: "a", label: "A", run: () => {} }]);

    service.open(fakeContext());
    service.open(fakeContext());

    expect(renderers).toHaveLength(2);
    expect(renderers[0]?.destroyCalls).toBe(1);
  });

  it("destroys the renderer on dispose()", () => {
    const { factory, renderers } = createFakeRendererFactory();
    const service = new ContextMenuService({ container: fakeContainer(), rendererFactory: factory });
    service.registerProvider({ name: "p" }, () => [{ type: "action", id: "a", label: "A", run: () => {} }]);

    service.open(fakeContext());
    service.dispose();

    expect(renderers[0]?.destroyCalls).toBe(1);
  });

  it("calls restoreFocus after the menu closes", () => {
    const { factory } = createFakeRendererFactory();
    const restoreFocus = vi.fn();
    const service = new ContextMenuService({ container: fakeContainer(), rendererFactory: factory, restoreFocus });
    service.registerProvider({ name: "p" }, () => [{ type: "action", id: "a", label: "A", run: () => {} }]);

    service.open(fakeContext());
    service.close();

    expect(restoreFocus).toHaveBeenCalledTimes(1);
  });
});

describe("ContextMenuService — async actions", () => {
  it("marks the action busy while its run() promise is pending, then closes the menu on success", async () => {
    const { factory, renderers } = createFakeRendererFactory();
    const service = new ContextMenuService({ container: fakeContainer(), rendererFactory: factory });
    let resolveRun: () => void = () => {};
    service.registerProvider({ name: "p" }, () => [
      {
        type: "action",
        id: "a",
        label: "A",
        run: () => new Promise<void>((resolve) => (resolveRun = resolve))
      }
    ]);

    service.open(fakeContext());
    renderers[0]?.mountCalls[0]?.actions.run("a");

    expect(renderers[0]?.updateCalls.at(-1)?.busyItemId).toBe("a");
    expect(service.isOpen()).toBe(true);

    resolveRun();
    await Promise.resolve();
    await Promise.resolve();

    expect(service.isOpen()).toBe(false);
  });

  it("does not commit anything once the signal has been aborted by closing the menu first", async () => {
    const { factory, renderers } = createFakeRendererFactory();
    const service = new ContextMenuService({ container: fakeContainer(), rendererFactory: factory });
    const captured: { signal: AbortSignal | null } = { signal: null };
    let resolveRun: () => void = () => {};
    service.registerProvider({ name: "p" }, () => [
      {
        type: "action",
        id: "a",
        label: "A",
        run: (_context: ContextMenuContext, signal: AbortSignal) => {
          captured.signal = signal;
          return new Promise<void>((resolve) => (resolveRun = resolve));
        }
      }
    ]);

    service.open(fakeContext());
    renderers[0]?.mountCalls[0]?.actions.run("a");
    const updateCallsBeforeClose = renderers[0]?.updateCalls.length ?? 0;

    service.close(); // user closes the menu while the action is still running

    expect(captured.signal?.aborted).toBe(true);

    resolveRun(); // the action eventually finishes, after the menu already closed
    await Promise.resolve();
    await Promise.resolve();

    // No further update()/renderer activity after the abort — nothing was committed post-close.
    expect(renderers[0]?.updateCalls.length).toBe(updateCallsBeforeClose);
  });

  it("clears the busy state and keeps the menu open when the action rejects", async () => {
    const { factory, renderers } = createFakeRendererFactory();
    const service = new ContextMenuService({ container: fakeContainer(), rendererFactory: factory });
    let rejectRun: (error: Error) => void = () => {};
    service.registerProvider({ name: "p" }, () => [
      {
        type: "action",
        id: "a",
        label: "A",
        run: () => new Promise<void>((_resolve, reject) => (rejectRun = reject))
      }
    ]);

    service.open(fakeContext());
    renderers[0]?.mountCalls[0]?.actions.run("a");
    rejectRun(new Error("action failed"));
    await Promise.resolve();
    await Promise.resolve();

    expect(service.isOpen()).toBe(true);
    expect(renderers[0]?.updateCalls.at(-1)?.busyItemId).toBeNull();
  });

  it("closes the menu immediately for a synchronous action", () => {
    const { factory, renderers } = createFakeRendererFactory();
    const service = new ContextMenuService({ container: fakeContainer(), rendererFactory: factory });
    const calls: string[] = [];
    service.registerProvider({ name: "p" }, () => [
      { type: "action", id: "a", label: "A", run: () => { calls.push("ran"); } }
    ]);

    service.open(fakeContext());
    renderers[0]?.mountCalls[0]?.actions.run("a");

    expect(calls).toEqual(["ran"]);
    expect(service.isOpen()).toBe(false);
  });

  it("catches a synchronous throw from run() instead of propagating it — closes the menu and reports to console.error", () => {
    const { factory, renderers } = createFakeRendererFactory();
    const service = new ContextMenuService({ container: fakeContainer(), rendererFactory: factory });
    const thrown = new Error("insertColumns blocks this");
    service.registerProvider({ name: "p" }, () => [
      {
        type: "action",
        id: "a",
        label: "A",
        run: () => {
          throw thrown;
        }
      }
    ]);
    const consoleError = vi.spyOn(console, "error").mockImplementation(() => {});

    service.open(fakeContext());
    expect(() => renderers[0]?.mountCalls[0]?.actions.run("a")).not.toThrow();

    expect(service.isOpen()).toBe(false);
    expect(consoleError).toHaveBeenCalledWith(expect.stringContaining("a"), thrown);

    consoleError.mockRestore();
  });

  it("forwards a synchronous throw to onActionError, in addition to console.error", () => {
    const { factory, renderers } = createFakeRendererFactory();
    const thrown = new Error("insertColumns blocks this");
    const onActionError = vi.fn();
    const service = new ContextMenuService({ container: fakeContainer(), rendererFactory: factory, onActionError });
    service.registerProvider({ name: "p" }, () => [
      {
        type: "action",
        id: "a",
        label: "A",
        run: () => {
          throw thrown;
        }
      }
    ]);
    const consoleError = vi.spyOn(console, "error").mockImplementation(() => {});

    service.open(fakeContext());
    renderers[0]?.mountCalls[0]?.actions.run("a");

    expect(onActionError).toHaveBeenCalledWith("a", thrown);

    consoleError.mockRestore();
  });

  it("forwards a rejected action's error to onActionError and clears the busy state", async () => {
    const { factory, renderers } = createFakeRendererFactory();
    const onActionError = vi.fn();
    const service = new ContextMenuService({ container: fakeContainer(), rendererFactory: factory, onActionError });
    const rejection = new Error("action failed");
    let rejectRun: (error: Error) => void = () => {};
    service.registerProvider({ name: "p" }, () => [
      {
        type: "action",
        id: "a",
        label: "A",
        run: () => new Promise<void>((_resolve, reject) => (rejectRun = reject))
      }
    ]);
    const consoleError = vi.spyOn(console, "error").mockImplementation(() => {});

    service.open(fakeContext());
    renderers[0]?.mountCalls[0]?.actions.run("a");
    rejectRun(rejection);
    await Promise.resolve();
    await Promise.resolve();

    expect(onActionError).toHaveBeenCalledWith("a", rejection);
    expect(service.isOpen()).toBe(true);
    expect(renderers[0]?.updateCalls.at(-1)?.busyItemId).toBeNull();

    consoleError.mockRestore();
  });

  it("ignores a disabled action", () => {
    const { factory, renderers } = createFakeRendererFactory();
    const service = new ContextMenuService({ container: fakeContainer(), rendererFactory: factory });
    const calls: string[] = [];
    service.registerProvider({ name: "p" }, () => [
      { type: "action", id: "a", label: "A", disabled: true, run: () => { calls.push("ran"); } }
    ]);

    service.open(fakeContext());
    renderers[0]?.mountCalls[0]?.actions.run("a");

    expect(calls).toEqual([]);
    expect(service.isOpen()).toBe(true);
  });
});

describe("ContextMenuService — beforeContextMenuOpen Hook", () => {
  it("does not open the menu when a tap cancels", () => {
    const hooks = new HookManager({ registry: PRODUCTION_HOOK_REGISTRY });
    hooks.tap("beforeContextMenuOpen", { name: "block" }, (ctx: { cancel(reason?: string): void }) => ctx.cancel());
    const { factory, renderers } = createFakeRendererFactory();
    const service = new ContextMenuService({ container: fakeContainer(), hooks, rendererFactory: factory });
    service.registerProvider({ name: "p" }, () => [{ type: "action", id: "a", label: "A", run: () => {} }]);

    const opened = service.open(fakeContext());

    expect(opened).toBe(false);
    expect(renderers).toHaveLength(0);
  });
});

describe("ContextMenuService — getBounds passthrough (menu clipping fix)", () => {
  it("includes the getBounds() result in the render context as `bounds`", () => {
    const { factory, renderers } = createFakeRendererFactory();
    const bounds = { left: 0, top: 0, right: 400, bottom: 300 } as DOMRect;
    const service = new ContextMenuService({
      container: fakeContainer(),
      rendererFactory: factory,
      getBounds: () => bounds
    });
    service.registerProvider({ name: "p" }, () => [{ type: "action", id: "a", label: "A", run: () => {} }]);

    service.open(fakeContext());

    expect(renderers[0]?.mountCalls[0]?.bounds).toBe(bounds);
  });

  it("bounds is null when getBounds is not provided — existing callers are unaffected", () => {
    const { factory, renderers } = createFakeRendererFactory();
    const service = new ContextMenuService({ container: fakeContainer(), rendererFactory: factory });
    service.registerProvider({ name: "p" }, () => [{ type: "action", id: "a", label: "A", run: () => {} }]);

    service.open(fakeContext());

    expect(renderers[0]?.mountCalls[0]?.bounds).toBeNull();
  });
});
