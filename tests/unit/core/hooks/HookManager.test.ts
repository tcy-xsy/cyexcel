import { describe, expect, it, vi } from "vitest";

import { CyExcelError, HookManager, createCancellableContext } from "../../../../src/core";
import type { HookRegistry } from "../../../../src/core";

function managerWith(registry: HookRegistry): HookManager {
  return new HookManager({ registry });
}

describe("HookManager — priority ordering", () => {
  it("runs same-priority taps in registration order", () => {
    const manager = managerWith({ demo: { mode: "series", async: false, isolate: false } });
    const calls: string[] = [];
    manager.tap("demo", { name: "a" }, () => calls.push("a"));
    manager.tap("demo", { name: "b" }, () => calls.push("b"));
    manager.tap("demo", { name: "c" }, () => calls.push("c"));

    manager.runSeries("demo", {});

    expect(calls).toEqual(["a", "b", "c"]);
  });

  it("runs higher-priority taps first, tie-breaking by registration order", () => {
    const manager = managerWith({ demo: { mode: "series", async: false, isolate: false } });
    const calls: string[] = [];
    manager.tap("demo", { name: "low", priority: 0 }, () => calls.push("low"));
    manager.tap("demo", { name: "high-1", priority: 100 }, () => calls.push("high-1"));
    manager.tap("demo", { name: "high-2", priority: 100 }, () => calls.push("high-2"));
    manager.tap("demo", { name: "mid", priority: 50 }, () => calls.push("mid"));

    manager.runSeries("demo", {});

    expect(calls).toEqual(["high-1", "high-2", "mid", "low"]);
  });
});

describe("HookManager — series mode", () => {
  it("runs every tap regardless of what earlier taps did", () => {
    const manager = managerWith({ demo: { mode: "series", async: false, isolate: false } });
    const calls: string[] = [];
    manager.tap("demo", { name: "a" }, () => calls.push("a"));
    manager.tap("demo", { name: "b" }, () => calls.push("b"));

    manager.runSeries("demo", {});

    expect(calls).toEqual(["a", "b"]);
  });

  it("propagates a non-isolated tap's exception and stops subsequent taps", () => {
    const manager = managerWith({ demo: { mode: "series", async: false, isolate: false } });
    const calls: string[] = [];
    manager.tap("demo", { name: "a" }, () => {
      calls.push("a");
      throw new Error("boom");
    });
    manager.tap("demo", { name: "b" }, () => calls.push("b"));

    expect(() => manager.runSeries("demo", {})).toThrow("boom");
    expect(calls).toEqual(["a"]);
  });
});

describe("HookManager — bail mode", () => {
  it("stops as soon as a tap cancels and does not run later taps", () => {
    const registry: HookRegistry = { demo: { mode: "bail", async: false, isolate: false } };
    const manager = managerWith(registry);
    const calls: string[] = [];
    manager.tap("demo", { name: "a" }, (ctx: ReturnType<typeof createCancellableContext<object>>) => {
      calls.push("a");
      ctx.cancel("nope");
    });
    manager.tap("demo", { name: "b" }, () => calls.push("b"));

    const context = createCancellableContext({});
    manager.runBail("demo", context);

    expect(calls).toEqual(["a"]);
    expect(context.cancelled).toBe(true);
    expect(context.reason).toBe("nope");
  });

  it("runs every tap when none of them cancel", () => {
    const registry: HookRegistry = { demo: { mode: "bail", async: false, isolate: false } };
    const manager = managerWith(registry);
    const calls: string[] = [];
    manager.tap("demo", { name: "a" }, () => calls.push("a"));
    manager.tap("demo", { name: "b" }, () => calls.push("b"));

    const context = createCancellableContext({});
    manager.runBail("demo", context);

    expect(calls).toEqual(["a", "b"]);
    expect(context.cancelled).toBe(false);
  });
});

describe("HookManager — waterfall mode", () => {
  it("chains each tap's return value into the next", () => {
    const manager = managerWith({ demo: { mode: "waterfall", async: false, isolate: false } });
    manager.tap<unknown, number>("demo", { name: "double" }, (value: number) => value * 2);
    manager.tap<unknown, number>("demo", { name: "plusOne" }, (value: number) => value + 1);

    const result = manager.runWaterfall("demo", 5, {});

    expect(result).toBe(11); // (5 * 2) + 1
  });

  it("returns the initial value unchanged when there are no taps", () => {
    const manager = managerWith({ demo: { mode: "waterfall", async: false, isolate: false } });

    expect(manager.runWaterfall("demo", "unchanged", {})).toBe("unchanged");
  });
});

describe("HookManager — parallel mode", () => {
  it("runs every tap against the same context, independent of each other", () => {
    const manager = managerWith({ demo: { mode: "parallel", async: false, isolate: false } });
    const seen: number[] = [];
    manager.tap("demo", { name: "a" }, (ctx: { value: number }) => seen.push(ctx.value));
    manager.tap("demo", { name: "b" }, (ctx: { value: number }) => seen.push(ctx.value * 10));

    manager.runParallel("demo", { value: 3 });

    expect(seen).toEqual([3, 30]);
  });
});

describe("HookManager — isolate: true (display/transform hooks)", () => {
  it("catches a throwing tap, reports it via onHookError, and keeps running later taps", () => {
    const registry: HookRegistry = {
      demo: { mode: "series", async: false, isolate: true },
      onHookError: { mode: "series", async: false, isolate: true }
    };
    const manager = managerWith(registry);
    const reported: unknown[] = [];
    manager.tap("onHookError", { name: "logger" }, (ctx: { error: unknown; hookName: string }) =>
      reported.push(ctx)
    );

    const calls: string[] = [];
    manager.tap("demo", { name: "a" }, () => {
      calls.push("a");
      throw new Error("display glitch");
    });
    manager.tap("demo", { name: "b" }, () => calls.push("b"));

    expect(() => manager.runSeries("demo", {})).not.toThrow();
    expect(calls).toEqual(["a", "b"]);
    expect(reported).toHaveLength(1);
    expect((reported[0] as { hookName: string }).hookName).toBe("demo");
  });

  it("leaves the waterfall value unchanged when a tap throws, instead of aborting the chain", () => {
    const manager = managerWith({ demo: { mode: "waterfall", async: false, isolate: true } });
    manager.tap<unknown, number>("demo", { name: "broken" }, () => {
      throw new Error("nope");
    });
    manager.tap<unknown, number>("demo", { name: "plusOne" }, (value: number) => value + 1);

    const result = manager.runWaterfall("demo", 5, {});

    expect(result).toBe(6); // broken tap skipped, only +1 applied
  });

  it("does not recurse when an onHookError handler itself throws", () => {
    const registry: HookRegistry = {
      demo: { mode: "series", async: false, isolate: true },
      onHookError: { mode: "series", async: false, isolate: true }
    };
    const manager = managerWith(registry);
    const consoleError = vi.spyOn(console, "error").mockImplementation(() => {});
    manager.tap("onHookError", { name: "broken-logger" }, () => {
      throw new Error("logger itself is broken");
    });
    manager.tap("demo", { name: "a" }, () => {
      throw new Error("original failure");
    });

    expect(() => manager.runSeries("demo", {})).not.toThrow();
    expect(consoleError).toHaveBeenCalledTimes(1);
    consoleError.mockRestore();
  });
});

describe("HookManager — tapAsync", () => {
  it("throws immediately when registering an async tap on a hook that does not support async", () => {
    const manager = managerWith({ demo: { mode: "bail", async: false, isolate: false } });

    expect(() => manager.tapAsync("demo", { name: "a" }, async () => {})).toThrow(CyExcelError);
  });

  it("allows tapAsync on a hook whose registry entry has async: true", () => {
    const manager = managerWith({ demo: { mode: "bail", async: true, isolate: false } });

    expect(() => manager.tapAsync("demo", { name: "a" }, async () => {})).not.toThrow();
  });

  it("throws for an unknown hook name", () => {
    const manager = managerWith({});

    expect(() => manager.tap("does-not-exist", { name: "a" }, () => {})).toThrow(CyExcelError);
  });
});

describe("HookManager — async run methods", () => {
  it("runSeriesAsync awaits both sync and async taps in order", async () => {
    const manager = managerWith({ demo: { mode: "series", async: true, isolate: false } });
    const calls: string[] = [];
    manager.tap("demo", { name: "sync" }, () => calls.push("sync"));
    manager.tapAsync("demo", { name: "async" }, async () => {
      await Promise.resolve();
      calls.push("async");
    });

    await manager.runSeriesAsync("demo", {}, new AbortController().signal);

    expect(calls).toEqual(["sync", "async"]);
  });

  it("runBailAsync stops once a tap cancels", async () => {
    const manager = managerWith({ demo: { mode: "bail", async: true, isolate: false } });
    const calls: string[] = [];
    manager.tapAsync("demo", { name: "a" }, async (ctx: ReturnType<typeof createCancellableContext<object>>) => {
      calls.push("a");
      ctx.cancel();
    });
    manager.tap("demo", { name: "b" }, () => calls.push("b"));

    const context = createCancellableContext({});
    await manager.runBailAsync("demo", context, new AbortController().signal);

    expect(calls).toEqual(["a"]);
  });

  it("runWaterfallAsync chains resolved values and stops once the signal aborts", async () => {
    const manager = managerWith({ demo: { mode: "waterfall", async: true, isolate: false } });
    manager.tapAsync<unknown, number>("demo", { name: "double" }, async (value: number) => value * 2);
    manager.tap<unknown, number>("demo", { name: "plusOne" }, (value: number) => value + 1);

    const result = await manager.runWaterfallAsync("demo", 5, {}, new AbortController().signal);

    expect(result).toBe(11);
  });
});

describe("HookManager — disposer idempotency", () => {
  it("stops the tap from running once disposed, and disposing twice is a no-op", () => {
    const manager = managerWith({ demo: { mode: "series", async: false, isolate: false } });
    const calls: string[] = [];
    const dispose = manager.tap("demo", { name: "a" }, () => calls.push("a"));

    manager.runSeries("demo", {});
    dispose();
    dispose();
    manager.runSeries("demo", {});

    expect(calls).toEqual(["a"]);
  });

  it("dispose() clears every tap for every hook", () => {
    const manager = managerWith({ demo: { mode: "series", async: false, isolate: false } });
    const calls: string[] = [];
    manager.tap("demo", { name: "a" }, () => calls.push("a"));

    manager.dispose();
    manager.runSeries("demo", {});

    expect(calls).toEqual([]);
  });
});
