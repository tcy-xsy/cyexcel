import { describe, expect, it, vi } from "vitest";

import { EventBus } from "../../../../src/core/events";

describe("EventBus — on/emit", () => {
  it("calls listeners in registration order with the emitted payload", () => {
    const bus = new EventBus();
    const calls: string[] = [];
    bus.on<{ value: string }>("change", (payload) => calls.push(`a:${payload.value}`));
    bus.on<{ value: string }>("change", (payload) => calls.push(`b:${payload.value}`));

    bus.emit("change", { value: "x" });

    expect(calls).toEqual(["a:x", "b:x"]);
  });

  it("only calls listeners registered for the emitted event name", () => {
    const bus = new EventBus();
    const changeCalls: unknown[] = [];
    const destroyCalls: unknown[] = [];
    bus.on("change", (p) => changeCalls.push(p));
    bus.on("destroy", (p) => destroyCalls.push(p));

    bus.emit("change", 1);

    expect(changeCalls).toEqual([1]);
    expect(destroyCalls).toEqual([]);
  });

  it("emitting an event with no listeners is a no-op", () => {
    const bus = new EventBus();
    expect(() => bus.emit("nothing-listening", {})).not.toThrow();
  });
});

describe("EventBus — disposer idempotency", () => {
  it("stops the listener from being called once disposed, and disposing twice is a no-op", () => {
    const bus = new EventBus();
    const calls: unknown[] = [];
    const dispose = bus.on("change", (p) => calls.push(p));

    bus.emit("change", 1);
    dispose();
    dispose();
    bus.emit("change", 2);

    expect(calls).toEqual([1]);
  });

  it("dispose() clears every listener for every event", () => {
    const bus = new EventBus();
    const calls: unknown[] = [];
    bus.on("change", (p) => calls.push(p));
    bus.on("destroy", (p) => calls.push(p));

    bus.dispose();
    bus.emit("change", 1);
    bus.emit("destroy", 1);

    expect(calls).toEqual([]);
  });
});

describe("EventBus — listener error isolation", () => {
  it("isolates a throwing listener: later listeners still run and emit() does not throw", () => {
    const bus = new EventBus();
    const calls: string[] = [];
    bus.on("change", () => {
      calls.push("a");
      throw new Error("listener blew up");
    });
    bus.on("change", () => calls.push("b"));

    expect(() => bus.emit("change", {})).not.toThrow();
    expect(calls).toEqual(["a", "b"]);
  });

  it("forwards a throwing listener's error to the 'error' event", () => {
    const bus = new EventBus();
    const errors: unknown[] = [];
    bus.on("error", (payload) => errors.push(payload));
    bus.on("change", () => {
      throw new Error("listener blew up");
    });

    bus.emit("change", {});

    expect(errors).toHaveLength(1);
    expect((errors[0] as { eventName: string }).eventName).toBe("change");
  });

  it("does not recurse when an 'error' listener itself throws", () => {
    const bus = new EventBus();
    const consoleError = vi.spyOn(console, "error").mockImplementation(() => {});
    bus.on("error", () => {
      throw new Error("error handler is broken too");
    });
    bus.on("change", () => {
      throw new Error("original failure");
    });

    expect(() => bus.emit("change", {})).not.toThrow();
    expect(consoleError).toHaveBeenCalledTimes(1);
    consoleError.mockRestore();
  });
});
