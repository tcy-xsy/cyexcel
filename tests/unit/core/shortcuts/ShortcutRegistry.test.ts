import { describe, expect, it, vi } from "vitest";

import { CyExcelErrorCode } from "../../../../src/core/errors";
import { ShortcutRegistry, parseShortcut } from "../../../../src/core/shortcuts";
import type { ShortcutMatchInput } from "../../../../src/core/shortcuts";

function input(overrides: Partial<ShortcutMatchInput> = {}): ShortcutMatchInput {
  return { key: "z", ctrlKey: false, metaKey: false, shiftKey: false, altKey: false, ...overrides };
}

describe("parseShortcut", () => {
  it("parses a bare key with no modifiers", () => {
    expect(parseShortcut("Enter")).toEqual({ mod: false, shift: false, alt: false, key: "enter" });
  });

  it("parses Mod+<key>, case-insensitively", () => {
    expect(parseShortcut("Mod+Z")).toEqual({ mod: true, shift: false, alt: false, key: "z" });
    expect(parseShortcut("mod+z")).toEqual({ mod: true, shift: false, alt: false, key: "z" });
  });

  it("parses multiple modifiers regardless of order", () => {
    expect(parseShortcut("Mod+Shift+Z")).toEqual({ mod: true, shift: true, alt: false, key: "z" });
    expect(parseShortcut("Shift+Mod+Z")).toEqual({ mod: true, shift: true, alt: false, key: "z" });
  });

  it("parses Alt", () => {
    expect(parseShortcut("Mod+Alt+Delete")).toEqual({ mod: true, shift: false, alt: true, key: "delete" });
  });

  it("throws CyExcelError(InvalidOptions) for an empty combo", () => {
    expect(() => parseShortcut("")).toThrow(/invalid-options|combo/i);
    try {
      parseShortcut("");
    } catch (error) {
      expect(error).toMatchObject({ code: CyExcelErrorCode.InvalidOptions });
    }
  });

  it("throws for a combo with no key (only modifiers)", () => {
    expect(() => parseShortcut("Mod+Shift+")).toThrow();
  });

  it("throws for an unknown modifier token", () => {
    expect(() => parseShortcut("Ctrl+Z")).toThrow();
  });
});

describe("ShortcutRegistry matching", () => {
  it("dispatches to a handler whose combo exactly matches", () => {
    const registry = new ShortcutRegistry();
    const handler = vi.fn();
    registry.register("Mod+Z", handler);

    const handled = registry.dispatch(input({ key: "z", ctrlKey: true }));

    expect(handled).toBe(true);
    expect(handler).toHaveBeenCalledTimes(1);
  });

  it("does not dispatch when an unlisted modifier is held (exact match, not a subset match)", () => {
    const registry = new ShortcutRegistry();
    const handler = vi.fn();
    registry.register("Mod+Z", handler);

    // Alt wasn't part of the registered combo, so Mod+Alt+Z must not match it.
    const handled = registry.dispatch(input({ key: "z", ctrlKey: true, altKey: true }));

    expect(handled).toBe(false);
    expect(handler).not.toHaveBeenCalled();
  });

  it("treats ctrlKey and metaKey as interchangeable for Mod", () => {
    const registry = new ShortcutRegistry();
    const handler = vi.fn();
    registry.register("Mod+Z", handler);

    expect(registry.dispatch(input({ key: "z", metaKey: true }))).toBe(true);
    expect(handler).toHaveBeenCalledTimes(1);
  });

  it("matches a bare key combo only when no modifiers are held", () => {
    const registry = new ShortcutRegistry();
    const handler = vi.fn();
    registry.register("Enter", handler);

    expect(registry.dispatch(input({ key: "Enter", shiftKey: true }))).toBe(false);
    expect(registry.dispatch(input({ key: "Enter" }))).toBe(true);
    expect(handler).toHaveBeenCalledTimes(1);
  });

  it("returns false from dispatch when no registered combo matches", () => {
    const registry = new ShortcutRegistry();
    registry.register("Mod+Z", vi.fn());

    expect(registry.dispatch(input({ key: "y", ctrlKey: true }))).toBe(false);
  });

  it("stops matching after the returned disposer is called", () => {
    const registry = new ShortcutRegistry();
    const handler = vi.fn();
    const dispose = registry.register("Mod+Z", handler);

    dispose();
    const handled = registry.dispatch(input({ key: "z", ctrlKey: true }));

    expect(handled).toBe(false);
    expect(handler).not.toHaveBeenCalled();
  });

  it("calling a disposer twice is a no-op", () => {
    const registry = new ShortcutRegistry();
    const handler = vi.fn();
    const dispose = registry.register("Mod+Z", handler);
    dispose();

    expect(() => dispose()).not.toThrow();
    expect(registry.dispatch(input({ key: "z", ctrlKey: true }))).toBe(false);
  });
});

describe("ShortcutRegistry conflict policy: last-registered wins, `return false` falls through", () => {
  it("the most recently registered handler for the same combo runs first", () => {
    const registry = new ShortcutRegistry();
    const older = vi.fn();
    const newer = vi.fn();
    registry.register("Mod+Z", older);
    registry.register("Mod+Z", newer);

    registry.dispatch(input({ key: "z", ctrlKey: true }));

    expect(newer).toHaveBeenCalledTimes(1);
    expect(older).not.toHaveBeenCalled();
  });

  it("a handler returning false falls through to the next older handler for the same combo", () => {
    const registry = new ShortcutRegistry();
    const older = vi.fn();
    const newer = vi.fn(() => false as const);
    registry.register("Mod+Z", older);
    registry.register("Mod+Z", newer);

    const handled = registry.dispatch(input({ key: "z", ctrlKey: true }));

    expect(newer).toHaveBeenCalledTimes(1);
    expect(older).toHaveBeenCalledTimes(1);
    expect(handled).toBe(true);
  });

  it("dispatch returns false when every handler for a combo falls through", () => {
    const registry = new ShortcutRegistry();
    registry.register("Mod+Z", () => false);
    registry.register("Mod+Z", () => false);

    expect(registry.dispatch(input({ key: "z", ctrlKey: true }))).toBe(false);
  });

  it("disposing the newer handler restores the older one as the active match", () => {
    const registry = new ShortcutRegistry();
    const older = vi.fn();
    const newer = vi.fn();
    registry.register("Mod+Z", older);
    const disposeNewer = registry.register("Mod+Z", newer);

    disposeNewer();
    registry.dispatch(input({ key: "z", ctrlKey: true }));

    expect(older).toHaveBeenCalledTimes(1);
    expect(newer).not.toHaveBeenCalled();
  });
});

describe("ShortcutRegistry.dispose()", () => {
  it("clears every registered handler; dispatch afterward always returns false", () => {
    const registry = new ShortcutRegistry();
    const handler = vi.fn();
    registry.register("Mod+Z", handler);
    registry.register("Enter", vi.fn());

    registry.dispose();

    expect(registry.dispatch(input({ key: "z", ctrlKey: true }))).toBe(false);
    expect(registry.dispatch(input({ key: "Enter" }))).toBe(false);
    expect(handler).not.toHaveBeenCalled();
  });

  it("is safe to call twice", () => {
    const registry = new ShortcutRegistry();
    registry.register("Mod+Z", vi.fn());
    registry.dispose();

    expect(() => registry.dispose()).not.toThrow();
  });
});
