import { describe, expect, it, vi } from "vitest";

import { StructureTransition } from "../../../../src/core/integration";

function createTransition() {
  const calls: string[] = [];
  const options = {
    cancelActiveEdit: vi.fn(() => calls.push("edit")),
    cancelBodyGesture: vi.fn(() => calls.push("body")),
    cancelHeaderGesture: vi.fn(() => calls.push("header")),
    closeContextMenu: vi.fn(() => calls.push("menu"))
  };
  return { transition: new StructureTransition(options), options, calls };
}

describe("StructureTransition", () => {
  it("starts at version 0 and touches nothing until applied", () => {
    const { transition, options } = createTransition();
    expect(transition.version).toBe(0);
    expect(options.cancelActiveEdit).not.toHaveBeenCalled();
    expect(options.cancelBodyGesture).not.toHaveBeenCalled();
  });

  it("advances the version by exactly one per transition", () => {
    const { transition } = createTransition();
    transition.apply();
    expect(transition.version).toBe(1);
    transition.apply();
    expect(transition.version).toBe(2);
  });

  it("tears down every consumer anchored to the old version", () => {
    const { transition, options } = createTransition();
    transition.apply();
    expect(options.cancelActiveEdit).toHaveBeenCalledTimes(1);
    expect(options.cancelBodyGesture).toHaveBeenCalledTimes(1);
    expect(options.cancelHeaderGesture).toHaveBeenCalledTimes(1);
    expect(options.closeContextMenu).toHaveBeenCalledTimes(1);
  });

  it("cancels the active edit before publishing the new version", () => {
    // A stale numeric anchor must not be able to commit into the newly
    // published row/column order, so the edit dies while the old version is
    // still current.
    let versionSeenByEditor: number | undefined;
    const transition: StructureTransition = new StructureTransition({
      cancelActiveEdit: () => {
        versionSeenByEditor = transition.version;
      },
      cancelBodyGesture: () => undefined,
      cancelHeaderGesture: () => undefined,
      closeContextMenu: () => undefined
    });

    transition.apply();

    expect(versionSeenByEditor).toBe(0);
    expect(transition.version).toBe(1);
  });

  it("bumping and tearing down cannot be separated", () => {
    // The whole reason this class exists: there is no way to advance the
    // version without the teardown, or to tear down without advancing.
    const { transition, calls } = createTransition();
    transition.apply();
    expect(calls).toEqual(["edit", "body", "header", "menu"]);
    expect(transition.version).toBe(1);
  });

  // The safe-integer guard in apply() is deliberately left untested: reaching
  // it needs 2^53 structural commits, and the only way to simulate that would
  // be a constructor seam that exists purely for the test. Noted rather than
  // faked with an assertion that does not exercise the guard.
});
