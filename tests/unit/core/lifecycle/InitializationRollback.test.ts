import { describe, expect, it, vi } from "vitest";

import { InitializationRollback } from "../../../../src/core/lifecycle";

describe("InitializationRollback", () => {
  it("rolls resources back in reverse acquisition order", () => {
    const calls: string[] = [];
    const rollback = new InitializationRollback();
    rollback.defer(() => calls.push("first"));
    rollback.defer(() => calls.push("second"));

    rollback.rollback();
    rollback.rollback();

    expect(calls).toEqual(["second", "first"]);
  });

  it("isolates cleanup failures and continues rolling back", () => {
    const first = vi.fn();
    const rollback = new InitializationRollback();
    rollback.defer(first);
    rollback.defer(() => {
      throw new Error("cleanup failed");
    });

    expect(() => rollback.rollback()).not.toThrow();
    expect(first).toHaveBeenCalledOnce();
  });

  it("disarms cleanup after a successful commit", () => {
    const dispose = vi.fn();
    const rollback = new InitializationRollback();
    rollback.defer(dispose);

    rollback.commit();
    rollback.rollback();

    expect(dispose).not.toHaveBeenCalled();
    expect(() => rollback.defer(dispose)).toThrow(
      "Initialization rollback is no longer active"
    );
  });
});
