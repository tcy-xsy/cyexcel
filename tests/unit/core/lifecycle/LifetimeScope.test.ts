import { describe, expect, it } from "vitest";

import { LifetimeScope } from "../../../../src/core/lifecycle";

describe("LifetimeScope", () => {
  it("cleans up in reverse registration order and isolates disposer errors", () => {
    const errors: unknown[] = [];
    const calls: string[] = [];
    const scope = new LifetimeScope({ onDisposeError: error => errors.push(error) });
    scope.defer(() => calls.push("first"));
    scope.defer(() => {
      calls.push("second");
      throw new Error("broken cleanup");
    });
    scope.defer(() => calls.push("third"));

    scope.dispose();

    expect(calls).toEqual(["third", "second", "first"]);
    expect(errors).toHaveLength(1);
    expect(scope.state).toBe("disposed");
    expect(scope.signal.aborted).toBe(true);
  });

  it("disposes child scopes through their parent", () => {
    const parent = new LifetimeScope();
    const child = parent.fork();
    const calls: string[] = [];
    child.defer(() => calls.push("child"));

    parent.dispose();

    expect(child.state).toBe("disposed");
    expect(calls).toEqual(["child"]);
  });
});
