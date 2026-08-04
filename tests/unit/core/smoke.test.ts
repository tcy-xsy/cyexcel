import { describe, expect, it } from "vitest";

import { createCyExcelCore, CyExcelCore } from "../../../src/core";

describe("CyExcelCore", () => {
  it("creates an empty core with immutable options", () => {
    const core = createCyExcelCore({ id: "orders" });

    expect(core).toBeInstanceOf(CyExcelCore);
    expect(core.options).toEqual({ id: "orders" });
    expect(Object.isFrozen(core.options)).toBe(true);
    expect(core.disposed).toBe(false);
  });

  it("can be disposed repeatedly", () => {
    const core = createCyExcelCore();

    core.dispose();
    core.dispose();

    expect(core.disposed).toBe(true);
  });
});
