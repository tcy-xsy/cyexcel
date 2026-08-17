import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { CyExcelErrorCode } from "../../../src/core/errors";

/**
 * Keeps `doc/api-design.md`'s error-code table honest.
 *
 * The table is a promise to consumers about which failures they can branch on,
 * and a promise that only a test can keep: when this gate was added, 12 of the
 * 20 codes had never been written down, including the two that tell a caller
 * their DataSource is missing a capability. Documentation drift is invisible
 * until someone needs the missing row.
 */
const DOC_PATH = fileURLToPath(new URL("../../../doc/api-design.md", import.meta.url));

function documentedCodes(): Set<string> {
  const table = readFileSync(DOC_PATH, "utf8");
  return new Set([...table.matchAll(/^\| `([a-z0-9-]+)` \|/gm)].map((match) => match[1] as string));
}

describe("doc/api-design.md error-code table", () => {
  it("documents every code the runtime can throw", () => {
    const documented = documentedCodes();
    const undocumented = Object.values(CyExcelErrorCode).filter((code) => !documented.has(code));
    expect(undocumented).toEqual([]);
  });

  it("does not document codes the runtime cannot throw", () => {
    const runtime = new Set<string>(Object.values(CyExcelErrorCode));
    const stale = [...documentedCodes()].filter((code) => !runtime.has(code));
    expect(stale).toEqual([]);
  });
});
