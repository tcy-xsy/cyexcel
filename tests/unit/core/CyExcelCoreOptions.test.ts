import { describe, expect, it, vi } from "vitest";

import {
  CyExcelCore,
  type CyExcelCoreOptions
} from "../../../src/core/CyExcelCore";
import { CyExcelError, CyExcelErrorCode } from "../../../src/core/errors";
import type { DataSource } from "../../../src/core/types";
import {
  validateCyExcelCoreDataDimensions,
  validateCyExcelCoreHeaderGeometry,
  validateCyExcelCoreOptions
} from "../../../src/core/validateCyExcelCoreOptions";

function expectInvalid(run: () => unknown, field: string): void {
  try {
    run();
    expect.unreachable("expected invalid options to throw");
  } catch (error) {
    expect(error).toBeInstanceOf(CyExcelError);
    expect((error as CyExcelError).code).toBe(CyExcelErrorCode.InvalidOptions);
    expect((error as Error).message).toContain(field);
  }
}

function dataSource(rowCount: number, columnCount: number): DataSource {
  return {
    getRowCount: () => rowCount,
    getColumnCount: () => columnCount,
    getCell: () => undefined,
    setCell: () => {}
  };
}

describe("validateCyExcelCoreOptions", () => {
  it("accepts the supported zero/positive geometry and policy values", () => {
    expect(() => validateCyExcelCoreOptions({
      columns: [{ id: "a", width: 0, minWidth: 0, maxWidth: Number.POSITIVE_INFINITY }],
      data: [{}],
      rowHeight: 24,
      rowHeights: new Map([[0, 36]]),
      headerHeight: 0,
      rowHeaderWidth: 0,
      freeze: { rows: 0, columns: 0 },
      overscanRows: 2,
      overscanColumns: 3,
      scrollCommitPolicy: "cancel",
      columnLetterFallback: false,
      rowReorder: true,
      showGridLines: false,
      clipboard: { maxCells: 0, maxBytes: 0, oversizedAction: "block" }
    })).not.toThrow();
  });

  it("rejects simultaneous data inputs instead of silently preferring dataSource", () => {
    expectInvalid(
      () => validateCyExcelCoreOptions({ data: [], dataSource: dataSource(0, 0) }),
      "data/dataSource"
    );
  });

  it.each<[string, CyExcelCoreOptions]>([
    ["rowHeight", { rowHeight: 0 }],
    ["rowHeight", { rowHeight: Number.POSITIVE_INFINITY }],
    ["rowHeights key", { rowHeights: new Map([[-1, 20]]) }],
    ["rowHeights value", { rowHeights: new Map([[0, Number.NaN]]) }],
    ["headerHeight", { headerHeight: -1 }],
    ["headerHeight", { headerHeight: Number.POSITIVE_INFINITY }],
    ["rowHeaderWidth", { rowHeaderWidth: -1 }],
    ["rowHeaderWidth", { rowHeaderWidth: Number.NaN }],
    ["freeze.rows", { freeze: { rows: 1.5 } }],
    ["freeze.columns", { freeze: { columns: -1 } }],
    ["overscanRows", { overscanRows: Number.MAX_SAFE_INTEGER + 1 }],
    ["overscanColumns", { overscanColumns: 0.5 }],
    ["columns[0].width", { columns: [{ id: "a", width: Number.NaN }] }],
    ["columns[0].maxWidth", { columns: [{ id: "a", minWidth: 50, maxWidth: 40 }] }],
    ["clipboard.maxCells", { clipboard: { maxCells: -1 } }],
    [
      "clipboard.oversizedAction",
      { clipboard: { oversizedAction: "prompt" } } as unknown as CyExcelCoreOptions
    ]
  ])("rejects invalid %s before integration setup", (field, options) => {
    expectInvalid(() => validateCyExcelCoreOptions(options), field);
  });

  it.each([
    ["plugins[0].name", { name: 1, setup: () => {} }],
    ["plugins[0].setup", { name: "broken", setup: null }]
  ])("rejects an invalid %s shape", (field, plugin) => {
    expectInvalid(
      () => validateCyExcelCoreOptions({ plugins: [plugin] } as unknown as CyExcelCoreOptions),
      field
    );
  });

  it.each([
    ["options", null],
    ["freeze", { freeze: 1 }],
    ["columnTypes", { columnTypes: [] }],
    ["clipboard", { clipboard: null }]
  ])("rejects an invalid %s object shape", (field, options) => {
    expectInvalid(
      () => validateCyExcelCoreOptions(options as unknown as CyExcelCoreOptions),
      field
    );
  });
});

describe("validateCyExcelCoreDataDimensions", () => {
  it("accepts dimensions matching the configured columns and row-height bounds", () => {
    expect(() => validateCyExcelCoreDataDimensions(
      { rowHeights: new Map([[1, 32]]) },
      { rowCount: 2, columnCount: 3 },
      3
    )).not.toThrow();
  });

  it.each([
    ["dataSource.getRowCount()", { rowCount: -1, columnCount: 1 }, 1],
    ["dataSource.getColumnCount()", { rowCount: 1, columnCount: 1.5 }, 1],
    ["dataSource.getColumnCount()", { rowCount: 1, columnCount: 2 }, 1]
  ] as const)("rejects invalid %s", (field, dimensions, configuredColumnCount) => {
    expectInvalid(
      () => validateCyExcelCoreDataDimensions({}, dimensions, configuredColumnCount),
      field
    );
  });

  it("rejects an explicit row height outside the initial source row count", () => {
    expectInvalid(
      () => validateCyExcelCoreDataDimensions(
        { rowHeights: new Map([[2, 32]]) },
        { rowCount: 2, columnCount: 1 },
        1
      ),
      "rowHeights key"
    );
  });
});

describe("CyExcelCore constructor preflight", () => {
  it("fails a conflicting data input before reading the source or setting up a plugin", () => {
    const getRowCount = vi.fn(() => 0);
    const setup = vi.fn();
    const source = { ...dataSource(0, 0), getRowCount };

    expectInvalid(
      () => new CyExcelCore({} as HTMLElement, {
        data: [],
        dataSource: source,
        plugins: [{ name: "must-not-run", setup }]
      }),
      "data/dataSource"
    );

    expect(getRowCount).not.toHaveBeenCalled();
    expect(setup).not.toHaveBeenCalled();
  });

  it("fails an out-of-range rowHeights entry before touching DOM or plugins", () => {
    const setup = vi.fn();
    expectInvalid(
      () => new CyExcelCore({} as HTMLElement, {
        columns: [{ id: "a" }],
        data: [{}],
        rowHeights: new Map([[1, 30]]),
        plugins: [{ name: "must-not-run", setup }]
      }),
      "rowHeights key"
    );
    expect(setup).not.toHaveBeenCalled();
  });

  it("fails an invalid plugin shape before touching DOM", () => {
    expectInvalid(
      () => new CyExcelCore(
        {} as HTMLElement,
        { plugins: [{ name: "broken", setup: null }] } as unknown as CyExcelCoreOptions
      ),
      "plugins[0].setup"
    );
  });

  it("fails overflowing nested header geometry before touching DOM or plugins", () => {
    const setup = vi.fn();
    expectInvalid(
      () => new CyExcelCore({} as HTMLElement, {
        columns: [{ id: "a" }],
        headerHeight: Number.MAX_VALUE,
        headerGroups: [{ id: "group", title: "Group", columns: [0] }],
        plugins: [{ name: "must-not-run", setup }]
      }),
      "headerHeight"
    );
    expect(setup).not.toHaveBeenCalled();
  });

  it("fails a DataSource/column dimension mismatch before touching DOM or plugins", () => {
    const setup = vi.fn();
    expectInvalid(
      () => new CyExcelCore({} as HTMLElement, {
        columns: [{ id: "a" }],
        dataSource: dataSource(1, 2),
        plugins: [{ name: "must-not-run", setup }]
      }),
      "dataSource.getColumnCount()"
    );
    expect(setup).not.toHaveBeenCalled();
  });
});

describe("validateCyExcelCoreHeaderGeometry", () => {
  it("accepts a finite total across nested header bands", () => {
    expect(() => validateCyExcelCoreHeaderGeometry(24, 3)).not.toThrow();
  });

  it("rejects a finite band height whose nested total overflows", () => {
    expectInvalid(
      () => validateCyExcelCoreHeaderGeometry(Number.MAX_VALUE, 1),
      "headerHeight"
    );
  });
});
