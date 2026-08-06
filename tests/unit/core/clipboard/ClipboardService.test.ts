import { describe, expect, it, vi } from "vitest";

import {
  ArrayDataSource,
  CellTypeRegistry,
  ClipboardService,
  ColumnManager,
  CommandManager,
  SetRangeCommand,
  TableStore,
  escapeTsvCell,
  registerBuiltinCellTypes
} from "../../../../src/core";
import type { ClipboardIO } from "../../../../src/core";

function fakeIo(overrides: Partial<ClipboardIO> = {}): ClipboardIO & {
  writeText: ReturnType<typeof vi.fn>;
  readText: ReturnType<typeof vi.fn>;
  download: ReturnType<typeof vi.fn>;
} {
  return {
    writeText: vi.fn(async () => {}),
    readText: vi.fn(async () => ""),
    download: vi.fn(),
    ...overrides
  } as never;
}

function setup(columnDefs: { id: string; type?: string }[] = [{ id: "name" }, { id: "age", type: "number" }]) {
  const registry = new CellTypeRegistry();
  registerBuiltinCellTypes(registry);
  const columnManager = new ColumnManager({ columns: columnDefs, registry });

  const rows = [
    { name: "Alice", age: 30 },
    { name: "Bob", age: 40 },
    { name: "Carol", age: 50 }
  ];
  const dataSource = new ArrayDataSource(
    rows,
    columnDefs.map((c) => ({ id: c.id, field: c.id }))
  );
  const store = new TableStore(dataSource);
  const commandManager = new CommandManager(store);

  return { registry, columnManager, store, commandManager, dataSource };
}

function makeService(
  setupResult: ReturnType<typeof setup>,
  overrides: {
    io?: ReturnType<typeof fakeIo>;
    options?: ConstructorParameters<typeof ClipboardService>[0]["options"];
  } = {}
) {
  const { columnManager, store, commandManager } = setupResult;
  const io = overrides.io ?? fakeIo();

  const service = new ClipboardService({
    getCellValue: (row, column) => store.getCell(row, column),
    getColumn: (column) => columnManager.getColumn(column),
    getRowCount: () => setupResult.dataSource.getRowCount(),
    getColumnCount: () => columnManager.columnCount,
    parseValue: (row, column, text) => {
      const runtimeColumn = columnManager.getColumn(column);
      return runtimeColumn.parse ? runtimeColumn.parse(text, { row, column: runtimeColumn }) : text;
    },
    getSelectionRange: () => ({ startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 }),
    getStoreVersion: () => store.version,
    applyRangeCommand: (range, values) => {
      commandManager.execute(new SetRangeCommand(store, range, values));
    },
    io,
    ...(overrides.options !== undefined ? { options: overrides.options } : {})
  });

  return { service, io };
}

describe("ClipboardService.planCopy", () => {
  it("defaults to the current selection when no range is given", () => {
    const s = setup();
    const { service } = makeService(s);

    const plan = service.planCopy();

    expect(plan.rows).toBe(1);
    expect(plan.columns).toBe(2);
  });
});

describe("ClipboardService.copy", () => {
  it("writes display-formatted text to the clipboard and reports copied", async () => {
    const s = setup();
    const { service, io } = makeService(s);

    const result = await service.copy({ range: { startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 } });

    expect(result.status).toBe("copied");
    expect(io.writeText).toHaveBeenCalledWith("Alice\t30\r\nBob\t40");
  });

  it("writes raw-serialized text when mode is 'raw'", async () => {
    const registry = new CellTypeRegistry();
    registerBuiltinCellTypes(registry);
    const columnManager = new ColumnManager({ columns: [{ id: "tags", type: "array" }], registry });
    const dataSource = new ArrayDataSource(
      [{ tags: ["a", "b", "c", "d"] }],
      [{ id: "tags", field: "tags" }]
    );
    const store = new TableStore(dataSource);
    const commandManager = new CommandManager(store);
    const io = fakeIo();
    const service = new ClipboardService({
      getCellValue: (row, column) => store.getCell(row, column),
      getColumn: (column) => columnManager.getColumn(column),
      getRowCount: () => dataSource.getRowCount(),
      getColumnCount: () => columnManager.columnCount,
      parseValue: (row, column, text) => {
        const c = columnManager.getColumn(column);
        return c.parse ? c.parse(text, { row, column: c }) : text;
      },
      getSelectionRange: () => ({ startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 }),
      getStoreVersion: () => store.version,
      applyRangeCommand: (range, values) => commandManager.execute(new SetRangeCommand(store, range, values)),
      io
    });

    const result = await service.copy({ mode: "raw" });

    expect(result.status).toBe("copied");
    expect(io.writeText).toHaveBeenCalledWith(escapeTsvCell(JSON.stringify(["a", "b", "c", "d"])));
  });

  it("returns 'oversized' without calling the raw serializer or writing to the clipboard", async () => {
    const s = setup();
    let serializeCalled = false;
    s.registry.register("watched", {
      extends: "text",
      serialize: (value) => {
        serializeCalled = true;
        return String(value);
      }
    });
    const columnManager = new ColumnManager({
      columns: [{ id: "name", type: "watched" }, { id: "age", type: "number" }],
      registry: s.registry
    });
    const { service, io } = makeService({ ...s, columnManager }, { options: { maxCells: 1 } });

    const result = await service.copy({
      mode: "raw",
      range: { startRow: 0, startColumn: 0, endRow: 2, endColumn: 1 }
    });

    expect(result.status).toBe("oversized");
    expect(serializeCalled).toBe(false);
    expect(io.writeText).not.toHaveBeenCalled();
  });

  it("proceeds with the copy when the caller confirms an oversized plan", async () => {
    const s = setup();
    const { service, io } = makeService(s, { options: { maxCells: 1 } });

    const result = await service.copy({
      confirmed: true,
      range: { startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 }
    });

    expect(result.status).toBe("copied");
    expect(io.writeText).toHaveBeenCalled();
  });

  it("downloads a Blob instead of writing to the clipboard when the plan is oversized and download is requested", async () => {
    const s = setup();
    const { service, io } = makeService(s, { options: { maxCells: 1 } });

    const result = await service.copy({
      download: true,
      range: { startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 }
    });

    expect(result.status).toBe("downloaded");
    expect(io.writeText).not.toHaveBeenCalled();
    expect(io.download).toHaveBeenCalledTimes(1);
    const [blob, filename] = io.download.mock.calls[0] as [Blob, string];
    expect(blob).toBeInstanceOf(Blob);
    expect(filename).toMatch(/\.tsv$/);
  });

  it("reports cancelled and never writes to the clipboard when the signal aborts before the worker responds", async () => {
    const s = setup();
    const io = fakeIo();
    let capturedWorker: { onmessage: unknown; terminate: () => void } | undefined;
    const service = new ClipboardService({
      getCellValue: (row, column) => s.store.getCell(row, column),
      getColumn: (column) => s.columnManager.getColumn(column),
      getRowCount: () => s.dataSource.getRowCount(),
      getColumnCount: () => s.columnManager.columnCount,
      parseValue: (_row, _column, text) => text,
      getSelectionRange: () => ({ startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 }),
      getStoreVersion: () => s.store.version,
      applyRangeCommand: (range, values) => s.commandManager.execute(new SetRangeCommand(s.store, range, values)),
      io,
      createWorker: () => {
        const worker = {
          onmessage: null as unknown,
          onerror: null as unknown,
          postMessage: () => {},
          terminate: vi.fn()
        };
        capturedWorker = worker as never;
        return worker as never;
      }
    });

    const controller = new AbortController();
    const resultPromise = service.copy({
      range: { startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 },
      signal: controller.signal
    });
    controller.abort();
    const result = await resultPromise;

    expect(result.status).toBe("cancelled");
    expect(io.writeText).not.toHaveBeenCalled();
    expect(capturedWorker?.terminate).toHaveBeenCalled();
  });
});

describe("ClipboardService.paste", () => {
  it("commits a pasted range as exactly one transaction / one History entry", async () => {
    const s = setup();
    const { service } = makeService(s);

    expect(s.commandManager.historySize).toBe(0);
    await service.paste({ range: { startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 }, text: "X\t99" });

    expect(s.commandManager.historySize).toBe(1);
    expect(s.store.getCell(0, 0)).toBe("X");
    expect(s.store.getCell(0, 1)).toBe(99);

    s.commandManager.undo();
    expect(s.store.getCell(0, 0)).toBe("Alice");
    expect(s.store.getCell(0, 1)).toBe(30);
  });

  it("pastes a multi-row multi-column TSV block starting at the given anchor", async () => {
    const s = setup();
    const { service } = makeService(s);

    await service.paste({
      range: { startRow: 1, startColumn: 0, endRow: 1, endColumn: 1 },
      text: "Zed\t1\r\nYak\t2"
    });

    expect(s.store.getCell(1, 0)).toBe("Zed");
    expect(s.store.getCell(1, 1)).toBe(1);
    expect(s.store.getCell(2, 0)).toBe("Yak");
    expect(s.store.getCell(2, 1)).toBe(2);
  });

  it("keeps the existing value for a column whose clipboard.paste policy is false", async () => {
    const registry = new CellTypeRegistry();
    registerBuiltinCellTypes(registry);
    const columnManager = new ColumnManager({
      columns: [{ id: "name" }, { id: "age", type: "number", clipboard: { paste: false } }],
      registry
    });
    const dataSource = new ArrayDataSource([{ name: "Alice", age: 30 }], [
      { id: "name", field: "name" },
      { id: "age", field: "age" }
    ]);
    const store = new TableStore(dataSource);
    const commandManager = new CommandManager(store);
    const service = new ClipboardService({
      getCellValue: (row, column) => store.getCell(row, column),
      getColumn: (column) => columnManager.getColumn(column),
      getRowCount: () => dataSource.getRowCount(),
      getColumnCount: () => columnManager.columnCount,
      parseValue: (row, column, text) => {
        const c = columnManager.getColumn(column);
        return c.parse ? c.parse(text, { row, column: c }) : text;
      },
      getSelectionRange: () => ({ startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 }),
      getStoreVersion: () => store.version,
      applyRangeCommand: (range, values) => commandManager.execute(new SetRangeCommand(store, range, values)),
      io: fakeIo()
    });

    await service.paste({ range: { startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 }, text: "Changed\t999" });

    expect(store.getCell(0, 0)).toBe("Changed");
    expect(store.getCell(0, 1)).toBe(30);
  });

  it("clamps a paste that would overflow the grid to the remaining rows/columns instead of throwing", async () => {
    const s = setup();
    const { service } = makeService(s);

    await expect(
      service.paste({
        range: { startRow: 2, startColumn: 0, endRow: 2, endColumn: 1 },
        text: "Row1\t1\r\nRow2\t2\r\nRow3\t3"
      })
    ).resolves.toBeUndefined();

    expect(s.store.getCell(2, 0)).toBe("Row1");
  });

  it("does not commit anything when the signal is already aborted", async () => {
    const s = setup();
    const { service } = makeService(s);
    const controller = new AbortController();
    controller.abort();

    await service.paste({
      range: { startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 },
      text: "X\t1",
      signal: controller.signal
    });

    expect(s.commandManager.historySize).toBe(0);
    expect(s.store.getCell(0, 0)).toBe("Alice");
  });

  it("reads from clipboard IO when no explicit text is given", async () => {
    const s = setup();
    const io = fakeIo({ readText: vi.fn(async () => "Read\t7") });
    const { service } = makeService(s, { io });

    await service.paste({ range: { startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 } });

    expect(io.readText).toHaveBeenCalled();
    expect(s.store.getCell(0, 0)).toBe("Read");
    expect(s.store.getCell(0, 1)).toBe(7);
  });
});
