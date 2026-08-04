import type {
  CellRange,
  Command,
  CommandContext,
  CyExcelOptions,
  DataSource,
  Patch
} from "../../src/core";

// --- CellRange: fixed shape and readonly fields -------------------------

const range: CellRange = {
  startRow: 0,
  startColumn: 0,
  endRow: 10,
  endColumn: 5
};

// @ts-expect-error readonly fields cannot be reassigned
range.startRow = 1;

// --- DataSource: fixed shape, setCell must stay synchronous -------------

const dataSource: DataSource = {
  getRowCount: () => 0,
  getColumnCount: () => 0,
  getCell: (row, column) => `${row}:${column}`,
  setCell: (row, column, value) => {
    void row;
    void column;
    void value;
  }
};

const setCellReturnsVoid: void = dataSource.setCell(0, 0, 1);
void setCellReturnsVoid;

// --- Command: fixed shape, execute/undo must stay synchronous -----------

const patch: Patch = { type: "cell-value", row: 0, column: 0 };

const command: Command = {
  type: "noop",
  execute: (context: CommandContext) => {
    void context;
    return patch;
  },
  undo: (context: CommandContext) => {
    void context;
    return patch;
  }
};

const executeReturnsPatch: Patch | Patch[] = command.execute({
  transactionId: "t-1"
});
void executeReturnsPatch;

// --- Patch: discriminated union stays exhaustive -------------------------

function describePatch(value: Patch): string {
  switch (value.type) {
    case "cell-value":
    case "cell-style":
      return `${value.type}:${value.row}:${value.column}`;
    case "range-value":
      return `range-value:${value.range.startRow}`;
    case "row-structure":
      return `row-structure:${value.operation}:${value.start}:${value.count}`;
    case "column-layout":
      return `column-layout:${value.column}`;
    case "full-reset":
      return "full-reset";
  }
}
void describePatch;

// --- CyExcelOptions: fixed shape and readonly fields ---------------------

const options: CyExcelOptions = {
  rowHeight: 32,
  headerHeight: 40,
  readonly: false,
  overscanRows: 5,
  overscanColumns: 2,
  freeze: { rows: 1, columns: 1 }
};

// @ts-expect-error readonly fields cannot be reassigned
options.rowHeight = 40;

// @ts-expect-error readonly fields cannot be reassigned
options.freeze = { rows: 2 };
