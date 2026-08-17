import type * as CoreApi from "../../src/core";
import {
  SetCellCommand,
  type CellRange,
  type Command,
  type CyExcelCoreOptions,
  type CyExcelPlugin,
  type DataSource,
  type SemanticChange,
  type SceneSnapshot,
  type SelectionSnapshot
} from "../../src/core";

// --- Stable public value contracts ---------------------------------------

const range: CellRange = {
  startRow: 0,
  startColumn: 0,
  endRow: 10,
  endColumn: 5
};

// @ts-expect-error public ranges are immutable values
range.startRow = 1;

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

const command: Command = new SetCellCommand(0, 1, "value");
void command;

function describeChange(value: SemanticChange): string {
  switch (value.type) {
    case "cell-value":
    case "cell-style":
      return `${value.type}:${value.row}:${value.column}`;
    case "range-value":
      return `range-value:${value.range.startRow}`;
    case "row-structure":
      return `row-structure:${value.operation}:${value.start}:${value.count}`;
    case "column-structure":
      return `column-structure:${value.operation}:${value.start}:${value.count}`;
    case "column-layout":
      return `column-layout:${value.column}`;
    case "row-layout":
      return `row-layout:${value.rows.join(",")}`;
    case "full-reset":
      return "full-reset";
  }
}
void describeChange;

const rowLayoutChange: SemanticChange = { type: "row-layout", rows: [2, 4], range };
// @ts-expect-error a row-layout fact always identifies at least one changed row
const emptyRowLayoutChange: SemanticChange = { type: "row-layout", rows: [], range };
if (rowLayoutChange.type === "row-layout") {
  // @ts-expect-error row-layout rows are immutable
  rowLayoutChange.rows.push(6);
}
void emptyRowLayoutChange;

const options: CyExcelCoreOptions = {
  columns: [{ id: "name", field: "name", width: 120 }],
  data: [{ name: "Ada" }],
  rowHeight: 32,
  rowHeights: new Map([[2, 48]]),
  headerHeight: 40,
  overscanRows: 5,
  overscanColumns: 2,
  freeze: { rows: 1, columns: 1 }
};

// @ts-expect-error public options are immutable
options.rowHeight = 40;

// @ts-expect-error nested public options are immutable
options.freeze = { rows: 2 };

const plugin: CyExcelPlugin = {
  name: "type-contract",
  setup(context) {
    const dispose = context.events.on("change", () => {});
    return dispose;
  }
};
void plugin;

declare const scene: SceneSnapshot;
const selection: SelectionSnapshot = {
  kind: "cell",
  active: { row: 0, column: 0 },
  anchor: { row: 0, column: 0 },
  range: { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 },
  ranges: [{ startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 }]
};
const sceneVersion: number = scene.version;
void sceneVersion;
void selection;

// --- The product barrel has no alternate mutable runtime -----------------

type IsPublic<Name extends PropertyKey> = Name extends keyof typeof CoreApi ? true : false;

const retainedSurface: [
  IsPublic<"CyExcelCore">,
  IsPublic<"ArrayDataSource">,
  IsPublic<"SceneSnapshot">,
  IsPublic<"SetCellCommand">,
  IsPublic<"createBasicContextMenuPlugin">
] = [true, true, true, true, true];

const internalSurface: [
  IsPublic<"TableStore">,
  IsPublic<"MutationCoordinator">,
  IsPublic<"LayoutEngine">,
  IsPublic<"GridStructureOwner">,
  IsPublic<"createGridCommandManager">,
  IsPublic<"createTableCommandManager">,
  IsPublic<"CommandManager">,
  IsPublic<"SelectionModel">,
  IsPublic<"SurfaceHost">,
  IsPublic<"RenderScheduler">,
  IsPublic<"ColumnManager">,
  IsPublic<"HeaderManager">,
  IsPublic<"PluginManager">,
  IsPublic<"HistoryManager">,
  IsPublic<"CyExcelOptions">,
  IsPublic<"assertValidCyExcelOptions">
] = [
  false,
  false,
  false,
  false,
  false,
  false,
  false,
  false,
  false,
  false,
  false,
  false,
  false,
  false,
  false,
  false
];

void retainedSurface;
void internalSurface;
