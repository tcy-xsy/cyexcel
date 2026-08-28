import type * as CoreApi from "../../src/core";
import {
  rowsToWorkbookSnapshot,
  SetCellCommand,
  type CellRange,
  type Command,
  type CyExcelCoreOptions,
  type CyExcelPlugin,
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
    case "row-reorder":
      return `row-reorder:${value.order.join(",")}`;
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
  snapshot: rowsToWorkbookSnapshot([{ name: "Ada" }], ["name"]),
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
  IsPublic<"SceneSnapshot">,
  IsPublic<"SetCellCommand">,
  IsPublic<"createBasicContextMenuPlugin">
] = [true, true, true, true];

// 数据层不再可替换（ADR 0004 §3.2）：`DataSource` 与它唯一的实现
// `ArrayDataSource` 一起删掉了，桶里不该再冒出第二条写数据的路。
// 这里把「不导出」钉成类型事实，重新导出会在这一行变红。
const removedDataSurface: [IsPublic<"DataSource">, IsPublic<"ArrayDataSource">] = [false, false];
void removedDataSurface;

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
