export {
  CyExcelCore,
  createCyExcelCore,
  type CyExcelCoreOptions,
  type CyExcelHeaderActionEvent,
  type SortState
} from "./core";
/*
 * 导入器必须跟着根入口一起出去：`CyExcelCoreOptions.data` 收成 `snapshot` 之后
 * （ADR 0004 §3.2），只 import `CyExcelCore` 的调用方**没有任何办法**把一份数据
 * 装进表里——裸数组不再是入口能吃的形状，而快照由 `matrixToWorkbookSnapshot`
 * / `rowsToWorkbookSnapshot` 造。少了这些导出，根入口就只能开空表。
 */
export {
  matrixToWorkbookSnapshot,
  rowsToWorkbookSnapshot,
  type ImportRow,
  type MatrixToWorkbookSnapshotOptions,
  type RowsToWorkbookSnapshotOptions,
  type WorkbookSnapshot
} from "./core";
export {
  DEFAULT_CYEXCEL_TAG_NAME,
  defineCyExcel
} from "./element/defineCyExcel";

export { createSheet, type CreateSheetOptions } from "./createSheet";
export {
  createWorkbook,
  type CreateWorkbookOptions,
  type WorkbookHandle,
  type WorkbookSheetOptions
} from "./createWorkbook";

// 工具栏与编辑栏也从根入口出去：用 createSheet() 的人多半还想改 schema 或单独挂一条。
export {
  createFormulaBar,
  createToolbar,
  defaultToolbarSchema,
  formulaBarWidget,
  toolbarWidget,
  type ToolbarItem,
  type ToolbarSchema
} from "./core/toolbar";
