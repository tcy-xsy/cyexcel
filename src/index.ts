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
 * 装进表里——对象数组不再是入口能吃的形状，而快照只能由 `rowsToWorkbookSnapshot`
 * 造。少了这一个导出，根入口就只能开空表。
 */
export {
  rowsToWorkbookSnapshot,
  type ImportRow,
  type RowsToWorkbookSnapshotOptions,
  type WorkbookSnapshot
} from "./core";
export {
  DEFAULT_CYEXCEL_TAG_NAME,
  defineCyExcel
} from "./element/defineCyExcel";
