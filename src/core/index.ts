/**
 * Product API for the headless-without-registration core entry point.
 *
 * Mutable runtime owners and integration factories deliberately stay behind
 * CyExcelCore. Public consumers receive immutable snapshots and narrow
 * extension contracts instead of alternate mutation paths.
 */
export {
  CyExcelCore,
  createCyExcelCore,
  type CyExcelColumnConfig,
  type CyExcelCoreOptions,
  type CyExcelHeaderActionEvent,
  type SortState
} from "./CyExcelCore";

export { CyExcelError, CyExcelErrorCode } from "./errors";

export type { CellRange, Command, Disposer } from "./types";

/**
 * 存储层的公开面：模型、快照与导入器。
 *
 * 替代了原来导出的 `DataSource` / `ArrayDataSource` 一族——调用方不再替换数据层
 * （ADR 0004 §3.2），而是交一份快照进来。
 */
export {
  rowsToWorkbookSnapshot,
  SheetModel,
  WorkbookModel,
  type CellRecord,
  type ImportRow,
  type RowsToWorkbookSnapshotOptions,
  type SheetSnapshot,
  type WorkbookSnapshot
} from "./sheet";
export type {
  ChangeSet,
  SemanticChange
} from "./data";
export type { ModelRevision } from "./model";
export type { TransactionReadView } from "./mutations";

// Frozen command intents are observable from command hooks. Planning,
// mutation creation, execution and history remain owned by CyExcelCore.
export {
  SetCellCommand,
  type SetCellCommandParams,
  SetRangeCommand,
  type SetRangeCommandParams,
  MergeCellsCommand,
  type MergeCellsCommandParams,
  UnmergeCellsCommand,
  type UnmergeCellsCommandParams,
  ClearRangeCommand,
  type ClearRangeCommandParams,
  InsertRowsCommand,
  type InsertRowsCommandParams,
  DeleteRowsCommand,
  type DeleteRowsCommandParams,
  MoveRowsCommand,
  type MoveRowsCommandParams,
  ReorderRowsCommand,
  type ReorderRowsCommandParams,
  InsertColumnsCommand,
  type InsertColumnsCommandParams,
  DeleteColumnsCommand,
  type DeleteColumnsCommandParams,
  MoveColumnsCommand,
  type MoveColumnsCommandParams,
  HideColumnCommand,
  type HideColumnCommandParams,
  ShowColumnCommand,
  type ShowColumnCommandParams,
  ResizeRowCommand,
  type ResizeRowCommandParams,
  SetRowHeightsCommand,
  type SetRowHeightsCommandParams,
  type SetRowHeightUpdate,
  type RowHeightCommandProvenance,
  ResizeColumnCommand,
  type ResizeColumnCommandParams,
  type BeforeCommandContext,
  type BeforeTransactionCommitContext,
  type BeforeRowsChangeContext,
  type BeforeColumnsChangeContext
} from "./commands";

// Immutable geometry and selection published by CyExcelCore.
export {
  SceneSnapshot,
  type SceneSurfaceMetrics,
  type SceneColumnHeaderBandSpec,
  type SceneColumnHeaderBandDescriptor,
  type SceneColumnHeaderLeafBandDescriptor,
  type SceneCellFragment,
  type SceneColumnHeaderFragment,
  type SceneRowHeaderFragment,
  type SceneSelectionHandleFragment,
  type SceneCellAddress,
  type SceneScrollTarget,
  type SceneLocalPoint,
  type SceneCellHit,
  type SceneColumnHeaderHit,
  type SceneRowHeaderHit,
  type SceneRootDescriptor,
  type SceneSurfaceDescriptor,
  type SceneLayerDescriptor,
  type ScenePaneDescriptor,
  type SceneNodeDescriptor,
  type SceneSurfaceId,
  type SceneLayerId,
  type SceneBodyPaneId,
  type ScenePaneId,
  type SceneAxisBehavior,
  type SceneTranslation
} from "./scene";
export type {
  CellRect,
  IndexRange,
  ViewportState
} from "./layout";
export type {
  SelectionKind,
  CellPosition,
  SelectionBounds,
  SelectionSnapshot
} from "./interaction";

// Sorting primitives, so a caller can build its own sort (a multi-column one,
// say) on the same permutation contract `sortByColumn()` commits.
export {
  computeSortOrder,
  createDefaultSortCompare,
  nextSortDirection,
  type SortRequest
} from "./sort";

// User-facing column, cell, header and editor extension contracts.
export type { Canvas2DLike } from "./render";
export {
  type CellType,
  type CellTypeDefinition,
  type CellFormatContext,
  type CellRenderContext,
  type CellHitTestContext,
  type CellAction,
  type CanvasCellRenderer,
  type DomCellContext,
  type DomCellRenderer,
  type DomCellRendererFactory,
  type ColumnDefinition,
  type RuntimeColumn,
  type SelectOption,
  type HeaderState,
  type HeaderRenderContext,
  type HeaderHitTestContext,
  type HeaderAction,
  type CanvasHeaderRenderer,
  type DomHeaderContext,
  type DomHeaderRenderer,
  type DomHeaderRendererFactory,
  type HeaderDefinition,
  type ColumnClipboardPolicy
} from "./cells";
export {
  defaultHeaderHitTest,
  DEFAULT_HEADER_SORT_ZONE_WIDTH,
  DEFAULT_HEADER_SORT_ZONE_MIN_WIDTH,
  type RuntimeHeader
} from "./headers";
export {
  createTextCellEditor,
  createBooleanCellEditor,
  createDateCellEditor,
  createSelectCellEditor,
  createTextAreaCellEditor,
  type CellEditor,
  type CellEditorContext,
  type CellEditorFactory,
  type ValidationResult,
  type SelectCellEditorOption,
  type TextAreaCellEditorOptions
} from "./editors";

export type {
  CopyMode,
  CopyPlan,
  ClipboardOptions,
  ClipboardCopyStatus,
  ClipboardCopyResult,
  ClipboardCopyOptions,
  ClipboardPasteOptions
} from "./clipboard";

// Extension services are exposed by CyExcelCore getters. Their classes are
// type-only here so the public barrel does not advertise standalone mutable
// runtimes with a second lifecycle.
export type {
  HookManager,
  CancellableContext,
  HookHandlerMeta
} from "./hooks";
export type { EventBus } from "./events";
export {
  createBasicContextMenuPlugin,
  type ContextMenuService,
  type ContextMenuTarget,
  type ContextMenuContext,
  type ContextMenuItem,
  type ContextMenuProvider,
  type ContextMenuProviderMeta
} from "./menu";
export type { ShortcutHandler } from "./shortcuts";
export type { CyExcelPlugin, PluginContext } from "./plugins";
