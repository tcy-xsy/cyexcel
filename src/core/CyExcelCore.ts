import {
  CellTypeRegistry,
  ColumnManager,
  DEFAULT_COLUMN_WIDTH,
  DOMCellManager,
  createColumnAwarePainter,
  registerBuiltinCellTypes,
  snapshotColumnDefinitions,
  type CellTypeDefinition,
  type ColumnDefinition,
  type RuntimeColumn
} from "./cells";
import {
  ClipboardService,
  buildCellGrid,
  encodeTsv,
  type ClipboardCopyOptions,
  type ClipboardCopyResult,
  type ClipboardIO,
  type ClipboardOptions,
  type ClipboardPasteOptions,
  type CopyPlan,
  type SerializeWorkerLike
} from "./clipboard";
import {
  SetCellCommand,
  SetRangeCommand,
  MergeCellsCommand,
  UnmergeCellsCommand,
  InsertRowsCommand,
  DeleteRowsCommand,
  MoveRowsCommand,
  ReorderRowsCommand,
  InsertColumnsCommand,
  DeleteColumnsCommand,
  MoveColumnsCommand,
  HideColumnCommand,
  ShowColumnCommand,
  ResizeRowCommand,
  SetRowHeightsCommand,
  ResizeColumnCommand,
  dispatchCellAction,
  type CommandManager
} from "./commands";
import {
  TableStore,
  compileDirtySet,
  createChangeSet,
  type SemanticChange,
  summarizeStructuralChanges
} from "./data";
import { WorkbookModel, type SheetModel, type WorkbookSnapshot } from "./sheet";
import { findMergeAt } from "./sheet/merges";
import type { CellEditorFactory } from "./editors";
import { EditorManager } from "./editors";
import { EventBus } from "./events";
import {
  DomHeaderManager,
  HeaderCanvasRenderer,
  HeaderManager,
  RowHeaderCanvasRenderer,
  defaultHeaderHitTest,
  type HeaderAction,
  type HeaderState,
  type RuntimeHeader
} from "./headers";
import { HookManager, PRODUCTION_HOOK_REGISTRY } from "./hooks";
import {
  FillHandleController,
  HeaderGestureController,
  HeaderResizeController,
  HeaderSelectionController,
  InputRouter,
  InteractionFrameProvider,
  KeyboardController,
  PointerController,
  BodyGestureController,
  RowReorderController,
  SelectionModel,
  computeFillPlan,
  computeAutoFillDownTargetRow,
  declineInput,
  getDeclaredDomPointerInputTarget,
  handleInput,
  nativeInputTarget,
  noInputTarget,
  surfaceInputTarget,
  isPointOnFillHandle,
  type InteractionFrame,
  type HeaderResizeIntent,
  type InputCandidateContext,
  type InputSurfacePolicyContext,
  type InputTargetResolution,
  type SelectionSnapshot
} from "./interaction";
import {
  LayoutEngine,
  SkeletonSnapshot,
  validateAxisChange,
  ViewportSnapshot,
  type LayoutRoot,
  type RowHeightProvenance,
  type ViewportState
} from "./layout";
import { InitializationRollback } from "./lifecycle";
import { ContextMenuService, classifyContextMenuTarget } from "./menu";
import { DeferredMutationQueue } from "./model";
import { PluginManager, type CyExcelPlugin } from "./plugins";
import {
  SceneSnapshot,
  ViewportManager,
  SurfaceHost,
  clientPoint,
  type SurfacePoint,
  type SurfaceId,
  type SurfaceSetSnapshot
} from "./scene";
import { WrapMeasurementService, type RowHeightProvenanceView } from "./measurement";
import { ShortcutRegistry, type ShortcutHandler } from "./shortcuts";
import { computeSortOrder, createDefaultSortCompare, nextSortDirection, type SortState } from "./sort";
import {
  ContentLayerRenderer,
  CanvasSurface,
  RenderScheduler,
  SelectionLayerRenderer,
  readDevicePixelRatio,
  type Canvas2DLike
} from "./render";
import { CELL_FONT } from "./render/theme";
import type { CellRange, Command, Disposer } from "./types";
import {
  GridStructureOwner,
  StructureTransition,
  createGridCommandManager,
  resolveInitialAxisGeometry,
  seedAxisMeta
} from "./integration";
import { columnWidthOf } from "./mutations/columnWidthMeta";
import {
  validateCyExcelCoreDataDimensions,
  validateCyExcelCoreOptions
} from "./validateCyExcelCoreOptions";

export type CyExcelColumnConfig = ColumnDefinition;

export type { SortState };

/** Emitted after a custom canvas header hitTest returns an action. */
export interface CyExcelHeaderActionEvent {
  readonly action: HeaderAction;
  readonly column: number;
  readonly header: RuntimeHeader;
  readonly structureVersion: number;
}

export interface CyExcelCoreOptions {
  readonly columns?: readonly CyExcelColumnConfig[];
  /**
   * 初始数据，**一份快照**（ADR 0004 §3.1「模型即快照」）。省略就是一张
   * 0 行 × `columns.length` 列的空表。
   *
   * 对象数组走导入器，它是入口之外的独立设施（§3.2）：
   *
   * ```ts
   * new CyExcelCore(container, {
   *   columns,
   *   snapshot: rowsToWorkbookSnapshot(rows, ["sku", "name", "price"])
   * });
   * ```
   *
   * **`fields` 是导入的参数，不是列配置的字段**——数据长什么形状是数据自己的
   * 事实（ADR §4.5 否决了拿 `field` 当数据格式开关），列配置只管宽度、类型这些
   * 展示属性。第 `i` 个字段就是第 `i` 列。
   *
   * **快照的列数必须等于 `columns.length`**，对不上在构造期就报
   * `invalid-options`：两者是同一根轴的两种说法，不一致只会以「渲染少画一列」
   * 的形式静默出现。
   *
   * **没有写回。** 快照被读进存储模型之后，与调用方手里的对象再无联系；所有
   * 编辑一律走 `setCell()`/`setRange()`/`paste()`，它们经过 Command/Patch/History。
   * 这是 §3.2 的直接后果，是意图之内的破坏性变更（4B 设计 §3.3）。
   */
  readonly snapshot?: WorkbookSnapshot;
  readonly rowHeight?: number;
  /** Row index -> explicit height, only for rows that differ from `rowHeight`. See LayoutOptions.rowHeights. */
  readonly rowHeights?: ReadonlyMap<number, number>;
  readonly freeze?: {
    readonly rows?: number;
    readonly columns?: number;
  };
  readonly overscanRows?: number;
  readonly overscanColumns?: number;
  /** What an in-progress edit does when the viewport scrolls. Default: "commit". */
  readonly scrollCommitPolicy?: "commit" | "cancel";
  readonly editorFactory?: CellEditorFactory;
  /**
   * Extra named CellTypes to register (alongside the built-ins) before
   * columns are resolved — `columns[].type` can reference these by name.
   * Registered after the built-ins, so `extends: "number"` etc. resolves;
   * a name that collides with a built-in overrides it. There's no
   * post-construction `registerColumnType()` method: columns resolve once
   * at construction (ColumnManager's cached-merge design, see the Step 8
   * plan), so registering a type afterwards would have nothing left to
   * apply it to until a dynamic column API exists (Step 12).
   */
  readonly columnTypes?: Readonly<Record<string, CellTypeDefinition>>;
  /**
   * Pixel height of the column title row. 0 (default) means no header band
   * at all — matches `freeze`/`overscanRows` defaulting to off, and keeps
   * every existing fixture/test that never set this unaffected (their
   * canvas content still starts at y=0). Columns still resolve their
   * `header` config either way; this only controls whether/how tall the
   * band that displays it is.
   */
  readonly headerHeight?: number;
  /**
   * When a column has no explicit `header`, show its spreadsheet-style
   * letter (A, B, ..., Z, AA, ...). **Default `true`** (ADR 0004 §3.8.1):
   * columns no longer have an `id`, so there is no other name to fall back
   * to. Set it `false` to leave such columns with no title at all — that is
   * now what "off" means, not "use the id". An
   * explicit `header` always overrides the letter for that one column —
   * letter and custom title are mutually exclusive per column (one row),
   * not a separate always-visible letter row stacked on top — matching
   * Univer's `customizeColumnHeader()` model.
   */
  readonly columnLetterFallback?: boolean;
  /** Pixel width of the row-number column on the left. 0 (default) means no row header band at all — same "unset is a no-op" precedent as headerHeight. */
  readonly rowHeaderWidth?: number;
  /**
   * Enables dragging the top/bottom edge of the current row selection to
   * reorder rows (insert-style, via the same `moveRows()` Command path as
   * the programmatic API and the row-header context menu — never an
   * Excel-style destructive overwrite move). Default `false`: the row
   * header's existing drag-to-select-multiple-rows gesture is completely
   * unaffected either way, and with this off it's the only gesture the row
   * header has, unchanged from before this option existed.
   */
  readonly rowReorder?: boolean;
  /** Draw a light gridline around every body cell. Default: true. */
  readonly showGridLines?: boolean;
  readonly clipboard?: ClipboardOptions;
  /** Installed via PluginManager at construction, in order — see doc/customization-and-hooks.md §11.2. Nothing is installed by default (e.g. createBasicContextMenuPlugin() must be listed explicitly to get a Copy/Paste right-click menu). */
  readonly plugins?: readonly CyExcelPlugin[];
}

const DEFAULT_ROW_HEIGHT = 24;
const AUTO_HEIGHT_MAX_SYNC_PASSES = 2;

interface DomLayer {
  readonly scroller: HTMLDivElement;
  readonly spacer: HTMLDivElement;
  readonly headerCanvas: HTMLCanvasElement;
  readonly headerDomLayer: HTMLDivElement;
  readonly rowHeaderCanvas: HTMLCanvasElement;
  readonly contentCanvas: HTMLCanvasElement;
  readonly selectionCanvas: HTMLCanvasElement;
  readonly domCellLayer: HTMLDivElement;
  readonly editorLayer: HTMLDivElement;
  readonly menuLayer: HTMLDivElement;
}

function createDom(container: HTMLElement, headerHeight: number, rowHeaderWidth: number): DomLayer {
  const scroller = document.createElement("div");
  scroller.style.position = "relative";
  scroller.style.width = "100%";
  scroller.style.height = "100%";
  scroller.style.overflow = "auto";

  const spacer = document.createElement("div");
  spacer.style.position = "relative";
  // Later in DOM order than stickyLayer (see the note below), so without
  // this it would paint on top of — and silently swallow every pointer
  // event over — the canvases despite being invisible.
  spacer.style.pointerEvents = "none";

  const stickyLayer = document.createElement("div");
  stickyLayer.style.position = "sticky";
  stickyLayer.style.top = "0";
  stickyLayer.style.left = "0";
  stickyLayer.style.width = "0";
  stickyLayer.style.height = "0";
  stickyLayer.style.overflow = "visible";

  // Header band occupies the fixed top strip [0, headerHeight) and the row
  // header occupies the fixed left strip [0, rowHeaderWidth); body layers
  // are shifted down/right by those same amounts so their own local
  // (0-based) coordinate space starts right where the header bands end —
  // no header offset needs to be threaded through PointerController/
  // EditorManager/DOMCellManager math, the CSS position alone accounts
  // for it.
  const bodyLeft = `${rowHeaderWidth}px`;

  const headerCanvas = document.createElement("canvas");
  headerCanvas.style.position = "absolute";
  headerCanvas.style.top = "0";
  headerCanvas.style.left = bodyLeft;

  const headerDomLayer = document.createElement("div");
  headerDomLayer.style.position = "absolute";
  headerDomLayer.style.top = "0";
  headerDomLayer.style.left = bodyLeft;

  // Spans the *full* viewport height (header band + body), unlike the body
  // layers below — RowHeaderCanvasRenderer uses that extra room to paint
  // the top-left corner cell itself instead of needing a third canvas.
  const rowHeaderCanvas = document.createElement("canvas");
  rowHeaderCanvas.style.position = "absolute";
  rowHeaderCanvas.style.top = "0";
  rowHeaderCanvas.style.left = "0";

  const bodyTop = `${headerHeight}px`;

  const contentCanvas = document.createElement("canvas");
  contentCanvas.style.position = "absolute";
  contentCanvas.style.top = bodyTop;
  contentCanvas.style.left = bodyLeft;
  contentCanvas.tabIndex = 0;

  const selectionCanvas = document.createElement("canvas");
  selectionCanvas.style.position = "absolute";
  selectionCanvas.style.top = bodyTop;
  selectionCanvas.style.left = bodyLeft;
  selectionCanvas.style.pointerEvents = "none";

  const domCellLayer = document.createElement("div");
  domCellLayer.style.position = "absolute";
  domCellLayer.style.top = bodyTop;
  domCellLayer.style.left = bodyLeft;

  const editorLayer = document.createElement("div");
  editorLayer.style.position = "absolute";
  editorLayer.style.top = bodyTop;
  editorLayer.style.left = bodyLeft;

  // Unlike the body layers above, menuLayer is NOT offset by bodyTop — a
  // right-click context menu can be anchored inside either the header band
  // or the body, so it needs the same un-offset coordinate space
  // #handleContextMenu computes its point in (relative to the scroller's
  // own visible top-left, header-inclusive).
  const menuLayer = document.createElement("div");
  menuLayer.style.position = "absolute";
  menuLayer.style.top = "0";
  menuLayer.style.left = "0";

  // domCellLayer before editorLayer so the (single, at most) active DOM
  // editor always renders on top of any custom DOM cell renderers. The
  // header pair comes next — the header band (y: [0, headerHeight)) and
  // the body band (y: [headerHeight, ...)) never visually overlap, so
  // relative paint order between the two groups makes no visual
  // difference, but existing code (and tests) that grabs "the first/second
  // canvas in #app" to mean content/selection needs that pair to stay
  // first in DOM order. menuLayer is appended last so a context menu
  // renders on top of literally everything else, including an active
  // editor.
  stickyLayer.append(
    contentCanvas,
    selectionCanvas,
    domCellLayer,
    editorLayer,
    headerCanvas,
    headerDomLayer,
    rowHeaderCanvas,
    menuLayer
  );
  // stickyLayer must come BEFORE spacer in flow order: a `position: sticky`
  // element only pins at `top/left: 0` once its natural flow position would
  // otherwise scroll past that point. If it started after the spacer (at
  // the bottom of the scrollable content), it would never need to "stick"
  // while scrolled to the top and would sit far below the viewport instead.
  scroller.append(stickyLayer, spacer);
  try {
    container.append(scroller);
  } catch (error) {
    scroller.remove();
    throw error;
  }

  return {
    scroller,
    spacer,
    headerCanvas,
    headerDomLayer,
    rowHeaderCanvas,
    contentCanvas,
    selectionCanvas,
    domCellLayer,
    editorLayer,
    menuLayer
  };
}

function isNativePointerPathEntry(entry: unknown): boolean {
  if (!(entry instanceof Element)) {
    return false;
  }
  if (["input", "select", "textarea", "button", "option"].includes(entry.localName)) {
    return true;
  }
  const contentEditable = entry.getAttribute("contenteditable");
  return contentEditable !== null && contentEditable.toLowerCase() !== "false";
}

/**
 * Off-screen, unfocusable textarea used only as the last-resort clipboard
 * fallback (old Safari, insecure contexts) — `document.execCommand("copy"/
 * "paste")` requires a real, briefly-focused text field, it can't operate
 * on an arbitrary string directly.
 */
function withHiddenTextarea<T>(run: (textarea: HTMLTextAreaElement) => T): T {
  const textarea = document.createElement("textarea");
  textarea.style.position = "fixed";
  textarea.style.top = "0";
  textarea.style.left = "0";
  textarea.style.opacity = "0";
  textarea.style.pointerEvents = "none";
  document.body.appendChild(textarea);
  try {
    return run(textarea);
  } finally {
    textarea.remove();
  }
}

/**
 * Tier 2 (async Clipboard API) with a tier 3 (hidden textarea +
 * execCommand) fallback for browsers/contexts where it's unavailable or
 * throws (permission denied, insecure context). Used by the programmatic
 * copy()/paste() API — the native copy/paste ClipboardEvent path (tier 1)
 * is handled separately in #handleCopyEvent/#handlePasteEvent below, since
 * that one needs to stay synchronous.
 */
function createBrowserClipboardIO(): ClipboardIO {
  return {
    async writeText(text) {
      try {
        await navigator.clipboard.writeText(text);
        return;
      } catch {
        // Falls through to the textarea fallback below.
      }
      withHiddenTextarea((textarea) => {
        textarea.value = text;
        textarea.focus();
        textarea.select();
        document.execCommand("copy");
      });
    },
    async readText() {
      try {
        return await navigator.clipboard.readText();
      } catch {
        return withHiddenTextarea((textarea) => {
          textarea.focus();
          document.execCommand("paste");
          return textarea.value;
        });
      }
    },
    download(blob, filename) {
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = filename;
      link.style.display = "none";
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url);
    }
  };
}

/**
 * 从入口的 `snapshot` 装载存储模型，没给就建一张空表（ADR 0004 §3.1）。
 *
 * **入口不再吃对象数组**：那条路是 `rowsToWorkbookSnapshot`，一个独立的、
 * 永久的导入器（§3.2；xlsx 导入将来落在同一个位置）。入口只认快照这一种形态，
 * 于是「装载」与「把外部数据翻译成格子」是两件分开的事，各自可测。
 *
 * 空表的列数取 `columns.length` 而不是 0：坐标空间是派生的
 * （`maxOccupied + 1 + blank`），一张没有任何格子的表若不把列数记进
 * `blankColumns`，`columnCount()` 会答 0，紧接着的
 * `validateCyExcelCoreDataDimensions` 就会把「没给数据」当成配置错误拒掉。
 */
function buildWorkbook(
  options: CyExcelCoreOptions,
  columns: readonly CyExcelColumnConfig[]
): WorkbookModel {
  if (options.snapshot !== undefined) {
    return WorkbookModel.fromSnapshot(options.snapshot);
  }
  return WorkbookModel.createEmpty({
    sheet: { blankRows: 0, blankColumns: columns.length }
  });
}

/**
 * 建布局，并把「配置说了、模型还没记」的那部分几何补回模型。
 *
 * 顺序是有讲究的（`axisGeometry.ts` 里写了理由）：先算合流结果，再让
 * `new LayoutEngine` 去校验，**最后**才写模型——校验失败时模型一个字节都没动。
 * 轴长在写之前就取好，写完钉回去，免得补出来的侧表条目把派生轴撑长（spec §3.3）。
 */
function buildLayout(
  options: CyExcelCoreOptions,
  columns: readonly CyExcelColumnConfig[],
  sheet: SheetModel,
  dimensions: { readonly rowCount: number; readonly columnCount: number }
): LayoutEngine {
  const defaultColumnWidth = DEFAULT_COLUMN_WIDTH;
  const geometry = resolveInitialAxisGeometry({
    sheet,
    rowCount: dimensions.rowCount,
    columnCount: columns.length,
    ...(options.rowHeights !== undefined ? { configuredRowHeights: options.rowHeights } : {}),
    configuredColumnWidths: columns.map((column) => column.width),
    defaultColumnWidth
  });
  const layout = new LayoutEngine({
    rowCount: dimensions.rowCount,
    columnCount: columns.length,
    rowHeight: options.rowHeight ?? DEFAULT_ROW_HEIGHT,
    ...(geometry.rowHeights.size > 0 ? { rowHeights: geometry.rowHeights } : {}),
    columnWidths: geometry.columnWidths,
    defaultColumnWidth,
    ...(options.freeze !== undefined ? { freeze: options.freeze } : {}),
    ...(options.overscanRows !== undefined ? { overscanRows: options.overscanRows } : {}),
    ...(options.overscanColumns !== undefined
      ? { overscanColumns: options.overscanColumns }
      : {})
  });
  seedAxisMeta(sheet, geometry, dimensions);
  return layout;
}

type CoreInteractionFrame = InteractionFrame<
  SceneSnapshot,
  SurfaceSetSnapshot,
  SelectionSnapshot
>;

type CorePointerInputContext = InputCandidateContext<
  CoreInteractionFrame,
  PointerEvent,
  EventTarget
>;

/**
 * The integration layer: SheetModel -> TableStore -> CommandManager -> Patch
 * -> RenderScheduler/Canvas -> DOM editor overlay, all wired to a real
 * container. Unlike src/core/{data,commands,layout,render,interaction,
 * editors}/, this class genuinely creates and manipulates DOM/Canvas — see
 * doc/code-reference.md §2 for why that doesn't violate the "core doesn't
 * depend on HTMLElement" ADR (that ADR is about not requiring the Custom
 * Elements registry, not about never touching the DOM).
 */
export class CyExcelCore {
  readonly #scroller: HTMLDivElement;
  readonly #headerCanvas: HTMLCanvasElement;
  readonly #rowHeaderCanvas: HTMLCanvasElement;
  readonly #contentCanvas: HTMLCanvasElement;
  readonly #selectionCanvas: HTMLCanvasElement;
  readonly #inputSurfaceOwners = new WeakMap<EventTarget, SurfaceId>();
  readonly #surfaceHost: SurfaceHost;
  /** Owns wrap-height measurement and the measured-row cache for the current data/column structure epoch. Core keeps only the commit half. */
  readonly #measurement: WrapMeasurementService;

  #columns: readonly CyExcelColumnConfig[];
  readonly #cellTypeRegistry: CellTypeRegistry;
  readonly #columnManager: ColumnManager;
  readonly #headerManager: HeaderManager;
  /**
   * The last sort *applied* — never one that is merely queued or that the
   * storage layer refused; `#runSortByColumn` writes it only after the reorder
   * commits. `null` before any sort, and again after any column-structure
   * change (see `#publishCommitReceipt`). Not the data's live order:
   * editing a cell afterwards does not clear it (ADR 0003 §3.4, matching
   * Excel). Backs both the header's visual indicator (via `#headerState()`)
   * and the public `sortState` getter.
   */
  #sortState: SortState | null = null;
  readonly #columnHeaderHeight: number;
  readonly #rowHeaderWidth: number;
  /**
   * 表格数据的唯一权威（ADR 0004 §3.1）。这里只留 sheet 不留 workbook：
   * 侧表资源登记在 sheet 上，workbook 那层壳目前只为 `toSnapshot()` 收拢
   * 资源用，等 T7 的 `snapshot?` 入口落地时再把它一起持有。
   */
  readonly #sheet: SheetModel;
  readonly #store: TableStore;
  readonly #commandManager: CommandManager;
  readonly #layout: LayoutEngine;
  /**
   * 交给测量服务的行高来源视图。
   *
   * 读的是权威——`SheetModel` 的 `rowData[row].hp`，不是布局根上的索引：ADR 0004
   * 阶段 4B T8 之后布局层连一份来源的副本都不留（设计稿 §4.1/§4.2）。两个箭头函数
   * 都只在被调用时才碰 `#sheet`，所以放在字段初始化里是安全的。
   */
  readonly #rowHeightProvenance: RowHeightProvenanceView = {
    provenanceOf: (row: number): RowHeightProvenance => this.#sheet.getRowMeta(row)?.hp ?? "default",
    autoRows: (): Iterable<number> => this.#autoHeightRows()
  };
  #scene!: SceneSnapshot;
  /**
   * Publication identity, independent from the model/layout identities held
   * by a Scene's Skeleton. Gesture previews and authoritative projections
   * share this sequence, so no consumer can mistake different geometry for
   * the same Scene merely because both came from one committed root.
  */
  #sceneGeneration = 0;
  #publishedSceneGeneration = 0;
  readonly #renderScheduler: RenderScheduler;
  readonly #contentRenderer: ContentLayerRenderer;
  readonly #selectionRenderer: SelectionLayerRenderer;
  readonly #headerCanvasRenderer: HeaderCanvasRenderer;
  readonly #rowHeaderCanvasRenderer: RowHeaderCanvasRenderer;
  readonly #headerCanvasSurface: CanvasSurface;
  readonly #rowHeaderCanvasSurface: CanvasSurface;
  readonly #contentCanvasSurface: CanvasSurface;
  readonly #selectionCanvasSurface: CanvasSurface;
  readonly #domHeaderManager: DomHeaderManager;
  readonly #selectionModel: SelectionModel;
  readonly #structureOwner: GridStructureOwner;
  readonly #pointerController: PointerController;
  readonly #fillHandleController: FillHandleController;
  readonly #headerResizeController: HeaderResizeController;
  readonly #headerSelectionController: HeaderSelectionController;
  readonly #inputRouter: InputRouter<
    CoreInteractionFrame,
    PointerEvent,
    void,
    EventTarget
  >;
  readonly #rowReorderController: RowReorderController | undefined;
  readonly #keyboardController: KeyboardController;
  readonly #editorManager: EditorManager;
  readonly #domCellManager: DOMCellManager;
  readonly #clipboardService: ClipboardService;
  readonly #hooks: HookManager;
  readonly #events: EventBus;
  readonly #contextMenuService: ContextMenuService;
  readonly #pluginManager: PluginManager;
  readonly #shortcuts: ShortcutRegistry;
  readonly #scrollCommitPolicy: "commit" | "cancel";
  readonly #overscanRows: number;
  readonly #overscanColumns: number;

  /** Owns the body pointer gesture's kind, structure version and capture session. */
  readonly #bodyGestures: BodyGestureController;
  /** Owns every header pointer gesture: resize, header selection, row reorder and the reorder auto-scroll loop. */
  readonly #headerGestures: HeaderGestureController;
  /** The sole owner of the scroller's native offsets and the spacer's size. */
  readonly #viewport: ViewportManager;
  #autoHeightContinuationFrame: number | null = null;
  #finalizationDepth = 0;
  readonly #mutationIntents: DeferredMutationQueue;
  readonly #interactionFrames: InteractionFrameProvider<
    SceneSnapshot,
    SurfaceSetSnapshot,
    SelectionSnapshot
  >;
  /** Owns the structure version and the teardown that must accompany every bump. */
  readonly #structure: StructureTransition;
  /**
   * Whether the focus about to land on the content canvas came from a pointer.
   *
   * The canvas is `tabIndex = 0` and gets `.focus()` on every click, so without
   * this the browser paints its default focus ring around the *entire* canvas
   * — which reads as "the whole table is selected" sitting behind the one-cell
   * selection box. Suppressing the ring unconditionally would be worse: it is
   * currently the only thing distinguishing a focused grid from an unfocused
   * one, since the active-cell box looks identical either way. So the ring is
   * kept for keyboard focus (Tab) and dropped only for pointer focus.
   */
  #focusFromPointer = false;
  #disposed = false;

  constructor(container: HTMLElement, options: CyExcelCoreOptions = {}) {
    validateCyExcelCoreOptions(options);

    // Copied, not aliased — insertColumns()/deleteColumns()/moveColumns()
    // splice this array in place, so without this much copying a structural
    // column change would silently resize/reorder the caller's own
    // `options.columns` array.
    // （`options.snapshot` 不需要同样的保护：它只被读一遍装进存储模型。）
    this.#columns = snapshotColumnDefinitions(options.columns ?? []);
    this.#scrollCommitPolicy = options.scrollCommitPolicy ?? "commit";
    this.#overscanRows = options.overscanRows ?? 0;
    this.#overscanColumns = options.overscanColumns ?? 0;

    this.#headerManager = new HeaderManager({
      columns: this.#columns,
      columnLetterFallback: options.columnLetterFallback ?? true
    });
    this.#columnHeaderHeight = options.headerHeight ?? 0;
    this.#rowHeaderWidth = options.rowHeaderWidth ?? 0;

    const sheet = buildWorkbook(options, this.#columns).sheet();
    const initialDimensions = {
      rowCount: sheet.rowCount(),
      columnCount: sheet.columnCount()
    };
    validateCyExcelCoreDataDimensions(
      options,
      initialDimensions,
      this.#columns.length
    );
    this.#sheet = sheet;
    this.#layout = buildLayout(options, this.#columns, sheet, initialDimensions);

    const initialization = new InitializationRollback();
    try {
    const dom = createDom(
      container,
      this.#headerHeight(), // #headerManager/#columnHeaderHeight are both assigned above
      this.#rowHeaderWidth
    );
    initialization.defer(() => dom.scroller.remove());
    initialization.defer(() => {
      if (this.#autoHeightContinuationFrame !== null) {
        cancelAnimationFrame(this.#autoHeightContinuationFrame);
        this.#autoHeightContinuationFrame = null;
      }
    });
    this.#scroller = dom.scroller;
    this.#viewport = new ViewportManager({
      scroller: dom.scroller,
      spacer: dom.spacer,
      // Read per call rather than captured: the column header band's height
      // depends on how many group levels #headerManager actually resolved.
      getBandSizes: () => ({
        rowHeaderWidth: this.#rowHeaderWidth,
        columnHeaderHeight: this.#headerHeight()
      })
    });
    this.#headerCanvas = dom.headerCanvas;
    this.#rowHeaderCanvas = dom.rowHeaderCanvas;
    this.#contentCanvas = dom.contentCanvas;
    this.#selectionCanvas = dom.selectionCanvas;
    this.#inputSurfaceOwners.set(dom.contentCanvas, "body");
    this.#inputSurfaceOwners.set(dom.domCellLayer, "body");
    this.#inputSurfaceOwners.set(dom.editorLayer, "body");
    this.#inputSurfaceOwners.set(dom.headerCanvas, "column-header");
    this.#inputSurfaceOwners.set(dom.headerDomLayer, "column-header");
    this.#inputSurfaceOwners.set(dom.menuLayer, "menu-overlay");
    this.#surfaceHost = new SurfaceHost();
    initialization.defer(() => this.#surfaceHost.dispose());
    this.#surfaceHost.attach("body", this.#contentCanvas);
    this.#surfaceHost.attach("column-header", this.#headerCanvas, { zIndex: 1 });
    this.#surfaceHost.attach("row-header", this.#rowHeaderCanvas, {
      region: { x: 0, y: this.#headerHeight() },
      zIndex: 1
    });
    this.#surfaceHost.attach("corner", this.#rowHeaderCanvas, {
      region: {
        x: 0,
        y: 0,
        width: this.#rowHeaderWidth,
        height: this.#headerHeight()
      },
      zIndex: 2
    });
    this.#surfaceHost.attach("menu-overlay", this.#scroller, { zIndex: 3 });

    // Constructed early — before the rest of the grid — so that any
    // `options.plugins` entry tapping "beforeInit" (fired a few lines
    // below, once every plugin is installed) actually gets to run before
    // the heavy setup that follows. `getCell`/`copy`/`paste` below close
    // over `this.#store`/`this.#clipboardService`, which don't exist yet
    // at this point — safe, because PluginManager only stores these
    // functions during `use()`; nothing calls them until a plugin's own
    // registered Hook/listener/menu action actually runs, by which time
    // construction has finished.
    this.#hooks = new HookManager({ registry: PRODUCTION_HOOK_REGISTRY });
    initialization.defer(() => this.#hooks.dispose());
    this.#events = new EventBus();
    initialization.defer(() => this.#events.dispose());
    this.#shortcuts = new ShortcutRegistry();
    initialization.defer(() => this.#shortcuts.dispose());
    this.#contextMenuService = new ContextMenuService({
      container: dom.menuLayer,
      hooks: this.#hooks,
      restoreFocus: () => this.#contentCanvas.focus(),
      // Keeps a menu opened near the right/bottom edge fully within the
      // visible scroller instead of overflowing into (and being silently
      // clipped, including its click handling, by) the scroller's own
      // `overflow: auto` — see DefaultMenuRenderer's #clampToBounds() doc
      // comment for the real user report this fixes.
      getBounds: () => {
        const surface = this.#surfaceHost.refresh().get("menu-overlay");
        return surface
          ? new DOMRect(
              surface.clientOrigin.x,
              surface.clientOrigin.y,
              surface.cssSize.width,
              surface.cssSize.height
            )
          : null;
      },
      // Forwards through the same public "error" event `EventBus`'s own
      // listener-threw path uses (see EventBus.#reportError) — a menu
      // action throwing is otherwise only visible via console.error, which
      // a host app has no way to respond to it.
      onActionError: (itemId, error) => this.#events.emit("error", { error, eventName: `contextMenu:${itemId}` })
    });
    initialization.defer(() => this.#contextMenuService.dispose());
    this.#pluginManager = new PluginManager({
      hooks: this.#hooks,
      events: this.#events,
      contextMenu: this.#contextMenuService,
      getCell: (row, column) => this.#store.getCell(row, column),
      copy: (copyOptions) => this.#clipboardService.copy(copyOptions),
      paste: (pasteOptions) => this.#clipboardService.paste(pasteOptions),
      insertRows: (start, count) => this.insertRows(start, count),
      deleteRows: (start, count) => this.deleteRows(start, count),
      moveRows: (start, count, target) => this.moveRows(start, count, target),
      insertColumns: (start, columns) => this.insertColumns(start, columns),
      deleteColumns: (start, count) => this.deleteColumns(start, count),
      moveColumns: (start, count, target) => this.moveColumns(start, count, target),
      hideColumn: (column) => this.hideColumn(column),
      registerShortcut: (combo, handler) => this.registerShortcut(combo, handler)
    });
    initialization.defer(() => this.#pluginManager.dispose());
    for (const plugin of options.plugins ?? []) {
      this.#pluginManager.use(plugin);
    }
    this.#hooks.runSeries("beforeInit", {});

    // Built-ins register first (stack bottom), so a plugin/user combo
    // registered later via registerShortcut() on the same combo naturally
    // wins by default — see ShortcutRegistry's own doc comment for the
    // conflict policy. Registered here (before #editorManager/#selectionModel
    // exist) because these are just closures — nothing calls them until a
    // real keydown arrives, by which time construction has finished; same
    // "safe to reference not-yet-assigned fields in a deferred closure"
    // reasoning #pluginManager's getCell/copy/paste already relies on above.
    this.#shortcuts.register("Mod+Z", () => {
      this.undo();
    });
    this.#shortcuts.register("Mod+Shift+Z", () => {
      this.redo();
    });
    this.#shortcuts.register("Mod+Y", () => {
      this.redo();
    });
    this.#shortcuts.register("Enter", () => {
      const { active } = this.#selectionModel.snapshot;
      this.#editorManager.startEdit(active.row, active.column);
    });

    this.#cellTypeRegistry = new CellTypeRegistry();
    registerBuiltinCellTypes(this.#cellTypeRegistry);
    for (const [name, definition] of Object.entries(options.columnTypes ?? {})) {
      this.#cellTypeRegistry.register(name, definition);
    }
    this.#columnManager = new ColumnManager({
      columns: this.#columns,
      registry: this.#cellTypeRegistry,
      hooks: this.#hooks
    });

    this.#store = new TableStore(this.#sheet);
    initialization.defer(() => this.#store.coordinator.dispose());

    this.#updateSpacerSize();
    this.#refreshScene();

    const headerCtx = this.#headerCanvas.getContext("2d") as Canvas2DLike;
    const rowHeaderCtx = this.#rowHeaderCanvas.getContext("2d") as Canvas2DLike;
    const contentCtx = this.#contentCanvas.getContext("2d") as Canvas2DLike;
    const selectionCtx = this.#selectionCanvas.getContext("2d") as Canvas2DLike;
    this.#headerCanvasSurface = new CanvasSurface(this.#headerCanvas, headerCtx);
    this.#rowHeaderCanvasSurface = new CanvasSurface(this.#rowHeaderCanvas, rowHeaderCtx);
    this.#contentCanvasSurface = new CanvasSurface(this.#contentCanvas, contentCtx);
    this.#selectionCanvasSurface = new CanvasSurface(this.#selectionCanvas, selectionCtx);
    // Never appended to any DOM tree — exists purely so wrap-height
    // measurement has a real `measureText()` to call without touching
    // #contentCanvas's own drawing state.
    const measureCtx = document.createElement("canvas").getContext("2d") as Canvas2DLike;
    measureCtx.font = CELL_FONT;
    this.#measurement = new WrapMeasurementService({
      measureContext: measureCtx,
      getColumns: () => this.#columnManager.getColumns(),
      getCell: (row, column) => this.#store.getCell(row, column),
      onError: (error, eventName) => this.#events.emit("error", { error, eventName })
    });

    this.#contentRenderer = new ContentLayerRenderer({
      ctx: contentCtx,
      getScene: () => this.#scene,
      getCellValue: (row, column) => this.#store.getCell(row, column),
      painter: createColumnAwarePainter(() => this.#columnManager.getColumns()),
      onError: (error, cell) =>
        this.#events.emit("error", { error, eventName: `finalize:canvas-cell:${cell.row}:${cell.column}` }),
      ...(options.showGridLines !== undefined ? { showGridLines: options.showGridLines } : {})
    });
    this.#selectionRenderer = new SelectionLayerRenderer({
      ctx: selectionCtx,
      getScene: () => this.#scene,
      getSelection: () => this.#selectionModel.snapshot,
      isVisible: () => !this.#editorManager.isEditing(),
      getFillPreviewRange: () => this.#fillHandleController.previewRange
    });
    this.#headerCanvasRenderer = new HeaderCanvasRenderer({
      ctx: headerCtx,
      getScene: () => this.#scene,
      headerManager: this.#headerManager,
      getHeaderState: (column) => this.#headerState(column)
    });
    this.#rowHeaderCanvasRenderer = new RowHeaderCanvasRenderer({
      ctx: rowHeaderCtx,
      getScene: () => this.#scene,
      // Safe to reference #rowReorderController before it's assigned below
      // (only constructed when options.rowReorder is on) — same "closure
      // over a not-yet-assigned field" pattern getFillPreviewRange already
      // relies on a few lines down, for the same reason: nothing calls this
      // until the first real paint(), by which point construction is done.
      getReorderIndicatorRow: () => this.#rowReorderController?.indicatorRow ?? null
    });
    this.#domHeaderManager = new DomHeaderManager({
      container: dom.headerDomLayer,
      headerManager: this.#headerManager,
      getHeaderState: (column) => this.#headerState(column)
    });
    initialization.defer(() => this.#domHeaderManager.dispose());

    this.#renderScheduler = new RenderScheduler(
      {
        paintPresentation: (frame) => {
          const { presentation } = frame.dirty;
          if (presentation.domCells) {
            this.#runFinalizer("dom-cells", () => this.#domCellManager.sync(frame.scene));
          }
          if (presentation.columnHeader) {
            this.#runFinalizer("column-header-canvas", () =>
              this.#headerCanvasRenderer.paint(frame.scene)
            );
          }
          if (presentation.rowHeader) {
            this.#runFinalizer("row-header-canvas", () =>
              this.#rowHeaderCanvasRenderer.paint(frame.scene)
            );
          }
          if (presentation.domHeaders) {
            this.#runFinalizer("dom-headers", () => this.#domHeaderManager.sync(frame.scene));
          }
        },
        paintContent: (frame) =>
          this.#contentRenderer.paintContent(frame.dirty.content, frame.scene),
        paintSelection: (frame) => this.#selectionRenderer.paintSelection(frame.scene)
      },
      {
        frameScheduler: {
          requestFrame: (callback) => requestAnimationFrame(callback),
          cancelFrame: (handle) => cancelAnimationFrame(handle)
        },
        revisionGate: this.#store.coordinator.revisionGate,
        getScene: () => this.#scene,
        onFrameError: (error) =>
          this.#events.emit("error", { error, eventName: "finalize:render-frame" })
      }
    );
    initialization.defer(() => this.#renderScheduler.dispose());
    this.#mutationIntents = new DeferredMutationQueue({
      revisionGate: this.#store.coordinator.revisionGate,
      isGuarded: () => this.#finalizationDepth > 0,
      beforeRun: () => this.#headerGestures.cancelResizeOnly(true),
      onError: (error) => this.#events.emit("error", { error, eventName: "deferred-mutation" })
    });
    initialization.defer(() => this.#mutationIntents.dispose());

    this.#selectionModel = new SelectionModel({
      rowCount: initialDimensions.rowCount,
      columnCount: this.#columns.length
    });
    this.#interactionFrames = new InteractionFrameProvider({
      revisionGate: this.#store.coordinator.revisionGate,
      intentQueue: this.#mutationIntents,
      getScene: () => this.#scene,
      // Client origins can change after ancestor scroll/reflow without a
      // resize event. Measure once inside every short interaction lease.
      getSurfaces: () => this.#surfaceHost.refresh(),
      getSelection: () => this.#selectionModel.snapshot,
      getStructureVersion: () => this.#structure.version
    });
    this.#structureOwner = new GridStructureOwner({
      getColumns: () => this.#columns,
      replaceColumns: (columns) => {
        this.#columns = columns;
      },
      columnManager: this.#columnManager,
      headerManager: this.#headerManager,
      layout: this.#layout,
      selection: this.#selectionModel
    });
    this.#commandManager = createGridCommandManager({
      store: this.#store,
      layout: this.#layout,
      structureOwner: this.#structureOwner,
      hooks: this.#hooks
    });

    // Off by default. Without this candidate, row-header routing falls
    // through from resize directly to ordinary row selection.
    this.#rowReorderController = options.rowReorder
      ? new RowReorderController({
          onIndicatorChange: () =>
            this.#renderScheduler.invalidatePresentation({ rowHeader: true })
        })
      : undefined;

    this.#headerResizeController = new HeaderResizeController({
      getColumnLimits: (column) => {
        const definition = this.#columns[column];
        return definition
          ? {
              ...(definition.minWidth !== undefined ? { minSize: definition.minWidth } : {}),
              ...(definition.maxWidth !== undefined ? { maxSize: definition.maxWidth } : {})
            }
          : undefined;
      }
    });
    this.#headerSelectionController = new HeaderSelectionController();

    this.#fillHandleController = new FillHandleController({
      onFillPreviewChange: () => this.#renderScheduler.invalidateSelection()
    });

    this.#pointerController = new PointerController({
      selection: this.#selectionModel,
      onSelectionChange: () => this.#renderScheduler.invalidateSelection(),
      findMergeAt: (row, column) => findMergeAt(this.#sheet.merges(), row, column)
    });
    this.#headerGestures = new HeaderGestureController({
      resizeController: this.#headerResizeController,
      selectionController: this.#headerSelectionController,
      rowReorderController: this.#rowReorderController,
      selectionModel: this.#selectionModel,
      columnCaptureTarget: this.#headerCanvas,
      rowCaptureTarget: this.#rowHeaderCanvas,
      bodyCursorTarget: document.body,
      host: {
        withFrame: (eventName, run) => this.#dispatchInteraction(eventName, run),
        enqueueIntent: (run) => this.#interactionFrames.enqueueIntent(run),
        // Preparing a Layout root stays here: a gesture controller that
        // reached for LayoutEngine itself would rebuild the second coordinate
        // truth ADR 0002 removed.
        publishResizePreview: (preview) => {
          this.#publishResizePreview(
            preview.axis === "column"
              ? this.#layout.prepareSetColumnWidth(preview.column, preview.width)
              : this.#layout.prepareSetRowHeight(preview.row, preview.height)
          );
        },
        restoreAuthoritativeScene: () => {
          this.#withFinalization(() => this.#restoreAuthoritativeScene());
        },
        invalidateSelection: () => this.#renderScheduler.invalidateSelection(),
        invalidateRowHeader: () =>
          this.#renderScheduler.invalidatePresentation({ rowHeader: true }),
        scrollVerticallyBy: (delta) => {
          this.#viewport.scrollNativeBy(delta);
          this.#syncViewport();
          this.#renderScheduler.invalidateAll();
        },
        commitResize: (intent) => this.#commitHeaderResize(intent),
        commitReorder: (intent) => {
          if (intent.structureVersion === this.#structure.version) {
            this.#commitRowReorder(intent.start, intent.count, intent.target);
          }
        },
        updateHoverAfterGesture: (event, surface) => this.#updateHeaderHover(event, surface),
        reportError: (error, eventName) => this.#events.emit("error", { error, eventName })
      }
    });
    this.#structure = new StructureTransition({
      cancelActiveEdit: () => this.#cancelActiveEdit(),
      cancelBodyGesture: () => this.#bodyGestures.cancel(),
      cancelHeaderGesture: () => this.#headerGestures.cancel(true),
      closeContextMenu: () => this.#contextMenuService.close()
    });
    this.#bodyGestures = new BodyGestureController({
      pointerController: this.#pointerController,
      fillHandleController: this.#fillHandleController,
      captureTarget: this.#contentCanvas,
      focusTarget: this.#contentCanvas,
      withFrame: (eventName, run) => this.#dispatchInteraction(eventName, run),
      onFillCommit: (intent) => {
        this.#interactionFrames.enqueueIntent(() => {
          if (intent.structureVersion === this.#structure.version) {
            this.#commitFillNow(intent.sourceRange, intent.targetRange);
          }
        });
      },
      onError: (error, eventName) => this.#events.emit("error", { error, eventName })
    });
    this.#inputRouter = new InputRouter({
      surfacePolicy: (context) => this.#resolvePointerInputTarget(context)
    });
    initialization.defer(() => this.#inputRouter.dispose());
    this.#inputRouter.register({
      id: "body-fill-handle",
      surface: "body",
      priority: 300,
      route: (context) => this.#routeBodyFill(context)
    });
    this.#inputRouter.register({
      id: "body-canvas-cell",
      surface: "body",
      priority: 200,
      route: (context) => this.#routeBodyCanvasAction(context)
    });
    this.#inputRouter.register({
      id: "body-selection",
      surface: "body",
      priority: 100,
      route: (context) => {
        this.#startBodySelection(context);
        return handleInput(undefined);
      }
    });
    this.#inputRouter.register({
      id: "column-header-resize",
      surface: "column-header",
      priority: 300,
      route: (context) => this.#routeHeaderResize(context, "column-resize")
    });
    this.#inputRouter.register({
      id: "column-header-canvas-action",
      surface: "column-header",
      priority: 200,
      route: (context) => this.#routeHeaderCanvasAction(context)
    });
    this.#inputRouter.register({
      id: "column-header-selection",
      surface: "column-header",
      priority: 100,
      route: (context) => this.#routeHeaderSelection(context, "column-selection")
    });
    this.#inputRouter.register({
      id: "row-header-resize",
      surface: "row-header",
      priority: 300,
      route: (context) => this.#routeHeaderResize(context, "row-resize")
    });
    this.#inputRouter.register({
      id: "row-header-reorder",
      surface: "row-header",
      priority: 200,
      route: (context) => this.#routeRowReorder(context)
    });
    this.#inputRouter.register({
      id: "row-header-selection",
      surface: "row-header",
      priority: 100,
      route: (context) => this.#routeHeaderSelection(context, "row-selection")
    });
    this.#inputRouter.register({
      id: "corner-selection",
      surface: "corner",
      priority: 100,
      route: (context) => this.#routeHeaderSelection(context, null)
    });

    this.#keyboardController = new KeyboardController({
      selection: this.#selectionModel,
      onSelectionChange: () => this.#renderScheduler.invalidateSelection(),
      findMergeAt: (row, column) => findMergeAt(this.#sheet.merges(), row, column)
    });

    this.#editorManager = new EditorManager({
      container: dom.editorLayer,
      getScene: () => this.#scene,
      findMergeAt: (row, column) => findMergeAt(this.#sheet.merges(), row, column),
      ...(options.editorFactory !== undefined ? { editorFactory: options.editorFactory } : {}),
      getCellValue: (row, column) => this.#store.getCell(row, column),
      commitValue: (row, column, value) =>
        this.#enqueueMutationPromise(() => {
          this.#applyCommandNow(new SetCellCommand(row, column, this.#parseEditedValue(row, column, value)));
        }),
      // Pure forwarding to whatever the column's own CellType declares —
      // EditorManager itself never decides which editor a column gets, and
      // neither does this integration layer; a column that doesn't declare
      // createEditor()/serialize() just falls back to EditorManager's own
      // defaults (see doc/progress.md's editor-pluggability patch).
      resolveEditor: (row, column) => {
        const runtimeColumn = this.#columnManager.getColumn(column);
        return runtimeColumn.createEditor?.({ row, column: runtimeColumn });
      },
      serializeCellValue: (row, column, value) => {
        const runtimeColumn = this.#columnManager.getColumn(column);
        return runtimeColumn.serialize
          ? runtimeColumn.serialize(value, { row, column: runtimeColumn })
          : String(value ?? "");
      },
      // The three transitions the selection box's isVisible() (above) needs
      // to actually take effect at the right moment — without these, the
      // box only hides/reappears whenever some unrelated repaint happens to
      // fire next, not exactly when editing starts or ends.
      onEditStart: () => this.#renderScheduler.invalidateSelection(),
      onEditCommit: () => {
        this.#selectionModel.moveActive(1, 0);
        this.#renderScheduler.invalidateSelection();
        this.#scrollActiveIntoView();
        // editor.destroy() (already run by this point) removes the
        // editor's DOM node — if it held keyboard focus, the browser drops
        // focus to document.body once it's gone. Without reclaiming focus
        // here, #handleKeyDown (bound to contentCanvas) stops receiving any
        // further Enter/arrow-key events after the very first edit cycle.
        this.#contentCanvas.focus();
      },
      onEditCancel: () => {
        this.#renderScheduler.invalidateSelection();
        this.#contentCanvas.focus();
      }
    });
    initialization.defer(() => this.#editorManager.dispose());

    this.#domCellManager = new DOMCellManager({
      container: dom.domCellLayer,
      getColumn: (column) => this.#columnManager.getColumn(column),
      getCellValue: (row, column) => this.#store.getCell(row, column),
      commitValue: (row, column, value) => {
        void this.#enqueueMutationPromise(() => {
          this.#applyCommandNow(new SetCellCommand(row, column, this.#parseEditedValue(row, column, value)));
        }).catch((error: unknown) => this.#events.emit("error", { error, eventName: "deferred-dom-commit" }));
      }
    });
    initialization.defer(() => this.#domCellManager.dispose());

    this.#clipboardService = new ClipboardService({
      getCellValue: (row, column) => this.#store.getCell(row, column),
      getColumn: (column) => this.#columnManager.getColumn(column),
      getRowCount: () => this.#sheet.rowCount(),
      getColumnCount: () => this.#columns.length,
      parseValue: (row, column, text) => this.#parseEditedValue(row, column, text),
      getSelectionRange: () => this.#selectionModel.snapshot.range,
      getStoreVersion: () => this.#store.version,
      applyRangeCommand: (range, values) =>
        this.#enqueueMutationPromise(() => this.#applyCommandNow(new SetRangeCommand(range, values))),
      io: createBrowserClipboardIO(),
      // `Worker`'s real `onmessage`/`onerror` setters are typed against
      // `MessageEvent`, narrower than the structural `SerializeWorkerLike`
      // ClipboardWorkerClient depends on (see that file's doc comment) —
      // the cast is safe because ClipboardWorkerClient only ever reads
      // `event.data`, which a real MessageEvent always has.
      createWorker: () =>
        new Worker(new URL("./clipboard/worker/serializeWorker.ts", import.meta.url), {
          type: "module"
        }) as unknown as SerializeWorkerLike,
      ...(options.clipboard !== undefined ? { options: options.clipboard } : {})
    });
    initialization.defer(() => this.#clipboardService.dispose());

    this.#resizeCanvases();
    this.#syncViewport();
    this.#renderScheduler.invalidateAll();

    initialization.defer(() => {
      this.#contentCanvas.removeEventListener("dblclick", this.#handleDblClick);
      this.#contentCanvas.removeEventListener("keydown", this.#handleKeyDown);
      this.#scroller.removeEventListener("pointerdown", this.#handleDelegatedPointerDown, true);
      this.#contentCanvas.removeEventListener("copy", this.#handleCopyEvent);
      this.#contentCanvas.removeEventListener("paste", this.#handlePasteEvent);
      this.#scroller.removeEventListener("scroll", this.#handleScroll);
      this.#scroller.removeEventListener("contextmenu", this.#handleContextMenu);
      this.#headerCanvas.removeEventListener("pointermove", this.#handleHeaderPointerHover);
      this.#rowHeaderCanvas.removeEventListener("pointermove", this.#handleRowHeaderPointerHover);
    });
    this.#contentCanvas.addEventListener("dblclick", this.#handleDblClick);
    this.#contentCanvas.addEventListener("keydown", this.#handleKeyDown);
    this.#contentCanvas.addEventListener("focus", this.#handleContentCanvasFocus);
    this.#contentCanvas.addEventListener("blur", this.#handleContentCanvasBlur);
    this.#scroller.addEventListener("pointerdown", this.#handleDelegatedPointerDown, true);
    this.#contentCanvas.addEventListener("copy", this.#handleCopyEvent);
    this.#contentCanvas.addEventListener("paste", this.#handlePasteEvent);
    this.#scroller.addEventListener("scroll", this.#handleScroll);
    // Attached to the scroller (not contentCanvas) so a right-click over
    // the header band — outside contentCanvas's own bounding box, see
    // createDom's bodyTop offset — can still open a "column-header"-target
    // menu, not just "cell"/"blank".
    this.#scroller.addEventListener("contextmenu", this.#handleContextMenu);

    // A 0-height header canvas (the default — see CyExcelCoreOptions.headerHeight)
    // never receives pointer events, so this is a no-op unless headers are enabled.
    this.#headerCanvas.addEventListener("pointermove", this.#handleHeaderPointerHover);
    // Same story for a 0-width row header canvas — see CyExcelCoreOptions.rowHeaderWidth.
    this.#rowHeaderCanvas.addEventListener("pointermove", this.#handleRowHeaderPointerHover);

    this.#pluginManager.activate();
    this.#events.emit("ready", undefined);
    initialization.commit();
    } catch (error) {
      initialization.rollback();
      throw error;
    }
  }

  get disposed(): boolean {
    return this.#disposed;
  }

  get hooks(): HookManager {
    return this.#hooks;
  }

  get events(): EventBus {
    return this.#events;
  }

  /** The immutable Scene currently published to paint, overlay and hit-test consumers. */
  get scene(): SceneSnapshot {
    return this.#scene;
  }

  /** The immutable selection snapshot currently shared by paint and input. */
  get selection(): SelectionSnapshot {
    return this.#selectionModel.snapshot;
  }

  get contextMenu(): ContextMenuService {
    return this.#contextMenuService;
  }

  /**
   * The most recent sort that was actually applied — via header click or
   * direct API call, both update this the same way — or `null` if no sort
   * has ever been applied. A sort that is still queued, that was cancelled
   * by its `AbortSignal`, or that the storage layer refused never appears
   * here. Reflects the *last sort applied*, not the data's current order:
   * editing a cell afterwards does not clear it, same one-shot indicator
   * semantics as Excel (ADR 0003 §3.4). A column insert/delete/move clears
   * it, because after the column axis moves the claim can no longer be made
   * about any particular index.
   */
  get sortState(): SortState | null {
    return this.#sortState;
  }

  /**
   * Registers a keyboard shortcut (`"Mod+Shift+Z"`-style combo — see
   * ShortcutRegistry.parseShortcut()). The built-in Mod+Z/Mod+Shift+Z/
   * Mod+Y/Enter shortcuts are registered the same way at construction, so a
   * combo registered here for the same combo overrides them by default
   * (return `false` from `handler` to fall through to the older
   * registration instead — see ShortcutRegistry's conflict policy).
   */
  registerShortcut(combo: string, handler: ShortcutHandler): Disposer {
    return this.#shortcuts.register(combo, handler);
  }

  getCell(row: number, column: number): unknown {
    return this.#store.getCell(row, column);
  }

  getColumn(column: number): RuntimeColumn {
    return this.#columnManager.getColumn(column);
  }

  getHeader(column: number): RuntimeHeader {
    return this.#headerManager.getHeader(column);
  }

  setCell(row: number, column: number, value: unknown): void {
    this.#enqueueOrRunMutation(() =>
      this.#applyCommandNow(new SetCellCommand(row, column, value))
    );
  }

  /**
   * Merges `range` into one visual cell. Overlap with existing merges follows
   * `resolveMergeCells` (ADR 0004 merge-cells design §0.2/§2): a range that
   * only partially overlaps an existing merge throws
   * `MergeOverlapsExistingMerge`; one that fully contains an existing merge
   * absorbs it. Every non-anchor cell in `range` is cleared (its prior value
   * is restored on undo) — the anchor (`range`'s top-left cell) keeps its
   * value. No geometry changes, so unlike the structural row/column methods
   * this does not cancel an in-progress edit — same as `setCell()`.
   */
  mergeCells(range: CellRange): void {
    this.#enqueueOrRunMutation(() => this.#applyCommandNow(new MergeCellsCommand(range)));
  }

  /**
   * Un-merges every merged range that intersects `range` — not just one
   * exactly equal to it, matching Excel. Only the merge table changes; cell
   * values already cleared by the original merge are not restored.
   */
  unmergeCells(range: CellRange): void {
    this.#enqueueOrRunMutation(() => this.#applyCommandNow(new UnmergeCellsCommand(range)));
  }

  /**
   * Inserts `count` blank rows (every cell `undefined`) at `start`. To
   * insert pre-filled rows, insert blank ones then setRange() over them.
   * The SheetModel edit and the matching Layout/Selection roots are
   * prepared together; live owners are changed only by the narrow commit.
   * New rows are always default-height (no third `heights` argument here);
   * a row wanting a custom height gets one via `setRowHeight()` afterward.
   */
  insertRows(start: number, count = 1): void {
    this.#enqueueOrRunMutation(() => {
      validateAxisChange({ axis: "row", operation: "insert", start, count }, this.#sheet.rowCount());
      this.#cancelActiveEdit();
      const rows = Array.from({ length: count }, () => new Array<unknown>(this.#columns.length).fill(undefined));
      this.#applyCommandNow(new InsertRowsCommand(start, rows));
    });
  }

  /** Deletes rows while the integration participant captures inverse layout state during prepare. */
  deleteRows(start: number, count = 1): void {
    this.#enqueueOrRunMutation(() => {
      validateAxisChange({ axis: "row", operation: "delete", start, count }, this.#sheet.rowCount());
      this.#cancelActiveEdit();
      this.#applyCommandNow(new DeleteRowsCommand(start, count));
    });
  }

  /** `target` is the insertion index in the pre-move index space (same convention `SheetModel.moveRows()` uses). */
  moveRows(start: number, count: number, target: number): void {
    this.#enqueueOrRunMutation(() => {
      validateAxisChange({ axis: "row", operation: "move", start, count, target }, this.#sheet.rowCount());
      this.#cancelActiveEdit();
      this.#applyCommandNow(new MoveRowsCommand(start, count, target));
    });
  }

  /**
   * Physically reorders every row by `column`. The data permutation and the
   * row geometry (each row's height and height-provenance) land in the same
   * transaction, so no published frame can show sorted rows against unsorted
   * heights, and one undo puts both back.
   *
   * Non-empty values are ordered by the column's cell type's `compare` when
   * it defines one, and by `createDefaultSortCompare`'s number-or-string
   * fallback otherwise. Empty cells (`null`/`undefined`) sort last in both
   * directions either way — see `createDefaultSortCompare`.
   *
   * Returns whether the sort was **applied before this call returned**. It
   * is `false` in two distinguishable-by-context cases:
   *
   * - `options.signal` aborted before the order was applied. A cancelled
   *   sort writes nothing, so there is nothing to undo.
   * - The call arrived while the mutation queue was busy — from inside an
   *   event listener or a header-click intent, say — and was queued behind
   *   the in-flight commit. The sort still runs, just later; watch
   *   `sortState` or the `change` event for it, and note that a queued sort
   *   that the storage layer then refuses reports through the `error` event
   *   rather than throwing back here.
   *
   * Either way `false` means "nothing had been applied yet when this
   * returned", never "the indicator moved anyway".
   */
  sortByColumn(
    column: number,
    direction: "asc" | "desc",
    options: { signal?: AbortSignal } = {}
  ): boolean {
    // Resolved eagerly so an out-of-range column throws from the call itself.
    this.#columnManager.getColumn(column);
    let applied = false;
    this.#enqueueOrRunMutation(() => {
      applied = this.#runSortByColumn(column, direction, options.signal);
    });
    return applied;
  }

  /**
   * The sort itself, run inline in whatever mutation slot the caller is
   * already standing in — `sortByColumn`'s queued closure, or the header
   * click intent, which runs during the mutation queue's own drain and so
   * can never get a second slot of its own.
   *
   * `#sortState` and the indicator repaint happen *after* `#applyCommandNow`
   * returns. A refused reorder throws out of here with
   * the indicator still describing the last sort that really was applied —
   * the ordering is the whole point, not an accident of layout.
   */
  #runSortByColumn(column: number, direction: "asc" | "desc", signal?: AbortSignal): boolean {
    const runtimeColumn = this.#columnManager.getColumn(column);
    const columnCompare = runtimeColumn.compare;
    const compare = createDefaultSortCompare(
      direction,
      columnCompare === undefined
        ? undefined
        // Invoked *on* the column, the way `column.format()` is called in
        // createColumnAwarePainter: a comparator written as an object method
        // may read `this.options` (rank by the column's configured order,
        // say), and detaching it would break that only when sorted.
        : (left, right) => columnCompare.call(runtimeColumn, left, right)
    );
    // Computed inside the mutation slot rather than before it: if the call
    // was deferred, the permutation has to be a permutation of the row count
    // that will actually be there when ReorderRowsCommand validates it.
    const order = computeSortOrder({
      rowCount: this.#store.getRowCount(),
      getSortKey: (row) => this.#store.getCell(row, column),
      compare,
      direction,
      ...(signal === undefined ? {} : { signal })
    });
    if (order === null) {
      return false;
    }
    // A sort moves the row under the editor out from under it.
    this.#cancelActiveEdit();
    this.#applyCommandNow(new ReorderRowsCommand(order));
    // Updated regardless of call site (header click or direct API call): the
    // indicator reflects the *last sort applied*, not how it was triggered.
    this.#sortState = { column, direction };
    this.#renderScheduler.invalidatePresentation({ columnHeader: true, domHeaders: true });
    return true;
  }

  /** Sets an exact row height through Command/Mutation/History. */
  setRowHeight(row: number, height: number): void {
    this.#enqueueOrRunMutation(() => {
      if (
        row >= 0 &&
        row < this.#layout.rowCount &&
        this.#layout.getRowHeight(row) === height &&
        this.#rowHeightProvenance.provenanceOf(row) === "manual"
      ) {
        return;
      }
      this.#cancelActiveEdit();
      this.#applyCommandNow(new ResizeRowCommand(row, height));
    });
  }

  setColumnWidth(column: number, width: number): void {
    this.#enqueueOrRunMutation(() => {
      if (
        column >= 0 &&
        column < this.#layout.columnCount &&
        // 与权威比而不是与索引比：隐藏列的几何宽度是 0，拿它比会把「把隐藏列的恢复
        // 宽度设成它已经是的那个值」当成一次真变更（ADR 0004 阶段 4B T8b）。
        columnWidthOf(this.#sheet.getColumnMeta(column), this.#layout.defaultColumnWidth) === width
      ) {
        return;
      }
      this.#cancelActiveEdit();
      this.#applyCommandNow(new ResizeColumnCommand(column, width));
    });
  }

  /** Measures whatever rows are about to be painted (already-visible + overscan, per the Scene's own body panes) that haven't been measured before — called from `#syncViewport()`, before painting. A no-op scan when the table has no wrap columns at all, and a no-op per-row once a row has been measured once (until an actual edit re-triggers it via `#publishCommitReceipt()`'s change-driven cache invalidation). */
  #measureVisibleWrapRows(
    scene: SceneSnapshot,
    remainingFollowUpPasses: number
  ): boolean {
    if (!this.#measurement.hasWrapColumns()) {
      return false;
    }
    return this.#measureAndResizeRows(
      new Set(this.#measurement.visibleUnmeasuredRows(scene)),
      scene,
      remainingFollowUpPasses
    );
  }

  /**
   * Inserts full column definitions at `start` — unlike `insertRows()`,
   * there's no reasonable default to synthesize: a column needs at least an
   * `id` (used as the `field` key unless overridden), and `insertRows()`'s
   * "every cell undefined" default has no column-shaped equivalent. See the
   * Step 17 plan's "key design decisions" §4.
   */
  insertColumns(start: number, columns: readonly CyExcelColumnConfig[]): void {
    const inserted = snapshotColumnDefinitions(columns);
    this.#enqueueOrRunMutation(() => {
      validateAxisChange({ axis: "column", operation: "insert", start, count: inserted.length }, this.#columns.length);
      this.#cancelActiveEdit();
      this.#applyCommandNow(new InsertColumnsCommand(start, inserted));
    });
  }

  deleteColumns(start: number, count = 1): void {
    this.#enqueueOrRunMutation(() => {
      validateAxisChange({ axis: "column", operation: "delete", start, count }, this.#columns.length);
      this.#cancelActiveEdit();
      this.#applyCommandNow(new DeleteColumnsCommand(start, count));
    });
  }

  /** `target` is the insertion index in the pre-move index space (same convention `SheetModel.moveColumns()` uses). */
  moveColumns(start: number, count: number, target: number): void {
    this.#enqueueOrRunMutation(() => {
      validateAxisChange({ axis: "column", operation: "move", start, count, target }, this.#columns.length);
      this.#cancelActiveEdit();
      this.#applyCommandNow(new MoveColumnsCommand(start, count, target));
    });
  }

  hideColumn(column: number): void {
    this.#enqueueOrRunMutation(() => {
      this.#cancelActiveEdit();
      this.#applyCommandNow(new HideColumnCommand(column));
    });
  }

  showColumn(column: number): void {
    this.#enqueueOrRunMutation(() => {
      this.#cancelActiveEdit();
      this.#applyCommandNow(new ShowColumnCommand(column));
    });
  }

  undo(): void {
    this.#enqueueOrRunMutation(() => this.#publishCommitReceipt(this.#commandManager.undo()));
  }

  redo(): void {
    this.#enqueueOrRunMutation(() => this.#publishCommitReceipt(this.#commandManager.redo()));
  }

  isEditing(): boolean {
    return this.#editorManager.isEditing();
  }

  startEdit(row: number, column: number): void {
    this.#editorManager.startEdit(row, column);
  }

  commitEdit(): Promise<void> {
    return this.#editorManager.commitEdit();
  }

  cancelEdit(): void {
    this.#editorManager.cancelEdit();
  }

  /** Cheap, synchronous size/complexity estimate — never constructs a full serialized string. `range` defaults to the current selection. */
  planCopy(range?: CellRange): CopyPlan {
    return this.#clipboardService.planCopy(range);
  }

  /**
   * The programmatic copy API — goes through the async Clipboard API (with
   * a hidden-textarea fallback) and, for large ranges, a Worker. Real
   * Ctrl+C keypresses are handled separately by #handleCopyEvent, which
   * needs to stay synchronous to satisfy the browser's native `copy` event
   * contract.
   */
  copy(options?: ClipboardCopyOptions): Promise<ClipboardCopyResult> {
    return this.#clipboardService.copy(options);
  }

  paste(options?: ClipboardPasteOptions): Promise<void> {
    return this.#clipboardService.paste(options);
  }

  /** Re-reads the container's current size and repaints. No automatic ResizeObserver — call this explicitly. */
  resize(): void {
    this.#enqueueOrRunMutation(() => {
      this.#headerGestures.cancel(true);
      this.#resizeCanvases();
      this.#syncViewport();
      this.#renderScheduler.invalidateAll();
    });
  }

  dispose(): void {
    if (this.#disposed) {
      return;
    }
    this.#headerGestures.cancel(false);
    this.#bodyGestures.cancel();
    this.#disposed = true;
    this.#mutationIntents.dispose();
    if (this.#autoHeightContinuationFrame !== null) {
      cancelAnimationFrame(this.#autoHeightContinuationFrame);
      this.#autoHeightContinuationFrame = null;
    }

    this.#hooks.runSeries("beforeDestroy", {});
    this.#events.emit("destroy", undefined);

    this.#pluginManager.dispose();
    this.#contextMenuService.dispose();
    this.#editorManager.dispose();
    this.#domCellManager.dispose();
    this.#domHeaderManager.dispose();
    this.#renderScheduler.dispose();
    this.#clipboardService.dispose();
    this.#store.coordinator.dispose();
    this.#inputRouter.dispose();

    this.#contentCanvas.removeEventListener("dblclick", this.#handleDblClick);
    this.#contentCanvas.removeEventListener("keydown", this.#handleKeyDown);
    this.#contentCanvas.removeEventListener("focus", this.#handleContentCanvasFocus);
    this.#contentCanvas.removeEventListener("blur", this.#handleContentCanvasBlur);
    this.#scroller.removeEventListener("pointerdown", this.#handleDelegatedPointerDown, true);
    this.#contentCanvas.removeEventListener("copy", this.#handleCopyEvent);
    this.#contentCanvas.removeEventListener("paste", this.#handlePasteEvent);
    this.#scroller.removeEventListener("scroll", this.#handleScroll);
    this.#scroller.removeEventListener("contextmenu", this.#handleContextMenu);
    this.#headerCanvas.removeEventListener("pointermove", this.#handleHeaderPointerHover);
    this.#rowHeaderCanvas.removeEventListener("pointermove", this.#handleRowHeaderPointerHover);

    this.#hooks.dispose();
    this.#events.dispose();
    this.#shortcuts.dispose();
    this.#surfaceHost.dispose();

    this.#scroller.remove();
  }

  #applyCommandNow(command: Command): void {
    this.#publishCommitReceipt(this.#commandManager.execute(command));
  }

  /** User renderers may dispatch commands; replay them after the current receipt is fully observed. */
  #enqueueOrRunMutation(run: () => void): void {
    this.#mutationIntents.enqueueOrRun(run);
  }

  #enqueueMutationPromise(run: () => void): Promise<void> {
    return this.#mutationIntents.enqueuePromise(run);
  }

  #dispatchInteraction<TResult>(
    eventName: string,
    handler: (
      frame: InteractionFrame<SceneSnapshot, SurfaceSetSnapshot, SelectionSnapshot>
    ) => TResult
  ): TResult | undefined {
    try {
      return this.#interactionFrames.dispatch(handler);
    } catch (error) {
      this.#events.emit("error", { error, eventName: `interaction:${eventName}` });
      return undefined;
    }
  }

  #withFinalization(run: () => void): void {
    this.#finalizationDepth += 1;
    try {
      run();
    } finally {
      this.#finalizationDepth -= 1;
      if (this.#finalizationDepth === 0) {
        this.#mutationIntents.drain();
      }
    }
  }

  /**
   * Structural row operations unconditionally cancel an in-progress edit
   * rather than trying to figure out whether the edited row specifically
   * survives — insert/move can shift which row a still-open editor is
   * anchored to just as easily as delete can remove it outright, and
   * "always cancel" is a simpler, corruption-proof default than tracking
   * whether the edit happens to still be valid after the structural
   * change.
   */
  #cancelActiveEdit(): void {
    if (this.#editorManager.isEditing()) {
      this.#editorManager.cancelEdit();
    }
  }

  /**
   * The built-in text editor (and any custom editor that just returns
   * whatever the user typed) always produces a raw string — nothing
   * converted it to the column's actual value type before this fix, so
   * committing "500" into a `type: "number"`/`"money"` column stored the
   * *string* "500". `CellType.format()` for those built-ins requires
   * `typeof value === "number"` and returns "" otherwise (see
   * src/core/cells/builtins.ts), so the cell would silently render blank
   * after every edit — a real, data-corrupting gap between Step 7's editor
   * and Step 8's `CellType.parse()`, which existed but was never called
   * anywhere. Only strings get parsed: a custom editor or DOM cell renderer
   * that already returns a properly-typed value (e.g. a boolean from a
   * checkbox) passes straight through unchanged.
   */
  #parseEditedValue(row: number, column: number, value: unknown): unknown {
    if (typeof value !== "string") {
      return value;
    }
    const runtimeColumn = this.#columnManager.getColumn(column);
    return runtimeColumn.parse ? runtimeColumn.parse(value, { row, column: runtimeColumn }) : value;
  }

  /**
   * Runs the post-commit receipt: structure transition, derived geometry,
   * Scene publication, measurement and the change event.
   *
   * Named `#invalidatePatches` until Phase 8 — a name from when it derived
   * render invalidation by walking the Patch list itself. It no longer
   * interprets changes at all: `compileDirtySet()` owns the dirty derivation,
   * `summarizeStructuralChanges()` the axis facts, and
   * `WrapMeasurementService` the measurement domain. What is left is
   * composition, which §12 explicitly allows CyExcelCore to hold.
   */
  #publishCommitReceipt(
    changes: readonly SemanticChange[] | undefined,
    measurementPassesRemaining = AUTO_HEIGHT_MAX_SYNC_PASSES
  ): void {
    if (!changes) {
      return;
    }
    // Capture the receipt-derived facts before any finalizer runs. A DOM
    // renderer or auto-height change may synchronously dispatch another
    // command and advance CommandManager.lastChangeSet reentrantly.
    const changeSet =
      this.#commandManager.lastChangeSet?.modelVersion === this.#store.version
        ? this.#commandManager.lastChangeSet
        : createChangeSet(this.#store.version, changes);
    // One pass over the change list instead of the three this method used to
    // make, and the measurement cache's own invalidation rules now live with
    // the cache rather than being spelled out here.
    const summary = summarizeStructuralChanges(changeSet.changes);
    const { hasRowStructure, hasColumnStructure, hasColumnLayout, hasRowLayout } = summary;
    const hasAxisChange = hasRowStructure || hasColumnStructure || hasColumnLayout;
    this.#withFinalization(() => {
      this.#measurement.invalidateFor(changeSet.changes);
      if (hasRowStructure || hasColumnStructure) {
        this.#structure.apply();
      }
      if (hasColumnStructure && this.#sortState !== null) {
        // Every other column index in the codebase is remapped through
        // `remapColumnIndex` on insert/delete/move; `#sortState.column` is
        // *cleared* instead. The indicator is one-shot — it claims "the last
        // sort applied", not "the data satisfies this order" (ADR 0003
        // §3.4) — and once the column axis moves underneath it, that claim
        // is no longer safe to make at all: the sorted column may be gone
        // entirely, and a stale index would both draw ▲ on the wrong header
        // and start a different column's first click at "desc".
        this.#sortState = null;
        this.#renderScheduler.invalidatePresentation({ columnHeader: true, domHeaders: true });
      }

      // The model revision and all authoritative structural roots are already
      // committed. Derived bookkeeping and user formatters are finalizers too:
      // failures are observable, but can neither escape nor suppress change.
      if (hasAxisChange || hasRowLayout) {
        this.#runFinalizer("spacer", () => this.#updateSpacerSize());
        this.#runFinalizer("selection-invalidation", () => this.#renderScheduler.invalidateSelection());
      }

      // Publish the exact receipt Scene before invoking formatter/user code.
      // Measurement, hit-test and every later presentation consumer can now
      // refer to the same model/layout basis instead of the previous receipt.
      let receiptScene: SceneSnapshot | undefined;
      const sceneReady = this.#runFinalizer("scene", () => {
        receiptScene = this.#refreshScene(this.#readViewport());
      });
      if (sceneReady && receiptScene) {
        this.#reprojectEditor(receiptScene);

        const measurement = this.#measurement.rowsToMeasureFor(receiptScene, {
          hasColumnStructureChange: hasColumnStructure,
          rowHeightProvenance: this.#rowHeightProvenance,
          defaultRowHeight: this.#layout.rowHeight
        });
        const rowResizeScheduled = measurementPassesRemaining > 0 && measurement.rows.size > 0
          ? this.#measureAndResizeRows(
              measurement.rows,
              receiptScene,
              measurementPassesRemaining - 1
            )
          : false;

        // One compiled invalidation for the whole receipt rather than one per
        // Patch: coalescing belongs to compileDirtySet(), and the renderer no
        // longer needs to know the model's Patch vocabulary at all.
        this.#runFinalizer("render-invalidation", () =>
          this.#renderScheduler.invalidateDirtySet(compileDirtySet(changeSet.changes))
        );
        if (!rowResizeScheduled) {
          // Presentation renderers are observers of this committed receipt.
          // Flush while finalization is guarded so renderer commands queue
          // until change publication; a derived row-height receipt instead
          // flushes both receipts against its final Scene.
          this.#runFinalizer("render-frame", () => this.#renderScheduler.flushNow());
        }
        if (measurementPassesRemaining === 0 && measurement.visibleUnmeasuredCount > 0) {
          this.#scheduleAutoHeightContinuation();
        }
      }
      // `changeSet` alone: it is already the frozen, model-versioned form of
      // the same facts, so emitting a second `patches` array only gave
      // consumers two names for one thing and two chances to disagree.
      this.#events.emit("change", { changeSet });
    });
  }

  #handleContentCanvasFocus = (): void => {
    this.#contentCanvas.style.outline = this.#focusFromPointer ? "none" : "";
  };

  #handleContentCanvasBlur = (): void => {
    this.#focusFromPointer = false;
    this.#contentCanvas.style.outline = "";
  };

  /** 来源为 `auto` 的行。只走稀疏表里真实存在的条目，不扫坐标空间。 */
  *#autoHeightRows(): IterableIterator<number> {
    for (const [row, meta] of this.#sheet.rowMetaEntries()) {
      if (meta.hp === "auto") {
        yield row;
      }
    }
  }

  #runFinalizer(name: string, run: () => void): boolean {
    try {
      run();
      return true;
    } catch (error) {
      this.#events.emit("error", { error, eventName: `finalize:${name}` });
      return false;
    }
  }

  /** Measures auto-owned rows via the measurement service, then queues one skip-history Command for the batch. Core owns only the commit half: which command, which history policy, and what happens when the basis moved underneath a formatter. */
  #measureAndResizeRows(
    rows: ReadonlySet<number>,
    basisScene: SceneSnapshot,
    remainingFollowUpPasses: number
  ): boolean {
    const basisModelVersion = this.#store.version;
    const basisLayoutRoot = this.#layout.root;
    const plan = this.#measurement.planUpdates(rows, {
      scene: basisScene,
      modelVersion: basisModelVersion,
      layoutVersion: basisLayoutRoot.version,
      defaultRowHeight: this.#layout.rowHeight,
      rowHeightProvenance: this.#rowHeightProvenance
    });
    if (plan.updates.length === 0) {
      return false;
    }
    const frozenUpdates = plan.updates;
    this.#enqueueOrRunMutation(() => {
      // A formatter is user code and may have queued an earlier mutation
      // while measurement was running. Never apply geometry measured from
      // a model or Layout root that is no longer authoritative.
      if (
        this.#store.version !== basisModelVersion ||
        this.#layout.root !== basisLayoutRoot
      ) {
        this.#measurement.forget(frozenUpdates.map(({ row }) => row));
        this.#scheduleAutoHeightContinuation();
        return;
      }

      let changes: readonly SemanticChange[];
      try {
        changes = this.#commandManager.execute(
          new SetRowHeightsCommand(frozenUpdates, "auto"),
          { history: "skip" }
        );
      } catch (error) {
        this.#events.emit("error", { error, eventName: "finalize:auto-height-command" });
        // The originating receipt deliberately deferred its paint. If the
        // derived commit is rejected, flush that already-published Scene so
        // presentation still catches up with the committed model.
        this.#withFinalization(() => {
          this.#runFinalizer("render-frame", () => this.#renderScheduler.flushNow());
        });
        return;
      }

      this.#measurement.markMeasured(frozenUpdates.map(({ row }) => row));
      this.#publishCommitReceipt(changes, remainingFollowUpPasses);
      const rowLayoutChanges = Object.freeze(
        changes.filter((change): change is Extract<SemanticChange, { type: "row-layout" }> =>
          change.type === "row-layout"
        )
      );
      if (rowLayoutChanges.length > 0) {
        this.#events.emit("layout-change", {
          layoutVersion: this.#layout.version,
          changes: rowLayoutChanges
        });
      }
    });
    return true;
  }

  #scheduleAutoHeightContinuation(): void {
    if (this.#disposed || this.#autoHeightContinuationFrame !== null) {
      return;
    }
    this.#autoHeightContinuationFrame = requestAnimationFrame(() => {
      this.#autoHeightContinuationFrame = null;
      if (this.#disposed) {
        return;
      }
      this.#enqueueOrRunMutation(() => this.#syncViewport());
    });
  }

  /**
   * Layout and Selection roots were already published by the structure
   * integration participant. This method only runs observer side effects.
   * Also drops the whole measurement epoch via
   * `WrapMeasurementService.forgetAll()` — an insert/delete/move shifts which
   * row every index refers to, so every cached entry is a stale "row N was
   * measured" claim about a row that may no longer be at N; clearing is
   * simpler and safer than trying to remap each entry the way
   * `SelectionModel`/`LayoutEngine` do for their own per-row state, and just
   * means the next `#syncViewport()` re-measures whatever's visible (cheap —
   * bounded by viewport size, not row count).
   */
  /**
   * The spacer's size is what determines the scroller's native
   * `scrollWidth`/`scrollHeight` — `#resizeCanvases()` already computes the
   * body's own width/height as `rect.width - rowHeaderWidth`/`rect.height -
   * headerHeight()` (the row/column header bands are pinned, sharing the
   * scroller's box rather than adding to it), so this needs the same
   * `+rowHeaderWidth`/`+headerHeight()` to make the scroller's total extent
   * match: without it, native max-scroll comes up exactly `rowHeaderWidth`px
   * short horizontally and `headerHeight()`px short vertically, permanently
   * hiding that many pixels of the last column/row behind nothing reachable
   * by scrolling. Found via a real user report against the playground
   * (10-column dataset, last column's text visibly clipped even scrolled
   * all the way right) — see doc/progress.md's Step 18 patch entry.
   */
  /**
   * Feeds the current content and viewport metrics to both scroll mappers and
   * sizes the spacer to the *native* extent they allow, which is capped at
   * whatever the browser can actually scroll. For any table below that cap the
   * native extent equals the logical one, so the spacer is sized exactly as it
   * was before segmenting existed.
   */
  #updateSpacerSize(root: LayoutRoot = this.#layout.root): void {
    this.#viewport.setContentSize(root.columnAxis.getTotalWidth(), root.rowMetrics.totalHeight);
  }

  /**
   * Lets each axis reconcile the native offset the scroller now reports: a
   * continuous scroll keeps the 1:1 windowed mapping and only recentres near
   * the window's ends, while a jump (scrollbar drag) is treated as a seek and
   * mapped proportionally across the whole logical axis.
   *
   * Deliberately has no "suppress our own scroll event" token: assigning
   * `scrollTop` fires its event *asynchronously*, so a synchronous flag would
   * already be cleared by the time the event lands and would suppress nothing.
   * Instead this relies on `recenter()` being idempotent — the follow-up event
   * re-runs this, gets `null` for both axes, and stops. The cost is one extra
   * Scene per recentre, which happens at most once per half window of
   * scrolling; the benefit is that no genuine user scroll can ever be
   * swallowed by a stale suppression token.
   */
  #recenterScroll(): void {
    this.#viewport.reconcile();
  }

  /** Moves the scroller to a logical position, picking a segment that can express it. */
  #applyLogicalScroll(logicalTop: number, logicalLeft: number): void {
    this.#viewport.scrollToLogical(logicalTop, logicalLeft);
  }

  /** Publishes gesture-only geometry without changing the authoritative Layout root or History. */
  #publishResizePreview(root: LayoutRoot): void {
    this.#withFinalization(() => {
      let preview: SceneSnapshot;
      try {
        preview = this.#createSceneSnapshot(this.#readViewport(), root);
      } catch (error) {
        this.#events.emit("error", { error, eventName: "finalize:resize-preview-scene" });
        return;
      }
      this.#publishScene(preview);
      this.#reprojectEditor(preview);
      this.#runFinalizer("spacer", () => this.#updateSpacerSize(root));
      this.#runFinalizer("render-invalidation", () => this.#renderScheduler.invalidateAll());
      this.#runFinalizer("selection-invalidation", () => this.#renderScheduler.invalidateSelection());
      this.#runFinalizer("render-frame", () => this.#renderScheduler.flushNow());
    });
  }

  /** Total pixel height of the column header band — a single title row. */
  #headerHeight(): number {
    return this.#columnHeaderHeight;
  }

  #resizeCanvases(): void {
    const rect = this.#scroller.getBoundingClientRect();
    const dpr = readDevicePixelRatio();
    const headerHeight = this.#headerHeight();
    const bodyHeight = Math.max(0, rect.height - headerHeight);
    const bodyWidth = Math.max(0, rect.width - this.#rowHeaderWidth);

    this.#headerCanvasSurface.resize(bodyWidth, headerHeight, dpr);
    this.#rowHeaderCanvasSurface.resize(this.#rowHeaderWidth, rect.height, dpr);
    this.#contentCanvasSurface.resize(bodyWidth, bodyHeight, dpr);
    this.#selectionCanvasSurface.resize(bodyWidth, bodyHeight, dpr);

    this.#selectionRenderer.setSize(bodyWidth, bodyHeight);
    this.#updateSpacerSize();
    this.#surfaceHost.refresh();
  }

  /**
   * `width`/`height` here are the body's own scrollable viewport size
   * (total minus the header band and row header band) — every downstream
   * consumer (LayoutEngine's getVisibleRowRange/getVisibleColumnRange,
   * DOMCellManager, EditorManager...) reasons in the body's local
   * coordinate space, which neither header band is physically part of
   * (see createDom's bodyTop/bodyLeft offsets).
   */
  #readViewport(): ViewportState {
    return this.#viewport.readState();
  }

  #createSceneSnapshot(viewportState: ViewportState, layoutRoot: LayoutRoot = this.#layout.root): SceneSnapshot {
    const generation = this.#sceneGeneration + 1;
    if (!Number.isSafeInteger(generation)) {
      throw new RangeError("Scene publication generation exhausted");
    }
    // Reserve before construction. If a later snapshot constructor rejects,
    // the generation remains consumed rather than being reused for different
    // geometry on a subsequent publication.
    this.#sceneGeneration = generation;
    return new SceneSnapshot({
      version: generation,
      skeleton: SkeletonSnapshot.fromRoot(layoutRoot, this.#store.version),
      viewport: new ViewportSnapshot({
        version: generation,
        ...viewportState,
        overscanRows: this.#overscanRows,
        overscanColumns: this.#overscanColumns
      }),
      surfaceMetrics: {
        columnHeaderHeight: this.#headerHeight(),
        rowHeaderWidth: this.#rowHeaderWidth
      },
      findMergeAt: (row, column) => findMergeAt(this.#sheet.merges(), row, column)
    });
  }

  #refreshScene(viewportState: ViewportState = this.#readViewport()): SceneSnapshot {
    const scene = this.#createSceneSnapshot(viewportState);
    this.#publishScene(scene);
    return scene;
  }

  #publishScene(scene: SceneSnapshot): void {
    if (scene.version <= this.#publishedSceneGeneration) {
      throw new RangeError(
        `Scene publication generation must increase: ${scene.version} <= ${this.#publishedSceneGeneration}`
      );
    }
    this.#scene = scene;
    this.#publishedSceneGeneration = scene.version;
  }

  #reprojectEditor(scene: SceneSnapshot): void {
    this.#runFinalizer("editor-geometry", () => this.#editorManager?.updateGeometry(scene));
  }

  #restoreAuthoritativeScene(): void {
    // Cancellation is itself a publication. Re-project the committed roots
    // instead of switching back to an older Scene object and making the
    // public generation go backwards after one or more gesture previews.
    const scene = this.#createSceneSnapshot(this.#readViewport());
    this.#publishScene(scene);
    this.#reprojectEditor(scene);
    this.#runFinalizer("spacer", () => this.#updateSpacerSize());
    this.#runFinalizer("render-invalidation", () => this.#renderScheduler.invalidateAll());
    this.#runFinalizer("selection-invalidation", () => this.#renderScheduler.invalidateSelection());
    this.#runFinalizer("render-frame", () => this.#renderScheduler.flushNow());
  }

  #syncViewport(): void {
    const viewport = this.#readViewport();
    // Runs before any painting below, not during it — growing a row here
    // only ever changes state `LayoutEngine`/the spacer read fresh at paint
    // time, never something already-painted content has to un-paint.
    // Lazy — only rows about to actually be painted this pass (construction,
    // scroll, resize all call #syncViewport()) ever get measured, matching
    // the same "only touch what's visible" principle the rest of this
    // codebase's virtual scrolling already follows; a 1,000,000-row table
    // never pays for wrap measurement on rows nobody has scrolled to yet.
    this.#withFinalization(() => {
      let viewportScene: SceneSnapshot | undefined;
      const sceneReady = this.#runFinalizer("scene", () => {
        viewportScene = this.#refreshScene(viewport);
      });
      if (sceneReady && viewportScene) {
        this.#reprojectEditor(viewportScene);
        const rowResizeScheduled = this.#measureVisibleWrapRows(
          viewportScene,
          AUTO_HEIGHT_MAX_SYNC_PASSES - 1
        );
        this.#runFinalizer("render-invalidation", () => this.#renderScheduler.invalidateAll());
        // The selection box's on-screen position depends on scroll position
        // just like content does — invalidateAll() (called by every caller of
        // #syncViewport(): construction, resize(), scroll) only marks the
        // content layer dirty, so without this the box would stay frozen at
        // whatever screen position it was last painted at.
        this.#runFinalizer("selection-invalidation", () => this.#renderScheduler.invalidateSelection());
        if (!rowResizeScheduled) {
          this.#runFinalizer("render-frame", () => this.#renderScheduler.flushNow());
        }
      }
    });
  }

  /**
   * Keyboard-driven selection changes (arrow keys, Tab, Home/End, and the
   * move-down-on-commit after Enter) can land the active cell outside the
   * currently scrolled-into-view area — nothing about SelectionModel or
   * KeyboardController touches scroll position, so without this the active
   * cell can end up completely off-screen with no visual feedback that
   * anything moved. Mouse-driven selection never needs this: a click can
   * only land on a cell that's already visible. A cell inside the frozen
   * band on a given axis is always visible on that axis regardless of
   * scroll, so that axis is left untouched.
   */
  #scrollActiveIntoView(): void {
    const { active } = this.#selectionModel.snapshot;
    const target = this.#scene.getScrollTargetForCell(active);
    if (!target) {
      return;
    }
    this.#applyLogicalScroll(target.logicalTop, target.logicalLeft);
  }

  #routeBodyFill(context: CorePointerInputContext) {
    const began = this.#bodyGestures.tryBeginFill(
      context.point as SurfacePoint<"body">,
      context.event,
      context.frame
    );
    return began ? handleInput(undefined) : declineInput();
  }

  /** A custom renderer only owns the input when its single hitTest call returns an action. */
  #routeBodyCanvasAction(context: CorePointerInputContext) {
    const { event, frame } = context;
    if (event.button !== 0 || frame.structureVersion === undefined) {
      return declineInput();
    }
    const hit = frame.scene.hitTestCellGeometry(context.point);
    if (!hit) {
      return declineInput();
    }

    const column = this.#columnManager.getColumn(hit.cell.column);
    const renderer = column.renderer;
    if (!renderer || renderer.kind !== "canvas" || !renderer.hitTest) {
      return declineInput();
    }
    let action;
    try {
      action = renderer.hitTest({
        ctx: this.#contentCanvas.getContext("2d") as Canvas2DLike,
        rect: hit.rect,
        row: hit.cell.row,
        column,
        value: this.#store.getCell(hit.cell.row, hit.cell.column),
        point: hit.localPoint
      });
    } catch (error) {
      this.#events.emit("error", {
        error,
        eventName: `interaction:canvas-cell-hit-test:${hit.cell.row}:${hit.cell.column}`
      });
      return declineInput();
    }
    if (!action) {
      return declineInput();
    }

    const cell = hit.cell;
    const structureVersion = frame.structureVersion;
    this.#interactionFrames.enqueueIntent(() => {
      if (structureVersion !== this.#structure.version) {
        return;
      }
      const dispatch = dispatchCellAction(action, cell.row, cell.column);
      if (dispatch.kind === "command") {
        this.#applyCommandNow(dispatch.command);
      } else {
        this.#editorManager.startEdit(cell.row, cell.column);
      }
    });
    return handleInput(undefined);
  }

  #startBodySelection(context: CorePointerInputContext): void {
    this.#bodyGestures.beginSelection(
      context.point as SurfacePoint<"body">,
      context.event,
      context.frame
    );
  }

  #handleDelegatedPointerDown = (event: PointerEvent): void => {
    if (!event.isPrimary || (event.button !== 0 && event.button !== 2)) {
      return;
    }
    // Runs before any routing calls contentCanvas.focus(), so the focus
    // handler below can tell a pointer-driven focus from a Tab.
    this.#focusFromPointer = true;
    if (this.#bodyGestures.isActive) {
      if (!this.#bodyGestures.matchesPointer(event.pointerId)) {
        return;
      }
      this.#bodyGestures.cancel();
    }
    if (this.#headerGestures.isActive) {
      if (!this.#headerGestures.matchesPointer(event.pointerId)) {
        return;
      }
      this.#headerGestures.cancel(true);
    }
    const handled = this.#dispatchInteraction("pointer-down", (frame) =>
      this.#inputRouter.dispatch(frame, {
        event,
        clientPoint: clientPoint(event.clientX, event.clientY),
        path: event.composedPath()
      }).handled
    );
    if (handled === undefined) {
      this.#bodyGestures.cancel();
      this.#headerGestures.cancel(true);
    }
  };

  #resolvePointerInputTarget(
    context: InputSurfacePolicyContext<CoreInteractionFrame, PointerEvent, EventTarget>
  ): InputTargetResolution {
    for (const entry of context.path) {
      const declaredTarget = getDeclaredDomPointerInputTarget(entry);
      if (declaredTarget) {
        return declaredTarget;
      }
      if (isNativePointerPathEntry(entry)) {
        return nativeInputTarget();
      }
      if (entry === this.#rowHeaderCanvas) {
        const corner = context.frame.surfaces.get("corner");
        const point = context.frame.surfaces.clientToSurface(context.clientPoint, "corner");
        return surfaceInputTarget(
          corner && point &&
            point.x >= 0 &&
            point.y >= 0 &&
            point.x < corner.cssSize.width &&
            point.y < corner.cssSize.height
            ? "corner"
            : "row-header"
        );
      }
      const surface = this.#inputSurfaceOwners.get(entry);
      if (surface) {
        return surfaceInputTarget(surface);
      }
    }
    return noInputTarget();
  }

  #handleDblClick = (event: MouseEvent): void => {
    if (this.#editorManager.isEditing()) {
      return;
    }
    this.#dispatchInteraction("double-click", (frame) => {
      const point = frame.surfaces.clientToSurface(
        clientPoint(event.clientX, event.clientY),
        "body"
      );
      const selection = frame.selection;
      if (!point || !selection) {
        return;
      }
      const structureVersion = frame.structureVersion;
      if (isPointOnFillHandle(point, selection.range, frame.scene)) {
        const range = selection.range;
        this.#interactionFrames.enqueueIntent(() => {
          if (structureVersion === this.#structure.version) {
            this.#autoFillDownNow(range);
          }
        });
        return;
      }
      const active = selection.active;
      this.#interactionFrames.enqueueIntent(() => {
        if (structureVersion === this.#structure.version) {
          this.#editorManager.startEdit(active.row, active.column);
        }
      });
    });
  };

  /** Double-clicking the fill handle extends the selection down to match the nearest non-empty neighbor column's run — see computeAutoFillDownTargetRow(). */
  #autoFillDownNow(range: CellRange): void {
    const targetRow = computeAutoFillDownTargetRow(
      range,
      { columnCount: this.#columns.length, rowCount: this.#sheet.rowCount() },
      (row, column) => this.#store.getCell(row, column)
    );
    if (targetRow === null) {
      return;
    }
    this.#commitFillNow(range, { ...range, endRow: targetRow });
  }

  /** Applies a fill-handle result (drag or double-click auto-fill) as one SetRangeCommand, then moves the selection onto what was just filled. The value matrix itself is computed by computeFillPlan(). */
  #commitFillNow(sourceRange: CellRange, targetRange: CellRange): void {
    const values = computeFillPlan(sourceRange, targetRange, {
      getCell: (row, column) => this.#store.getCell(row, column),
      isPasteBlocked: (column) => this.#columnManager.getColumn(column).clipboard.paste === false
    });

    this.#applyCommandNow(new SetRangeCommand(targetRange, values));
    this.#selectionModel.selectCell(targetRange.startRow, targetRange.startColumn);
    this.#selectionModel.extendTo(targetRange.endRow, targetRange.endColumn);
    this.#renderScheduler.invalidateSelection();
  }

  /**
   * Point is computed relative to the scroller (not contentCanvas) — see
   * the listener registration comment — so it's already in the same
   * header-inclusive coordinate space `classifyContextMenuTarget()` and
   * menuLayer (also un-offset by bodyTop, see createDom) both expect.
   * "row-header"/"corner" targets are now actually reachable once
   * `rowHeaderWidth` is nonzero (see the row-header patch in
   * doc/progress.md — Step 9's "不做行号头" scope note no longer applies).
   */
  #handleContextMenu = (event: MouseEvent): void => {
    event.preventDefault();
    this.#dispatchInteraction("context-menu", (frame) => {
      const point = frame.surfaces.clientToSurface(
        clientPoint(event.clientX, event.clientY),
        "menu-overlay"
      );
      if (!point || !frame.selection) {
        return;
      }
      const context = classifyContextMenuTarget(
        { x: point.x, y: point.y },
        frame.scene,
        frame.selection
      );
      this.#contextMenuService.open(context);
    });
  };

  /**
   * Real Ctrl+C on the focused grid — handled separately from the
   * programmatic copy() (which goes through ClipboardService's async
   * Worker/oversized-plan machinery) because the browser's native `copy`
   * event contract requires `clipboardData.setData()` to be called
   * synchronously within the event, before it returns. Always copies the
   * current selection in "display" mode, the same default `sheet.copy()`
   * uses; the CopyPlan/oversized-confirmation flow only applies to the
   * programmatic API.
   */
  #handleCopyEvent = (event: ClipboardEvent): void => {
    if (this.#editorManager.isEditing()) {
      return;
    }
    const range = this.#selectionModel.snapshot.range;
    const grid = buildCellGrid(range, "display", {
      getCellValue: (row, column) => this.#store.getCell(row, column),
      getColumn: (column) => this.#columnManager.getColumn(column)
    });
    event.clipboardData?.setData("text/plain", encodeTsv(grid));
    event.preventDefault();
  };

  /** Real Ctrl+V — the pasted text is already synchronously available on the event, so this skips ClipboardService's async io.readText() entirely. */
  #handlePasteEvent = (event: ClipboardEvent): void => {
    if (this.#editorManager.isEditing()) {
      return;
    }
    const text = event.clipboardData?.getData("text/plain");
    if (text === undefined) {
      return;
    }
    event.preventDefault();
    void this.#clipboardService.paste({ text });
  };

  #handleKeyDown = (event: KeyboardEvent): void => {
    // Keydown events from the mounted editor's own DOM (a sibling of the
    // canvas, not a descendant) never reach this listener, but this guard
    // stays as cheap insurance against that invariant changing later.
    if (this.#editorManager.isEditing()) {
      return;
    }

    // Built-in Mod+Z/Mod+Shift+Z/Mod+Y/Enter are themselves registered
    // through #shortcuts (see the constructor) — a KeyboardEvent already has
    // the exact shape ShortcutRegistry.dispatch() expects, so it's passed
    // straight through, no wrapping needed.
    if (this.#shortcuts.dispatch(event)) {
      event.preventDefault();
      return;
    }

    if (event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
      event.preventDefault();
      const { active } = this.#selectionModel.snapshot;
      this.#editorManager.startEdit(active.row, active.column, event.key);
      return;
    }

    if (this.#keyboardController.handleKeyDown(event.key, { shift: event.shiftKey })) {
      event.preventDefault();
      this.#scrollActiveIntoView();
    }
  };

  #handleScroll = (): void => {
    this.#headerGestures.cancelResizeOnly(true);
    if (this.#editorManager.isEditing()) {
      if (this.#scrollCommitPolicy === "commit") {
        void this.#editorManager.commitEdit();
      } else {
        this.#editorManager.cancelEdit();
      }
    }
    this.#recenterScroll();
    this.#syncViewport();
    this.#renderScheduler.invalidateAll();
  };

  #routeHeaderResize(
    context: CorePointerInputContext,
    kind: "column-resize" | "row-resize"
  ) {
    const { event, frame } = context;
    if (
      event.button !== 0 ||
      !event.isPrimary ||
      frame.structureVersion === undefined ||
      context.surface !== (kind === "column-resize" ? "column-header" : "row-header")
    ) {
      return declineInput();
    }
    const point = context.point as
      | SurfacePoint<"column-header">
      | SurfacePoint<"row-header">;
    if (!this.#headerResizeController.handlePointerDown(
      point,
      frame.scene,
      frame.structureVersion
    )) {
      return declineInput();
    }

    event.preventDefault();
    this.#contentCanvas.focus();
    this.#headerGestures.begin(kind, event, frame.structureVersion);
    if (kind === "column-resize") {
      this.#headerCanvas.style.cursor = "col-resize";
    } else {
      this.#rowHeaderCanvas.style.cursor = "row-resize";
    }
    return handleInput(undefined);
  }

  #routeHeaderCanvasAction(context: CorePointerInputContext) {
    const { event, frame } = context;
    if (
      event.button !== 0 ||
      !event.isPrimary ||
      frame.structureVersion === undefined ||
      context.surface !== "column-header"
    ) {
      return declineInput();
    }
    const hit = frame.scene.hitTestColumnHeader(
      context.point as SurfacePoint<"column-header">
    );
    if (!hit || hit.band.role !== "leaf") {
      return declineInput();
    }
    const header = this.#headerManager.getHeader(hit.column);
    const renderer = header.renderer;

    let action: HeaderAction | void;
    if (renderer === undefined) {
      // No custom renderer -> the default painter's own reserved right-edge
      // strip is click-to-sort (see defaultHeaderHitTest's doc comment).
      action = defaultHeaderHitTest({
        rect: hit.rect,
        point: hit.localPoint,
        sortable: this.#columnManager.getColumn(hit.column).sortable
      });
    } else if (renderer.kind === "canvas" && renderer.hitTest) {
      try {
        action = renderer.hitTest({
          ctx: this.#headerCanvas.getContext("2d") as Canvas2DLike,
          rect: hit.rect,
          column: hit.column,
          title: header.title,
          state: this.#headerState(hit.column),
          point: hit.localPoint
        });
      } catch (error) {
        this.#events.emit("error", {
          error,
          eventName: `interaction:canvas-header-hit-test:${hit.column}`
        });
        return declineInput();
      }
    } else {
      // A "dom" renderer, or a canvas renderer with no hitTest at all:
      // neither participates in this action-hitTest path.
      return declineInput();
    }
    if (!action) {
      return declineInput();
    }

    const payload: CyExcelHeaderActionEvent = Object.freeze({
      action: Object.freeze({ ...action }),
      column: hit.column,
      header,
      structureVersion: frame.structureVersion
    });
    // Only the *default* painter's reserved strip means "core, sort this".
    // A custom renderer's `toggle-sort` is what it has always been — a
    // notification the host translates into its own command (see
    // doc/customization-and-hooks.md §7) — and sorting it here as well would
    // give every host following that doc two reorders and two history
    // entries per click.
    const coreOwnsTheSort = renderer === undefined;
    this.#interactionFrames.enqueueIntent(() => {
      if (payload.structureVersion !== this.#structure.version) {
        return;
      }
      if (coreOwnsTheSort && payload.action.type === "toggle-sort") {
        // Applied before emit(), and `#sortState` written only once the
        // reorder has committed: a listener that synchronously reads
        // `sortState` (the playground's own sortHint update, for instance)
        // must see the sort this exact action produced — and must never see
        // one that has not touched the data. Emitting first left every
        // listener one click behind; writing `#sortState` first (what this
        // used to do, from `sortByColumn`'s deferring call site) announced
        // sorts that had not run yet and could still fail.
        if (!this.#applyHeaderSort(payload.column)) {
          // Refused or cancelled. The `error` event already carries why, the
          // indicator still describes the last sort that really happened,
          // and emitting `header-action` here would tell every listener a
          // sort just did.
          return;
        }
      }
      this.#events.emit("header-action", payload);
    });
    event.preventDefault();
    this.#contentCanvas.focus();
    return handleInput(undefined);
  }

  /** `HeaderState` for one column, as shown by its header's sort/filter indicator. Backs both `HeaderCanvasRenderer`/`DomHeaderManager`'s `getHeaderState` and the `state` a custom canvas renderer's `hitTest`/`paint` receive. Filtering is Part B — always `false` for now. */
  #headerState(column: number): HeaderState {
    return {
      sorted: this.#sortState !== null && this.#sortState.column === column ? this.#sortState.direction : false,
      filtered: false,
      sortable: this.#columnManager.getColumn(column).sortable
    };
  }

  /**
   * Applies a default header's `toggle-sort`: cycles asc/desc via
   * `nextSortDirection` — never a third "restore original order" state,
   * because a physical sort destroys the original order and only undo can
   * recover it (ADR 0003 §3.4).
   *
   * Runs the sort *inline* rather than through `sortByColumn`. The caller is
   * an interaction intent, which the mutation queue runs during its own
   * drain, so `sortByColumn` would only queue the work behind this drain and
   * return before it happened — leaving nothing here to report on.
   *
   * Returns whether the indicator moved. A reorder the storage layer rejects
   * (a permutation that would cut a merge region, say) reports an `error`
   * event instead of throwing out of the pointer-routing pipeline, and
   * leaves the indicator exactly where it was — which is only true because
   * `#runSortByColumn` writes `#sortState` after the command commits.
   */
  #applyHeaderSort(column: number): boolean {
    const direction = nextSortDirection(this.#sortState, column);
    try {
      return this.#runSortByColumn(column, direction);
    } catch (error) {
      this.#events.emit("error", { error, eventName: `interaction:toggle-sort:${column}` });
      return false;
    }
  }

  #routeHeaderSelection(
    context: CorePointerInputContext,
    gestureKind: "column-selection" | "row-selection" | null
  ) {
    const { event, frame } = context;
    if (frame.structureVersion === undefined) {
      return declineInput();
    }
    const result = this.#headerSelectionController.down(
      context.point as
        | SurfacePoint<"column-header">
        | SurfacePoint<"row-header">
        | SurfacePoint<"corner">,
      frame.scene,
      this.#selectionModel,
      frame.structureVersion,
      { shift: event.shiftKey, button: event.button, isPrimary: event.isPrimary }
    );
    if (!result) {
      return declineInput();
    }
    this.#renderScheduler.invalidateSelection();
    this.#contentCanvas.focus();
    if (gestureKind && event.button === 0 && this.#headerSelectionController.dragging) {
      event.preventDefault();
      this.#headerGestures.begin(gestureKind, event, frame.structureVersion);
    }
    return handleInput(undefined);
  }

  #routeRowReorder(context: CorePointerInputContext) {
    const { event, frame } = context;
    if (
      !this.#rowReorderController ||
      event.button !== 0 ||
      !event.isPrimary ||
      !frame.selection ||
      frame.structureVersion === undefined ||
      context.surface !== "row-header"
    ) {
      return declineInput();
    }
    const point = context.point as SurfacePoint<"row-header">;
    if (!this.#rowReorderController.handlePointerDown(
      point.y,
      frame.selection,
      frame.scene,
      frame.structureVersion
    )) {
      return declineInput();
    }

    event.preventDefault();
    this.#contentCanvas.focus();
    document.body.style.cursor = "grabbing";
    this.#headerGestures.begin("row-reorder", event, frame.structureVersion);
    return handleInput(undefined);
  }

  #commitHeaderResize(intent: HeaderResizeIntent): void {
    if (intent.structureVersion !== this.#structure.version) {
      this.#withFinalization(() => this.#restoreAuthoritativeScene());
      return;
    }
    try {
      this.#cancelActiveEdit();
      if (intent.axis === "column") {
        this.#applyCommandNow(
          new ResizeColumnCommand(intent.column, intent.width)
        );
      } else {
        this.#applyCommandNow(new ResizeRowCommand(intent.row, intent.height));
      }
    } catch (error) {
      this.#withFinalization(() => this.#restoreAuthoritativeScene());
      this.#events.emit("error", {
        error,
        eventName: `interaction:${intent.axis}-resize`
      });
    }
  }

  #commitRowReorder(start: number, count: number, target: number): void {
    try {
      validateAxisChange(
        { axis: "row", operation: "move", start, count, target },
        this.#sheet.rowCount()
      );
      this.#cancelActiveEdit();
      this.#applyCommandNow(new MoveRowsCommand(start, count, target));
    } catch (error) {
      this.#events.emit("error", { error, eventName: "interaction:row-reorder" });
    }
  }

  #handleHeaderPointerHover = (event: PointerEvent): void => {
    if (!this.#headerGestures.isActive) {
      this.#updateHeaderHover(event, "column-header");
    }
  };

  #handleRowHeaderPointerHover = (event: PointerEvent): void => {
    if (!this.#headerGestures.isActive) {
      this.#updateHeaderHover(event, "row-header");
    }
  };

  #updateHeaderHover(
    event: PointerEvent,
    explicitSurface?: "column-header" | "row-header"
  ): void {
    const surface = explicitSurface ?? (
      event.currentTarget === this.#headerCanvas ? "column-header" : "row-header"
    );
    this.#dispatchInteraction(`${surface}-pointer-hover`, (frame) => {
      const snapshot = frame.surfaces.get(surface);
      const point = frame.surfaces.clientToSurface(
        clientPoint(event.clientX, event.clientY),
        surface
      );
      if (
        !snapshot ||
        !point ||
        point.x < 0 ||
        point.y < 0 ||
        point.x >= snapshot.cssSize.width ||
        point.y >= snapshot.cssSize.height
      ) {
        if (surface === "column-header") {
          this.#headerCanvas.style.cursor = "";
        } else {
          this.#rowHeaderCanvas.style.cursor = "";
        }
        return;
      }
      const resize = this.#headerResizeController.hitTest(point, frame.scene);
      if (surface === "column-header") {
        this.#headerCanvas.style.cursor = resize ? "col-resize" : "";
        return;
      }
      if (resize) {
        this.#rowHeaderCanvas.style.cursor = "row-resize";
        return;
      }
      const reorder = frame.selection
        ? this.#rowReorderController?.isPointOnHandle(
            point.y,
            frame.selection,
            frame.scene
          ) ?? false
        : false;
      this.#rowHeaderCanvas.style.cursor = reorder ? "grab" : "";
    });
  }

}

export function createCyExcelCore(
  container: HTMLElement,
  options: CyExcelCoreOptions = {}
): CyExcelCore {
  return new CyExcelCore(container, options);
}
