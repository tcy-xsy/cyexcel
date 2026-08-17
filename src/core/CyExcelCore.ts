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
  InsertRowsCommand,
  DeleteRowsCommand,
  MoveRowsCommand,
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
import { ArrayDataSource, TableStore, createChangeSet, type ArrayDataSourceColumn } from "./data";
import type { CellEditorFactory } from "./editors";
import { EditorManager } from "./editors";
import { CyExcelError, CyExcelErrorCode } from "./errors";
import { EventBus } from "./events";
import {
  DomHeaderManager,
  HeaderCanvasRenderer,
  HeaderManager,
  RowHeaderCanvasRenderer,
  type HeaderAction,
  type HeaderGroupDefinition,
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
  type ViewportState
} from "./layout";
import { InitializationRollback } from "./lifecycle";
import { ContextMenuService, classifyContextMenuTarget } from "./menu";
import { DeferredMutationQueue } from "./model";
import { PluginManager, type CyExcelPlugin } from "./plugins";
import {
  SceneSnapshot,
  SurfaceHost,
  clientPoint,
  type SurfacePoint,
  type SurfaceId,
  type SurfaceSetSnapshot
} from "./scene";
import { WrapMeasurementService } from "./measurement";
import { ShortcutRegistry, type ShortcutHandler } from "./shortcuts";
import {
  ContentLayerRenderer,
  CanvasSurface,
  RenderScheduler,
  SelectionLayerRenderer,
  readDevicePixelRatio,
  type Canvas2DLike
} from "./render";
import { CELL_FONT } from "./render/theme";
import type {
  CellRange,
  Command,
  DataSource,
  Disposer,
  Patch,
  RowData
} from "./types";
import { GridStructureOwner, createGridCommandManager } from "./integration";
import {
  validateCyExcelCoreDataDimensions,
  validateCyExcelCoreHeaderGeometry,
  validateCyExcelCoreOptions
} from "./validateCyExcelCoreOptions";

export type CyExcelColumnConfig = ColumnDefinition;

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
   * Row objects are kept by reference, not deep-cloned (matches
   * ArrayDataSource's own contract when constructed directly via
   * `dataSource`, and ADR 0001 §5's "complex values are atomic, no deep
   * observation" applied one level up to the row wrapper itself — see
   * Univer's `Worksheet` constructor / `ObjectMatrix` for the same choice
   * in an established OSS spreadsheet). Two consequences: `setCell()` etc.
   * mutate the exact row objects passed in here, and — the one to actually
   * watch for — mutating a row object in this array yourself *after*
   * construction changes what the sheet reads back, but bypasses Command/
   * Patch/History entirely, so nothing repaints, nothing becomes
   * undoable, and no `change` event fires. Always go through
   * `setCell()`/`setRange()`/`paste()` for edits; never mutate `data`'s
   * row objects directly once they're handed to the constructor.
   */
  readonly data?: readonly RowData[];
  readonly dataSource?: DataSource;
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
  /** Arbitrary-depth nested column groups (Step 19), drawn in rows above the column titles — only shown when `headerHeight > 0`. */
  readonly headerGroups?: readonly HeaderGroupDefinition[];
  /**
   * When a column has no explicit `header`, show its spreadsheet-style
   * letter (A, B, ..., Z, AA, ...) instead of falling back to its `id`.
   * Default `false` — unset behavior (fall back to `id`) is unchanged. An
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

function buildDataSource(
  options: CyExcelCoreOptions,
  columns: readonly CyExcelColumnConfig[]
): DataSource {
  if (options.dataSource) {
    return options.dataSource;
  }
  // Shallow-copies only the *array* (cheap: one reference per row), not
  // each row object — ArrayDataSource.insertRows()/deleteRows()/moveRows()
  // splice() this array in place, so without even this much copying a
  // structural edit would silently resize/reorder the caller's own array.
  // Row objects themselves are intentionally left un-cloned: setCell()
  // mutating them in place is ArrayDataSource's own tested contract
  // (tests/unit/core/data/ArrayDataSource.test.ts), so cloning them here
  // only for the `data` convenience path would make that path behave
  // differently from constructing ArrayDataSource directly — for no
  // documented reason, and at O(rows * columns) cost (the previous
  // `options.data.map((row) => ({ ...row }))` here was the dominant cost
  // behind Step 15's 1,000,000 x 50 first-paint benchmark taking 6-10s to
  // construct; see doc/progress.md's Step 15 entry).
  const rows: RowData[] = options.data ? [...options.data] : [];
  const arrayColumns: ArrayDataSourceColumn[] = columns.map((column) => ({
    id: column.id,
    field: column.field ?? column.id
  }));
  return new ArrayDataSource(rows, arrayColumns);
}

function buildLayout(
  options: CyExcelCoreOptions,
  columns: readonly CyExcelColumnConfig[],
  rowCount: number
): LayoutEngine {
  return new LayoutEngine({
    rowCount,
    columnCount: columns.length,
    rowHeight: options.rowHeight ?? DEFAULT_ROW_HEIGHT,
    ...(options.rowHeights !== undefined ? { rowHeights: options.rowHeights } : {}),
    columnWidths: columns.map((column) => column.width ?? DEFAULT_COLUMN_WIDTH),
    ...(options.freeze !== undefined ? { freeze: options.freeze } : {}),
    ...(options.overscanRows !== undefined ? { overscanRows: options.overscanRows } : {}),
    ...(options.overscanColumns !== undefined
      ? { overscanColumns: options.overscanColumns }
      : {})
  });
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
 * The integration layer: DataSource -> TableStore -> CommandManager -> Patch
 * -> RenderScheduler/Canvas -> DOM editor overlay, all wired to a real
 * container. Unlike src/core/{data,commands,layout,render,interaction,
 * editors}/, this class genuinely creates and manipulates DOM/Canvas — see
 * doc/code-reference.md §2 for why that doesn't violate the "core doesn't
 * depend on HTMLElement" ADR (that ADR is about not requiring the Custom
 * Elements registry, not about never touching the DOM).
 */
export class CyExcelCore {
  readonly #scroller: HTMLDivElement;
  readonly #spacer: HTMLDivElement;
  readonly #headerCanvas: HTMLCanvasElement;
  readonly #rowHeaderCanvas: HTMLCanvasElement;
  readonly #contentCanvas: HTMLCanvasElement;
  readonly #selectionCanvas: HTMLCanvasElement;
  readonly #inputSurfaceOwners = new WeakMap<EventTarget, SurfaceId>();
  readonly #surfaceHost: SurfaceHost;
  /** Owns wrap-height measurement and the measured-row cache for the current data/column structure epoch. Core keeps only the commit half. */
  readonly #measurement: WrapMeasurementService;

  #columns: readonly CyExcelColumnConfig[];
  readonly #headerGroups: readonly HeaderGroupDefinition[];
  readonly #cellTypeRegistry: CellTypeRegistry;
  readonly #columnManager: ColumnManager;
  readonly #headerManager: HeaderManager;
  readonly #columnHeaderHeight: number;
  readonly #groupHeaderHeight: number;
  readonly #rowHeaderWidth: number;
  readonly #dataSource: DataSource;
  readonly #store: TableStore;
  readonly #commandManager: CommandManager;
  readonly #layout: LayoutEngine;
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
  #autoHeightContinuationFrame: number | null = null;
  #finalizationDepth = 0;
  readonly #mutationIntents: DeferredMutationQueue;
  readonly #interactionFrames: InteractionFrameProvider<
    SceneSnapshot,
    SurfaceSetSnapshot,
    SelectionSnapshot
  >;
  #structureGeneration = 0;
  #disposed = false;

  constructor(container: HTMLElement, options: CyExcelCoreOptions = {}) {
    validateCyExcelCoreOptions(options);

    // Copied, not aliased — insertColumns()/deleteColumns()/moveColumns()
    // splice this array in place (mirrors buildDataSource()'s identical
    // reasoning for `options.data`, a few lines below), so without this
    // much copying a structural column change would silently resize/
    // reorder the caller's own `options.columns` array.
    this.#columns = snapshotColumnDefinitions(options.columns ?? []);
    this.#headerGroups = options.headerGroups ?? [];
    this.#scrollCommitPolicy = options.scrollCommitPolicy ?? "commit";
    this.#overscanRows = options.overscanRows ?? 0;
    this.#overscanColumns = options.overscanColumns ?? 0;

    this.#headerManager = new HeaderManager({
      columns: this.#columns,
      groups: this.#headerGroups,
      columnLetterFallback: options.columnLetterFallback ?? false
    });
    this.#columnHeaderHeight = options.headerHeight ?? 0;
    this.#groupHeaderHeight =
      this.#columnHeaderHeight > 0 && (options.headerGroups?.length ?? 0) > 0
        ? this.#columnHeaderHeight
        : 0;
    this.#rowHeaderWidth = options.rowHeaderWidth ?? 0;
    validateCyExcelCoreHeaderGeometry(
      this.#columnHeaderHeight,
      this.#headerManager.groupDepth
    );

    const dataSource = buildDataSource(options, this.#columns);
    const initialDimensions = {
      rowCount: dataSource.getRowCount(),
      columnCount: dataSource.getColumnCount()
    };
    validateCyExcelCoreDataDimensions(
      options,
      initialDimensions,
      this.#columns.length
    );
    this.#dataSource = dataSource;
    this.#layout = buildLayout(options, this.#columns, initialDimensions.rowCount);

    const initialization = new InitializationRollback();
    try {
    const dom = createDom(
      container,
      this.#headerHeight(), // #headerManager/#columnHeaderHeight/#groupHeaderHeight are all assigned above
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
    this.#spacer = dom.spacer;
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
      // action throwing (e.g. insertColumns() blocked by headerGroups) is
      // otherwise only visible via console.error, which a host app has no
      // way to respond to it.
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
      hasHeaderGroups: () => this.#headerGroups.length > 0,
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

    this.#store = new TableStore(this.#dataSource);
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
      headerManager: this.#headerManager
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
      headerManager: this.#headerManager
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
      getStructureVersion: () => this.#structureGeneration
    });
    this.#structureOwner = new GridStructureOwner({
      getColumns: () => this.#columns,
      replaceColumns: (columns) => {
        this.#columns = columns;
      },
      headerGroups: this.#headerGroups,
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
      onSelectionChange: () => this.#renderScheduler.invalidateSelection()
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
          this.#scroller.scrollTop += delta;
          this.#syncViewport();
          this.#renderScheduler.invalidateAll();
        },
        commitResize: (intent) => this.#commitHeaderResize(intent),
        commitReorder: (intent) => {
          if (intent.structureVersion === this.#structureGeneration) {
            this.#commitRowReorder(intent.start, intent.count, intent.target);
          }
        },
        updateHoverAfterGesture: (event, surface) => this.#updateHeaderHover(event, surface),
        reportError: (error, eventName) => this.#events.emit("error", { error, eventName })
      }
    });
    this.#bodyGestures = new BodyGestureController({
      pointerController: this.#pointerController,
      fillHandleController: this.#fillHandleController,
      captureTarget: this.#contentCanvas,
      focusTarget: this.#contentCanvas,
      withFrame: (eventName, run) => this.#dispatchInteraction(eventName, run),
      onFillCommit: (intent) => {
        this.#interactionFrames.enqueueIntent(() => {
          if (intent.structureVersion === this.#structureGeneration) {
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
      onSelectionChange: () => this.#renderScheduler.invalidateSelection()
    });

    this.#editorManager = new EditorManager({
      container: dom.editorLayer,
      getScene: () => this.#scene,
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
      getRowCount: () => this.#dataSource.getRowCount(),
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
   * Inserts `count` blank rows (every cell `undefined`) at `start`. To
   * insert pre-filled rows, insert blank ones then setRange() over them.
   * The DataSource edit and the matching Layout/Selection roots are
   * prepared together; live owners are changed only by the narrow commit.
   * New rows are always default-height (no third `heights` argument here);
   * a row wanting a custom height gets one via `setRowHeight()` afterward.
   */
  insertRows(start: number, count = 1): void {
    this.#enqueueOrRunMutation(() => {
      validateAxisChange({ axis: "row", operation: "insert", start, count }, this.#dataSource.getRowCount());
      this.#cancelActiveEdit();
      const rows = Array.from({ length: count }, () => new Array<unknown>(this.#columns.length).fill(undefined));
      this.#applyCommandNow(new InsertRowsCommand(start, rows));
    });
  }

  /** Deletes rows while the integration participant captures inverse layout state during prepare. */
  deleteRows(start: number, count = 1): void {
    this.#enqueueOrRunMutation(() => {
      validateAxisChange({ axis: "row", operation: "delete", start, count }, this.#dataSource.getRowCount());
      this.#cancelActiveEdit();
      this.#applyCommandNow(new DeleteRowsCommand(start, count));
    });
  }

  /** `target` is the insertion index in the pre-move index space (same convention `ArrayDataSource.moveRows()` uses). */
  moveRows(start: number, count: number, target: number): void {
    this.#enqueueOrRunMutation(() => {
      validateAxisChange({ axis: "row", operation: "move", start, count, target }, this.#dataSource.getRowCount());
      this.#cancelActiveEdit();
      this.#applyCommandNow(new MoveRowsCommand(start, count, target));
    });
  }

  /** Sets an exact row height through Command/Mutation/History. */
  setRowHeight(row: number, height: number): void {
    this.#enqueueOrRunMutation(() => {
      if (
        row >= 0 &&
        row < this.#layout.rowCount &&
        this.#layout.getRowHeight(row) === height &&
        this.#layout.getRowHeightProvenance(row) === "manual"
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
        this.#layout.getColumnRestoreWidth(column) === width
      ) {
        return;
      }
      this.#cancelActiveEdit();
      this.#applyCommandNow(new ResizeColumnCommand(column, width));
    });
  }

  /** Measures whatever rows are about to be painted (already-visible + overscan, per the Scene's own body panes) that haven't been measured before — called from `#syncViewport()`, before painting. A no-op scan when the table has no wrap columns at all, and a no-op per-row once a row has been measured once (until an actual edit re-triggers it via `#invalidatePatches()`'s cell-value/range-value handling). */
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
      this.#assertNoHeaderGroups("insert");
      validateAxisChange({ axis: "column", operation: "insert", start, count: inserted.length }, this.#columns.length);
      this.#cancelActiveEdit();
      this.#applyCommandNow(new InsertColumnsCommand(start, inserted));
    });
  }

  deleteColumns(start: number, count = 1): void {
    this.#enqueueOrRunMutation(() => {
      this.#assertNoHeaderGroups("delete");
      validateAxisChange({ axis: "column", operation: "delete", start, count }, this.#columns.length);
      this.#cancelActiveEdit();
      this.#applyCommandNow(new DeleteColumnsCommand(start, count));
    });
  }

  /** `target` is the insertion index in the pre-move index space (same convention `ArrayDataSource.moveColumns()` uses). */
  moveColumns(start: number, count: number, target: number): void {
    this.#enqueueOrRunMutation(() => {
      this.#assertNoHeaderGroups("move");
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
    this.#enqueueOrRunMutation(() => this.#invalidatePatches(this.#commandManager.undo()));
  }

  redo(): void {
    this.#enqueueOrRunMutation(() => this.#invalidatePatches(this.#commandManager.redo()));
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
    this.#invalidatePatches(this.#commandManager.execute(command));
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
   * `headerGroups`' `HeaderGroupDefinition.columns` are plain column
   * indices — a structural change would need to remap or invalidate them,
   * which is deferred to Step 19 (arbitrary-depth grouped headers) when the
   * group data structure gets redesigned anyway (see the Step 17 plan's
   * "key design decisions" §5). Checked before `#cancelActiveEdit()` so a
   * doomed call doesn't cancel an in-progress edit for nothing, and well
   * before any Command runs, so a rejected call never leaves the Store/
   * History committed while `#columns`/`ColumnManager`/`LayoutEngine` are
   * still out of sync.
   */
  #assertNoHeaderGroups(operation: "insert" | "delete" | "move"): void {
    if (this.#headerGroups.length > 0) {
      throw new CyExcelError(
        CyExcelErrorCode.HeaderGroupsBlockColumnStructureChange,
        `Cannot ${operation} columns while headerGroups is configured — grouped headers reference column indices that a structural change would invalidate`
      );
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

  #invalidatePatches(
    patches: readonly Patch[] | undefined,
    measurementPassesRemaining = AUTO_HEIGHT_MAX_SYNC_PASSES
  ): void {
    if (!patches) {
      return;
    }
    // Capture the receipt-derived facts before any finalizer runs. A DOM
    // renderer or auto-height change may synchronously dispatch another
    // command and advance CommandManager.lastChangeSet reentrantly.
    const changeSet =
      this.#commandManager.lastChangeSet?.modelVersion === this.#store.version
        ? this.#commandManager.lastChangeSet
        : createChangeSet(this.#store.version, patches);
    this.#withFinalization(() => {
      let hasRowStructure = false;
      let hasColumnStructure = false;
      let hasColumnLayout = false;
      for (const patch of patches) {
        if (patch.type === "row-structure") {
          hasRowStructure = true;
        } else if (patch.type === "column-structure") {
          hasColumnStructure = true;
        } else if (patch.type === "column-layout") {
          hasColumnLayout = true;
        } else if (patch.type === "cell-value") {
          this.#measurement.forgetRow(patch.row);
        } else if (patch.type === "range-value") {
          this.#measurement.forgetRowRange(patch.range.startRow, patch.range.endRow);
        } else if (patch.type === "full-reset") {
          this.#measurement.forgetAll();
        }
      }

      if (hasRowStructure || hasColumnStructure) {
        // Undo/redo reaches this branch without going through the public
        // structure facade. Cancel before any stale numeric editor anchor
        // can commit into the newly published row/column order.
        this.#cancelActiveEdit();
        const nextStructureGeneration = this.#structureGeneration + 1;
        if (!Number.isSafeInteger(nextStructureGeneration)) {
          throw new RangeError("Structure generation exhausted");
        }
        this.#structureGeneration = nextStructureGeneration;
        this.#bodyGestures.cancel();
        this.#headerGestures.cancel(true);
        this.#contextMenuService.close();
      }

      // The model revision and all authoritative structural roots are already
      // committed. Derived bookkeeping and user formatters are finalizers too:
      // failures are observable, but can neither escape nor suppress change.
      if (hasRowStructure || hasColumnStructure || hasColumnLayout || patches.some((patch) => patch.type === "row-layout")) {
        this.#runFinalizer("spacer", () => this.#updateSpacerSize());
      }
      if (hasRowStructure || hasColumnStructure || hasColumnLayout) {
        this.#measurement.forgetAll();
      }
      if (
        hasRowStructure ||
        hasColumnStructure ||
        hasColumnLayout ||
        patches.some((patch) => patch.type === "row-layout")
      ) {
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

        const hasWrapColumns = this.#measurement.hasWrapColumns();
        const visibleUnmeasuredRows = hasWrapColumns
          ? this.#measurement.visibleUnmeasuredRows(receiptScene)
          : [];
        const rowsToMeasure = new Set(visibleUnmeasuredRows);
        if (!hasWrapColumns && hasColumnStructure) {
          // Deleting the last wrap column has to shrink auto-owned rows back;
          // they are already "measured", so the visible scan never finds them.
          for (const row of this.#measurement.rowsNeedingDefaultReset(
            receiptScene,
            this.#layout.root.rowHeightProvenance,
            this.#layout.rowHeight
          )) {
            rowsToMeasure.add(row);
          }
        }
        const rowResizeScheduled = measurementPassesRemaining > 0 && rowsToMeasure.size > 0
          ? this.#measureAndResizeRows(
              rowsToMeasure,
              receiptScene,
              measurementPassesRemaining - 1
            )
          : false;

        for (const patch of patches) {
          this.#runFinalizer("render-invalidation", () => this.#renderScheduler.invalidatePatch(patch));
        }
        if (!rowResizeScheduled) {
          // Presentation renderers are observers of this committed receipt.
          // Flush while finalization is guarded so renderer commands queue
          // until change publication; a derived row-height receipt instead
          // flushes both receipts against its final Scene.
          this.#runFinalizer("render-frame", () => this.#renderScheduler.flushNow());
        }
        if (
          measurementPassesRemaining === 0 &&
          visibleUnmeasuredRows.length > 0
        ) {
          this.#scheduleAutoHeightContinuation();
        }
      }
      this.#events.emit("change", {
        patches: Object.freeze([...patches]),
        changeSet
      });
    });
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
      rowHeightProvenance: basisLayoutRoot.rowHeightProvenance
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

      let patches: readonly Patch[];
      try {
        patches = this.#commandManager.execute(
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
      this.#invalidatePatches(patches, remainingFollowUpPasses);
      const rowLayoutPatches = Object.freeze(
        patches.filter((patch): patch is Extract<Patch, { type: "row-layout" }> =>
          patch.type === "row-layout"
        )
      );
      if (rowLayoutPatches.length > 0) {
        this.#events.emit("layout-change", {
          layoutVersion: this.#layout.version,
          patches: rowLayoutPatches
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
  #updateSpacerSize(root: LayoutRoot = this.#layout.root): void {
    const totalWidth = root.columnAxis.getTotalWidth();
    this.#spacer.style.width = `${totalWidth + this.#rowHeaderWidth}px`;
    this.#spacer.style.height = `${root.rowMetrics.totalHeight + this.#headerHeight()}px`;
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

  /** Total pixel height of the column header band: the leaf title row plus every group row (Step 19 — `groupHeaderHeight` is a per-level height, multiplied by however many levels `#headerManager` actually resolved). */
  #headerHeight(): number {
    return this.#columnHeaderHeight + this.#groupHeaderHeight * this.#headerManager.groupDepth;
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
    return {
      scrollTop: this.#scroller.scrollTop,
      scrollLeft: this.#scroller.scrollLeft,
      width: Math.max(0, this.#scroller.clientWidth - this.#rowHeaderWidth),
      height: Math.max(0, this.#scroller.clientHeight - this.#headerHeight())
    };
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
        columnHeaderBands: [
          ...Array.from({ length: this.#headerManager.groupDepth }, (_unused, index) => {
            const level = this.#headerManager.groupDepth - 1 - index;
            return {
              role: "group" as const,
              level,
              y: index * this.#groupHeaderHeight,
              height: this.#groupHeaderHeight
            };
          }),
          {
            role: "leaf" as const,
            y: this.#groupHeaderHeight * this.#headerManager.groupDepth,
            height: this.#columnHeaderHeight
          }
        ],
        rowHeaderWidth: this.#rowHeaderWidth
      }
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
    this.#scroller.scrollTop = target.scrollTop;
    this.#scroller.scrollLeft = target.scrollLeft;
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
      if (structureVersion !== this.#structureGeneration) {
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
          if (structureVersion === this.#structureGeneration) {
            this.#autoFillDownNow(range);
          }
        });
        return;
      }
      const active = selection.active;
      this.#interactionFrames.enqueueIntent(() => {
        if (structureVersion === this.#structureGeneration) {
          this.#editorManager.startEdit(active.row, active.column);
        }
      });
    });
  };

  /** Double-clicking the fill handle extends the selection down to match the nearest non-empty neighbor column's run — see computeAutoFillDownTargetRow(). */
  #autoFillDownNow(range: CellRange): void {
    const targetRow = computeAutoFillDownTargetRow(
      range,
      { columnCount: this.#columns.length, rowCount: this.#dataSource.getRowCount() },
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
    if (!renderer || renderer.kind !== "canvas" || !renderer.hitTest) {
      return declineInput();
    }

    let action: HeaderAction | void;
    try {
      action = renderer.hitTest({
        ctx: this.#headerCanvas.getContext("2d") as Canvas2DLike,
        rect: hit.rect,
        column: hit.column,
        title: header.title,
        state: {},
        point: hit.localPoint
      });
    } catch (error) {
      this.#events.emit("error", {
        error,
        eventName: `interaction:canvas-header-hit-test:${hit.column}`
      });
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
    this.#interactionFrames.enqueueIntent(() => {
      if (payload.structureVersion === this.#structureGeneration) {
        this.#events.emit("header-action", payload);
      }
    });
    event.preventDefault();
    this.#contentCanvas.focus();
    return handleInput(undefined);
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
    if (intent.structureVersion !== this.#structureGeneration) {
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
        this.#dataSource.getRowCount()
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
