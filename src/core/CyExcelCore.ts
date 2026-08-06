import {
  CellTypeRegistry,
  ColumnManager,
  DOMCellManager,
  createColumnAwarePainter,
  dispatchCellAction,
  registerBuiltinCellTypes,
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
  CommandManager
} from "./commands";
import { ArrayDataSource, TableStore, type ArrayDataSourceColumn } from "./data";
import type { CellEditorFactory } from "./editors";
import { EditorManager } from "./editors";
import { EventBus } from "./events";
import {
  DomHeaderManager,
  HeaderCanvasRenderer,
  HeaderManager,
  RowHeaderCanvasRenderer,
  type HeaderGroupDefinition,
  type RuntimeHeader
} from "./headers";
import { HookManager, PRODUCTION_HOOK_REGISTRY } from "./hooks";
import {
  FillHandleController,
  KeyboardController,
  PointerController,
  SelectionModel,
  computeFillValues,
  isPointOnFillHandle
} from "./interaction";
import { LayoutEngine, toViewportRect, type ViewportState } from "./layout";
import { ContextMenuService, classifyContextMenuTarget } from "./menu";
import { PluginManager, type CyExcelPlugin } from "./plugins";
import {
  ContentLayerRenderer,
  RenderScheduler,
  SelectionLayerRenderer,
  configureCanvasForDpr,
  type Canvas2DLike
} from "./render";
import type { CellRange, Command, DataSource, Disposer, Patch, RowData } from "./types";

export type CyExcelColumnConfig = ColumnDefinition;

export interface CyExcelCoreOptions {
  readonly columns?: readonly CyExcelColumnConfig[];
  readonly data?: readonly RowData[];
  readonly dataSource?: DataSource;
  readonly rowHeight?: number;
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
  /** Two-level column groups drawn in a row above the column titles — only shown when `headerHeight > 0`. */
  readonly headerGroups?: readonly HeaderGroupDefinition[];
  /** Pixel width of the row-number column on the left. 0 (default) means no row header band at all — same "unset is a no-op" precedent as headerHeight. */
  readonly rowHeaderWidth?: number;
  /** Draw a light gridline around every body cell. Default: true. */
  readonly showGridLines?: boolean;
  readonly clipboard?: ClipboardOptions;
  /** Installed via PluginManager at construction, in order — see doc/customization-and-hooks.md §11.2. Nothing is installed by default (e.g. createBasicContextMenuPlugin() must be listed explicitly to get a Copy/Paste right-click menu). */
  readonly plugins?: readonly CyExcelPlugin[];
}

const DEFAULT_ROW_HEIGHT = 24;
const DEFAULT_COLUMN_WIDTH = 120;
const DEFAULT_MIN_COLUMN_WIDTH = 20;
const RESIZE_HANDLE_TOLERANCE = 4;

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
  container.append(scroller);

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
  const rows: RowData[] = options.data ? options.data.map((row) => ({ ...row })) : [];
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
    columnWidths: columns.map((column) => column.width ?? DEFAULT_COLUMN_WIDTH),
    ...(options.freeze !== undefined ? { freeze: options.freeze } : {}),
    ...(options.overscanRows !== undefined ? { overscanRows: options.overscanRows } : {}),
    ...(options.overscanColumns !== undefined
      ? { overscanColumns: options.overscanColumns }
      : {})
  });
}

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

  readonly #columns: readonly CyExcelColumnConfig[];
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
  readonly #renderScheduler: RenderScheduler;
  readonly #contentRenderer: ContentLayerRenderer;
  readonly #selectionRenderer: SelectionLayerRenderer;
  readonly #headerCanvasRenderer: HeaderCanvasRenderer;
  readonly #rowHeaderCanvasRenderer: RowHeaderCanvasRenderer;
  readonly #domHeaderManager: DomHeaderManager;
  readonly #selectionModel: SelectionModel;
  readonly #pointerController: PointerController;
  readonly #fillHandleController: FillHandleController;
  readonly #keyboardController: KeyboardController;
  readonly #editorManager: EditorManager;
  readonly #domCellManager: DOMCellManager;
  readonly #clipboardService: ClipboardService;
  readonly #hooks: HookManager;
  readonly #events: EventBus;
  readonly #contextMenuService: ContextMenuService;
  readonly #pluginManager: PluginManager;
  readonly #scrollCommitPolicy: "commit" | "cancel";
  readonly #detachPointer: Disposer;
  readonly #detachFillHandle: Disposer;

  #resizeDrag: { readonly column: number; readonly startX: number; readonly startWidth: number } | null =
    null;
  #disposed = false;

  constructor(container: HTMLElement, options: CyExcelCoreOptions = {}) {
    this.#columns = options.columns ?? [];
    this.#scrollCommitPolicy = options.scrollCommitPolicy ?? "commit";

    this.#headerManager = new HeaderManager({
      columns: this.#columns,
      groups: options.headerGroups ?? []
    });
    this.#columnHeaderHeight = options.headerHeight ?? 0;
    this.#groupHeaderHeight =
      this.#columnHeaderHeight > 0 && (options.headerGroups?.length ?? 0) > 0
        ? this.#columnHeaderHeight
        : 0;
    this.#rowHeaderWidth = options.rowHeaderWidth ?? 0;

    const dom = createDom(
      container,
      this.#columnHeaderHeight + this.#groupHeaderHeight,
      this.#rowHeaderWidth
    );
    this.#scroller = dom.scroller;
    this.#spacer = dom.spacer;
    this.#headerCanvas = dom.headerCanvas;
    this.#rowHeaderCanvas = dom.rowHeaderCanvas;
    this.#contentCanvas = dom.contentCanvas;
    this.#selectionCanvas = dom.selectionCanvas;

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
    this.#events = new EventBus();
    this.#contextMenuService = new ContextMenuService({
      container: dom.menuLayer,
      hooks: this.#hooks,
      restoreFocus: () => this.#contentCanvas.focus()
    });
    this.#pluginManager = new PluginManager({
      hooks: this.#hooks,
      events: this.#events,
      contextMenu: this.#contextMenuService,
      getCell: (row, column) => this.#store.getCell(row, column),
      copy: (copyOptions) => this.#clipboardService.copy(copyOptions),
      paste: (pasteOptions) => this.#clipboardService.paste(pasteOptions),
      insertRows: (start, count) => this.insertRows(start, count),
      deleteRows: (start, count) => this.deleteRows(start, count),
      moveRows: (start, count, target) => this.moveRows(start, count, target)
    });
    for (const plugin of options.plugins ?? []) {
      this.#pluginManager.use(plugin);
    }
    this.#hooks.runSeries("beforeInit", {});

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

    this.#dataSource = buildDataSource(options, this.#columns);
    this.#store = new TableStore(this.#dataSource);
    this.#commandManager = new CommandManager(this.#store, { hooks: this.#hooks });

    this.#layout = buildLayout(options, this.#columns, this.#dataSource.getRowCount());
    this.#updateSpacerSize();

    const headerCtx = this.#headerCanvas.getContext("2d") as Canvas2DLike;
    const rowHeaderCtx = this.#rowHeaderCanvas.getContext("2d") as Canvas2DLike;
    const contentCtx = this.#contentCanvas.getContext("2d") as Canvas2DLike;
    const selectionCtx = this.#selectionCanvas.getContext("2d") as Canvas2DLike;

    this.#contentRenderer = new ContentLayerRenderer({
      ctx: contentCtx,
      layout: this.#layout,
      getCellValue: (row, column) => this.#store.getCell(row, column),
      painter: createColumnAwarePainter(this.#columnManager.getColumns()),
      ...(options.showGridLines !== undefined ? { showGridLines: options.showGridLines } : {})
    });
    this.#selectionRenderer = new SelectionLayerRenderer({
      ctx: selectionCtx,
      layout: this.#layout,
      getSelection: () => this.#selectionModel.snapshot,
      isVisible: () => !this.#editorManager.isEditing(),
      getFillPreviewRange: () => this.#fillHandleController.previewRange
    });
    this.#headerCanvasRenderer = new HeaderCanvasRenderer({
      ctx: headerCtx,
      layout: this.#layout,
      headerManager: this.#headerManager,
      columnHeaderHeight: this.#columnHeaderHeight,
      groupHeaderHeight: this.#groupHeaderHeight
    });
    this.#rowHeaderCanvasRenderer = new RowHeaderCanvasRenderer({
      ctx: rowHeaderCtx,
      layout: this.#layout,
      rowHeaderWidth: this.#rowHeaderWidth,
      headerHeight: this.#columnHeaderHeight + this.#groupHeaderHeight
    });
    this.#domHeaderManager = new DomHeaderManager({
      container: dom.headerDomLayer,
      layout: this.#layout,
      headerManager: this.#headerManager,
      columnHeaderHeight: this.#columnHeaderHeight,
      groupHeaderHeight: this.#groupHeaderHeight
    });

    this.#renderScheduler = new RenderScheduler(
      {
        paintContent: (dirty) => this.#contentRenderer.paintContent(dirty),
        paintSelection: () => this.#selectionRenderer.paintSelection()
      },
      {
        frameScheduler: {
          requestFrame: (callback) => requestAnimationFrame(callback),
          cancelFrame: (handle) => cancelAnimationFrame(handle)
        }
      }
    );

    this.#selectionModel = new SelectionModel({
      rowCount: this.#dataSource.getRowCount(),
      columnCount: this.#columns.length
    });

    // Attached (and its own pointerdown registered) BEFORE PointerController
    // below, so a handle hit can stopImmediatePropagation() and pre-empt
    // PointerController's click-to-select for that gesture — DOM listeners
    // on the same element+event type fire in registration order. See
    // FillHandleController.attach()'s doc comment.
    this.#fillHandleController = new FillHandleController({
      layout: this.#layout,
      getSelection: () => this.#selectionModel.snapshot,
      onFillPreviewChange: () => this.#renderScheduler.invalidateSelection(),
      onFillCommit: (sourceRange, targetRange) => this.#commitFill(sourceRange, targetRange)
    });
    this.#detachFillHandle = this.#fillHandleController.attach(this.#contentCanvas);

    this.#pointerController = new PointerController({
      layout: this.#layout,
      selection: this.#selectionModel,
      onSelectionChange: () => this.#renderScheduler.invalidateSelection()
    });
    this.#detachPointer = this.#pointerController.attach(this.#contentCanvas);

    this.#keyboardController = new KeyboardController({
      selection: this.#selectionModel,
      onSelectionChange: () => this.#renderScheduler.invalidateSelection()
    });

    this.#editorManager = new EditorManager({
      container: dom.editorLayer,
      layout: this.#layout,
      ...(options.editorFactory !== undefined ? { editorFactory: options.editorFactory } : {}),
      getCellValue: (row, column) => this.#store.getCell(row, column),
      commitValue: (row, column, value) => {
        this.#applyCommand(new SetCellCommand(this.#store, row, column, this.#parseEditedValue(row, column, value)));
      },
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

    this.#domCellManager = new DOMCellManager({
      container: dom.domCellLayer,
      layout: this.#layout,
      getColumn: (column) => this.#columnManager.getColumn(column),
      getCellValue: (row, column) => this.#store.getCell(row, column),
      commitValue: (row, column, value) => {
        this.#applyCommand(new SetCellCommand(this.#store, row, column, this.#parseEditedValue(row, column, value)));
      }
    });

    this.#clipboardService = new ClipboardService({
      getCellValue: (row, column) => this.#store.getCell(row, column),
      getColumn: (column) => this.#columnManager.getColumn(column),
      getRowCount: () => this.#dataSource.getRowCount(),
      getColumnCount: () => this.#columns.length,
      parseValue: (row, column, text) => this.#parseEditedValue(row, column, text),
      getSelectionRange: () => this.#selectionModel.snapshot.range,
      getStoreVersion: () => this.#store.version,
      applyRangeCommand: (range, values) => this.#applyCommand(new SetRangeCommand(this.#store, range, values)),
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

    this.#resizeCanvases();
    this.#syncViewport();
    this.#renderScheduler.invalidateAll();

    this.#contentCanvas.addEventListener("dblclick", this.#handleDblClick);
    this.#contentCanvas.addEventListener("keydown", this.#handleKeyDown);
    this.#contentCanvas.addEventListener("pointerdown", this.#handleCellHitTest);
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
    this.#headerCanvas.addEventListener("pointerdown", this.#handleHeaderPointerDown);
    // Same story for a 0-width row header canvas — see CyExcelCoreOptions.rowHeaderWidth.
    this.#rowHeaderCanvas.addEventListener("pointerdown", this.#handleRowHeaderPointerDown);
    window.addEventListener("pointermove", this.#handleHeaderPointerMove);
    window.addEventListener("pointerup", this.#handleHeaderPointerUp);

    this.#events.emit("ready", undefined);
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

  get contextMenu(): ContextMenuService {
    return this.#contextMenuService;
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
    this.#applyCommand(new SetCellCommand(this.#store, row, column, value));
  }

  /** Inserts `count` blank rows (every cell `undefined`) at `start`. To insert pre-filled rows, insert blank ones then setRange() over them. */
  insertRows(start: number, count = 1): void {
    this.#cancelActiveEdit();
    const rows = Array.from({ length: count }, () => new Array<unknown>(this.#columns.length).fill(undefined));
    this.#applyCommand(new InsertRowsCommand(this.#store, start, rows, this.#hooks));
  }

  deleteRows(start: number, count = 1): void {
    this.#cancelActiveEdit();
    this.#applyCommand(new DeleteRowsCommand(this.#store, start, count, this.#hooks));
  }

  /** `target` is the insertion index in the pre-move index space (same convention `ArrayDataSource.moveRows()` uses). */
  moveRows(start: number, count: number, target: number): void {
    this.#cancelActiveEdit();
    this.#applyCommand(new MoveRowsCommand(this.#store, start, count, target, this.#hooks));
  }

  undo(): void {
    this.#invalidatePatches(this.#commandManager.undo());
  }

  redo(): void {
    this.#invalidatePatches(this.#commandManager.redo());
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
    this.#resizeCanvases();
    this.#syncViewport();
    this.#renderScheduler.invalidateAll();
  }

  dispose(): void {
    if (this.#disposed) {
      return;
    }
    this.#disposed = true;

    this.#hooks.runSeries("beforeDestroy", {});
    this.#events.emit("destroy", undefined);

    this.#pluginManager.dispose();
    this.#contextMenuService.dispose();
    this.#editorManager.dispose();
    this.#domCellManager.dispose();
    this.#domHeaderManager.dispose();
    this.#renderScheduler.dispose();
    this.#clipboardService.dispose();
    this.#detachPointer();
    this.#detachFillHandle();

    this.#contentCanvas.removeEventListener("dblclick", this.#handleDblClick);
    this.#contentCanvas.removeEventListener("keydown", this.#handleKeyDown);
    this.#contentCanvas.removeEventListener("pointerdown", this.#handleCellHitTest);
    this.#contentCanvas.removeEventListener("copy", this.#handleCopyEvent);
    this.#contentCanvas.removeEventListener("paste", this.#handlePasteEvent);
    this.#scroller.removeEventListener("scroll", this.#handleScroll);
    this.#scroller.removeEventListener("contextmenu", this.#handleContextMenu);
    this.#headerCanvas.removeEventListener("pointerdown", this.#handleHeaderPointerDown);
    this.#rowHeaderCanvas.removeEventListener("pointerdown", this.#handleRowHeaderPointerDown);
    window.removeEventListener("pointermove", this.#handleHeaderPointerMove);
    window.removeEventListener("pointerup", this.#handleHeaderPointerUp);

    this.#hooks.dispose();
    this.#events.dispose();

    this.#scroller.remove();
  }

  #applyCommand(command: Command): void {
    this.#invalidatePatches(this.#commandManager.execute(command));
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

  #invalidatePatches(patches: readonly Patch[] | undefined): void {
    if (!patches) {
      return;
    }
    for (const patch of patches) {
      this.#renderScheduler.invalidatePatch(patch);
      if (patch.type === "row-structure") {
        this.#applyRowStructureChange(patch);
      }
    }
    // A patch may have changed a DOM-rendered cell's value; sync() is cheap
    // when nothing needs it (empty needed-set) and DOMCellManager only ever
    // holds a viewport-bounded number of instances either way.
    this.#domCellManager.sync(this.#readViewport());
    this.#events.emit("change", { patches });
  }

  /**
   * `RenderScheduler.invalidatePatch()` already treats "row-structure" as
   * "repaint everything" (content layer), but it doesn't know about
   * `LayoutEngine`/`SelectionModel` — those need to actually catch up
   * with the new row count, and the selection needs to follow whichever
   * row it was on through the insert/delete/move (see
   * `remapRowIndex()`/`SelectionModel.setBounds()`), not just get
   * silently clamped.
   */
  #applyRowStructureChange(patch: Extract<Patch, { type: "row-structure" }>): void {
    const rowCount = this.#dataSource.getRowCount();
    this.#layout.setRowCount(rowCount);
    this.#selectionModel.setBounds(
      { rowCount, columnCount: this.#columns.length },
      {
        operation: patch.operation,
        start: patch.start,
        count: patch.count,
        ...(patch.target !== undefined ? { target: patch.target } : {})
      }
    );
    this.#updateSpacerSize();
    // The selection box's on-screen position/size depends on the row it's
    // anchored to, which may have just moved — invalidateAll() (already
    // triggered by invalidatePatch() above) only marks the content layer
    // dirty, same reasoning as #syncViewport()'s identical call after a
    // scroll.
    this.#renderScheduler.invalidateSelection();
  }

  #updateSpacerSize(): void {
    this.#spacer.style.width = `${this.#layout.getTotalWidth()}px`;
    this.#spacer.style.height = `${this.#layout.getTotalHeight()}px`;
  }

  #headerHeight(): number {
    return this.#columnHeaderHeight + this.#groupHeaderHeight;
  }

  #resizeCanvases(): void {
    const rect = this.#scroller.getBoundingClientRect();
    const dpr = window.devicePixelRatio || 1;
    const headerHeight = this.#headerHeight();
    const bodyHeight = Math.max(0, rect.height - headerHeight);
    const bodyWidth = Math.max(0, rect.width - this.#rowHeaderWidth);

    configureCanvasForDpr(this.#headerCanvas, bodyWidth, headerHeight, dpr);
    configureCanvasForDpr(this.#rowHeaderCanvas, this.#rowHeaderWidth, rect.height, dpr);
    configureCanvasForDpr(this.#contentCanvas, bodyWidth, bodyHeight, dpr);
    configureCanvasForDpr(this.#selectionCanvas, bodyWidth, bodyHeight, dpr);

    this.#headerCanvas.getContext("2d")?.scale(dpr, dpr);
    this.#rowHeaderCanvas.getContext("2d")?.scale(dpr, dpr);
    this.#contentCanvas.getContext("2d")?.scale(dpr, dpr);
    this.#selectionCanvas.getContext("2d")?.scale(dpr, dpr);

    this.#selectionRenderer.setSize(bodyWidth, bodyHeight);
    this.#updateSpacerSize();
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

  #syncViewport(): void {
    const viewport = this.#readViewport();
    this.#contentRenderer.setViewport(viewport);
    this.#selectionRenderer.setViewport(viewport);
    this.#pointerController.setViewport(viewport);
    this.#fillHandleController.setViewport(viewport);
    this.#editorManager.setViewport(viewport);
    this.#domCellManager.sync(viewport);
    this.#headerCanvasRenderer.setViewport(viewport);
    this.#headerCanvasRenderer.paint();
    this.#rowHeaderCanvasRenderer.setViewport(viewport);
    this.#rowHeaderCanvasRenderer.paint();
    this.#domHeaderManager.sync(viewport);
    // The selection box's on-screen position depends on scroll position
    // just like content does — invalidateAll() (called by every caller of
    // #syncViewport(): construction, resize(), scroll) only marks the
    // content layer dirty, so without this the box would stay frozen at
    // whatever screen position it was last painted at.
    this.#renderScheduler.invalidateSelection();
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
    const rect = this.#layout.getCellRect(active.row, active.column);
    const viewport = this.#readViewport();

    if (active.row >= this.#layout.frozenRowCount) {
      const frozenHeight = this.#layout.frozenRowCount * this.#layout.rowHeight;
      const top = rect.y - viewport.scrollTop;
      const bottom = top + rect.height;
      if (top < frozenHeight) {
        this.#scroller.scrollTop = Math.max(0, rect.y - frozenHeight);
      } else if (bottom > viewport.height) {
        this.#scroller.scrollTop = Math.max(0, rect.y + rect.height - viewport.height);
      }
    }

    if (active.column >= this.#layout.frozenColumnCount) {
      const frozenWidth = this.#frozenColumnWidth();
      const left = rect.x - viewport.scrollLeft;
      const right = left + rect.width;
      if (left < frozenWidth) {
        this.#scroller.scrollLeft = Math.max(0, rect.x - frozenWidth);
      } else if (right > viewport.width) {
        this.#scroller.scrollLeft = Math.max(0, rect.x + rect.width - viewport.width);
      }
    }
  }

  #frozenColumnWidth(): number {
    const frozenColumnCount = this.#layout.frozenColumnCount;
    if (frozenColumnCount === 0) {
      return 0;
    }
    const lastFrozenColumn = frozenColumnCount - 1;
    return this.#layout.getColumnLeft(lastFrozenColumn) + this.#layout.getColumnWidth(lastFrozenColumn);
  }

  /**
   * Separate listener from PointerController's own pointerdown handling
   * (registered after it, so selection updates first) — PointerController
   * is pure geometry and shouldn't know about renderers/Commands. Only
   * acts when the clicked cell's resolved column has a CanvasCellRenderer
   * with hitTest(); a DOM-renderer cell's own mounted element receives the
   * real click natively and never reaches this listener.
   */
  #handleCellHitTest = (event: PointerEvent): void => {
    // Right-click (button 2) is handled by #handleContextMenu instead — a
    // custom renderer's click action (e.g. "increment this counter") has
    // no business firing just because the user opened the context menu.
    if (event.button !== 0) {
      return;
    }
    const canvasRect = this.#contentCanvas.getBoundingClientRect();
    const point = { x: event.clientX - canvasRect.left, y: event.clientY - canvasRect.top };
    const viewport = this.#readViewport();
    const cell = this.#layout.getCellAtPoint(point, viewport);
    if (!cell) {
      return;
    }

    const column = this.#columnManager.getColumn(cell.column);
    const renderer = column.renderer;
    if (!renderer || renderer.kind !== "canvas" || !renderer.hitTest) {
      return;
    }

    const rect = toViewportRect(this.#layout.getCellRect(cell.row, cell.column), viewport);
    const action = renderer.hitTest({
      ctx: this.#contentCanvas.getContext("2d") as Canvas2DLike,
      rect,
      row: cell.row,
      column,
      value: this.#store.getCell(cell.row, cell.column),
      point: { x: point.x - rect.x, y: point.y - rect.y }
    });
    if (!action) {
      return;
    }

    const dispatch = dispatchCellAction(action, this.#store, cell.row, cell.column);
    if (dispatch.kind === "command") {
      this.#applyCommand(dispatch.command);
    } else {
      this.#editorManager.startEdit(cell.row, cell.column);
    }
  };

  #handleDblClick = (event: MouseEvent): void => {
    if (this.#editorManager.isEditing()) {
      return;
    }
    const canvasRect = this.#contentCanvas.getBoundingClientRect();
    const point = { x: event.clientX - canvasRect.left, y: event.clientY - canvasRect.top };
    if (isPointOnFillHandle(point, this.#selectionModel.snapshot.range, this.#layout, this.#readViewport())) {
      this.#autoFillDown();
      return;
    }
    const { active } = this.#selectionModel.snapshot;
    this.#editorManager.startEdit(active.row, active.column);
  };

  /**
   * Double-clicking the fill handle extends the selection straight down to
   * match the length of the nearest non-empty neighbor column (left, or
   * right if the selection is already at column 0) — Excel's own
   * double-click-fill-handle behavior. Scans down from the selection's own
   * first row until an empty cell ends the neighbor's run.
   */
  #autoFillDown(): void {
    const { range } = this.#selectionModel.snapshot;
    const neighborColumn = range.startColumn > 0 ? range.startColumn - 1 : range.startColumn + 1;
    if (neighborColumn < 0 || neighborColumn >= this.#columns.length) {
      return;
    }

    let targetRow = range.endRow;
    for (let row = range.startRow; row < this.#dataSource.getRowCount(); row += 1) {
      const value = this.#store.getCell(row, neighborColumn);
      if (value === null || value === undefined || value === "") {
        break;
      }
      targetRow = row;
    }
    if (targetRow <= range.endRow) {
      return;
    }

    this.#commitFill(range, { ...range, endRow: targetRow });
  }

  /**
   * Applies a fill-handle result (drag or double-click auto-fill) as one
   * SetRangeCommand — a single undo step, same convention paste (Step 10)
   * and the row-structure commands (Step 12) use. `targetRange` always
   * extends `sourceRange` down or right, never both — see
   * FillHandleController. Columns with `clipboard.paste === false` keep
   * their existing values in the extended region, the same semantic
   * ClipboardService.paste() uses for read-only columns.
   */
  #commitFill(sourceRange: CellRange, targetRange: CellRange): void {
    const rows = targetRange.endRow - targetRange.startRow + 1;
    const columns = targetRange.endColumn - targetRange.startColumn + 1;
    const values: unknown[][] = Array.from({ length: rows }, () => new Array<unknown>(columns).fill(undefined));

    for (let row = sourceRange.startRow; row <= sourceRange.endRow; row += 1) {
      for (let column = sourceRange.startColumn; column <= sourceRange.endColumn; column += 1) {
        values[row - targetRange.startRow]![column - targetRange.startColumn] = this.#store.getCell(row, column);
      }
    }

    if (targetRange.endRow > sourceRange.endRow) {
      const newRowCount = targetRange.endRow - sourceRange.endRow;
      for (let column = sourceRange.startColumn; column <= sourceRange.endColumn; column += 1) {
        const relativeColumn = column - targetRange.startColumn;
        const readOnly = this.#columnManager.getColumn(column).clipboard.paste === false;
        const known = Array.from({ length: sourceRange.endRow - sourceRange.startRow + 1 }, (_unused, index) =>
          this.#store.getCell(sourceRange.startRow + index, column)
        );
        const filled = readOnly ? null : computeFillValues(known, newRowCount);
        for (let index = 0; index < newRowCount; index += 1) {
          const row = sourceRange.endRow + 1 + index;
          values[row - targetRange.startRow]![relativeColumn] = readOnly
            ? this.#store.getCell(row, column)
            : filled![index];
        }
      }
    } else if (targetRange.endColumn > sourceRange.endColumn) {
      const newColumnCount = targetRange.endColumn - sourceRange.endColumn;
      for (let row = sourceRange.startRow; row <= sourceRange.endRow; row += 1) {
        const known = Array.from(
          { length: sourceRange.endColumn - sourceRange.startColumn + 1 },
          (_unused, index) => this.#store.getCell(row, sourceRange.startColumn + index)
        );
        const filled = computeFillValues(known, newColumnCount);
        for (let index = 0; index < newColumnCount; index += 1) {
          const column = sourceRange.endColumn + 1 + index;
          const relativeColumn = column - targetRange.startColumn;
          const readOnly = this.#columnManager.getColumn(column).clipboard.paste === false;
          values[row - targetRange.startRow]![relativeColumn] = readOnly
            ? this.#store.getCell(row, column)
            : filled[index];
        }
      }
    }

    this.#applyCommand(new SetRangeCommand(this.#store, targetRange, values));
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
    const scrollerRect = this.#scroller.getBoundingClientRect();
    const point = { x: event.clientX - scrollerRect.left, y: event.clientY - scrollerRect.top };
    const context = classifyContextMenuTarget(
      point,
      this.#readViewport(),
      { rowHeaderWidth: this.#rowHeaderWidth, columnHeaderHeight: this.#headerHeight() },
      this.#layout,
      this.#selectionModel.snapshot
    );
    this.#contextMenuService.open(context);
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

    const { active } = this.#selectionModel.snapshot;

    if (event.key === "Enter") {
      event.preventDefault();
      this.#editorManager.startEdit(active.row, active.column);
      return;
    }

    // Ctrl+Z / Cmd+Z undoes; Ctrl+Shift+Z, Cmd+Shift+Z, and Ctrl+Y all redo
    // (the same three shortcuts Excel/Sheets/most editors accept).
    if ((event.ctrlKey || event.metaKey) && !event.altKey) {
      const key = event.key.toLowerCase();
      if (key === "z" && !event.shiftKey) {
        event.preventDefault();
        this.undo();
        return;
      }
      if ((key === "z" && event.shiftKey) || key === "y") {
        event.preventDefault();
        this.redo();
        return;
      }
    }

    if (event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
      event.preventDefault();
      this.#editorManager.startEdit(active.row, active.column, event.key);
      return;
    }

    if (this.#keyboardController.handleKeyDown(event.key, { shift: event.shiftKey })) {
      event.preventDefault();
      this.#scrollActiveIntoView();
    }
  };

  #handleScroll = (): void => {
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

  /**
   * Column width is clamped to [minWidth ?? DEFAULT_MIN_COLUMN_WIDTH,
   * maxWidth ?? Infinity] from that column's own ColumnDefinition. A click
   * that doesn't land on a resize handle selects the whole column instead
   * — reuses SelectionModel.selectColumn(), written in Step 6 but never
   * wired to a real listener until this row-header patch (see
   * doc/progress.md — leaving column headers non-clickable while row
   * headers select rows would be a new, confusing inconsistency).
   */
  #handleHeaderPointerDown = (event: PointerEvent): void => {
    const rect = this.#headerCanvas.getBoundingClientRect();
    const x = event.clientX - rect.left;
    const viewport = this.#readViewport();
    const column = this.#columnAtResizeHandle(x, viewport);
    if (column !== null) {
      this.#resizeDrag = { column, startX: event.clientX, startWidth: this.#layout.getColumnWidth(column) };
      this.#headerCanvas.style.cursor = "col-resize";
      return;
    }

    const hit = this.#layout.getCellAtPoint({ x, y: 0 }, viewport);
    if (hit) {
      this.#selectionModel.selectColumn(hit.column);
      this.#renderScheduler.invalidateSelection();
      this.#contentCanvas.focus();
    }
  };

  #handleHeaderPointerMove = (event: PointerEvent): void => {
    if (!this.#resizeDrag) {
      this.#updateHeaderResizeCursor(event);
      return;
    }
    const { column, startX, startWidth } = this.#resizeDrag;
    const columnConfig = this.#columns[column];
    const minWidth = columnConfig?.minWidth ?? DEFAULT_MIN_COLUMN_WIDTH;
    const maxWidth = columnConfig?.maxWidth ?? Number.POSITIVE_INFINITY;
    const nextWidth = Math.min(maxWidth, Math.max(minWidth, startWidth + (event.clientX - startX)));

    this.#layout.setColumnWidth(column, nextWidth);
    this.#updateSpacerSize();
    // Every layer whose positions depend on column offsets needs to catch
    // up — content/selection via the scheduler, the header canvas
    // directly (it isn't scheduler-driven), and both DOM overlays (a
    // DOM-rendered body or header cell in a later column just moved).
    const viewport = this.#readViewport();
    this.#renderScheduler.invalidateAll();
    this.#renderScheduler.invalidateSelection();
    this.#headerCanvasRenderer.paint();
    this.#domCellManager.sync(viewport);
    this.#domHeaderManager.sync(viewport);
  };

  #handleHeaderPointerUp = (event: PointerEvent): void => {
    this.#resizeDrag = null;
    this.#updateHeaderResizeCursor(event);
  };

  /** Shows a col-resize cursor while hovering a resize handle, clears it otherwise — the 4px handle zone (RESIZE_HANDLE_TOLERANCE) has no other visual affordance. */
  #updateHeaderResizeCursor(event: PointerEvent): void {
    const rect = this.#headerCanvas.getBoundingClientRect();
    const x = event.clientX - rect.left;
    const y = event.clientY - rect.top;
    if (x < 0 || x > rect.width || y < 0 || y > rect.height) {
      this.#headerCanvas.style.cursor = "";
      return;
    }
    const viewport = this.#readViewport();
    const overHandle = this.#columnAtResizeHandle(x, viewport) !== null;
    this.#headerCanvas.style.cursor = overHandle ? "col-resize" : "";
  }

  /**
   * A click in the corner band (y < headerHeight) selects everything;
   * otherwise selects the clicked row. `x` is forced to 0 in the
   * getCellAtPoint() call — same trick classifyContextMenuTarget() uses —
   * since only the resolved row matters here, and column 0 is always a
   * valid x for that lookup regardless of how wide the row header is.
   */
  #handleRowHeaderPointerDown = (event: PointerEvent): void => {
    const rect = this.#rowHeaderCanvas.getBoundingClientRect();
    const y = event.clientY - rect.top;
    const headerHeight = this.#headerHeight();

    if (headerHeight > 0 && y < headerHeight) {
      this.#selectionModel.selectAll();
    } else {
      const viewport = this.#readViewport();
      const hit = this.#layout.getCellAtPoint({ x: 0, y: y - headerHeight }, viewport);
      if (!hit) {
        return;
      }
      this.#selectionModel.selectRow(hit.row);
    }
    this.#renderScheduler.invalidateSelection();
    this.#contentCanvas.focus();
  };

  /** Finds the column whose right edge (translated for its own frozen status) is within RESIZE_HANDLE_TOLERANCE of `x`. Only considers frozen + currently-visible columns, same set the header actually renders. */
  #columnAtResizeHandle(x: number, viewport: ViewportState): number | null {
    const frozenColumnCount = this.#layout.frozenColumnCount;
    const visible = this.#layout.getVisibleColumnRange(viewport);

    const candidates = new Set<number>();
    for (let column = 0; column < frozenColumnCount; column += 1) {
      candidates.add(column);
    }
    for (let i = 0; i < visible.count; i += 1) {
      candidates.add(visible.start + i);
    }

    for (const column of candidates) {
      const frozen = column < frozenColumnCount;
      const left = this.#layout.getColumnLeft(column) - (frozen ? 0 : viewport.scrollLeft);
      const right = left + this.#layout.getColumnWidth(column);
      if (Math.abs(x - right) <= RESIZE_HANDLE_TOLERANCE) {
        return column;
      }
    }
    return null;
  }
}

export function createCyExcelCore(
  container: HTMLElement,
  options: CyExcelCoreOptions = {}
): CyExcelCore {
  return new CyExcelCore(container, options);
}
