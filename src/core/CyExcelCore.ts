import { SetCellCommand, CommandManager } from "./commands";
import { ArrayDataSource, TableStore, type ArrayDataSourceColumn } from "./data";
import type { CellEditorFactory } from "./editors";
import { EditorManager } from "./editors";
import { KeyboardController, PointerController, SelectionModel } from "./interaction";
import { LayoutEngine, type ViewportState } from "./layout";
import {
  ContentLayerRenderer,
  RenderScheduler,
  SelectionLayerRenderer,
  configureCanvasForDpr,
  type Canvas2DLike
} from "./render";
import type { Command, DataSource, Disposer, Patch, RowData } from "./types";

export interface CyExcelColumnConfig {
  readonly id: string;
  /** Field name in each row object; defaults to `id`. */
  readonly field?: string;
  /** Column width in logical pixels; defaults to 120. */
  readonly width?: number;
}

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
}

const DEFAULT_ROW_HEIGHT = 24;
const DEFAULT_COLUMN_WIDTH = 120;

interface DomLayer {
  readonly scroller: HTMLDivElement;
  readonly spacer: HTMLDivElement;
  readonly contentCanvas: HTMLCanvasElement;
  readonly selectionCanvas: HTMLCanvasElement;
  readonly editorLayer: HTMLDivElement;
}

function createDom(container: HTMLElement): DomLayer {
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

  const contentCanvas = document.createElement("canvas");
  contentCanvas.style.position = "absolute";
  contentCanvas.style.top = "0";
  contentCanvas.style.left = "0";
  contentCanvas.tabIndex = 0;

  const selectionCanvas = document.createElement("canvas");
  selectionCanvas.style.position = "absolute";
  selectionCanvas.style.top = "0";
  selectionCanvas.style.left = "0";
  selectionCanvas.style.pointerEvents = "none";

  const editorLayer = document.createElement("div");
  editorLayer.style.position = "absolute";
  editorLayer.style.top = "0";
  editorLayer.style.left = "0";

  stickyLayer.append(contentCanvas, selectionCanvas, editorLayer);
  // stickyLayer must come BEFORE spacer in flow order: a `position: sticky`
  // element only pins at `top/left: 0` once its natural flow position would
  // otherwise scroll past that point. If it started after the spacer (at
  // the bottom of the scrollable content), it would never need to "stick"
  // while scrolled to the top and would sit far below the viewport instead.
  scroller.append(stickyLayer, spacer);
  container.append(scroller);

  return { scroller, spacer, contentCanvas, selectionCanvas, editorLayer };
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
  readonly #contentCanvas: HTMLCanvasElement;
  readonly #selectionCanvas: HTMLCanvasElement;

  readonly #columns: readonly CyExcelColumnConfig[];
  readonly #dataSource: DataSource;
  readonly #store: TableStore;
  readonly #commandManager: CommandManager;
  readonly #layout: LayoutEngine;
  readonly #renderScheduler: RenderScheduler;
  readonly #contentRenderer: ContentLayerRenderer;
  readonly #selectionRenderer: SelectionLayerRenderer;
  readonly #selectionModel: SelectionModel;
  readonly #pointerController: PointerController;
  readonly #keyboardController: KeyboardController;
  readonly #editorManager: EditorManager;
  readonly #scrollCommitPolicy: "commit" | "cancel";
  readonly #detachPointer: Disposer;

  #disposed = false;

  constructor(container: HTMLElement, options: CyExcelCoreOptions = {}) {
    this.#columns = options.columns ?? [];
    this.#scrollCommitPolicy = options.scrollCommitPolicy ?? "commit";

    const dom = createDom(container);
    this.#scroller = dom.scroller;
    this.#spacer = dom.spacer;
    this.#contentCanvas = dom.contentCanvas;
    this.#selectionCanvas = dom.selectionCanvas;

    this.#dataSource = buildDataSource(options, this.#columns);
    this.#store = new TableStore(this.#dataSource);
    this.#commandManager = new CommandManager(this.#store);

    this.#layout = buildLayout(options, this.#columns, this.#dataSource.getRowCount());
    this.#updateSpacerSize();

    const contentCtx = this.#contentCanvas.getContext("2d") as Canvas2DLike;
    const selectionCtx = this.#selectionCanvas.getContext("2d") as Canvas2DLike;

    this.#contentRenderer = new ContentLayerRenderer({
      ctx: contentCtx,
      layout: this.#layout,
      getCellValue: (row, column) => this.#store.getCell(row, column)
    });
    this.#selectionRenderer = new SelectionLayerRenderer({ ctx: selectionCtx });

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
        this.#applyCommand(new SetCellCommand(this.#store, row, column, value));
      },
      onEditCommit: () => {
        this.#selectionModel.moveActive(1, 0);
        this.#renderScheduler.invalidateSelection();
      }
    });

    this.#resizeCanvases();
    this.#syncViewport();
    this.#renderScheduler.invalidateAll();

    this.#contentCanvas.addEventListener("dblclick", this.#handleDblClick);
    this.#contentCanvas.addEventListener("keydown", this.#handleKeyDown);
    this.#scroller.addEventListener("scroll", this.#handleScroll);
  }

  get disposed(): boolean {
    return this.#disposed;
  }

  getCell(row: number, column: number): unknown {
    return this.#store.getCell(row, column);
  }

  setCell(row: number, column: number, value: unknown): void {
    this.#applyCommand(new SetCellCommand(this.#store, row, column, value));
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

    this.#editorManager.dispose();
    this.#renderScheduler.dispose();
    this.#detachPointer();

    this.#contentCanvas.removeEventListener("dblclick", this.#handleDblClick);
    this.#contentCanvas.removeEventListener("keydown", this.#handleKeyDown);
    this.#scroller.removeEventListener("scroll", this.#handleScroll);

    this.#scroller.remove();
  }

  #applyCommand(command: Command): void {
    this.#invalidatePatches(this.#commandManager.execute(command));
  }

  #invalidatePatches(patches: readonly Patch[] | undefined): void {
    if (!patches) {
      return;
    }
    for (const patch of patches) {
      this.#renderScheduler.invalidatePatch(patch);
    }
  }

  #updateSpacerSize(): void {
    this.#spacer.style.width = `${this.#layout.getTotalWidth()}px`;
    this.#spacer.style.height = `${this.#layout.getTotalHeight()}px`;
  }

  #resizeCanvases(): void {
    const rect = this.#scroller.getBoundingClientRect();
    const dpr = window.devicePixelRatio || 1;

    configureCanvasForDpr(this.#contentCanvas, rect.width, rect.height, dpr);
    configureCanvasForDpr(this.#selectionCanvas, rect.width, rect.height, dpr);

    this.#contentCanvas.getContext("2d")?.scale(dpr, dpr);
    this.#selectionCanvas.getContext("2d")?.scale(dpr, dpr);

    this.#selectionRenderer.setSize(rect.width, rect.height);
    this.#updateSpacerSize();
  }

  #syncViewport(): void {
    const viewport: ViewportState = {
      scrollTop: this.#scroller.scrollTop,
      scrollLeft: this.#scroller.scrollLeft,
      width: this.#scroller.clientWidth,
      height: this.#scroller.clientHeight
    };
    this.#contentRenderer.setViewport(viewport);
    this.#pointerController.setViewport(viewport);
  }

  #handleDblClick = (): void => {
    if (this.#editorManager.isEditing()) {
      return;
    }
    const { active } = this.#selectionModel.snapshot;
    this.#editorManager.startEdit(active.row, active.column);
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

    if (event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
      event.preventDefault();
      this.#editorManager.startEdit(active.row, active.column, event.key);
      return;
    }

    if (this.#keyboardController.handleKeyDown(event.key, { shift: event.shiftKey })) {
      event.preventDefault();
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
}

export function createCyExcelCore(
  container: HTMLElement,
  options: CyExcelCoreOptions = {}
): CyExcelCore {
  return new CyExcelCore(container, options);
}
