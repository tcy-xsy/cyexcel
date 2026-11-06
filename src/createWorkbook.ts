import {
  CyExcelCore,
  rowsToWorkbookSnapshot,
  type CyExcelCoreOptions,
  type TopWidgetFactory
} from "./core";
import { DEFAULT_BLANK_COLUMNS } from "./core/sheet";
import { excelSourceFromSheet, exportWorkbookToExcel, formulaBarWidget, toolbarWidget } from "./core/toolbar";
import type { ExcelExportOptions, ExcelSheetExportSource, ToolbarOptions } from "./core/toolbar";
import { createFormulaEngine } from "./formula";
import { WorkbookModel } from "./core/sheet/WorkbookModel";

export interface WorkbookSheetOptions extends CyExcelCoreOptions {
  readonly name?: string;
}

export interface CreateWorkbookOptions extends CyExcelCoreOptions {
  readonly sheets?: readonly WorkbookSheetOptions[];
  readonly toolbar?: boolean | ToolbarOptions;
  readonly formulaBar?: boolean;
  /** Number of blank rows in a newly added sheet. Defaults to 1000. */
  readonly newSheetRows?: number;
}

export interface WorkbookHandle {
  getActiveSheet(): CyExcelCore;
  setActiveSheet(sheetId: string): void;
  addSheet(options?: WorkbookSheetOptions): string;
  /** Permanently delete a sheet. Throws if it is unknown or the last sheet. */
  deleteSheet(sheetId: string): void;
  /** Rename a sheet. Leading/trailing whitespace is trimmed; blank names are rejected. */
  renameSheet(sheetId: string, name: string): void;
  /** Export all sheets in tab order, including sheets that have never been opened. */
  exportExcel(options?: ExcelExportOptions): Promise<void | Blob>;
  dispose(): void;
}

interface SheetSpec {
  readonly id: string;
  readonly name: string;
  readonly options: WorkbookSheetOptions;
  readonly inheritWorkbookSnapshot: boolean;
}

interface SheetRuntime {
  readonly id: string;
  readonly name: string;
  readonly host: HTMLDivElement;
  readonly sheet: CyExcelCore;
}

function element<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  className?: string,
  text?: string
): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  if (className !== undefined) {
    node.className = className;
  }
  if (text !== undefined) {
    node.textContent = text;
  }
  return node;
}

function createTopWidgets(
  toolbar: boolean | ToolbarOptions | undefined,
  formulaBar: boolean | undefined,
  exportExcel: () => Promise<void>
): TopWidgetFactory[] {
  const widgets: TopWidgetFactory[] = [];
  if (toolbar !== false) {
    const toolbarOptions = typeof toolbar === "object" ? toolbar : {};
    widgets.push(toolbarWidget({ ...toolbarOptions, exportExcel: toolbarOptions.exportExcel ?? exportExcel }));
  }
  if (formulaBar !== false) {
    widgets.push(formulaBarWidget());
  }
  return widgets;
}

function normalizeSheetName(name: string | undefined, index: number): string {
  const trimmed = name?.trim();
  return trimmed || `工作表${index + 1}`;
}

/**
 * A workbook shell that owns tabs and delegates each visible sheet to a real
 * CyExcelCore. A core is created only while its host is visible because the
 * grid measures its viewport during construction.
 */
export function createWorkbook(
  container: HTMLElement,
  options: CreateWorkbookOptions = {}
): WorkbookHandle {
  const {
    toolbar,
    formulaBar,
    newSheetRows = 1000,
    headerHeight: defaultHeaderHeight = 28,
    rowHeaderWidth: defaultRowHeaderWidth = 48,
    ...workbookCoreOptions
  } = options;
  const defaultFormulaEngine = createFormulaEngine();
  const {
    snapshot: defaultSnapshot,
    ...defaultSheetOptions
  } = workbookCoreOptions;

  const root = element("div", "cyx-workbook");
  root.style.display = "flex";
  root.style.flexDirection = "column";
  root.style.height = "100%";

  const sheetHost = element("div", "cyx-workbook-sheets");
  sheetHost.style.position = "relative";
  sheetHost.style.flex = "1 1 auto";
  sheetHost.style.minHeight = "0";

  const tabBar = element("div", "cyx-workbook-tabs");
  tabBar.style.display = "flex";
  tabBar.style.alignItems = "center";
  tabBar.style.gap = "2px";
  tabBar.style.height = "36px";
  tabBar.style.padding = "0 6px";
  tabBar.style.borderTop = "1px solid #e0e3e8";
  tabBar.style.background = "#f8f9fb";
  tabBar.style.overflowX = "auto";

  const addSheetButton = element("button", "cyx-workbook-add-sheet", "+");
  addSheetButton.type = "button";
  addSheetButton.setAttribute("aria-label", "添加工作表");
  addSheetButton.title = "添加工作表";
  addSheetButton.style.flex = "0 0 auto";
  addSheetButton.style.width = "28px";
  addSheetButton.style.height = "28px";
  addSheetButton.style.border = "0";
  addSheetButton.style.borderRadius = "6px";
  addSheetButton.style.background = "transparent";
  addSheetButton.style.color = "#5f6b7a";
  addSheetButton.style.cursor = "pointer";
  addSheetButton.style.fontSize = "18px";
  addSheetButton.style.display = "grid";
  addSheetButton.style.placeItems = "center";
  addSheetButton.addEventListener("mouseenter", () => {
    addSheetButton.style.background = "#e9edf3";
  });
  addSheetButton.addEventListener("mouseleave", () => {
    addSheetButton.style.background = "transparent";
  });
  tabBar.append(addSheetButton);

  const specs = new Map<string, SheetSpec>();
  const sheetIdOrder: string[] = [];
  const runtimes = new Map<string, SheetRuntime>();
  const tabButtons = new Map<string, HTMLButtonElement>();
  let activeSheetId: string | null = null;
  let sheetSequence = 0;
  let disposed = false;
  const closeButtons = new Map<string, HTMLButtonElement>();

  function updateCloseButtons(): void {
    for (const button of closeButtons.values()) {
      button.disabled = sheetIdOrder.length <= 1;
      button.style.opacity = button.disabled ? "0.3" : "1";
      button.style.cursor = button.disabled ? "default" : "pointer";
    }
  }

  function setTabState(sheetId: string, active: boolean): void {
    const button = tabButtons.get(sheetId);
    if (!button) {
      return;
    }
    button.setAttribute("aria-selected", active ? "true" : "false");
    const item = button.parentElement as HTMLElement;
    item.style.background = active ? "#ffffff" : "transparent";
    item.style.color = active ? "#17223b" : "#5f6b7a";
    item.style.borderColor = active ? "#d7dbe1" : "transparent";
    item.style.fontWeight = active ? "650" : "500";
    item.style.boxShadow = active ? "0 1px 2px rgba(23, 34, 59, 0.06)" : "none";
  }

  function createTabButton(sheetId: string, name: string): HTMLDivElement {
    const item = element("div", "cyx-workbook-tab-item");
    item.style.display = "inline-flex";
    item.style.alignItems = "center";
    item.style.flex = "0 0 auto";
    item.style.border = "1px solid transparent";
    item.style.borderRadius = "6px";
    const button = element("button", "cyx-workbook-tab", name);
    button.type = "button";
    button.setAttribute("role", "tab");
    button.style.display = "inline-flex";
    button.style.alignItems = "center";
    button.style.height = "28px";
    button.style.padding = "0 6px 0 12px";
    button.style.border = "0";
    button.style.borderRadius = "6px";
    button.style.background = "transparent";
    button.style.color = "inherit";
    button.style.fontWeight = "inherit";
    button.style.fontSize = "12px";
    button.style.whiteSpace = "nowrap";
    button.style.cursor = "pointer";
    button.addEventListener("click", () => {
      setActive(sheetId);
    });
    button.title = "双击重命名工作表";
    button.addEventListener("dblclick", () => startRename(sheetId));
    button.addEventListener("keydown", event => {
      if (event.key === "F2") {
        event.preventDefault();
        startRename(sheetId);
      }
    });
    const close = element("button", "cyx-workbook-close-sheet");
    close.type = "button";
    close.setAttribute("aria-label", `删除工作表 ${name}`);
    close.title = `删除工作表 ${name}`;
    close.innerHTML = '<svg width="12" height="12" viewBox="0 0 12 12" fill="none" aria-hidden="true"><path d="m3 3 6 6m0-6-6 6" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/></svg>';
    close.style.display = "grid";
    close.style.placeItems = "center";
    close.style.width = "22px";
    close.style.height = "22px";
    close.style.marginRight = "3px";
    close.style.padding = "0";
    close.style.border = "0";
    close.style.borderRadius = "4px";
    close.style.background = "transparent";
    close.style.color = "inherit";
    close.addEventListener("mouseenter", () => {
      if (!close.disabled) close.style.background = "#e9edf3";
    });
    close.addEventListener("mouseleave", () => { close.style.background = "transparent"; });
    close.addEventListener("click", () => {
      if (disposed || sheetIdOrder.length <= 1) return;
      if (window.confirm(`删除工作表“${specs.get(sheetId)?.name}”？此操作无法撤销。`)) {
        deleteSheet(sheetId);
        if (activeSheetId) tabButtons.get(activeSheetId)?.focus();
      }
    });
    tabButtons.set(sheetId, button);
    closeButtons.set(sheetId, close);
    item.append(button, close);
    return item;
  }

  function renameSheet(sheetId: string, name: string): void {
    const spec = specs.get(sheetId);
    if (disposed || !spec) throw new RangeError(`unknown sheet "${sheetId}"`);
    const trimmed = name.trim();
    if (!trimmed) throw new RangeError("sheet name cannot be blank");
    specs.set(sheetId, { ...spec, name: trimmed });
    const runtime = runtimes.get(sheetId);
    if (runtime) runtimes.set(sheetId, { ...runtime, name: trimmed });
    const button = tabButtons.get(sheetId);
    if (button) button.textContent = trimmed;
    const close = closeButtons.get(sheetId);
    if (close) {
      close.setAttribute("aria-label", `删除工作表 ${trimmed}`);
      close.title = `删除工作表 ${trimmed}`;
    }
  }

  function startRename(sheetId: string): void {
    const button = tabButtons.get(sheetId);
    const spec = specs.get(sheetId);
    if (disposed || !button || !spec || button.hidden) return;
    const input = element("input", "cyx-workbook-rename-sheet");
    input.type = "text";
    input.value = spec.name;
    input.setAttribute("aria-label", "工作表名称");
    input.style.width = `${Math.max(80, button.offsetWidth)}px`;
    input.style.height = "26px";
    input.style.boxSizing = "border-box";
    input.style.border = "1px solid #8da5c6";
    input.style.borderRadius = "4px";
    input.style.padding = "0 6px";
    input.style.fontSize = "12px";
    let finished = false;
    const finish = (commit: boolean, focus: boolean): void => {
      if (finished) return;
      finished = true;
      if (commit && input.value.trim() && !disposed && specs.has(sheetId)) {
        renameSheet(sheetId, input.value);
      }
      input.remove();
      button.hidden = false;
      button.style.display = "inline-flex";
      if (focus) button.focus();
    };
    input.addEventListener("keydown", event => {
      if (event.key === "Enter" && !event.isComposing) {
        event.preventDefault();
        finish(true, true);
      } else if (event.key === "Escape") {
        event.preventDefault();
        finish(false, true);
      }
    });
    input.addEventListener("blur", () => finish(true, false));
    button.before(input);
    button.hidden = true;
    button.style.display = "none";
    input.focus();
    input.select();
  }

  function registerSheet(
    sheetOptions: WorkbookSheetOptions = {},
    inheritWorkbookSnapshot = false
  ): string {
    const index = sheetSequence;
    sheetSequence += 1;
    const id = `sheet-${String(index + 1).padStart(3, "0")}`;
    const name = normalizeSheetName(sheetOptions.name, index);
    const spec: SheetSpec = {
      id,
      name,
      options: sheetOptions,
      inheritWorkbookSnapshot
    };

    specs.set(id, spec);
    sheetIdOrder.push(id);
    tabBar.insertBefore(createTabButton(id, name), addSheetButton);
    updateCloseButtons();
    return id;
  }

  function resolveSheetOptions(spec: SheetSpec): CyExcelCoreOptions {
    const { name: _unusedName, snapshot: sheetSnapshot, ...overrides } = spec.options;
    const inheritedSnapshot = spec.inheritWorkbookSnapshot
      ? (sheetSnapshot ?? defaultSnapshot)
      : sheetSnapshot;
    const columnCount =
      (overrides.columns ?? options.columns ?? []).length || DEFAULT_BLANK_COLUMNS;
    return {
      ...defaultSheetOptions,
      ...overrides,
      formulaEngine: overrides.formulaEngine ?? options.formulaEngine ?? defaultFormulaEngine,
      contextMenu: overrides.contextMenu ?? options.contextMenu ?? true,
      headerHeight: overrides.headerHeight ?? defaultHeaderHeight,
      rowHeaderWidth: overrides.rowHeaderWidth ?? defaultRowHeaderWidth,
      snapshot: inheritedSnapshot ?? rowsToWorkbookSnapshot([], [], { blankRows: newSheetRows, blankColumns: columnCount })
    };
  }

  async function exportExcel(exportOptions: ExcelExportOptions = {}): Promise<void | Blob> {
    if (disposed) throw new Error("workbook is disposed");
    const sources: ExcelSheetExportSource[] = sheetIdOrder.map(id => {
      const spec = specs.get(id) as SheetSpec;
      const runtime = runtimes.get(id);
      if (runtime) return excelSourceFromSheet(runtime.sheet, spec.name);
      const sheetOptions = resolveSheetOptions(spec);
      let snapshot = sheetOptions.snapshot!;
      if (sheetOptions.formulaEngine && Object.values(snapshot.sheets).some(sheet =>
        Object.values(sheet.cellData).some(row => Object.values(row).some(cell => cell.f !== undefined)))) {
        const model = WorkbookModel.fromSnapshot(snapshot);
        model.sheet().enableFormulas(sheetOptions.formulaEngine);
        try { snapshot = model.toSnapshot(); } finally { model.sheet().disposeFormulas(); }
      }
      return {
        name: spec.name,
        snapshot,
        ...(sheetOptions.columns === undefined ? {} : { columns: sheetOptions.columns.map(column => {
          const title = (typeof column.header === "string" ? column.header : column.header?.title) ??
            (sheetOptions.columnLetterFallback === false ? "" : undefined);
          return {
            ...(title === undefined ? {} : { title }),
            ...(column.width === undefined ? {} : { width: column.width }),
            ...(column.wrap === undefined ? {} : { wrap: column.wrap })
          };
        }) })
      };
    });
    return exportWorkbookToExcel(sources, exportOptions);
  }

  function mountSheet(spec: SheetSpec): CyExcelCore {
    const existing = runtimes.get(spec.id);
    if (existing) {
      return existing.sheet;
    }

    const host = element("div", "cyx-workbook-sheet");
    host.style.position = "absolute";
    host.style.inset = "0";
    host.style.display = "flex";
    // The core measures its viewport in its constructor. Keep this append and
    // display before construction; do not move them below the constructor.
    sheetHost.append(host);

    const resolved = resolveSheetOptions(spec);
    try {
      const sheet = new CyExcelCore(host, {
        ...resolved,
        topWidgets:
          resolved.topWidgets ??
          createTopWidgets(toolbar, formulaBar, async () => { await exportExcel(); })
      });
      const runtime: SheetRuntime = { id: spec.id, name: spec.name, host, sheet };
      runtimes.set(spec.id, runtime);

      const previous = activeSheetId ? runtimes.get(activeSheetId) : undefined;
      if (previous && previous.id !== spec.id) {
        previous.host.style.display = "none";
      }
      activeSheetId = spec.id;
      for (const id of sheetIdOrder) {
        setTabState(id, id === spec.id);
      }
      sheet.resize();
      return sheet;
    } catch (error) {
      host.remove();
      throw error;
    }
  }

  function setActive(sheetId: string): void {
    const spec = specs.get(sheetId);
    if (disposed || !spec) {
      throw new RangeError(`unknown sheet "${sheetId}"`);
    }

    const existing = runtimes.get(sheetId);
    if (!existing) {
      mountSheet(spec);
      return;
    }

    const previous = activeSheetId ? runtimes.get(activeSheetId) : undefined;
    if (previous && previous.id !== sheetId) {
      previous.host.style.display = "none";
    }
    activeSheetId = sheetId;
    existing.host.style.display = "flex";
    for (const id of sheetIdOrder) {
      setTabState(id, id === sheetId);
    }
    existing.sheet.resize();
  }

  function addSheet(sheetOptions: WorkbookSheetOptions = {}): string {
    const id = registerSheet(sheetOptions, false);
    setActive(id);
    return id;
  }

  function deleteSheet(sheetId: string): void {
    const index = sheetIdOrder.indexOf(sheetId);
    if (disposed || index < 0) {
      throw new RangeError(`unknown sheet "${sheetId}"`);
    }
    if (sheetIdOrder.length === 1) {
      throw new RangeError("cannot delete the last sheet");
    }
    if (activeSheetId === sheetId) {
      // Mount the replacement first so a construction failure preserves this sheet.
      const replacementId = sheetIdOrder[index + 1] ?? sheetIdOrder[index - 1];
      setActive(replacementId as string);
    }
    runtimes.get(sheetId)?.sheet.dispose();
    runtimes.get(sheetId)?.host.remove();
    runtimes.delete(sheetId);
    specs.delete(sheetId);
    tabButtons.get(sheetId)?.parentElement?.remove();
    tabButtons.delete(sheetId);
    closeButtons.delete(sheetId);
    sheetIdOrder.splice(index, 1);
    updateCloseButtons();
  }

  addSheetButton.addEventListener("click", () => {
    addSheet();
  });

  container.style.display = "flex";
  container.style.flexDirection = "column";
  root.append(sheetHost, tabBar);
  container.append(root);

  const initialSheets = options.sheets?.length
    ? options.sheets
    : [{} as WorkbookSheetOptions];
  for (const sheetOptions of initialSheets) {
    registerSheet(sheetOptions, true);
  }
  setActive(sheetIdOrder[0] ?? "");

  return {
    getActiveSheet: () => {
      const sheetId = activeSheetId;
      if (!sheetId) {
        throw new Error("workbook has no active sheet");
      }
      const runtime = runtimes.get(sheetId);
      if (runtime) {
        return runtime.sheet;
      }
      const spec = specs.get(sheetId);
      if (!spec) {
        throw new Error("workbook has no active sheet");
      }
      return mountSheet(spec);
    },
    setActiveSheet: setActive,
    addSheet,
    deleteSheet,
    renameSheet,
    exportExcel,
    dispose: () => {
      if (disposed) {
        return;
      }
      disposed = true;
      for (const runtime of runtimes.values()) {
        runtime.sheet.dispose();
      }
      runtimes.clear();
      specs.clear();
      tabButtons.clear();
      closeButtons.clear();
      sheetIdOrder.length = 0;
      root.remove();
    }
  };
}
