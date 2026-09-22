import {
  CyExcelCore,
  rowsToWorkbookSnapshot,
  type CyExcelCoreOptions,
  type TopWidgetFactory
} from "./core";
import { DEFAULT_BLANK_COLUMNS } from "./core/sheet";
import { formulaBarWidget, toolbarWidget } from "./core/toolbar";
import type { ToolbarOptions } from "./core/toolbar";

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
  formulaBar: boolean | undefined
): TopWidgetFactory[] {
  const widgets: TopWidgetFactory[] = [];
  if (toolbar !== false) {
    widgets.push(toolbarWidget(toolbar === true ? {} : toolbar));
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

  function setTabState(sheetId: string, active: boolean): void {
    const button = tabButtons.get(sheetId);
    if (!button) {
      return;
    }
    button.setAttribute("aria-selected", active ? "true" : "false");
    button.style.background = active ? "#ffffff" : "transparent";
    button.style.color = active ? "#17223b" : "#5f6b7a";
    button.style.borderColor = active ? "#d7dbe1" : "transparent";
    button.style.fontWeight = active ? "650" : "500";
    button.style.boxShadow = active ? "0 1px 2px rgba(23, 34, 59, 0.06)" : "none";
  }

  function createTabButton(sheetId: string, name: string): HTMLButtonElement {
    const button = element("button", "cyx-workbook-tab", name);
    button.type = "button";
    button.setAttribute("role", "tab");
    button.style.display = "inline-flex";
    button.style.alignItems = "center";
    button.style.height = "28px";
    button.style.padding = "0 12px";
    button.style.border = "1px solid transparent";
    button.style.borderRadius = "6px";
    button.style.background = "transparent";
    button.style.color = "#5f6b7a";
    button.style.fontSize = "12px";
    button.style.whiteSpace = "nowrap";
    button.style.cursor = "pointer";
    button.addEventListener("click", () => {
      setActive(sheetId);
    });
    return button;
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
    const button = createTabButton(id, name);
    tabButtons.set(id, button);
    tabBar.insertBefore(button, addSheetButton);
    return id;
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

    const { name: _unusedName, snapshot: sheetSnapshot, ...overrides } = spec.options;
    const inheritedSnapshot = spec.inheritWorkbookSnapshot
      ? (sheetSnapshot ?? defaultSnapshot)
      : sheetSnapshot;
    const columnCount =
      (overrides.columns ?? options.columns ?? []).length || DEFAULT_BLANK_COLUMNS;
    const effectiveSnapshot = inheritedSnapshot ?? rowsToWorkbookSnapshot(
      [],
      [],
      { blankRows: newSheetRows, blankColumns: columnCount }
    );
    try {
      const sheet = new CyExcelCore(host, {
        ...defaultSheetOptions,
        ...overrides,
        headerHeight: overrides.headerHeight ?? defaultHeaderHeight,
        rowHeaderWidth: overrides.rowHeaderWidth ?? defaultRowHeaderWidth,
        snapshot: effectiveSnapshot,
        topWidgets:
          overrides.topWidgets ??
          options.topWidgets ??
          createTopWidgets(toolbar, formulaBar)
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
      sheetIdOrder.length = 0;
      root.remove();
    }
  };
}
