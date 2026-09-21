import {
  CyExcelCore,
  CyExcelError,
  createBasicContextMenuPlugin,
  createBooleanCellEditor,
  createDateCellEditor,
  createSelectCellEditor,
  createTextAreaCellEditor,
  rowsToWorkbookSnapshot
} from "/src/core/index.ts";

const PRODUCTS = [
  "工业温度传感器",
  "适用于高温高湿工业现场的远程温度与振动一体化监测传感器",
  "户外电源模组",
  "智能门锁套件",
  "车载充电底座",
  "空气质量探头",
  "仓储扫码终端",
  "低功耗控制器",
  "便携标签打印机"
];

const REGIONS = ["华东", "华南", "华北", "西南", "东北"];
const STATUSES = ["待处理", "已完成", "已取消"];
const TAG_POOL = ["爆款", "新品", "促销", "预售", "清仓"];

// A custom CanvasCellRenderer (Step 8) — display is independent from the
// column's CellType ("boolean"'s own format() stays the plain "true"/
// "false" it always was; this only changes how *this one column* draws,
// not the built-in type). Only touches non-edit-mode painting — the
// checkbox editor (wired below) is unaffected.
const inStockRenderer = {
  kind: "canvas",
  paint({ ctx, rect, value }) {
    ctx.font = "16px -apple-system, BlinkMacSystemFont, \"Segoe UI\", sans-serif";
    ctx.textBaseline = "middle";
    ctx.textAlign = "center";
    ctx.fillStyle = value ? "#16a34a" : "#9ca3af";
    ctx.fillText(value ? "✓" : "✗", rect.x + rect.width / 2, rect.y + rect.height / 2);
  }
};

// `type` is a real CyExcelCore ColumnDefinition field (Step 8) — it picks a
// registered CellType, which drives how the column is *drawn* (via
// format()). Every built-in CellType (src/core/cells/builtins.ts) plus the
// custom "money" type (extends "number", see CUSTOM_COLUMN_TYPES below) is
// represented here so the playground actually shows what's implemented.
// Step 19（Univer 参考）：columnLetterFallback 打开后，没配 header 的列显示
// 电子表格字母（B/D/E/...），配了 header 的列照常显示自定义标题——同一行，
// 字母和自定义标题二选一，不是两行堆叠。这里故意只给一半列配 header，让
// playground 里同时看到两种效果。
const COLUMNS = [
  // sortable:false on the column itself — demonstrates the per-column override.
  // The "标签" (array) column is unsortable too, but by its *type*'s own
  // declaration rather than here, which is the other half of the mechanism.
  { width: 140, type: "text", label: "SKU", header: "SKU", sortable: false },
  { width: 200, type: "text", label: "产品名称", wrap: true },
  { width: 110, type: "money", label: "单价", header: "单价" },
  { width: 90, type: "number", label: "数量" },
  { width: 110, type: "text", label: "地区" },
  // 数据类型（boolean/date/select…）和编辑控件是两个独立维度，不会被自动绑定——
  // 这三列显式通过 createEditor 接上对应控件，是业务方要自己做的事，不是"boolean
  // 类型天生等于 checkbox"。不接的类型（下面的 tags/meta）仍然用默认文本编辑器，
  // 靠 CellType.parse()/serialize() 保证往返正确。
  {
    width: 90,
    type: "boolean",
    label: "有库存",
    header: "有库存",
    renderer: inStockRenderer,
    createEditor: () => createBooleanCellEditor()
  },
  {
    width: 130,
    type: "date",
    label: "入库日期",
    createEditor: () => createDateCellEditor()
  },
  {
    width: 110,
    type: "select",
    label: "状态",
    header: "状态",
    options: STATUSES.map((value) => ({ value })),
    createEditor: (context) => createSelectCellEditor(context.column.options ?? [])
  },
  {
    width: 220,
    type: "array",
    label: "标签",
    createEditor: () =>
      createTextAreaCellEditor({
        validate: (text) => {
          try {
            const parsed = JSON.parse(text);
            return Array.isArray(parsed) ? { valid: true } : { valid: false, message: "需要 JSON 数组，例如 [\"a\",\"b\"]" };
          } catch {
            return { valid: false, message: "不是合法 JSON" };
          }
        }
      })
  },
  {
    width: 240,
    type: "json",
    label: "元数据",
    createEditor: () =>
      createTextAreaCellEditor({
        validate: (text) => {
          try {
            JSON.parse(text);
            return { valid: true };
          } catch {
            return { valid: false, message: "不是合法 JSON" };
          }
        }
      })
  }
];

// A custom ColumnType extending the built-in "number" — the "price" column
// above references it by name. Demonstrates Step 8's registerColumnType-
// style extends chain: format() overridden, everything else ("parse"/
// "compare") inherited from "number".
const CUSTOM_COLUMN_TYPES = {
  money: {
    extends: "number",
    format: (value) =>
      typeof value === "number" && Number.isFinite(value)
        ? `¥${value.toFixed(2)}`
        : ""
  }
};

function buildRow(index) {
  const quantity = ((index * 7) % 24) + 1;
  const price = Math.round((48 + ((index * 37) % 520) + (index % 3) * 0.5) * 100) / 100;
  const tagCount = (index % 3) + 1;
  const tags = Array.from(
    { length: tagCount },
    (_unused, tagIndex) => TAG_POOL[(index + tagIndex) % TAG_POOL.length]
  );
  return {
    sku: `CY-${String(index + 1).padStart(6, "0")}`,
    name: PRODUCTS[index % PRODUCTS.length],
    price,
    quantity,
    region: REGIONS[index % REGIONS.length],
    inStock: index % 4 !== 0,
    restockedAt: new Date(Date.UTC(2026, 0, 1 + (index % 365))),
    status: STATUSES[index % STATUSES.length],
    tags,
    meta: { sku: `CY-${index + 1}`, warehouse: `W${(index % 5) + 1}` }
  };
}

function buildRows(count) {
  return Array.from({ length: count }, (_unused, index) => buildRow(index));
}

/**
 * 第 i 个字段就是第 i 列——导入那一刻的参数，不是列配置的一部分
 * （ADR 0004 §3.2、§4.5）。顺序必须与 COLUMNS 一致。
 */
const FIELDS = [
  "sku",
  "name",
  "price",
  "quantity",
  "region",
  "inStock",
  "restockedAt",
  "status",
  "tags",
  "meta"
];

function clampNumber(value, min, max, fallback) {
  if (!Number.isFinite(value)) {
    return fallback;
  }
  return Math.min(max, Math.max(min, Math.round(value)));
}

/** Renders a cell value back into the "读取" text field, per column type. */
function displayValue(value) {
  if (value === undefined || value === null) {
    return "";
  }
  if (value instanceof Date) {
    return value.toISOString().slice(0, 10);
  }
  if (typeof value === "object") {
    return JSON.stringify(value);
  }
  return String(value);
}

/**
 * Parses the "写入" text field into the value shape each CellType expects —
 * mirrors what a real CellType.parse()/editor would do (only "number" has a
 * built-in parse(); the rest are playground-local since there's no other
 * editor wired to these types yet).
 */
function parseInputForType(type, raw) {
  switch (type) {
    case "number":
    case "money": {
      const parsed = Number(raw);
      return Number.isNaN(parsed) ? { error: "需要数字" } : { value: parsed };
    }
    case "boolean":
      return { value: raw === "true" || raw === "1" };
    case "date": {
      const date = new Date(raw);
      return Number.isNaN(date.getTime()) ? { error: "需要合法日期，例如 2026-08-05" } : { value: date };
    }
    case "array":
      return {
        value: raw
          .split(",")
          .map((item) => item.trim())
          .filter(Boolean)
      };
    case "json":
      try {
        return { value: JSON.parse(raw) };
      } catch {
        return { error: "不是合法 JSON" };
      }
    default:
      return { value: raw };
  }
}

const sheetMount = document.getElementById("sheetMount");
const statusBar = document.getElementById("statusBar");
const themeButton = document.getElementById("themeButton");
const undoButton = document.getElementById("undoButton");
const redoButton = document.getElementById("redoButton");
const rowCountSelect = document.getElementById("rowCountSelect");
const rowHeightInput = document.getElementById("rowHeightInput");
const overscanRowsInput = document.getElementById("overscanRowsInput");
const overscanColumnsInput = document.getElementById("overscanColumnsInput");
const rebuildButton = document.getElementById("rebuildButton");
const cellRow = document.getElementById("cellRow");
const cellColumn = document.getElementById("cellColumn");
const cellValue = document.getElementById("cellValue");
const readCellButton = document.getElementById("readCellButton");
const writeCellButton = document.getElementById("writeCellButton");
const cellFormResult = document.getElementById("cellFormResult");
const mergeStartRow = document.getElementById("mergeStartRow");
const mergeStartColumn = document.getElementById("mergeStartColumn");
const mergeEndRow = document.getElementById("mergeEndRow");
const mergeEndColumn = document.getElementById("mergeEndColumn");
const mergeCellsButton = document.getElementById("mergeCellsButton");
const unmergeCellsButton = document.getElementById("unmergeCellsButton");
const mergeFormResult = document.getElementById("mergeFormResult");
const sortColumn = document.getElementById("sortColumn");
const sortAscButton = document.getElementById("sortAscButton");
const sortDescButton = document.getElementById("sortDescButton");
const sortHint = document.getElementById("sortHint");
const filterColumn = document.getElementById("filterColumn");
const filterValue = document.getElementById("filterValue");
const applyFilterButton = document.getElementById("applyFilterButton");
const clearFilterButton = document.getElementById("clearFilterButton");
const filterHint = document.getElementById("filterHint");
const searchQuery = document.getElementById("searchQuery");
const searchMatchCase = document.getElementById("searchMatchCase");
const searchUseRegex = document.getElementById("searchUseRegex");
const searchFindButton = document.getElementById("searchFindButton");
const searchPrevButton = document.getElementById("searchPrevButton");
const searchNextButton = document.getElementById("searchNextButton");
const searchReplaceValue = document.getElementById("searchReplaceValue");
const searchReplaceButton = document.getElementById("searchReplaceButton");
const searchReplaceAllButton = document.getElementById("searchReplaceAllButton");
const searchClearButton = document.getElementById("searchClearButton");
const searchHint = document.getElementById("searchHint");
const metricConstruct = document.getElementById("metricConstruct");
const metricFps = document.getElementById("metricFps");

let core = null;
let currentRowCount = 0;
let statusBarDefaultText = "";
let statusBarErrorTimeoutId = null;

function populateColumnSelect(select, accepts = () => true) {
  select.innerHTML = "";
  COLUMNS.forEach((column, index) => {
    if (!accepts(index)) {
      return;
    }
    const option = document.createElement("option");
    option.value = String(index);
    // 列没有 id 了（ADR 0004 §3.8.2），一列就是它的下标——所有公开 API
    // 吃的也正是这个数字，所以括号里显示的比原来的 id 更贴近调用方要传的东西。
    option.textContent = `${column.label}（第 ${index} 列）`;
    select.append(option);
  });
}

function populateCellColumnSelect() {
  populateColumnSelect(cellColumn);
}

/**
 * A host building its own sort UI asks the sheet rather than re-deriving the
 * rule: `sortable` can come from the column *or* from its cell type (`array`
 * and `json` declare themselves unsortable), and only the resolved
 * RuntimeColumn knows which one won.
 */
function populateSortColumnSelect() {
  populateColumnSelect(sortColumn, (index) => core.getColumn(index).sortable);
}

/** Same reasoning as populateSortColumnSelect(): filterable can come from the column or its cell type. */
function populateFilterColumnSelect() {
  populateColumnSelect(filterColumn, (index) => core.getColumn(index).filterable);
}

function readConfigFromForm() {
  return {
    rowCount: Number(rowCountSelect.value),
    rowHeight: clampNumber(Number(rowHeightInput.value), 16, 80, 28),
    overscanRows: clampNumber(Number(overscanRowsInput.value), 0, 50, 4),
    overscanColumns: clampNumber(Number(overscanColumnsInput.value), 0, 20, 2)
  };
}

function rebuild() {
  const config = readConfigFromForm();
  rowHeightInput.value = String(config.rowHeight);
  overscanRowsInput.value = String(config.overscanRows);
  overscanColumnsInput.value = String(config.overscanColumns);

  if (core) {
    core.dispose();
    core = null;
  }
  if (statusBarErrorTimeoutId !== null) {
    clearTimeout(statusBarErrorTimeoutId);
    statusBarErrorTimeoutId = null;
  }
  statusBar.classList.remove("status-bar--error");
  // A fresh instance starts unfiltered — a stale "0 rows matched" overlay
  // left over from the previous instance would be actively wrong, not just
  // out of date (unlike sortHint's text, which is merely stale until the
  // next sort).
  emptyFilterState.hidden = true;
  // Same reasoning: a fresh instance has no search state either, and
  // currentMatch would otherwise point at coordinates in the previous
  // instance's now-disposed sheet.
  currentMatch = null;

  const rows = buildRows(config.rowCount);
  // Pass every real ColumnDefinition field through as-is — only "label" is
  // playground-local metadata (used by the read/write form below), so strip
  // just that one instead of hand-picking fields (which previously silently
  // dropped "options"/"createEditor" and made the typed-editor columns fall
  // back to the plain text editor).
  const columnConfigs = COLUMNS.map((column) =>
    Object.fromEntries(Object.entries(column).filter(([key]) => key !== "label"))
  );

  const start = performance.now();
  core = new CyExcelCore(sheetMount, {
    columns: columnConfigs,
    // ADR 0004 §3.2：入口只吃一份快照。对象数组在这里导进稀疏矩阵，导完这
    // 一遍之后 `rows` 与表格再无联系——改表格不会改到它，改它也不会改表格。
    snapshot: rowsToWorkbookSnapshot(rows, FIELDS),
    rowHeight: config.rowHeight,
    overscanRows: config.overscanRows,
    overscanColumns: config.overscanColumns,
    columnTypes: CUSTOM_COLUMN_TYPES,
    headerHeight: 28,
    rowHeaderWidth: 48,
    // Step 18：点行号选中整行后，把鼠标移到选区的上/下边缘（会出现一条插入
    // 指示线）按住拖动即可重新排列行顺序——插入式移动，不会覆盖/丢失目标位置
    // 原有数据；边缘之外的拖拽仍然是原有的多选。
    rowReorder: true,
    // Univer 参考实现：没配 header 的列显示电子表格字母（B/D/E/...），配了
    // header 的列（sku/price/inStock/status）照常显示自定义标题——同一行
    // 二选一，不是字母行 + 标题行两行堆叠。
    columnLetterFallback: true,
    plugins: [createBasicContextMenuPlugin()]
  });
  const elapsed = performance.now() - start;

  registerPlaygroundShortcuts(core);

  currentRowCount = config.rowCount;
  metricConstruct.textContent = `${elapsed.toFixed(1)} ms`;
  statusBarDefaultText =
    `${currentRowCount.toLocaleString("zh-CN")} 行 × ${COLUMNS.length} 列 · ` +
    `行高 ${config.rowHeight}px · overscan ${config.overscanRows}/${config.overscanColumns}`;
  statusBar.textContent = statusBarDefaultText;

  // 右键菜单 action 抛错现在会通过 core.events 的 "error" 事件转发出来——
  // 之前只有 console.error，用户在页面上完全看不出"点了没反应"到底是不是
  // 出错了。这里临时把错误信息顶到 statusBar 上几秒钟，然后再切回正常的
  // 行列统计文案。
  core.events.on("error", ({ error, eventName }) => {
    if (statusBarErrorTimeoutId !== null) {
      clearTimeout(statusBarErrorTimeoutId);
    }
    const message = error instanceof Error ? error.message : String(error);
    statusBar.textContent = `⚠ ${eventName}：${message}`;
    statusBar.classList.add("status-bar--error");
    statusBarErrorTimeoutId = setTimeout(() => {
      statusBar.textContent = statusBarDefaultText;
      statusBar.classList.remove("status-bar--error");
      statusBarErrorTimeoutId = null;
    }, 5000);
  });

  // Step 22 Part A：表头右侧窄条也能直接触发排序（HeaderCanvasRenderer 的
  // defaultHeaderHitTest），不只是上面的升序/降序按钮。header-action 事件在
  // 那条路径下触发，这里同步 sortHint 和排序列下拉框，让两种触发方式在界面
  // 上看起来是同一件事——core.sortState 是两条路径共用的同一份状态。
  core.events.on("header-action", (event) => {
    if (event.action.type !== "toggle-sort") {
      return;
    }
    const state = core.sortState;
    if (!state) {
      return;
    }
    const column = COLUMNS[state.column];
    sortColumn.value = String(state.column);
    sortHint.textContent =
      `上次动作：点击表头「${column.label}」排序为` +
      `${state.direction === "asc" ? "升序" : "降序"}（共 ` +
      `${currentRowCount.toLocaleString("zh-CN")} 行，撤销可一步还原顺序和行高）。` +
      "再点一次表头在升/降序间切换，不会回到原始顺序";
  });

  // Step 22 Part B：表头点击 toggle-filter 是纯通知——core 不知道该用什么谓词，
  // 这里（宿主）拿"过滤值"输入框当前的值对被点的那一列应用过滤，让表头点击和
  // 下面的按钮看起来是同一件事。
  core.events.on("header-action", (event) => {
    if (event.action.type !== "toggle-filter") {
      return;
    }
    filterColumn.value = String(event.column);
    filterBy("apply");
  });

  const maxRow = Math.max(0, currentRowCount - 1);
  if (Number(cellRow.value) > maxRow) {
    cellRow.value = String(maxRow);
  }
  cellRow.max = String(maxRow);
  // Only now: the answer comes off the resolved RuntimeColumn, which does not
  // exist until this rebuild has constructed the sheet.
  populateSortColumnSelect();
  populateFilterColumnSelect();
}

const shortcutHint = document.getElementById("shortcutHint");
const exportButton = document.getElementById("exportButton");

/** Shared by the toolbar button and the Ctrl/Cmd+S shortcut below — same call, two triggers. */
function exportSelection(sheetCore) {
  return sheetCore.copy({ download: true }).then(() => {
    if (shortcutHint) {
      shortcutHint.textContent =
        "上次动作：把当前选区导出成 TSV 文件下载了（sheet.copy({ download: true })）";
    }
  });
}

/**
 * Step 18：sheet.registerShortcut() 演示——每次 rebuild() 重建 core 都要重新
 * 注册一遍（dispose() 会清空旧实例的 ShortcutRegistry，不需要手动注销旧的）。
 */
function registerPlaygroundShortcuts(sheetCore) {
  sheetCore.registerShortcut("Mod+K", () => {
    themeButton.click();
    if (shortcutHint) {
      shortcutHint.textContent = "上次动作：Ctrl/Cmd+K 切换了主题（registerShortcut 演示）";
    }
  });
  sheetCore.registerShortcut("Mod+S", () => {
    void exportSelection(sheetCore);
  });
  exportButton.onclick = () => {
    void exportSelection(sheetCore);
  };
}

undoButton.addEventListener("click", () => {
  core?.undo();
});

redoButton.addEventListener("click", () => {
  core?.redo();
});

readCellButton.addEventListener("click", () => {
  if (!core) {
    return;
  }
  const row = clampNumber(Number(cellRow.value), 0, Math.max(0, currentRowCount - 1), 0);
  cellRow.value = String(row);
  const columnIndex = Number(cellColumn.value);
  const value = core.getCell(row, columnIndex);
  cellValue.value = displayValue(value);
  cellFormResult.textContent = `已读取 (${row}, ${columnIndex}) = ${JSON.stringify(value)}`;
});

writeCellButton.addEventListener("click", () => {
  if (!core) {
    return;
  }
  const row = clampNumber(Number(cellRow.value), 0, Math.max(0, currentRowCount - 1), 0);
  cellRow.value = String(row);
  const columnIndex = Number(cellColumn.value);
  const column = COLUMNS[columnIndex];
  const raw = cellValue.value;

  const result = parseInputForType(column.type, raw);
  if (result.error) {
    cellFormResult.textContent = `写入失败：${column.label} ${result.error}`;
    return;
  }
  core.setCell(row, columnIndex, result.value);
  cellFormResult.textContent =
    `已写入 (${row}, ${columnIndex}) = ${JSON.stringify(result.value)}，可以用撤销按钮回退`;
});

/** Reads the four merge-form inputs into a CellRange. */
function readMergeRangeFromForm() {
  return {
    startRow: clampNumber(Number(mergeStartRow.value), 0, Math.max(0, currentRowCount - 1), 0),
    startColumn: clampNumber(Number(mergeStartColumn.value), 0, COLUMNS.length - 1, 0),
    endRow: clampNumber(Number(mergeEndRow.value), 0, Math.max(0, currentRowCount - 1), 0),
    endColumn: clampNumber(Number(mergeEndColumn.value), 0, COLUMNS.length - 1, 0)
  };
}

mergeCellsButton.addEventListener("click", () => {
  if (!core) {
    return;
  }
  const range = readMergeRangeFromForm();
  try {
    core.mergeCells(range);
    mergeFormResult.textContent =
      `已合并 (${range.startRow}, ${range.startColumn}) - (${range.endRow}, ${range.endColumn})，` +
      "可以用撤销按钮回退";
  } catch (error) {
    if (error instanceof CyExcelError) {
      mergeFormResult.textContent = `合并失败：${error.code} — ${error.message}`;
    } else {
      throw error;
    }
  }
});

unmergeCellsButton.addEventListener("click", () => {
  if (!core) {
    return;
  }
  const range = readMergeRangeFromForm();
  try {
    core.unmergeCells(range);
    mergeFormResult.textContent =
      `已取消 (${range.startRow}, ${range.startColumn}) - (${range.endRow}, ${range.endColumn}) ` +
      "范围内的合并，可以用撤销按钮回退";
  } catch (error) {
    if (error instanceof CyExcelError) {
      mergeFormResult.textContent = `取消合并失败：${error.code} — ${error.message}`;
    } else {
      throw error;
    }
  }
});

/**
 * Step 22：sheet.sortByColumn() 演示——物理重排整根行轴。数据顺序和行高在同一个
 * transaction 里提交，所以给某一行设过的行高会跟着那一行走，一次撤销两者一起回退。
 */
function sortBy(direction) {
  if (!core) {
    return;
  }
  const columnIndex = Number(sortColumn.value);
  const column = COLUMNS[columnIndex];
  const started = performance.now();
  const applied = core.sortByColumn(columnIndex, direction);
  const elapsed = Math.round(performance.now() - started);
  sortHint.textContent = applied
    ? `上次动作：按「${column.label}」${direction === "asc" ? "升序" : "降序"}排序 ` +
      `${currentRowCount.toLocaleString("zh-CN")} 行，用时 ${elapsed}ms（撤销可一步还原顺序和行高）`
    : "上次动作：排序被取消，什么都没写入";
}

sortAscButton.addEventListener("click", () => sortBy("asc"));
sortDescButton.addEventListener("click", () => sortBy("desc"));

const emptyFilterState = document.getElementById("emptyFilterState");

/**
 * Step 22 Part B：sheet.filterByColumn()/clearFilter() 演示——预测值是这里
 * 手写的精确字符串匹配（`String(value) === 输入框的值`），不是核心提供的默认
 * 值列表 UI（那一层特意留给宿主，见 filterHint 的说明）。
 *
 * 匹配行数不额外扫描一遍——core 本来就会对每一行调用一次谓词，这里在同一个
 * 谓词闭包里顺手计数，比"先自己扫一遍数、再调 filterByColumn 让它再扫一遍"
 * 省一半开销，10 万行规模下这个差别是看得见的。
 */
function filterBy(kind) {
  if (!core) {
    return;
  }
  if (kind === "clear") {
    core.clearFilter();
    filterHint.textContent = "上次动作：已清除过滤，所有行的高度逐像素复原";
    emptyFilterState.hidden = true;
    return;
  }
  const columnIndex = Number(filterColumn.value);
  const column = COLUMNS[columnIndex];
  const target = filterValue.value;
  let matched = 0;
  const started = performance.now();
  core.filterByColumn(columnIndex, (value) => {
    const hit = String(value ?? "") === target;
    if (hit) {
      matched += 1;
    }
    return hit;
  });
  const elapsed = Math.round(performance.now() - started);
  filterHint.textContent = matched === 0
    ? `上次动作：按「${column.label}」等于「${target}」过滤，没有任何行匹配（用时 ${elapsed}ms）——` +
      "表格看起来是空的，这是预期行为，不是加载失败，清除过滤即可恢复"
    : `上次动作：只看「${column.label}」等于「${target}」的行，共匹配 ` +
      `${matched.toLocaleString("zh-CN")} / ${currentRowCount.toLocaleString("zh-CN")} 行，` +
      `用时 ${elapsed}ms（过滤不进撤销栈，清除过滤即可复原）`;
  emptyFilterState.hidden = matched !== 0;
}

applyFilterButton.addEventListener("click", () => filterBy("apply"));
clearFilterButton.addEventListener("click", () => filterBy("clear"));

/**
 * Step 23：sheet.find()/findNext()/findPrevious()/replace()/replaceAll()
 * 演示——匹配数量与"第几个"直接读 core 返回的结果，这里不重复统计。
 * `currentMatch` 只是宿主自己记的"上一次 findNext/findPrevious 停在哪"，
 * core 本身不暴露这个 getter（它已经体现在选区上了），「替换当前」要用
 * 具体坐标调 replace()，所以这里记一份。
 */
let currentMatch = null;

async function runSearchFind() {
  if (!core) {
    return;
  }
  const query = searchQuery.value;
  if (query === "") {
    searchHint.textContent = "查找框是空的——请先输入要找的文本";
    return;
  }
  currentMatch = null;
  const started = performance.now();
  const matches = await core.find(query, {
    matchCase: searchMatchCase.checked,
    useRegex: searchUseRegex.checked
  });
  const elapsed = Math.round(performance.now() - started);
  if (matches === null) {
    searchHint.textContent = "上次动作：查找已取消";
    return;
  }
  searchHint.textContent = matches.length === 0
    ? `上次动作：查找「${query}」，没有匹配（用时 ${elapsed}ms）`
    : `上次动作：查找「${query}」，共 ${matches.length.toLocaleString("zh-CN")} 个匹配，` +
      `用时 ${elapsed}ms——点「下一个」/「上一个」在匹配间跳转`;
}

function runSearchStep(direction) {
  if (!core) {
    return;
  }
  const match = direction === "next" ? core.findNext() : core.findPrevious();
  currentMatch = match;
  searchHint.textContent = match === null
    ? "上次动作：没有可跳转的匹配——请先点「查找」"
    : `上次动作：跳到第 ${match.row} 行第 ${match.column} 列的匹配`;
}

function runSearchReplaceCurrent() {
  if (!core || !currentMatch) {
    searchHint.textContent = "上次动作：还没有当前匹配——先点「查找」再点「下一个」定位一个";
    return;
  }
  const replaced = currentMatch;
  core.replace(replaced, searchReplaceValue.value);
  searchHint.textContent =
    `上次动作：把第 ${replaced.row} 行第 ${replaced.column} 列替换为「${searchReplaceValue.value}」（可撤销）`;
  currentMatch = null;
}

searchFindButton.addEventListener("click", () => {
  runSearchFind();
});
searchNextButton.addEventListener("click", () => runSearchStep("next"));
searchPrevButton.addEventListener("click", () => runSearchStep("previous"));
searchReplaceButton.addEventListener("click", runSearchReplaceCurrent);
searchClearButton.addEventListener("click", () => {
  if (!core) {
    return;
  }
  core.clearSearch();
  currentMatch = null;
  searchHint.textContent = "上次动作：已清除查找结果";
});
searchReplaceAllButton.addEventListener("click", async () => {
  if (!core) {
    return;
  }
  const query = searchQuery.value;
  if (query === "") {
    searchHint.textContent = "查找框是空的——请先输入要找的文本";
    return;
  }
  const started = performance.now();
  const count = await core.replaceAll(query, searchReplaceValue.value, {
    matchCase: searchMatchCase.checked,
    useRegex: searchUseRegex.checked
  });
  const elapsed = Math.round(performance.now() - started);
  currentMatch = null;
  searchHint.textContent = count === null
    ? "上次动作：全部替换已取消"
    : `上次动作：全部替换「${query}」→「${searchReplaceValue.value}」，共 ` +
      `${count.toLocaleString("zh-CN")} 个格子，用时 ${elapsed}ms（一次撤销即可复原）`;
});

rebuildButton.addEventListener("click", rebuild);

themeButton.addEventListener("click", () => {
  const dark = document.documentElement.dataset.theme === "dark";
  document.documentElement.dataset.theme = dark ? "light" : "dark";
  themeButton.textContent = dark ? "深色" : "浅色";
});

// CyExcelCore has no automatic ResizeObserver — resize() must be called
// explicitly whenever the mount container's size changes (documented risk
// in doc/progress.md's Step 7 entry).
const resizeObserver = new ResizeObserver(() => {
  core?.resize();
});
resizeObserver.observe(sheetMount);

let frameCount = 0;
let lastSample = performance.now();
function tick(now) {
  frameCount += 1;
  const elapsed = now - lastSample;
  if (elapsed >= 500) {
    metricFps.textContent = `${((frameCount * 1000) / elapsed).toFixed(0)} fps`;
    frameCount = 0;
    lastSample = now;
  }
  requestAnimationFrame(tick);
}
requestAnimationFrame(tick);

/**
 * Step 24：样式面板。
 *
 * 这一组按钮是本步**主要的验证路径**——渲染效果（字体、颜色、边框画出来对不对）
 * 只能靠眼睛，这个环境里没有 Chromium 跑浏览器测试。所以面板要能单独验证每一个
 * 样式键，外加一块显示 intern 状态的调试信息。
 */
const styleBold = document.getElementById("styleBold");
const styleItalic = document.getElementById("styleItalic");
const styleUnderline = document.getElementById("styleUnderline");
const styleStrike = document.getElementById("styleStrike");
const styleFontSize = document.getElementById("styleFontSize");
const styleColor = document.getElementById("styleColor");
const styleBackground = document.getElementById("styleBackground");
const styleAlign = document.getElementById("styleAlign");
const styleBorder = document.getElementById("styleBorder");
const styleNumFmt = document.getElementById("styleNumFmt");
const styleApplyNumFmt = document.getElementById("styleApplyNumFmt");
const styleClear = document.getElementById("styleClear");
const styleHint = document.getElementById("styleHint");
const styleDebug = document.getElementById("styleDebug");

function applyStyle(patch, label) {
  if (!core) {
    return;
  }
  const range = core.selection.range;
  try {
    core.setRangeStyle(range, patch);
    styleHint.textContent =
      `上次动作：${label}，作用于 R${range.startRow}C${range.startColumn}` +
      `–R${range.endRow}C${range.endColumn}（一次撤销即可全部回退）`;
  } catch (error) {
    // 非法 numFmt 走这里——错误在设置时抛出，一格都没写进去
    styleHint.textContent = `上次动作失败：${error instanceof Error ? error.message : String(error)}`;
  }
  updateStyleDebug();
}

/**
 * toggle 语义：当前格已经是粗体就取消，否则加上。这正是"合并语义 + null 清键"
 * 存在的理由——工具栏按钮只能表达"改这一个键"，不能整体替换样式。
 */
function toggleStyle(key, label) {
  if (!core) {
    return;
  }
  const { startRow, startColumn } = core.selection.range;
  const current = core.getCellStyle(startRow, startColumn);
  applyStyle({ [key]: current?.[key] === true ? null : true }, label);
}

function updateStyleDebug() {
  if (!core || !styleDebug) {
    return;
  }
  const { startRow, startColumn } = core.selection.range;
  const style = core.getCellStyle(startRow, startColumn);
  const summary = style === undefined ? "无样式" : JSON.stringify(style);
  styleDebug.textContent = `当前格 R${startRow}C${startColumn}：${summary}`;
}

styleBold?.addEventListener("click", () => toggleStyle("bold", "切换加粗"));
styleItalic?.addEventListener("click", () => toggleStyle("italic", "切换斜体"));
styleUnderline?.addEventListener("click", () => toggleStyle("underline", "切换下划线"));
styleStrike?.addEventListener("click", () => toggleStyle("strike", "切换删除线"));
styleFontSize?.addEventListener("change", () => {
  const size = Number(styleFontSize.value);
  applyStyle({ fontSize: Number.isFinite(size) && size > 0 ? size : null }, `字号 ${size}`);
});
styleColor?.addEventListener("change", () => applyStyle({ color: styleColor.value }, `字色 ${styleColor.value}`));
styleBackground?.addEventListener("change", () =>
  applyStyle({ background: styleBackground.value }, `底色 ${styleBackground.value}`)
);
styleAlign?.addEventListener("change", () =>
  applyStyle({ align: styleAlign.value === "" ? null : styleAlign.value }, `对齐 ${styleAlign.value || "默认"}`)
);
styleBorder?.addEventListener("click", () => {
  const side = { color: "#3b82f6", width: 1 };
  applyStyle({ border: { top: side, right: side, bottom: side, left: side } }, "加四边框");
});
styleApplyNumFmt?.addEventListener("click", () =>
  applyStyle({ numFmt: styleNumFmt.value === "" ? null : styleNumFmt.value }, `数字格式 ${styleNumFmt.value || "清除"}`)
);
styleClear?.addEventListener("click", () => {
  if (!core) {
    return;
  }
  core.clearRangeStyle(core.selection.range);
  styleHint.textContent = "上次动作：清除选区全部样式";
  updateStyleDebug();
});

// core 没有 selection-change 事件（只有 change/layout-change/ready 等），所以这块
// 调试信息由宿主在交互之后自己刷新——点选、键盘移动选区都会走到这两个事件。
sheetMount.addEventListener("click", () => updateStyleDebug());
sheetMount.addEventListener("keyup", () => updateStyleDebug());

populateCellColumnSelect();
rebuild();
updateStyleDebug();
