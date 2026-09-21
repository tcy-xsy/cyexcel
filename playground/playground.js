import {
  CyExcelError,
  createBasicContextMenuPlugin,
  createBooleanCellEditor,
  createDateCellEditor,
  createSelectCellEditor,
  createTextAreaCellEditor,
  rowsToWorkbookSnapshot
} from "/src/core/index.ts";
// 开箱即用的入口：createSheet 默认装上工具栏与编辑栏，高度布局由 core 负责。
import { createSheet } from "/src/createSheet.ts";

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
  // out of date. 搜索状态不必在这里清：查找面板现在归工具栏管，而工具栏会随
  // 实例一起重建（工具栏归 core 管，随实例一起 dispose）。
  emptyFilterState.hidden = true;

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
  core = createSheet(sheetMount, {
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
  // defaultHeaderHitTest），不只是工具栏上的升序/降序按钮。排序状态本身由 core
  // 持有（core.sortState），两条路径共用，所以这里只把它显示到状态栏。
  core.events.on("header-action", (event) => {
    if (event.action.type !== "toggle-sort") {
      return;
    }
    const state = core.sortState;
    if (!state) {
      return;
    }
    setStatus(
      `点击表头「${COLUMNS[state.column].label}」排序为` +
        `${state.direction === "asc" ? "升序" : "降序"}` +
        "（再点一次在升/降序间切换，撤销可一步还原顺序和行高）"
    );
  });

  // Step 22 Part B：表头点击 toggle-filter 是**纯通知**——core 不知道该用什么谓词。
  // 工具栏的过滤面板是同一件事的另一个入口，两边都由宿主决定条件；这里选择最朴素的
  // 一条："只看这一列非空的行"，并把它说清楚，免得用户以为点了没反应。
  core.events.on("header-action", (event) => {
    if (event.action.type !== "toggle-filter") {
      return;
    }
    core.filterByColumn(event.column, (value) => value !== null && value !== undefined && value !== "");
    setStatus(
      `点击表头过滤：只看第 ${event.column} 列非空的行` +
        "（工具栏的过滤面板里有更多条件，清除过滤也在那里）"
    );
  });

  const maxRow = Math.max(0, currentRowCount - 1);
  if (Number(cellRow.value) > maxRow) {
    cellRow.value = String(maxRow);
  }
  cellRow.max = String(maxRow);
  // Step 25 之前这块调试信息是靠 sheetMount 的 click/keyup 猜"选区可能变了"，
  // 现在 core 有 selection-change 事件了（工具栏的按钮高亮也走它），订阅即可。
  core.events.on("selection-change", () => updateStyleDebug());
  updateStyleDebug();
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

const emptyFilterState = document.getElementById("emptyFilterState");

// Step 22 Part A/B 与 Step 23 的排序、过滤、查找替换控件，以及 Step 24 的样式面板，
// 全部撤掉了——它们现在是工具栏（Step 25）data / format / align 三组里的项。
// 少了一层"同一个功能两个入口、改一处忘一处"，也让 playground 成为工具栏本身的
// 验证场：这一步没有浏览器测试，工具栏的渲染、点击、面板、dispose 全靠在这里点。
//
// 没有收进工具栏、仍然留在页面上的只有两类：读写单元格与合并区那两个表单（它们
// 演示的是**带参数的 API 调用**，不是工具栏那种"作用于当前选区"的动作），以及
// 下面这块调试信息。
const styleDebug = document.getElementById("styleDebug");

/** 当前格的样式与 styles 表条目数——肉眼验证 intern 去重的唯一直观方式。 */
function updateStyleDebug() {
  if (!core || !styleDebug) {
    return;
  }
  const { startRow, startColumn } = core.selection.range;
  const style = core.getCellStyle(startRow, startColumn);
  styleDebug.textContent =
    `当前格 R${startRow}C${startColumn}：${style === undefined ? "无样式" : JSON.stringify(style)}`;
}

/**
 * 把一句话顶到状态栏上，几秒后退回默认文案。
 *
 * 工具栏自己不解释它做了什么（那是产品级组件该有的克制），但 playground 作为演示
 * 场需要让人看见"刚才那一下发生了什么"——尤其是表头点击这类没有按钮反馈的路径。
 */
let statusResetTimeoutId = null;
function setStatus(text) {
  statusBar.textContent = text;
  if (statusResetTimeoutId !== null) {
    clearTimeout(statusResetTimeoutId);
  }
  statusResetTimeoutId = setTimeout(() => {
    statusBar.textContent = statusBarDefaultText;
    statusResetTimeoutId = null;
  }, 6000);
}

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

populateCellColumnSelect();
rebuild();
updateStyleDebug();
