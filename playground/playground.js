import {
  CyExcelCore,
  createBasicContextMenuPlugin,
  createBooleanCellEditor,
  createDateCellEditor,
  createSelectCellEditor,
  createTextAreaCellEditor
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
  { id: "sku", field: "sku", width: 140, type: "text", label: "SKU", header: "SKU", sortable: false },
  { id: "name", field: "name", width: 200, type: "text", label: "产品名称", wrap: true },
  { id: "price", field: "price", width: 110, type: "money", label: "单价", header: "单价" },
  { id: "quantity", field: "quantity", width: 90, type: "number", label: "数量" },
  { id: "region", field: "region", width: 110, type: "text", label: "地区" },
  // 数据类型（boolean/date/select…）和编辑控件是两个独立维度，不会被自动绑定——
  // 这三列显式通过 createEditor 接上对应控件，是业务方要自己做的事，不是"boolean
  // 类型天生等于 checkbox"。不接的类型（下面的 tags/meta）仍然用默认文本编辑器，
  // 靠 CellType.parse()/serialize() 保证往返正确。
  {
    id: "inStock",
    field: "inStock",
    width: 90,
    type: "boolean",
    label: "有库存",
    header: "有库存",
    renderer: inStockRenderer,
    createEditor: () => createBooleanCellEditor()
  },
  {
    id: "restockedAt",
    field: "restockedAt",
    width: 130,
    type: "date",
    label: "入库日期",
    createEditor: () => createDateCellEditor()
  },
  {
    id: "status",
    field: "status",
    width: 110,
    type: "select",
    label: "状态",
    header: "状态",
    options: STATUSES.map((value) => ({ value })),
    createEditor: (context) => createSelectCellEditor(context.column.options ?? [])
  },
  {
    id: "tags",
    field: "tags",
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
    id: "meta",
    field: "meta",
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

// Step 19: arbitrary-depth nested grouped headers — a group node either
// gives `columns` directly (leaf-adjacent row) or `children` (this group's
// span is the union of its children's spans, one row further out). Three
// levels here (the old Step 9 fixture only ever showed two), covering every
// column in COLUMNS above so the grouping is visually obvious end-to-end.
const HEADER_GROUPS = [
  {
    id: "all",
    title: "全部字段",
    children: [
      {
        id: "product",
        title: "商品信息",
        children: [
          { id: "identity", title: "标识", columns: [0, 1] },
          { id: "pricing", title: "价格数量", columns: [2, 3] }
        ]
      },
      {
        id: "ops",
        title: "运营信息",
        children: [
          { id: "logistics", title: "地区库存", columns: [4, 5, 6] },
          { id: "meta-group", title: "状态与元数据", columns: [7, 8, 9] }
        ]
      }
    ]
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
const headerGroupsToggle = document.getElementById("headerGroupsToggle");
const rebuildButton = document.getElementById("rebuildButton");
const cellRow = document.getElementById("cellRow");
const cellColumn = document.getElementById("cellColumn");
const cellValue = document.getElementById("cellValue");
const readCellButton = document.getElementById("readCellButton");
const writeCellButton = document.getElementById("writeCellButton");
const cellFormResult = document.getElementById("cellFormResult");
const sortColumn = document.getElementById("sortColumn");
const sortAscButton = document.getElementById("sortAscButton");
const sortDescButton = document.getElementById("sortDescButton");
const sortHint = document.getElementById("sortHint");
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
    option.textContent = `${column.label} (${column.id})`;
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
    data: rows,
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
    // Step 19：三层嵌套分组表头。配置了 headerGroups 后，列头右键菜单里
    // createBasicContextMenuPlugin() 贡献的"插入/删除/隐藏列"这几项会被
    // headerGroups 阻塞（分组引用了列索引，结构变了索引就错位了）——现在这
    // 几项会直接显示为禁用态（basicMenuPlugin.ts 的 hasHeaderGroups() 检
    // 查），且真的被点到时也会把错误抛到 statusBar，不会再是"什么反应都没
    // 有"。下面这个复选框可以直接关掉 headerGroups，用来实际体验插入/删除
    // 列这个功能。
    headerGroups: headerGroupsToggle.checked ? HEADER_GROUPS : undefined,
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

  // 右键菜单 action 抛错（比如 headerGroups 开着时点了插入/删除列）现在会
  // 通过 core.events 的 "error" 事件转发出来——之前只有 console.error，用
  // 户在页面上完全看不出"点了没反应"到底是不是出错了。这里临时把错误信息
  // 顶到 statusBar 上几秒钟，然后再切回正常的行列统计文案。
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

  const maxRow = Math.max(0, currentRowCount - 1);
  if (Number(cellRow.value) > maxRow) {
    cellRow.value = String(maxRow);
  }
  cellRow.max = String(maxRow);
  // Only now: the answer comes off the resolved RuntimeColumn, which does not
  // exist until this rebuild has constructed the sheet.
  populateSortColumnSelect();
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
