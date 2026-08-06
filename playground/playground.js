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
const COLUMNS = [
  { id: "sku", field: "sku", width: 140, type: "text", label: "SKU" },
  { id: "name", field: "name", width: 200, type: "text", label: "产品名称" },
  { id: "price", field: "price", width: 110, type: "money", label: "单价" },
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
const rebuildButton = document.getElementById("rebuildButton");
const cellRow = document.getElementById("cellRow");
const cellColumn = document.getElementById("cellColumn");
const cellValue = document.getElementById("cellValue");
const readCellButton = document.getElementById("readCellButton");
const writeCellButton = document.getElementById("writeCellButton");
const cellFormResult = document.getElementById("cellFormResult");
const metricConstruct = document.getElementById("metricConstruct");
const metricFps = document.getElementById("metricFps");

let core = null;
let currentRowCount = 0;

function populateCellColumnSelect() {
  cellColumn.innerHTML = "";
  COLUMNS.forEach((column, index) => {
    const option = document.createElement("option");
    option.value = String(index);
    option.textContent = `${column.label} (${column.id})`;
    cellColumn.append(option);
  });
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
    plugins: [createBasicContextMenuPlugin()]
  });
  const elapsed = performance.now() - start;

  currentRowCount = config.rowCount;
  metricConstruct.textContent = `${elapsed.toFixed(1)} ms`;
  statusBar.textContent =
    `${currentRowCount.toLocaleString("zh-CN")} 行 × ${COLUMNS.length} 列 · ` +
    `行高 ${config.rowHeight}px · overscan ${config.overscanRows}/${config.overscanColumns}`;

  const maxRow = Math.max(0, currentRowCount - 1);
  if (Number(cellRow.value) > maxRow) {
    cellRow.value = String(maxRow);
  }
  cellRow.max = String(maxRow);
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
