import {
  CyExcelError,
  createBasicContextMenuPlugin,
  createImageCellRenderer,
  matrixToWorkbookSnapshot
} from "/src/core/index.ts";
// 开箱即用的入口：createSheet 默认装上工具栏与编辑栏，高度布局由 core 负责。
import { createWorkbook } from "/src/createWorkbook.ts";

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

// Step 19（Univer 参考）：columnLetterFallback 打开后，没配 header 的列显示电子
// 表格字母。这一版一个 header 都不配，于是列头整排是 A/B/C/D——就是 Excel 的样子。
// （那个"配了 header 的列显示自定义标题"的分支机制还在，只是这里不演示了。）
//
// **这一版按 Excel 的模式走：不给列配 type。** 每个格子都是独立的，值是什么类型就
// 是什么类型，显示成什么样由数字格式（numFmt）决定——这正是 Excel 的分工，也是
// 导出 xlsx 时唯一需要的信息。列类型那套（整列统一的解析/编辑器/比较）留给将来的
// table 版兄弟组件，core 里机制还在，这里只是不用它。
//
// 同理，列头上也不放排序/筛选入口（Excel 的列标没有那两个箭头）——这个是在下面
// 组装 columnConfigs 时统一关掉的，不在这张表里逐列写。
//
// 于是列配置只剩下真正属于"列"的东西：宽度、标题、换行。
// 列配置只剩宽度和换行——列没有名字，就像 Excel 的列没有名字。
// （原先每列还挂着一个 playground 自用的 `label`，唯一的读者是表头排序/筛选的状态
// 栏提示；那两个入口去掉之后它就没有消费者了，一并删掉。）
const COLUMNS = [
  { width: 140 },
  { width: 200, wrap: true },
  { width: 110 },
  { width: 90 },
  { width: 110 },
  { width: 90 },
  { width: 130 },
  { width: 110 },
  { width: 220 },
  { width: 240 },
  // Step 26：这两列不再是"值本身就是文本"——最后一列敲进去的是纯文本 URL，格子
  // 里没有任何配置让它变成链接，靠的是默认绘制路径自己认出"整格内容就是一个
  // URL"然后画成蓝色下划线、点击在新标签页打开（core/cells/linkDetection.ts +
  // CyExcelCore 的 #routeBodyLinkFallback）。图片列反过来是显式的：列配了
  // `renderer: createImageCellRenderer()`，格子里放的 URL/data URI 被当图片画出
  // 来，不会因为文本"看起来像图片"就自动生效——两条机制刻意不对称，见 Step 26
  // 的 progress.md 条目。
  { width: 220 },
  { width: 64, renderer: createImageCellRenderer() },
  // 单元格级图片：这一列什么 renderer 都不配（普通文字列），图片是不是显示由
  // 这一格自己决定——`setCellImage()` 写在下面 renderRows() 里，只对部分行调用，
  // 演示"同一列里有的格是图片、有的格是空白"，而不是整列统一配置。这跟上面那列
  // 的列级 `createImageCellRenderer()` 是两条并存的路径，见 setCellImage 的文档。
  { width: 64 }
];

/** `yyyy-mm-dd`，按墙上时钟算，不绕 `Date`（那会按 UTC 解析再按本地读，差一天）。 */
function isoDate(year, month, dayOfYear) {
  const date = new Date(Date.UTC(year, month - 1, dayOfYear));
  return date.toISOString().slice(0, 10);
}

const AVATAR_COLORS = ["#f97316", "#3b82f6", "#10b981", "#a855f7", "#ef4444", "#0ea5e9", "#eab308"];

/**
 * 一个纯本地生成的 data: URI 小图标（按行号选色 + 取产品名首字），不依赖任何外部
 * 图片服务——playground 离线也能演示图片单元格，不会因为没有网络而只看到占位图。
 */
function productIconDataUri(index) {
  const color = AVATAR_COLORS[index % AVATAR_COLORS.length];
  const letter = PRODUCTS[index % PRODUCTS.length].charAt(0);
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="64" height="64">` +
    `<rect width="64" height="64" rx="12" fill="${color}"/>` +
    `<text x="32" y="42" font-size="26" text-anchor="middle" fill="#fff" font-family="sans-serif">${letter}</text>` +
    `</svg>`;
  return `data:image/svg+xml,${encodeURIComponent(svg)}`;
}

function buildRow(index) {
  const quantity = ((index * 7) % 24) + 1;
  const price = Math.round((48 + ((index * 37) % 520) + (index % 3) * 0.5) * 100) / 100;
  const tagCount = (index % 3) + 1;
  const tags = Array.from(
    { length: tagCount },
    (_unused, tagIndex) => TAG_POOL[(index + tagIndex) % TAG_POOL.length]
  );
  return [
    `CY-${String(index + 1).padStart(6, "0")}`,
    PRODUCTS[index % PRODUCTS.length],
    price,
    quantity,
    REGIONS[index % REGIONS.length],
    index % 4 !== 0,
    // 格子里只能放 string | number | boolean（对齐 Excel/xlsx 的 CellValue），
    // 所以这三列不再塞 Date / 数组 / 对象：日期给 ISO 串，标签和元数据先摊平成
    // 文本。要让日期列变成"序列号 + 日期 numFmt"（真正的 Excel 表示），得给这
    // 一列刷一个 numFmt——那是按格写样式，一百万行的 demo 里代价太大，留待列级
    // 样式有了之后再接。
    isoDate(2026, 1, 1 + (index % 365)),
    STATUSES[index % STATUSES.length],
    tags.join("、"),
    `CY-${index + 1} / W${(index % 5) + 1}`,
    // 整格就是一个 URL——不需要任何配置，默认绘制路径会自己认出来并画成链接。
    `https://www.cyexcel.example/products/CY-${String(index + 1).padStart(6, "0")}`,
    productIconDataUri(index)
  ];
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
  // 只剩 string | number | boolean 和"空"这四种形态——Date / 数组 / 对象那三个
  // 分支随 CellValue 的收紧一起没了。
  return value === undefined || value === null ? "" : String(value);
}

/**
 * 按**值本身**推断类型，这是 Excel 的做法：没有列类型可问，敲进来的是什么就存什么。
 *
 * `"123"` → 数字，`"true"` → 布尔，其余留作字符串。Excel 还会识别日期串并转成
 * 序列号，那一步要等日期改存序列号之后再补。
 */
function parseLikeExcel(raw) {
  const trimmed = raw.trim();
  if (trimmed === "") {
    return { value: "" };
  }
  if (trimmed === "true" || trimmed === "false") {
    return { value: trimmed === "true" };
  }
  // 前导零要保留（"00123" 不是 123），与 Excel 的文本格式一致。
  const numeric = Number(trimmed);
  if (!Number.isNaN(numeric) && String(numeric) === trimmed) {
    return { value: numeric };
  }
  return { value: raw };
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
const cellImageFile = document.getElementById("cellImageFile");
const cellImageFormResult = document.getElementById("cellImageFormResult");
const mergeStartRow = document.getElementById("mergeStartRow");
const mergeStartColumn = document.getElementById("mergeStartColumn");
const mergeEndRow = document.getElementById("mergeEndRow");
const mergeEndColumn = document.getElementById("mergeEndColumn");
const mergeCellsButton = document.getElementById("mergeCellsButton");
const unmergeCellsButton = document.getElementById("unmergeCellsButton");
const mergeFormResult = document.getElementById("mergeFormResult");
const dvStartRow = document.getElementById("dvStartRow");
const dvStartColumn = document.getElementById("dvStartColumn");
const dvEndRow = document.getElementById("dvEndRow");
const dvEndColumn = document.getElementById("dvEndColumn");
const dvType = document.getElementById("dvType");
const dvValue1 = document.getElementById("dvValue1");
const dvValue2 = document.getElementById("dvValue2");
const dvErrorStyle = document.getElementById("dvErrorStyle");
const dvRenderMode = document.getElementById("dvRenderMode");
const dvSetButton = document.getElementById("dvSetButton");
const dvClearButton = document.getElementById("dvClearButton");
const dvFindButton = document.getElementById("dvFindButton");
const dvResult = document.getElementById("dvResult");
const metricConstruct = document.getElementById("metricConstruct");
const metricFps = document.getElementById("metricFps");

let workbook = null;
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

  if (workbook) {
    workbook.dispose();
    workbook = null;
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
  // COLUMNS 里已经只剩真正的 ColumnDefinition 字段了（`label` 那个 playground 自用
  // 的元数据连同表头排序/筛选一起删了），所以整个原样传过去，不再逐字段挑——早先
  // 挑字段的写法静默漏掉过 options/createEditor，让配了编辑器的列退回纯文本编辑。
  const columnConfigs = COLUMNS.map((column) => ({
    ...column,
    // 照 Excel：列头上不放排序/筛选的入口。Excel 的列标就是 A/B/C，点它是选整列，
    // 排序和筛选走功能区——工具栏里 sort-asc / sort-desc / filter 三个按钮就是那
    // 个入口，所以关掉这两个不会让功能失去入口，只是把入口收归一处。
    //
    // 关掉之后列头那条"尾部热区"也一并让出来了（见 ColumnDefinition.sortable 的
    // 注释），整列宽度都是点击选列区，不会点到边上就变成排序。
    sortable: false,
    filterable: false
  }));

  const start = performance.now();
  workbook = createWorkbook(sheetMount, {
    columns: columnConfigs,
    // ADR 0004 §3.2：入口只吃一份快照。二维数组在这里导进稀疏矩阵，导完这
    // 一遍之后 `rows` 与表格再无联系——改表格不会改到它，改它也不会改表格。
    snapshot: matrixToWorkbookSnapshot(rows),
    rowHeight: config.rowHeight,
    overscanRows: config.overscanRows,
    overscanColumns: config.overscanColumns,
    headerHeight: 28,
    rowHeaderWidth: 48,
    // Step 18：点行号选中整行后，把鼠标移到选区的上/下边缘（会出现一条插入
    // 指示线）按住拖动即可重新排列行顺序——插入式移动，不会覆盖/丢失目标位置
    // 原有数据；边缘之外的拖拽仍然是原有的多选。
    rowReorder: true,
    // Univer 参考实现：没配 header 的列显示电子表格字母。这一版一列都没配，
    // 于是整排是 A/B/C/D。字母与自定义标题是同一行二选一，不是两行堆叠。
    columnLetterFallback: true,
    plugins: [createBasicContextMenuPlugin()]
  });
  const elapsed = performance.now() - start;

  core = workbook.getActiveSheet();

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

  // 表头上没有排序/筛选入口了（照 Excel：列标就是 A/B/C，点它是选整列），所以
  // 原先那两个 `header-action` 监听器一并删掉——列配成 sortable:false /
  // filterable:false 之后 defaultHeaderHitTest 会拒绝那两条热区，事件根本不再发出，
  // 留着就是永远跑不到的代码。排序和筛选的入口在工具栏（sort-asc / sort-desc /
  // filter 三个按钮），core 里 header-action 这套机制本身没动。

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

  const result = parseLikeExcel(raw);
  if (result.error) {
    cellFormResult.textContent = `写入失败：${column.label} ${result.error}`;
    return;
  }
  core.setCell(row, columnIndex, result.value);
  cellFormResult.textContent =
    `已写入 (${row}, ${columnIndex}) = ${JSON.stringify(result.value)}，可以用撤销按钮回退`;
});

cellImageFile.addEventListener("change", () => {
  if (!core) {
    return;
  }
  const file = cellImageFile.files?.[0];
  if (!file) {
    return;
  }
  const row = clampNumber(Number(cellRow.value), 0, Math.max(0, currentRowCount - 1), 0);
  cellRow.value = String(row);
  const columnIndex = Number(cellColumn.value);
  const reader = new FileReader();
  reader.onload = () => {
    core.setCellImage(row, columnIndex, String(reader.result));
    cellImageFormResult.textContent = `已写入 (${row}, ${columnIndex}) 的图片，可以用撤销按钮回退`;
    cellImageFile.value = "";
  };
  reader.readAsDataURL(file);
});

/** Reads the four data-validation-form inputs into a CellRange. */
function readDvRangeFromForm() {
  return {
    startRow: clampNumber(Number(dvStartRow.value), 0, Math.max(0, currentRowCount - 1), 0),
    startColumn: clampNumber(Number(dvStartColumn.value), 0, COLUMNS.length - 1, 0),
    endRow: clampNumber(Number(dvEndRow.value), 0, Math.max(0, currentRowCount - 1), 0),
    endColumn: clampNumber(Number(dvEndColumn.value), 0, COLUMNS.length - 1, 0)
  };
}

function buildDvRuleFromForm() {
  const type = dvType.value;
  if (type === "list") {
    // 每项可以写成 名称=颜色（#rgb / #rrggbb），颜色用作胶囊底色或单元格底色
    const items = [];
    const colors = {};
    for (const entry of dvValue1.value.split(",").map((part) => part.trim()).filter((part) => part !== "")) {
      const at = entry.lastIndexOf("=");
      const name = at === -1 ? entry : entry.slice(0, at).trim();
      if (at !== -1) {
        colors[name] = entry.slice(at + 1).trim();
      }
      if (!items.includes(name)) {
        items.push(name);
      }
    }
    return {
      criteria: { type: "list", items, ...(Object.keys(colors).length > 0 ? { colors } : {}) },
      renderMode: dvRenderMode.value,
      errorStyle: dvErrorStyle.value,
      errorMessage: `请从 ${items.join("、")} 中选择`
    };
  }
  if (type === "custom") {
    // 校验函数按名字注册在实例上，不进快照；规则里只存 { type: "custom", name }
    core.registerDataValidator("even", (value) =>
      typeof value === "number" && value % 2 === 0 ? true : "需要偶数"
    );
    return { criteria: { type: "custom", name: "even" }, errorStyle: dvErrorStyle.value };
  }
  const value1 = type === "date" ? dvValue1.value : Number(dvValue1.value);
  const value2 = type === "date" ? dvValue2.value : Number(dvValue2.value);
  return { criteria: { type, operator: "between", value1, value2 }, errorStyle: dvErrorStyle.value };
}

dvSetButton.addEventListener("click", () => {
  if (!core) {
    return;
  }
  const range = readDvRangeFromForm();
  try {
    const id = core.setDataValidation(range, buildDvRuleFromForm());
    dvResult.textContent =
      `已设置规则 ${id}：(${range.startRow}, ${range.startColumn}) - (${range.endRow}, ${range.endColumn})，` +
      "可以用撤销按钮回退；点「找出无效格」看现有数据是否合规";
  } catch (error) {
    if (error instanceof CyExcelError) {
      dvResult.textContent = `设置失败：${error.code} — ${error.message}`;
    } else {
      throw error;
    }
  }
});

dvClearButton.addEventListener("click", () => {
  if (!core) {
    return;
  }
  const range = readDvRangeFromForm();
  core.clearDataValidation(range);
  dvResult.textContent =
    `已清除 (${range.startRow}, ${range.startColumn}) - (${range.endRow}, ${range.endColumn}) 上的规则，可以撤销`;
});

dvFindButton.addEventListener("click", () => {
  if (!core) {
    return;
  }
  const invalid = core.findInvalidCells();
  const rules = core.getDataValidations().length;
  dvResult.textContent = invalid.length === 0
    ? `${rules} 条规则，没有无效格`
    : `${rules} 条规则，${invalid.length} 个无效格：` +
      invalid.slice(0, 20).map(({ row, column }) => `(${row}, ${column})`).join(" ") +
      (invalid.length > 20 ? " …" : "");
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
