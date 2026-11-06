import type { DocElement } from "../../components/content";
import { createSheet } from "@cyexcel/createSheet";
import { matrixToWorkbookSnapshot, serialFromISODate, type DataValidationRule } from "@cyexcel/core";
import {
  setText,
  actionButton,
  apiTable,
  callout,
  codeBlock,
  demoBlock,
  pageHeader,
  paragraph,
  section
} from "../../components/doc";

const columns = [
  { width: 150, header: "任务" },
  { width: 90, header: "数量", type: "number" },
  { width: 120, header: "状态" },
  { width: 100, header: "优先级" },
  { width: 120, header: "交付日期" },
  { width: 200, header: "备注" }
];

// 日期列使用 ISO 字符串，日期规则按日期解析，无需额外配置数字格式。
const rows: (string | number | undefined)[][] = [
  ["采购传感器", 40, "进行中", "高", "2026-03-15", "首批到货"],
  ["更换门锁", 12, "待处理", "中", "2025-06-01", ""],
  ["安装扫码终端", 250, "已完成", "低", "2026-02-01", "数量超出上限"],
  ["巡检电源模组", 8, "进行中", "中", "2026-04-20", "这是一段明显超过十个字的备注文字"],
  ["标签打印机保养", 3, "待处理", "低", "2026-05-08", ""]
];

const wholeColumn = (column: number) => ({ startRow: 0, startColumn: column, endRow: 49, endColumn: column });

/**
 * 初始规则通过快照导入，不产生撤销记录。
 * 日期规则的比较值使用 serialFromISODate 转换为序列号。
 */
const rules: readonly DataValidationRule[] = [
  {
    id: "dv-qty",
    ranges: [wholeColumn(1)],
    criteria: { type: "whole", operator: "between", value1: 1, value2: 100 },
    errorStyle: "stop",
    errorTitle: "数量无效",
    errorMessage: "数量必须是 1 到 100 的整数"
  },
  {
    id: "dv-status",
    ranges: [wholeColumn(2)],
    criteria: {
      type: "list",
      items: ["待处理", "进行中", "已完成"],
      colors: { 待处理: "#e8eaed", 进行中: "#d2e3fc", 已完成: "#ceead6" }
    },
    renderMode: "custom",
    errorMessage: "请从列表中选择状态"
  },
  {
    id: "dv-priority",
    ranges: [wholeColumn(3)],
    criteria: { type: "list", items: ["高", "中", "低"], colors: { 高: "#fce8e6", 中: "#fef7e0", 低: "#e6f4ea" } },
    renderMode: "arrow"
  },
  {
    id: "dv-date",
    ranges: [wholeColumn(4)],
    criteria: { type: "date", operator: "greaterThanOrEqual", value1: serialFromISODate("2026-01-01") ?? 0 },
    errorStyle: "warning",
    errorTitle: "日期偏早",
    errorMessage: "交付日期早于 2026-01-01，请核对日期"
  },
  {
    id: "dv-note",
    ranges: [wholeColumn(5)],
    criteria: { type: "textLength", operator: "lessThanOrEqual", value1: 10 },
    errorStyle: "information",
    errorMessage: "备注超过 10 个字，建议精简"
  }
];

function buildSnapshot() {
  const base = matrixToWorkbookSnapshot(rows);
  const sheetId = base.sheetOrder[0]!;
  return {
    ...base,
    sheets: { ...base.sheets, [sheetId]: { ...base.sheets[sheetId]!, dataValidations: rules } }
  };
}

export default function validationExamplePage(root: DocElement): void {
  let sheet: ReturnType<typeof createSheet>;
  root.append(
    pageHeader(
      "EXAMPLES",
      "数据验证",
      "按区域配置验证规则，支持无效输入拒绝、单元格标记和悬停提示。规则作为区域数据存储在快照中。"
    )
  );

  const setRule = section("设置规则");
  setRule.body.append(
    paragraph(
      "setDataValidation(range, rule) 为一个或多个区域设置规则，并返回规则 ID。每个单元格至多属于一条规则，" +
        "新规则覆盖的区域会从旧规则中移除，覆盖区域为空的旧规则会被删除。规则修改支持撤销，覆盖范围随行列的" +
        "插入、删除和移动更新。示例通过初始快照的 sheet.dataValidations 配置规则，因此初始规则不产生撤销记录。"
    )
  );
  root.append(setRule.root);
  root.append(
    codeBlock(`const id = sheet.setDataValidation(
  { startRow: 0, startColumn: 1, endRow: 49, endColumn: 1 },
  {
    criteria: { type: "whole", operator: "between", value1: 1, value2: 100 },
    errorStyle: "stop",                       // 默认 "stop"
    errorMessage: "数量必须是 1 到 100 的整数"
  }
);

sheet.clearDataValidation(range);             // 清除区域规则，删除覆盖范围为空的规则
sheet.getDataValidations();                   // 当前全部规则
sheet.getDataValidationAt(row, column);       // 指定单元格所属的规则`)
  );

  const types = section("规则类型");
  root.append(types.root);
  root.append(
    apiTable([
      { name: "whole", type: "整数", description: "值必须是整数（不做文本到数字的隐式转换），再按 operator 与 value1 / value2 比较。" },
      { name: "decimal", type: "小数", description: "任意有限数字，比较方式同上。" },
      { name: "textLength", type: "文本长度", description: "按值的字符串形式取长度再比较，value 必须是非负整数。" },
      { name: "date", type: "日期", description: "比较序列号。value1 / value2 可以写 \"2026-01-15\"，创建时会换算成序列号；单元格值是字符串时按日期解析。" },
      { name: "time", type: "时间", description: "只比序列号的小数部分，与是哪一天无关。" },
      { name: "list", type: "序列", description: "值必须属于 items 列表，支持下拉选项、显示模式和选项颜色。" },
      { name: "custom", type: "自定义", description: "按名称调用已注册的验证函数，见下文。" }
    ])
  );
  root.append(
    callout(
      "operator 取 between / notBetween / equal / notEqual / greaterThan / greaterThanOrEqual / lessThan / lessThanOrEqual；" +
        "between 需要 value2。allowBlank 默认为 true，允许空值。"
    )
  );

  const styles = section("三种错误样式");
  styles.body.append(
    paragraph(
      "errorStyle 定义编辑器提交无效值时的处理方式。停止级拒绝输入，警告级和信息级允许提交，并显示无效标记。" +
        "输入拒绝仅作用于编辑器提交；粘贴、填充、setCell() 和查找替换允许写入，无效值通过标记提示。"
    )
  );
  root.append(styles.root);
  root.append(
    apiTable([
      { name: "stop（默认）", type: "拒绝", description: "拒绝提交并显示模态对话框。点击「确定」或按 Esc 后放弃本次无效修改，保留原值并退出编辑。对话框显示期间不接受重复提交。" },
      { name: "warning", type: "仅标记", description: "允许提交，无效单元格右上角显示红色三角标记。" },
      { name: "information", type: "仅标记", description: "同 warning。" }
    ])
  );

  const list = section("下拉与显示模式");
  list.body.append(
    paragraph(
      "list 规则默认在单元格右侧显示下拉箭头，点击可打开选项列表，选择选项后通过单元格写入流程提交。" +
        "点击列表外部、按 Esc 或切换活动单元格时关闭列表。renderMode 控制显示方式："
    )
  );
  root.append(list.root);
  root.append(
    apiTable([
      { name: "arrow（默认）", type: "文字加箭头", description: "普通文字，右侧一个箭头；文字给箭头留出位置，过长会在箭头前截断。" },
      { name: "text", type: "仅文字", description: "仅显示文字，不预留箭头空间；双击单元格可编辑值。" },
      { name: "custom", type: "胶囊标签", description: "使用带底色的圆角标签显示文字和箭头，点击标签可打开列表；单元格空间不足时使用 arrow 模式。" }
    ])
  );
  root.append(
    codeBlock(`sheet.setDataValidation(range, {
  criteria: {
    type: "list",
    items: ["待处理", "进行中", "已完成"],
    // 选项颜色：custom 模式是胶囊底色，arrow / text 模式是单元格底色。
    // 支持 #rgb / #rrggbb / #rrggbbaa。
    colors: { 待处理: "#e8eaed", 进行中: "#d2e3fc", 已完成: "#ceead6" }
  },
  renderMode: "custom",
  showDropDown: true                            // false 表示隐藏下拉箭头
});`)
  );

  const custom = section("自定义规则");
  custom.body.append(
    paragraph(
      "自定义规则仅存储名称和可 JSON 序列化的 params，支持快照保存、撤销和行列移动。验证函数不写入快照，" +
        "需在创建实例时按名称注册。验证函数必须是同步纯函数，并控制计算开销，因为绘制时会对可见单元格调用。" +
        "快照引用未注册的验证函数时，相关单元格按有效处理。"
    )
  );
  root.append(custom.root);
  root.append(
    codeBlock(`sheet.registerDataValidator("even", (value, params, context) => {
  // 返回 true 有效；false 无效（默认文案）；字符串无效并作为提示文案
  // context.getCell(row, column) 只读当前已提交的状态，可以做跨格规则
  return typeof value === "number" && value % 2 === 0 ? true : "需要偶数";
});

sheet.setDataValidation(range, { criteria: { type: "custom", name: "even" } });`)
  );
  root.append(
    callout(
      "对于依赖其他单元格的自定义规则（如列内唯一），依赖值变化后的验证标记在下一次重绘时更新。"
    )
  );

  const check = section("事后检查");
  check.body.append(
    paragraph(
      "设置规则不会拒绝或修改已有数据。findInvalidCells() 返回无效单元格坐标，validateCell() 检查指定单元格。" +
        "dataValidation: { showInvalidMarks: false } 可关闭无效标记。" +
        "鼠标在无效单元格上停留 400ms 后显示规则说明，移开、滚动或开始编辑时关闭。" +
        "停止级提示标题为“输入无效”，警告级为“注意”，信息级为“提示”。"
    )
  );
  root.append(check.root);
  root.append(
    codeBlock(`sheet.findInvalidCells();          // [{ row: 2, column: 1 }, ...]
sheet.validateCell(2, 1);            // { valid: false, errorStyle: "stop", message: "…" }`)
  );

  const write = actionButton("写入一个违规值（只标记）", () => {
    sheet.setCell(4, 1, -5);
    setText(log, "已通过 setCell() 将第 5 行数量设置为 -5，并显示无效标记。此操作支持撤销。");
  });
  const find = actionButton("列出无效单元格", () => {
    const invalid = sheet.findInvalidCells();
    setText(log, invalid.length === 0
        ? "未发现无效单元格"
        : `${invalid.length} 个无效单元格：` + invalid.map(({ row, column }) => `(${row}, ${column})`).join(" "));
  });
  const clear = actionButton("清除数量列的规则", () => {
    sheet.clearDataValidation({ startRow: 0, startColumn: 1, endRow: 49, endColumn: 1 });
    setText(log, "已清除数量列的规则，红色小三角消失；可以撤销。");
  });
  const undo = actionButton("撤销", () => {
    sheet.undo();
    setText(log, "已撤销。");
  });

  const demo = demoBlock({
    title: "数据验证示例",
    description:
      "「数量」列拒绝 abc 或 500 等无效输入；「交付日期」早于 2026-01-01 或「备注」超过 10 个字时允许提交，并显示无效标记。" +
      "「状态」与「优先级」支持下拉选择。示例包含预设无效值，悬停可查看规则说明。",
    actions: [write, find, clear, undo]
  });
  root.append(demo.root);
  if (demo.log === undefined) {
    throw new Error("demo log missing");
  }
  const log = demo.log;

  demo.mount.addClass("with-toolbar");
  demo.mount.onMount = () => {
    sheet = createSheet(demo.mount.element, {
      columns,
      snapshot: buildSnapshot(),
      headerHeight: 28,
      rowHeaderWidth: 48,
      rowHeight: 30
    });
    return () => sheet.dispose();
  };
}
