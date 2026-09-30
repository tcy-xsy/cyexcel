import { createSheet } from "@cyexcel/createSheet";
import { matrixToWorkbookSnapshot, serialFromISODate, type DataValidationRule } from "@cyexcel/core";
import {
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
  { width: 90, header: "数量" },
  { width: 120, header: "状态" },
  { width: 100, header: "优先级" },
  { width: 120, header: "交付日期" },
  { width: 200, header: "备注" }
];

// 日期列放 ISO 字符串：日期规则对字符串会按日期解析，演示里也就不用再给这一列设数字格式。
const rows: (string | number | undefined)[][] = [
  ["采购传感器", 40, "进行中", "高", "2026-03-15", "首批到货"],
  ["更换门锁", 12, "待处理", "中", "2025-06-01", ""],
  ["安装扫码终端", 250, "已完成", "低", "2026-02-01", "数量超出上限"],
  ["巡检电源模组", 8, "进行中", "中", "2026-04-20", "这是一段明显超过十个字的备注文字"],
  ["标签打印机保养", 3, "待处理", "低", "2026-05-08", ""]
];

const wholeColumn = (column: number) => ({ startRow: 0, startColumn: column, endRow: 49, endColumn: column });

/**
 * 演示表的规则。放进**快照**而不是启动后再 setDataValidation()：规则本来就是快照的一部分
 * （`dataValidations`），这样初始规则也不会出现在撤销历史里——点“撤销”只会撤销你自己做的事。
 * 日期规则的值在快照里就是序列号，所以这里用 serialFromISODate 换算。
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
    errorMessage: "交付日期早于 2026-01-01，仍要保存吗？"
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

export default function validationExamplePage(root: HTMLElement): () => void {
  root.append(
    pageHeader(
      "EXAMPLES",
      "数据验证",
      "给一片区域设置规则，编辑时按规则拦截或提示，违规的格子用红色小三角标出来。规则是随快照走的区域数据，不是格子的属性。"
    )
  );

  const setRule = section("设置规则");
  setRule.body.append(
    paragraph(
      "setDataValidation(range, rule) 给一个或多个区域设置一条规则，返回规则 id。一格至多属于一条规则：" +
        "后设置的规则会从旧规则里挖掉重叠的部分，被挖空的旧规则整条消失。规则可以撤销，也会随行列的" +
        "插入、删除、移动一起走。规则也是快照的一部分（sheet.dataValidations），下面的演示表就是把" +
        "规则直接写在初始快照里，所以它们不在撤销历史里。"
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

sheet.clearDataValidation(range);             // 清除；被挖空的规则整条消失
sheet.getDataValidations();                   // 当前全部规则
sheet.getDataValidationAt(row, column);       // 这一格所在的规则`)
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
      { name: "list", type: "序列", description: "值必须在 items 里；可带下拉、显示模式和选项颜色，见下文。" },
      { name: "custom", type: "自定义", description: "按名字调用你注册的函数，见下文。" }
    ])
  );
  root.append(
    callout(
      "operator 取 between / notBetween / equal / notEqual / greaterThan / greaterThanOrEqual / lessThan / lessThanOrEqual；" +
        "between 需要 value2。allowBlank 默认 true，空值直接通过。"
    )
  );

  const styles = section("三种错误样式");
  styles.body.append(
    paragraph(
      "errorStyle 决定编辑器提交违规值时怎么办。与 Univer 一致，只有停止级会拒绝输入，警告和信息直接放行、" +
        "只把违规的格子标红。拒绝只作用于编辑器提交这一条路径：粘贴、填充柄、setCell()、查找替换都不拦截，" +
        "只把违规的格子标出来。"
    )
  );
  root.append(styles.root);
  root.append(
    apiTable([
      { name: "stop（默认）", type: "拒绝", description: "不提交，弹出模态对话框（标题、错误信息、一个「确定」按钮）；点确定回到编辑器，输入内容还在。对话框开着的时候，回车、失焦等提交动作都被挡住。" },
      { name: "warning", type: "仅标记", description: "值照常提交，违规的格子右上角出现红色小三角。" },
      { name: "information", type: "仅标记", description: "同 warning。" }
    ])
  );

  const list = section("下拉与显示模式");
  list.body.append(
    paragraph(
      "list 规则的单元格右侧常驻一个下拉箭头，点它打开列表，选一项就是一次普通的单元格写入。" +
        "箭头画在画布上，悬停时光标变手型；列表点外面、按 Esc 或换到别的格子都会关闭。" +
        "renderMode 控制怎么画："
    )
  );
  root.append(list.root);
  root.append(
    apiTable([
      { name: "arrow（默认）", type: "文字加箭头", description: "普通文字，右侧一个箭头；文字给箭头留出位置，过长会在箭头前截断。" },
      { name: "text", type: "仅文字", description: "不画箭头，也不预留空间；选值靠双击输入。" },
      { name: "custom", type: "胶囊标签", description: "带底色的圆角胶囊，文字和箭头都在里面，点整个胶囊都能打开列表；格子放不下时退回 arrow。" }
    ])
  );
  root.append(
    codeBlock(`sheet.setDataValidation(range, {
  criteria: {
    type: "list",
    items: ["待处理", "进行中", "已完成"],
    // 选项颜色：custom 模式是胶囊底色，arrow / text 模式是单元格底色。
    // 只收 #rgb / #rrggbb / #rrggbbaa。
    colors: { 待处理: "#e8eaed", 进行中: "#d2e3fc", 已完成: "#ceead6" }
  },
  renderMode: "custom",
  showDropDown: true                            // 设成 false 就不画箭头
});`)
  );

  const custom = section("自定义规则");
  custom.body.append(
    paragraph(
      "规则里只存名字和可 JSON 序列化的 params，所以规则能进快照、能撤销、能随行列移动；判定函数不进快照，" +
        "由每次创建实例的代码按名字注册。函数必须同步、纯函数，而且要足够便宜——绘制时每帧会对可见格调用。" +
        "快照里出现没注册的名字时按“有效”处理，别的环境打开同一份快照不会满屏报错。"
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
      "依赖其他单元格的自定义规则（比如“列内唯一”），另一格变化时这一格的红色小三角要等到下一次重绘才更新。"
    )
  );

  const check = section("事后检查");
  check.body.append(
    paragraph(
      "设置规则不会检查已有数据是否合规。findInvalidCells() 列出违规的格子，对应 Excel 的“圈释无效数据”；" +
        "validateCell() 检查单格。关掉红色小三角用 dataValidation: { showInvalidMarks: false }。" +
        "鼠标停在违规的格子上时，会在格子旁弹出这条规则的错误信息（停留 100ms 后出现，移开、滚动、开始编辑都会收起）。"
    )
  );
  root.append(check.root);
  root.append(
    codeBlock(`sheet.findInvalidCells();          // [{ row: 2, column: 1 }, ...]
sheet.validateCell(2, 1);            // { valid: false, errorStyle: "stop", message: "…" }`)
  );

  const write = actionButton("写入一个违规值（只标记）", () => {
    sheet.setCell(4, 1, -5);
    log.textContent = "已用 setCell() 把第 5 行数量写成 -5：没有被拦截，右上角出现红色小三角；可以撤销。";
  });
  const find = actionButton("列出无效格", () => {
    const invalid = sheet.findInvalidCells();
    log.textContent =
      invalid.length === 0
        ? "没有无效格"
        : `${invalid.length} 个无效格：` + invalid.map(({ row, column }) => `(${row}, ${column})`).join(" ");
  });
  const clear = actionButton("清除数量列的规则", () => {
    sheet.clearDataValidation({ startRow: 0, startColumn: 1, endRow: 49, endColumn: 1 });
    log.textContent = "已清除数量列的规则，红色小三角消失；可以撤销。";
  });
  const undo = actionButton("撤销", () => {
    sheet.undo();
    log.textContent = "已撤销。";
  });

  const demo = demoBlock({
    title: "试一试",
    description:
      "双击「数量」输入 abc 或 500（停止：弹出对话框拒绝输入）；双击「交付日期」输入 2025-01-01、" +
      "「备注」写超过 10 个字（警告、信息：放行，只标红三角）；点「状态」「优先级」的下拉箭头选值；" +
      "鼠标停在带红三角的格子上看错误提示。表里已经有几处违规值。",
    actions: [write, find, clear, undo]
  });
  root.append(demo.root);
  if (demo.log === undefined) {
    throw new Error("demo log missing");
  }
  const log = demo.log;

  demo.mount.classList.add("with-toolbar");
  const sheet = createSheet(demo.mount, {
    columns,
    snapshot: buildSnapshot(),
    headerHeight: 28,
    rowHeaderWidth: 48,
    rowHeight: 30
  });

  return () => sheet.dispose();
}
