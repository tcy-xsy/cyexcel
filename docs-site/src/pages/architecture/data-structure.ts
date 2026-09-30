import { pageHeader, section, codeBlock, apiTable, callout } from "../../components/doc";

export default function dataStructurePage(root: HTMLElement): void {
  root.append(
    pageHeader(
      "ARCHITECTURE",
      "数据结构",
      "cyexcel 的数据是一份稀疏 WorkbookSnapshot；外部对象数组只负责导入，不参与绑定。"
    )
  );

  const overview = section("顶层结构");
  overview.body.append(
    codeBlock(`{
  version: 1,
  sheetOrder: ["sheet-1"],
  sheets: {
    "sheet-1": {
      id: "sheet-1",
      name: "Sheet1",
      blankRows: 1000,
      blankColumns: 26,
      cellData: { /* row -> column -> CellRecord */ },
      rowData: { /* row -> RowMeta */ },
      columnData: { /* column -> ColumnMeta */ },
      mergeData: [],
      dataValidations: [ /* 数据验证规则，缺省视为空 */ ]
    }
  },
  styles: { /* styleId -> CellStyle */ },
  resources: []
}`)
  );
  root.append(overview.root);

  const cells = section("单元格");
  cells.body.append(
    codeBlock(`// 只存非空格子。
"0,0": { v: "CY-100001" }
"0,1": { v: "工业温度传感器" }
"0,2": { v: 128 }`)
  );
  root.append(cells.root);
  root.append(
    apiTable([
      { name: "v", type: "string | number | boolean", description: "单元格值，必需。" },
      { name: "t", type: "string", description: "语义类型名，可选；例如 text、number、date。" },
      { name: "s", type: "string", description: "样式 id，指向 workbook 顶层 styles 表。" },
      { name: "f", type: "string", description: "公式字符串，可选。" }
    ])
  );

  const coordinates = section("坐标空间");
  coordinates.body.append(
    codeBlock(`rowCount    = maxOccupiedRow    + 1 + blankRows
columnCount = maxOccupiedColumn + 1 + blankColumns`)
  );
  root.append(coordinates.root);
  root.append(
    apiTable([
      { name: "getUsedRowCount()", type: "number", description: "实际有数据的行数，不含尾部空白。" },
      { name: "getRowCount()", type: "number", description: "包含尾部空白行的坐标空间行数。" },
      { name: "getColumnCount()", type: "number", description: "包含尾部空白列的坐标空间列数。" }
    ])
  );

  const validation = section("数据验证规则");
  validation.body.append(
    codeBlock(`dataValidations: [
  {
    id: "dv-1",
    // 区域列表（sqref）：可以有多个矩形，规则内、规则之间都不重叠
    ranges: [{ startRow: 1, startColumn: 1, endRow: 99, endColumn: 1 }],
    criteria: { type: "whole", operator: "between", value1: 1, value2: 100 },
    errorStyle: "stop",
    errorMessage: "数量必须是 1 到 100 的整数"
  }
]`)
  );
  root.append(validation.root);
  root.append(
    apiTable([
      { name: "ranges", type: "CellRange[]", description: "规则覆盖的区域。规则是区域数据，不是单元格的属性，所以不进 cellData，行列变化时整块搬移。" },
      { name: "criteria", type: "object", description: "判据。type 取 whole / decimal / textLength / date / time / list / custom。date 与 time 的值是序列号；custom 只存 name 和 params，函数不进快照。" },
      { name: "allowBlank", type: "boolean", description: "默认 true。" },
      { name: "errorStyle", type: "\"stop\" | \"warning\" | \"information\"", description: "编辑器提交违规值时的处理，默认 stop。" },
      { name: "renderMode / showDropDown", type: "list 专用", description: "text / arrow / custom（默认 arrow）与是否画下拉；list.colors 是选项到十六进制颜色的映射。" },
      { name: "prompt / promptTitle", type: "string", description: "选中单元格时显示的输入提示。" }
    ])
  );
  root.append(
    callout(
      "一格至多属于一条规则；没用数据验证的表，快照里没有 dataValidations 这个字段，与之前逐字节一致。"
    )
  );

  root.append(
    callout(
      "matrixToWorkbookSnapshot() 是一次性导入。之后改外部 rows 不会影响表格；表格编辑也不会写回外部数组。"
    )
  );
}
