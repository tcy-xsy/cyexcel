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
      mergeData: []
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

  root.append(
    callout(
      "matrixToWorkbookSnapshot() 是一次性导入。之后改外部 rows 不会影响表格；表格编辑也不会写回外部数组。"
    )
  );
}
