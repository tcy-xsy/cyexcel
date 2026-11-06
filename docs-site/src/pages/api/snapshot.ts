import { pageHeader, paragraph, section, apiTable, callout, codeBlock } from "../../components/doc";

export default function snapshotApiPage(root: HTMLElement): void {
  root.append(
    pageHeader("API", "matrixToWorkbookSnapshot()", "将二维数组转换为 WorkbookSnapshot，导入后的表格数据与原始数组相互独立。")
  );

  const signature = section("签名");
  signature.body.append(
    codeBlock(`function matrixToWorkbookSnapshot(
  matrix: readonly (readonly (CellValue | null | undefined)[])[],
  options?: MatrixToWorkbookSnapshotOptions
): WorkbookSnapshot`)
  );
  root.append(signature.root);
  root.append(
    apiTable([
      { name: "matrix", type: "readonly (readonly CellValue[])?[]", description: "第 i 行、第 j 列映射到单元格 (i, j)。" },
      { name: "options.blankRows", type: "number", description: "在数据行之后额外保留的空白行。" },
      { name: "options.blankColumns", type: "number", description: "在最宽数据列之后额外保留的空白列。" }
    ])
  );

  root.append(
    callout(
      "undefined 和 null 都不会生成单元格记录。快照是一次性导入，不是响应式绑定。"
    )
  );

  const objectRows = section("对象数组（特殊导入）");
  objectRows.body.append(
    paragraph("对象数组可使用 rowsToWorkbookSnapshot(rows, fields) 导入。fields 指定本次导入的字段顺序，导入后的数据统一存储为快照。"),
    codeBlock(`function rowsToWorkbookSnapshot(
  rows: readonly ImportRow[],
  fields: readonly string[],
  options?: RowsToWorkbookSnapshotOptions
): WorkbookSnapshot`)
  );
  root.append(objectRows.root);
}
