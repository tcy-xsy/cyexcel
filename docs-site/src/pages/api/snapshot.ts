import { pageHeader, section, apiTable, callout, codeBlock } from "../../components/doc";

export default function snapshotApiPage(root: HTMLElement): void {
  root.append(
    pageHeader("API", "rowsToWorkbookSnapshot()", "把外部对象数组导入成 WorkbookSnapshot；导入后两者不再互相影响。")
  );

  const signature = section("签名");
  signature.body.append(
    codeBlock(`function rowsToWorkbookSnapshot(
  rows: readonly ImportRow[],
  fields: readonly string[],
  options?: RowsToWorkbookSnapshotOptions
): WorkbookSnapshot`)
  );
  root.append(signature.root);
  root.append(
    apiTable([
      { name: "rows", type: "readonly ImportRow[]", description: "字段名到单元格值的对象数组。" },
      { name: "fields", type: "readonly string[]", description: "导入字段顺序。第 i 个字段就是第 i 列。" },
      { name: "options.blankRows", type: "number", description: "在数据行之后额外保留的空白行。" },
      { name: "options.blankColumns", type: "number", description: "在字段列之后额外保留的空白列。" }
    ])
  );

  root.append(
    callout(
      "undefined 和 null 都不会生成单元格记录。快照是一次性导入，不是响应式绑定。"
    )
  );
}
