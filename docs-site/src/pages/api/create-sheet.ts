import { createSheet } from "@cyexcel/createSheet";
import { columns, snapshot } from "../../demo/data";
import { apiTable, codeBlock, demoBlock, pageHeader, section } from "../../components/doc";

export default function createSheetApiPage(root: HTMLElement): () => void {
  root.append(
    pageHeader(
      "API",
      "createSheet()",
      "开箱即用入口。内部会创建 CyExcelCore，并把工具栏、编辑栏挂到表格上方。"
    )
  );

  const signature = section("签名");
  signature.body.append(
    codeBlock(`function createSheet(
  container: HTMLElement,
  options?: CreateSheetOptions
): CyExcelCore`)
  );
  root.append(signature.root);
  root.append(
    apiTable([
      { name: "container", type: "HTMLElement", description: "表格挂载容器，必须有可计算宽高。" },
      { name: "options.toolbar", type: "boolean | ToolbarOptions", description: "默认 true。传 false 关闭工具栏，传对象可替换或追加 schema。" },
      { name: "options.formulaBar", type: "boolean", description: "默认 true。控制名称框和公式内容框。" },
      { name: "options.columns", type: "CyExcelColumnConfig[]", description: "列显示配置。下标即列身份。" },
      { name: "options.snapshot", type: "WorkbookSnapshot", description: "初始数据快照，推荐由 matrixToWorkbookSnapshot() 创建。" }
    ])
  );

  const demo = demoBlock({ title: "示例" });
  demo.mount.classList.add("with-toolbar");
  root.append(demo.root);
  const sheet = createSheet(demo.mount, {
    columns,
    snapshot,
    toolbar: true,
    formulaBar: true,
    headerHeight: 28,
    rowHeaderWidth: 48
  });

  return () => sheet.dispose();
}
