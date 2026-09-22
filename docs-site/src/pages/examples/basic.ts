import { createSheet } from "@cyexcel/createSheet";
import { columns, snapshot } from "../../demo/data";
import { codeBlock, demoBlock, pageHeader, paragraph, section } from "../../components/doc";

export default function basicExamplePage(root: HTMLElement): () => void {
  root.append(pageHeader("EXAMPLES", "基础表格", "用 createSheet() 创建一个带工具栏的可编辑表格。"));

  const intro = section("只挂核心");
  intro.body.append(
    paragraph("createSheet() 也不依赖 Custom Elements registry；默认会同时挂上工具栏和编辑栏。")
  );
  root.append(intro.root);
  root.append(
    codeBlock(`import { CyExcelCore, rowsToWorkbookSnapshot } from "cyexcel/core";

const sheet = createSheet(document.querySelector("#sheet"), {
  columns,
  snapshot: rowsToWorkbookSnapshot(rows, fields),
  headerHeight: 28,
  rowHeaderWidth: 48
});`)
  );

  const demo = demoBlock({
    title: "真实运行",
    description: "可以编辑单元格、调整列宽、多选单元格和使用键盘导航。"
  });
  root.append(demo.root);
  demo.mount.classList.add("with-toolbar");
  const sheet = createSheet(demo.mount, {
    columns,
    snapshot,
    headerHeight: 28,
    rowHeaderWidth: 48,
    rowHeight: 30
  });

  return () => sheet.dispose();
}
