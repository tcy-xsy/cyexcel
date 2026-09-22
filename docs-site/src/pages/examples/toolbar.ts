import { createSheet } from "@cyexcel/createSheet";
import { columns, snapshot } from "../../demo/data";
import { codeBlock, demoBlock, pageHeader, paragraph, section } from "../../components/doc";

export default function toolbarExamplePage(root: HTMLElement): () => void {
  root.append(
    pageHeader("EXAMPLES", "工具栏", "createSheet() 提供开箱即用的工具栏和编辑栏，也可以整体替换 schema。")
  );

  const intro = section("开箱即用");
  intro.body.append(paragraph("createSheet() 默认同时启用工具栏和编辑栏。"));
  root.append(intro.root);
  root.append(
    codeBlock(`const sheet = createSheet(container, {
  columns,
  snapshot,
  toolbar: true,
  formulaBar: true
});`)
  );

  const demo = demoBlock({
    title: "完整 UI",
    description: "工具栏被收进溢出菜单时仍然可用；查找替换是非模态对话框。"
  });
  demo.mount.classList.add("with-toolbar");
  root.append(demo.root);
  const sheet = createSheet(demo.mount, {
    columns,
    snapshot,
    toolbar: true,
    formulaBar: true,
    headerHeight: 28,
    rowHeaderWidth: 48,
    rowHeight: 30
  });

  return () => sheet.dispose();
}
