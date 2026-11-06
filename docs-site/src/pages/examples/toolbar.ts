import { createSheet } from "@cyexcel/createSheet";
import { columns, snapshot } from "../../demo/data";
import { codeBlock, demoBlock, pageHeader, paragraph, section } from "../../components/doc";

export default function toolbarExamplePage(root: HTMLElement): () => void {
  root.append(
    pageHeader("EXAMPLES", "工具栏", "createSheet() 默认提供工具栏和编辑栏，支持通过 schema 自定义工具栏内容。")
  );

  const intro = section("默认配置");
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
    description: "容器宽度不足时，工具栏操作显示在溢出菜单中。查找替换使用非模态对话框。"
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
