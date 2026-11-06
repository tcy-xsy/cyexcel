import type { DocElement } from "../../components/content";
import { createSheet } from "@cyexcel/createSheet";
import { columns, snapshot } from "../../demo/data";
import {
  setText,
  actionButton,
  codeBlock,
  demoBlock,
  pageHeader,
  paragraph,
  section
} from "../../components/doc";

export default function stylesExamplePage(root: DocElement): void {
  let sheet: ReturnType<typeof createSheet>;
  root.append(
    pageHeader("EXAMPLES", "样式与格式", "样式以 patch 合并写入，可以只改 bold、颜色、背景或数字格式。")
  );

  const intro = section("按 patch 合并");
  intro.body.append(
    paragraph("setRangeStyle() 只覆盖传入的字段。传 null 会清除对应样式，其他字段保持不变。")
  );
  root.append(intro.root);
  root.append(
    codeBlock(`sheet.setRangeStyle(sheet.selection.range, {
  bold: true,
  color: "#0b6bcb",
  background: "#eaf3ff"
});`)
  );

  const bold = actionButton("当前选区加粗", () => {
    sheet.setRangeStyle(sheet.selection.range, { bold: true });
    setText(log, "已把当前选区设为 bold: true");
  });
  const highlight = actionButton("高亮当前选区", () => {
    sheet.setRangeStyle(sheet.selection.range, {
      color: "#0b6bcb",
      background: "#eaf3ff"
    });
    setText(log, "已给当前选区应用颜色和背景");
  });
  const clear = actionButton("清除当前选区样式", () => {
    sheet.clearRangeStyle(sheet.selection.range);
    setText(log, "已清除当前选区样式");
  });

  const demo = demoBlock({
    title: "选中区域应用样式",
    description: "选择单元格区域后，可使用操作按钮设置样式。",
    actions: [bold, highlight, clear]
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
      snapshot,
      headerHeight: 28,
      rowHeaderWidth: 48,
      rowHeight: 30
    });
    return () => sheet.dispose();
  };
}
