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

export default function searchExamplePage(root: DocElement): void {
  let sheet: ReturnType<typeof createSheet>;
  root.append(
    pageHeader("EXAMPLES", "查找与替换", "find() 扫描显示文本，replace() 只替换命中的那一段。")
  );

  const intro = section("查找和部分替换");
  intro.body.append(
    paragraph("搜索匹配的是用户看到的文本，而不是底层存储值。替换时保留命中文本前后的内容。")
  );
  root.append(intro.root);
  root.append(
    codeBlock(`const matches = await sheet.find("east");
sheet.findNext();

await sheet.replaceAll("east", "西部");`)
  );

  const find = actionButton("查找 East", async () => {
    const matches = await sheet.find("east");
    sheet.findNext();
    setText(log, `找到 ${matches?.length ?? 0} 处，已选中第一处`);
  });
  const replace = actionButton("全部替换为“西部”", async () => {
    const count = await sheet.replaceAll("east", "西部");
    setText(log, `替换了 ${count ?? 0} 处`);
  });
  const reset = actionButton("恢复 East", async () => {
    await sheet.replaceAll("西部", "East");
    setText(log, "已恢复原始区域文本");
  });

  const demo = demoBlock({
    title: "分词区域示例",
    description: "East 是单元格文本的一部分，替换后不会覆盖整格。",
    actions: [find, replace, reset]
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
