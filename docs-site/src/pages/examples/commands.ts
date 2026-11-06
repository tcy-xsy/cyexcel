import { createSheet } from "@cyexcel/createSheet";
import { columns, snapshot } from "../../demo/data";
import {
  actionButton,
  codeBlock,
  demoBlock,
  pageHeader,
  paragraph,
  section
} from "../../components/doc";

export default function commandsExamplePage(root: HTMLElement): () => void {
  root.append(
    pageHeader("EXAMPLES", "编辑与历史", "setCell() 通过命令层提交修改，支持通过 undo() 和 redo() 撤销与重做。")
  );

  const intro = section("提交可撤销命令");
  intro.body.append(
    paragraph("单元格写入通过命令层执行，历史记录由组件统一管理。")
  );
  root.append(intro.root);
  root.append(
    codeBlock(`sheet.setCell(0, 1, "已确认订单");
sheet.undo();
sheet.redo();`)
  );

  const writeCell = actionButton("写入第 1 行第 2 列", () => {
    sheet.setCell(0, 1, "已确认订单");
    log.textContent = `第 1 行第 2 列已写入，当前值：${String(sheet.getCell(0, 1))}`;
  });
  const undo = actionButton("撤销", () => {
    sheet.undo();
    log.textContent = `已撤销，当前值：${String(sheet.getCell(0, 1))}`;
  });
  const redo = actionButton("重做", () => {
    sheet.redo();
    log.textContent = `已重做，当前值：${String(sheet.getCell(0, 1))}`;
  });

  const demo = demoBlock({
    title: "撤销与重做",
    description: "写入单元格后，可通过撤销与重做恢复对应状态。",
    actions: [writeCell, undo, redo]
  });
  root.append(demo.root);
  if (demo.log === undefined) {
    throw new Error("demo log missing");
  }
  const log = demo.log;

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
