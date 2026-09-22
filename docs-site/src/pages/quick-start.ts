import { createSheet } from "@cyexcel/createSheet";
import { columns, snapshot } from "../demo/data";
import {
  callout,
  codeBlock,
  demoBlock,
  pageHeader,
  paragraph,
  section
} from "../components/doc";

export default function quickStartPage(root: HTMLElement): () => void {
  root.append(
    pageHeader(
      "QUICK START",
      "快速开始",
      "三步接入：安装包、把 Excel 式二维数组导成快照、把快照交给表格组件。"
    )
  );

  const install = section("安装");
  install.body.append(paragraph("在已有前端工程中安装当前包："));
  root.append(install.root);
  root.append(codeBlock("pnpm add cyexcel", "shell"));

  const create = section("创建表格");
  create.body.append(
    paragraph("最直接的入口是 createSheet()，它会同时挂上工具栏和编辑栏。"),
    paragraph("数据不是通过二维数组直接绑定进组件，而是先导入成一份 WorkbookSnapshot。")
  );
  root.append(create.root);
  root.append(
    codeBlock(`import { createSheet } from "cyexcel";
import { matrixToWorkbookSnapshot } from "cyexcel/core";

const rows = [
  ["CY-100001", "工业温度传感器", 128],
  ["CY-100002", "便携标签打印机", 465]
];

const sheet = createSheet(document.querySelector("#sheet"), {
  snapshot: matrixToWorkbookSnapshot(rows),
  headerHeight: 28,
  rowHeaderWidth: 48
});`)
  );

  root.append(
    callout(
      "row i / column j 直接落到单元格 (i, j)；没有自定义表头时默认显示 A/B/C。导入后外部数组不会再被写回。"
    )
  );

  const demo = demoBlock({
    title: "带工具栏的表格",
    description: "使用 createSheet() 创建，包含撤销重做、格式、数据操作和结构操作入口。"
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
    rowHeight: 28
  });

  return () => sheet.dispose();
}
