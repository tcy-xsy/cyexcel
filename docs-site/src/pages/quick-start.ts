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
      "安装组件，将二维数组转换为数据快照，并使用快照创建表格实例。"
    )
  );

  const install = section("安装");
  install.body.append(paragraph("在已有前端工程中安装当前包："));
  root.append(install.root);
  root.append(codeBlock("pnpm add cyexcel", "shell"));

  const create = section("创建表格");
  create.body.append(
    paragraph("createSheet() 用于创建包含工具栏和编辑栏的表格实例。"),
    paragraph("初始数据通过 WorkbookSnapshot 导入，组件不与原始数组建立响应式绑定。")
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
      "二维数组的第 i 行、第 j 列映射到单元格 (i, j)。未设置自定义表头时，列标题使用 A、B、C 等字母。表格编辑不会修改原始数组。"
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
