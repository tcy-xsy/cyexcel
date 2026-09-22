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
      "三步接入：安装包、把对象数组导成快照、把快照交给表格组件。"
    )
  );

  const install = section("安装");
  install.body.append(paragraph("在已有前端工程中安装当前包："));
  root.append(install.root);
  root.append(codeBlock("pnpm add cyexcel", "shell"));

  const create = section("创建表格");
  create.body.append(
    paragraph("最直接的入口是 createSheet()，它会同时挂上工具栏和编辑栏。"),
    paragraph("数据不是通过数组直接绑定进组件，而是先导入成一份 WorkbookSnapshot。")
  );
  root.append(create.root);
  root.append(
    codeBlock(`import { createSheet } from "cyexcel";
import { rowsToWorkbookSnapshot } from "cyexcel/core";

const rows = [
  { sku: "CY-100001", name: "工业温度传感器", price: 128 },
  { sku: "CY-100002", name: "便携标签打印机", price: 465 }
];

const columns = [
  { width: 130, header: "SKU", sortable: false, filterable: false },
  { width: 220, header: "名称", sortable: false, filterable: false },
  { width: 100, header: "价格", sortable: false, filterable: false }
];

const sheet = createSheet(document.querySelector("#sheet"), {
  columns,
  snapshot: rowsToWorkbookSnapshot(rows, ["sku", "name", "price"]),
  headerHeight: 28,
  rowHeaderWidth: 48
});`)
  );

  root.append(
    callout(
      "columns 的下标和 fields 的下标一一对应。导入之后，表格只使用 (row, column) 坐标；外部数据对象不会再被写回。"
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
