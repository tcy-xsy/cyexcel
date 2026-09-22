import { createWorkbook } from "@cyexcel/createWorkbook";
import {
  columns,
  rows,
  summaryColumns,
  summarySnapshot
} from "../../demo/data";
import { matrixToWorkbookSnapshot } from "@cyexcel/core";
import { demoBlock, pageHeader, paragraph, section, codeBlock } from "../../components/doc";

export default function workbookExamplePage(root: HTMLElement): () => void {
  root.append(
    pageHeader(
      "EXAMPLES",
      "多 Sheet 工作簿",
      "createWorkbook() 提供底部 Sheet 标签、活跃 Sheet 切换和添加 Sheet 按钮。"
    )
  );

  const intro = section("一个工作簿，多个独立 Sheet");
  intro.body.append(
    paragraph("每个 Sheet 都有独立的 CyExcelCore，因此列配置、数据、选区和撤销栈都可以完全不同。")
  );
  root.append(intro.root);
  root.append(
    codeBlock(`const workbook = createWorkbook(container, {
  newSheetRows: 1000,
  sheets: [
    { name: "订单明细", columns, snapshot: orderSnapshot },
    { name: "区域汇总", columns: summaryColumns, snapshot: summarySnapshot }
  ]
});

workbook.getActiveSheet().insertRows(0, 1);`)
  );

  const demo = demoBlock({
    title: "底部 Sheet 标签",
    description: "切换底部标签；左下角的“+”用于添加 Sheet，行操作走当前 Sheet 的核心 API。"
  });
  demo.mount.classList.add("with-toolbar");
  root.append(demo.root);
  const workbook = createWorkbook(demo.mount, {
    toolbar: true,
    formulaBar: true,
    sheets: [
      {
        name: "订单明细",
        columns,
        snapshot: matrixToWorkbookSnapshot(rows),
        headerHeight: 28,
        rowHeaderWidth: 48
      },
      {
        name: "区域汇总",
        columns: summaryColumns,
        snapshot: summarySnapshot,
        headerHeight: 28,
        rowHeaderWidth: 48
      }
    ]
  });

  return () => workbook.dispose();
}
