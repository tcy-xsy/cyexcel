import { createWorkbook } from "@cyexcel/createWorkbook";
import { matrixToWorkbookSnapshot } from "@cyexcel/core";
import { codeBlock, demoBlock, pageHeader, paragraph } from "../../components/doc";

export default function formulasExamplePage(root: HTMLElement): () => void {
  root.append(pageHeader("EXAMPLES", "公式计算", "修改数量或单价，金额与合计自动重算。双击结果或使用公式栏查看和编辑原文。"));
  root.append(paragraph("首版支持同表 A1 引用及 12 个常用函数。复制与填充使用当前值；含公式时暂不支持行列增删、移动和排序。"));
  const demo = demoBlock({ title: "采购金额", description: "D1 = B1*C1，D4 = SUM(D1:D3)。修改 B1 后更新计算结果，支持撤销与重做。" });
  demo.mount.classList.add("with-toolbar");
  root.append(demo.root);
  const workbook = createWorkbook(demo.mount, {
    toolbar: true,
    columns: [
      { header: "物料", width: 180 }, { header: "数量", width: 120, type: "number" },
      { header: "单价", width: 120, type: "number" }, { header: "金额", width: 160 }
    ],
    snapshot: matrixToWorkbookSnapshot([["传感器", 3, 120, 0], ["电源", 5, 80, 0], ["线缆", 10, 15, 0], ["合计", undefined, undefined, 0]], { blankRows: 20 }),
    rowHeight: 32
  });
  const sheet = workbook.getActiveSheet();
  for (let row = 0; row < 3; row++) sheet.setCell(row, 3, `=B${row + 1}*C${row + 1}`);
  sheet.setCell(3, 3, "=SUM(D1:D3)");
  sheet.setCellStyle(3, 3, { bold: true });
  root.append(codeBlock(`sheet.setCell(0, 3, "=B1*C1");
sheet.setCell(3, 3, "=SUM(D1:D3)");
sheet.getCell(3, 3); // 910
sheet.getCellFormula(3, 3); // "=SUM(D1:D3)"`));
  return () => workbook.dispose();
}
