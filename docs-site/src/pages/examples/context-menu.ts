import { createWorkbook } from "@cyexcel/createWorkbook";
import { matrixToWorkbookSnapshot } from "@cyexcel/core";
import { codeBlock, demoBlock, pageHeader, paragraph } from "../../components/doc";

export default function contextMenuExamplePage(root: HTMLElement): () => void {
  root.append(pageHeader("EXAMPLES", "右键菜单", "右键单击单元格、行号或列头，可打开与当前位置及选区对应的操作菜单。"));
  root.append(paragraph("连续选区支持清除内容、合并和格式设置；整行或整列选区支持批量插入与删除。上述操作支持撤销。"));
  const demo = demoBlock({ title: "库存表", description: "悬停“插入”“删除”“单元格格式”“排序和筛选”展开子菜单。Shift+F10 也可打开菜单。" });
  demo.mount.classList.add("with-toolbar");
  root.append(demo.root);
  const workbook = createWorkbook(demo.mount, {
    columns: [
      { header: "商品", width: 180, sortable: true, filterable: true },
      { header: "数量", width: 120, type: "number", sortable: true, filterable: true },
      { header: "分类", width: 140, sortable: true, filterable: true },
      { header: "备注", width: 200 }
    ],
    rowHeight: 30,
    snapshot: matrixToWorkbookSnapshot([["传感器", 12, "器件", "首批到货"], ["电源", 8, "器件", "待检"], ["线缆", 30, "耗材", ""], ["标签", 200, "耗材", "备用"]], { blankRows: 26 })
  });
  root.append(codeBlock(`createWorkbook(container, { contextMenu: true, snapshot });
// 预设默认开启；contextMenu: false 可关闭。
// 核心入口显式启用：new CyExcelCore(container, { contextMenu: true, snapshot });
sheet.clearContents(); // 清除当前连续选区，保留格式，可撤销`));
  return () => workbook.dispose();
}
