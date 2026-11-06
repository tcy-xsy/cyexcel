import { createWorkbook } from "@cyexcel/createWorkbook";
import { matrixToWorkbookSnapshot } from "@cyexcel/core";
import { codeBlock, demoBlock, pageHeader, paragraph } from "../../components/doc";

export default function contextMenuExamplePage(root: HTMLElement): () => void {
  root.append(pageHeader("EXAMPLES", "右键菜单", "在单元格、行号、列头右键，菜单会按当前位置和选区提供对应操作。"));
  root.append(paragraph("试试框选多个单元格后清除内容、合并或设置格式；选中多行/多列后右键可批量插入、删除。操作后可用撤销恢复。"));
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
