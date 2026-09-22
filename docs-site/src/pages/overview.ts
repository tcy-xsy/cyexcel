import { createSheet } from "@cyexcel/createSheet";
import { columns, snapshot } from "../demo/data";
import { demoBlock, featureGrid, pageHeader, paragraph, section } from "../components/doc";

export default function overviewPage(root: HTMLElement): () => void {
  root.append(
    pageHeader(
      "OVERVIEW",
      "cyexcel",
      "一个面向大规模数据的表格组件：数据进入稀疏矩阵，所有变更都走 Command、事务和历史栈。"
    )
  );

  const intro = section("为什么用 cyexcel");
  intro.body.append(
    paragraph(
      "cyexcel 不是把表格包在响应式框架里的 UI 壳，它把数据、布局、渲染和变更记录放在一层可校验的核心里。滚动路径保持轻，编辑路径保持可追溯。"
    )
  );
  root.append(intro.root);
  root.append(
    featureGrid([
      { title: "Canvas 虚拟化", text: "渲染只处理视口附近的单元格，适合几十万行级别的数据浏览。" },
      { title: "Command 事务", text: "单元格、结构、样式和行高变化都经过统一命令层。" },
      { title: "撤销 / 重做", text: "编辑、合并、插入、删除和样式修改可以进入同一份历史记录。" },
      { title: "框架无关", text: "可以直接用 CyExcelCore，也可以在 React、Vue 或原生页面中接入。" },
      { title: "多 Sheet 工作簿", text: "createWorkbook() 提供底部 Sheet 标签和添加工作表入口。" },
      { title: "十万行示例", text: "文档内直接跑 100,000 行 × 5 列的真实数据表。" }
    ])
  );

  const demo = demoBlock({
    title: "基础表格",
    description: "下面是真实运行的表格，可以点击单元格、滚动、编辑和拖拽列宽。"
  });
  root.append(demo.root);
  demo.mount.classList.add("with-toolbar");
  const sheet = createSheet(demo.mount, {
    columns,
    snapshot,
    headerHeight: 28,
    rowHeaderWidth: 48,
    rowHeight: 28
  });

  return () => sheet.dispose();
}
