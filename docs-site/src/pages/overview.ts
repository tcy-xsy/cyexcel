import { createSheet } from "@cyexcel/createSheet";
import { columns, snapshot } from "../demo/data";
import { demoBlock, featureGrid, pageHeader, paragraph, section } from "../components/doc";

export default function overviewPage(root: HTMLElement): () => void {
  root.append(
    pageHeader(
      "OVERVIEW",
      "cyexcel",
      "面向浏览器的表格组件，采用稀疏矩阵存储和 Canvas 虚拟化渲染，支持命令事务与编辑历史管理。"
    )
  );

  const intro = section("组件概述");
  intro.body.append(
    paragraph(
      "cyexcel 将数据模型、布局计算、渲染和变更记录划分为独立模块。绘制范围限制在视口附近，编辑操作通过命令与事务提交，并记录可撤销的变更。"
    )
  );
  root.append(intro.root);
  root.append(
    featureGrid([
      { title: "Canvas 虚拟化", text: "仅绘制视口附近的单元格，减少滚动时的绘制开销。" },
      { title: "Command 事务", text: "单元格、结构、样式和行高变化都经过统一命令层。" },
      { title: "撤销 / 重做", text: "编辑、合并、插入、删除和样式修改可以进入同一份历史记录。" },
      { title: "框架无关", text: "支持通过 CyExcelCore 接入原生页面及 React、Vue 等框架。" },
      { title: "多 Sheet 工作簿", text: "createWorkbook() 提供工作表添加、删除、切换与重命名功能。" },
      { title: "十万行示例", text: "提供 100,000 行 × 5 列的数据示例，用于评估导入、编辑和滚动表现。" }
    ])
  );

  const demo = demoBlock({
    title: "基础表格",
    description: "支持单元格选择、滚动、编辑和列宽调整。"
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
