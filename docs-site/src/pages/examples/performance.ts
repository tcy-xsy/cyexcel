import { createSheet } from "@cyexcel/createSheet";
import type { CyExcelCore } from "@cyexcel/core";
import {
  buildPerformanceSnapshot,
  performanceColumns
} from "../../demo/performance-data";
import {
  actionButton,
  callout,
  demoBlock,
  pageHeader,
  paragraph,
  section
} from "../../components/doc";

export default function performanceExamplePage(root: HTMLElement): () => void {
  root.append(
    pageHeader(
      "EXAMPLES",
      "十万行性能",
      "100,000 行 × 5 列的真实运行示例。数据仍在内存里，但渲染只处理当前视口附近的单元格。"
    )
  );

  const intro = section("看的是滚动，不是全量绘制");
  intro.body.append(
    paragraph(
      "下面这张表每次加载都是完整导入 100,000 行。你可以直接拖到末尾、编辑一格、再滚回来，观察是否卡顿。"
    )
  );
  root.append(intro.root);

  root.append(
    callout(
      "性能提示：CyExcelCore 只虚拟化绘制和命中测试，不会虚拟化数据导入与内存。100,000 行仍会一次性导入快照并常驻内存；移动端或低端设备建议先从 10,000–50,000 行验证，百万级数据请用真实业务数据单独压测导入耗时、内存占用和滚动帧率。"
    )
  );

  let sheet: CyExcelCore | null = null;
  const logRef = { current: undefined as HTMLDivElement | undefined };
  const setLog = (text: string): void => {
    if (logRef.current) {
      logRef.current.textContent = text;
    }
  };
  const rebuild = actionButton("重新加载 100,000 行", () => {
    const importStart = performance.now();
    const snapshot = buildPerformanceSnapshot(100_000);
    const importMs = performance.now() - importStart;

    sheet?.dispose();
    const mountStart = performance.now();
    sheet = createSheet(demo.mount, {
      columns: performanceColumns,
      snapshot,
      headerHeight: 28,
      rowHeaderWidth: 48,
      rowHeight: 28,
      overscanRows: 4,
      overscanColumns: 2
    });
    const mountMs = performance.now() - mountStart;
    setLog([
      `导入快照：${importMs.toFixed(1)} ms`,
      `创建核心：${mountMs.toFixed(1)} ms`,
      `数据行：${sheet.getUsedRowCount().toLocaleString("zh-CN")}`,
      `坐标空间：${sheet.getRowCount().toLocaleString("zh-CN")} 行 × ${sheet.getColumnCount()} 列`
    ].join("\n"));
  });

  const demo = demoBlock({
    title: "100,000 rows",
    description: "默认就是十万行，不是占位图。",
    actions: [rebuild]
  });
  root.append(demo.root);
  if (demo.log === undefined) {
    throw new Error("demo log missing");
  }
  logRef.current = demo.log;

  const bootstrapStart = performance.now();
  const bootstrapSnapshot = buildPerformanceSnapshot(100_000);
  const bootstrapImportMs = performance.now() - bootstrapStart;
  const bootstrapMountStart = performance.now();
  sheet = createSheet(demo.mount, {
    columns: performanceColumns,
    snapshot: bootstrapSnapshot,
    headerHeight: 28,
    rowHeaderWidth: 48,
    rowHeight: 28,
    overscanRows: 4,
    overscanColumns: 2
  });
  const bootstrapMountMs = performance.now() - bootstrapMountStart;
  setLog([
    `导入快照：${bootstrapImportMs.toFixed(1)} ms`,
    `创建核心：${bootstrapMountMs.toFixed(1)} ms`,
    `数据行：${sheet.getUsedRowCount().toLocaleString("zh-CN")}`,
    `坐标空间：${sheet.getRowCount().toLocaleString("zh-CN")} 行 × ${sheet.getColumnCount()} 列`
  ].join("\n"));

  root.append(
    callout(
      "getUsedRowCount() 返回实际有数据的行数；getRowCount() 返回包含尾部空白行的坐标空间行数。"
    )
  );

  return () => sheet?.dispose();
}
