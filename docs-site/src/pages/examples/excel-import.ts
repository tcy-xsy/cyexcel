import type { DocElement } from "../../components/content";
import { createWorkbook } from "@cyexcel/createWorkbook";
import { matrixToWorkbookSnapshot } from "@cyexcel/core";
import { defaultToolbarSchema } from "@cyexcel/core/toolbar";
import { createExcelImportPlugin } from "../../../../plugins/excel-import/src";
import { createExcelImportSample } from "../../../../plugins/excel-import/examples/sample";
import { actionButton, demoBlock, pageHeader, paragraph, section, codeBlock, setText } from "../../components/doc";

export default function excelImportExamplePage(root: DocElement): void {
  root.append(
    pageHeader(
      "EXAMPLES",
      "Excel 导入",
      "createExcelImportPlugin() 提供可选的 .xlsx 导入，在浏览器内解析并替换当前工作簿。"
    )
  );

  const intro = section("Excel 导入插件");
  intro.body.append(
    paragraph("Excel 导入是独立的可选插件，主包不包含解析器。文件仅在浏览器内读取，导入成功后替换示例中的全部工作表。"),
    paragraph("支持 .xlsx 的数据、公式、常用样式、合并区域及行列尺寸。暂不导入 .xls、加密文件、图片、图表和数据验证规则。"),
    paragraph("示例文件包含六张工作表：数据类型、数字格式、公式、样式、布局和隐藏工作表。涵盖文本、数值、布尔值、日期、错误值、共享公式、字体与边框、合并单元格、冻结窗格及隐藏行列。超链接和富文本保留显示文本，隐藏工作表导入后显示为可见标签。")
  );
  root.append(intro.root);
  const sampleSheets = () => [
    { name: "Sales", snapshot: matrixToWorkbookSnapshot([["Product", "Quantity", "Amount"], ["Keyboard", 4, 320], ["Mouse", 8, 240]]) },
    { name: "Inventory", snapshot: matrixToWorkbookSnapshot([["Product", "Stock"], ["Keyboard", 24], ["Mouse", 36]]) }
  ];
  let resetImportExample: (() => void) | undefined;
  const importDemo = demoBlock({
    title: "导入 .xlsx 文件",
    description: "可先下载示例文件，再点击工具栏的“导入 Excel”。勾选“替换当前工作簿”后执行导入。",
    actions: [
      actionButton("下载示例 Excel", async () => {
        const url = URL.createObjectURL(await createExcelImportSample());
        const link = document.createElement("a");
        link.href = url;
        link.download = "cyexcel-import-example.xlsx";
        link.click();
        setTimeout(() => URL.revokeObjectURL(url), 1000);
      }),
      actionButton("重置示例", () => resetImportExample?.())
    ]
  });
  importDemo.mount.addClass("with-toolbar");
  root.append(importDemo.root);
  importDemo.mount.onMount = () => {
    let cleanup = () => {};
    const mount = () => {
      const plugin = createExcelImportPlugin({
        onWarning: message => importDemo.log?.setText(message)
      });
      const workbook = createWorkbook(importDemo.mount.element, {
        toolbar: {
          schema: defaultToolbarSchema.filter(item => ["undo", "redo", "export-excel"].includes(item.id)),
          items: plugin.toolbarItems
        },
        headerHeight: 28,
        rowHeaderWidth: 48,
        sheets: sampleSheets()
      });
      plugin.attach(workbook);
      cleanup = () => { plugin.dispose(); workbook.dispose(); };
    };
    resetImportExample = () => {
      cleanup();
      mount();
      if (importDemo.log) setText(importDemo.log, "等待操作…");
    };
    mount();
    return () => { resetImportExample = undefined; cleanup(); };
  };
  root.append(codeBlock(`import { createWorkbook } from 'cyexcel';
import { createExcelImportPlugin } from 'cyexcel-plugin-excel-import';

const excelImport = createExcelImportPlugin();
const workbook = createWorkbook(container, {
  toolbar: { items: excelImport.toolbarItems }
});
excelImport.attach(workbook);

// On unmount:
excelImport.dispose();
workbook.dispose();`));
}
