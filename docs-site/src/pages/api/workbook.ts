import type { DocElement } from "../../components/content";
import { createWorkbook } from "@cyexcel/createWorkbook";
import { columns, rows } from "../../demo/data";
import { matrixToWorkbookSnapshot } from "@cyexcel/core";
import { apiTable, codeBlock, demoBlock, pageHeader, section } from "../../components/doc";

export default function workbookApiPage(root: DocElement): void {
  root.append(
    pageHeader(
      "API",
      "createWorkbook()",
      "创建带底部 Sheet 标签和添加 Sheet 按钮的多 Sheet 工作簿。"
    )
  );

  const signature = section("签名");
  signature.body.append(
    codeBlock(`function createWorkbook(
  container: HTMLElement,
  options?: CreateWorkbookOptions
): WorkbookHandle`)
  );
  root.append(signature.root);
  root.append(
    apiTable([
      { name: "sheets", type: "WorkbookSheetOptions[]", description: "初始 Sheet 列表；每项包含自己的 columns 和 snapshot。" },
      { name: "toolbar", type: "boolean | ToolbarOptions", description: "默认 true；控制每个 Sheet 的工具栏。" },
      { name: "formulaBar", type: "boolean", description: "默认 true；控制每个 Sheet 的编辑栏。" },
      { name: "newSheetRows", type: "number", description: "新增空白 Sheet 的初始空白行数，默认 1000。" },
      { name: "...CyExcelCoreOptions", type: "object", description: "作为所有初始 Sheet 的默认核心配置；Sheet 内同名选项优先。" }
    ])
  );

  const handle = section("WorkbookHandle");
  root.append(handle.root);
  root.append(
    apiTable([
      { name: "getActiveSheet()", type: "CyExcelCore", description: "返回当前活跃 Sheet 的核心实例。" },
      { name: "setActiveSheet(sheetId)", type: "void", description: "切换活跃 Sheet。" },
      { name: "addSheet(options?)", type: "string", description: "追加 Sheet 并返回新 id。" },
      { name: "—", type: "—", description: "行级添加/删除不在 WorkbookHandle 上；使用活跃 CyExcelCore 的 insertRows() / deleteRows()。" },
      { name: "dispose()", type: "void", description: "销毁所有 Sheet 和底部标签栏。" }
    ])
  );

  const demo = demoBlock({ title: "示例" });
  demo.mount.addClass("with-toolbar");
  root.append(demo.root);
  demo.mount.onMount = () => {
    const workbook = createWorkbook(demo.mount.element, {
      toolbar: true,
      formulaBar: true,
      sheets: [
        {
          name: "订单",
          columns,
          snapshot: matrixToWorkbookSnapshot(rows),
          headerHeight: 28,
          rowHeaderWidth: 48
        },
        {
          name: "汇总",
          columns,
          headerHeight: 28,
          rowHeaderWidth: 48
        }
      ]
    });
    return () => workbook.dispose();
  };
}
