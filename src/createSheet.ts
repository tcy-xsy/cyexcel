import { CyExcelCore, type TopWidgetFactory, type CyExcelCoreOptions } from "./core";
import { formulaBarWidget, toolbarWidget } from "./core/toolbar";
import type { ToolbarOptions } from "./core/toolbar";

export interface CreateSheetOptions extends CyExcelCoreOptions {
  /** 工具栏。`false` 关掉，对象则作为 `createToolbar` 的选项（换 schema、加项）。 */
  readonly toolbar?: boolean | ToolbarOptions;
  /** 编辑栏（名称框 + fx 内容框）。 */
  readonly formulaBar?: boolean;
}

/**
 * 带工具栏和编辑栏的表格——开箱即用的那个入口。
 *
 * 对应 Univer 的 preset：`cyexcel/core` 是不含 UI 的瘦核心，这里把常用件装齐。
 * 想精简就直接用 `CyExcelCore`，想换工具栏内容就传 `toolbar: { schema }`。
 *
 * 显式传 `topWidgets` 时以它为准，这两个开关不再生效。
 */
export function createSheet(container: HTMLElement, options: CreateSheetOptions = {}): CyExcelCore {
  const { toolbar = true, formulaBar = true, ...coreOptions } = options;
  if (coreOptions.topWidgets !== undefined) {
    return new CyExcelCore(container, { contextMenu: true, ...coreOptions });
  }

  const topWidgets: TopWidgetFactory[] = [];
  if (toolbar !== false) {
    topWidgets.push(toolbarWidget(toolbar === true ? {} : toolbar));
  }
  if (formulaBar !== false) {
    topWidgets.push(formulaBarWidget());
  }
  return new CyExcelCore(container, { contextMenu: true, ...coreOptions, topWidgets });
}
