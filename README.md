# cyexcel

cyexcel 是一个与前端框架无关的表格组件，使用 Canvas 进行视口虚拟化渲染，以稀疏矩阵存储单元格数据，并通过命令、事务和历史记录管理编辑操作。

## 功能

- 单元格编辑、样式设置、合并与取消合并。
- 行列插入、删除、移动及尺寸调整。
- 撤销重做、剪贴板、查找替换、排序与筛选。
- 数据验证、下拉选项、链接与图片单元格。
- 多工作表管理，支持添加、删除、切换和重命名。
- 基础公式计算及 XLSX 导出。

数据在创建表格时导入快照，导入后的编辑与原始数组相互独立。虚拟化用于绘制和命中测试，数据仍完整保存在内存中。文档站提供 100,000 行 × 5 列示例；实际容量取决于数据内容、设备性能与内存，需结合业务数据评估。

## 安装

```bash
npm install cyexcel
```

## 快速开始

挂载容器需要具有可计算的宽度和高度。

```html
<div id="sheet" style="width: 100%; height: 480px;"></div>
```

```ts
import { createSheet, matrixToWorkbookSnapshot } from "cyexcel";

const container = document.querySelector<HTMLElement>("#sheet");
if (container === null) {
  throw new Error("Missing sheet container");
}

const sheet = createSheet(container, {
  snapshot: matrixToWorkbookSnapshot([
    ["A100", "传感器", 9.99],
    ["A101", "电源模组", 14.5]
  ])
});

sheet.setCell(0, 2, 12.99);
sheet.undo();

// 页面或组件卸载时调用 sheet.dispose() 释放资源。
```

`createSheet()` 默认启用工具栏、编辑栏和右键菜单。可通过 `toolbar: false`、`formulaBar: false` 和 `contextMenu: false` 分别关闭。

`matrixToWorkbookSnapshot()` 将二维数组的第 i 行、第 j 列映射到单元格 `(i, j)`。未配置表头时，列标题使用 A、B、C 等字母。对象数组可通过 `rowsToWorkbookSnapshot(rows, fields)` 按字段顺序导入。

## 模块入口

| 入口 | 说明 |
| --- | --- |
| `cyexcel` | 表格、工作簿、快照导入器及工具栏等常用 API |
| `cyexcel/core` | 核心类、类型及扩展 API，不访问 Custom Elements 注册表，不引入公式解析库 |
| `cyexcel/toolbar` | 工具栏与编辑栏 API |
| `cyexcel/formula` | 可选公式引擎 |
| `cyexcel/register` | 导入时注册 `<cy-excel>` 自定义元素；当前元素不提供完整的表格配置接口 |

`CyExcelCore` 可挂载到普通 HTML 容器，也可在 React、Vue 等框架的组件生命周期中创建和释放。

## 工作簿

`createWorkbook(container, options)` 提供工作表标签及添加、切换、删除和重命名操作。

- 点击标签的关闭按钮，在确认后删除工作表。删除不可撤销，且至少保留一张工作表。
- 删除当前工作表后，优先切换到右侧相邻工作表；右侧不存在时切换到左侧。
- 双击标签或在标签获得焦点时按 F2 可编辑名称。Enter 或失焦保存，Escape 取消，空白输入保留原名称。
- `workbook.deleteSheet(sheetId)` 和 `workbook.renameSheet(sheetId, name)` 提供程序化操作。无效 ID、删除最后一张工作表或传入空白名称时抛出 `RangeError`。

默认工作簿工具栏支持导出全部工作表，包括尚未打开的工作表：

```ts
await workbook.exportExcel({ fileName: "report.xlsx" });
```

导出使用当前工作表名称，支持样式、合并区域和内置数据验证规则。`download: false` 返回 Blob，`includeHeaders: false` 不导出自定义列标题。图片、自定义 JavaScript 验证规则及部分高级设置暂不支持导出。

## 公式与使用限制

工作簿默认通过 MIT 许可的 `fast-formula-parser` 启用公式。核心类可通过 `formulaEngine` 接入可选计算引擎，配置示例见文档站的「公式计算」页面。

- 当前支持 12 个基础函数及同工作表引用，暂不支持跨工作表引用。
- 含公式的工作表暂时限制行列插入、删除、移动及排序，尚未实现结构变更后的引用调整。
- 复制与填充按值处理，暂不支持公式引用自动调整。
- 支持 XLSX 导出，暂不支持 XLSX 导入。

## 文档与示例

本地启动文档站：

```bash
pnpm install
pnpm docs:dev
```

文档站包含快速开始、功能示例、API、数据结构及扩展说明，源码位于 [docs-site/](docs-site/)。

`examples/` 包含原生 HTML、ESM、自定义渲染器及 React、Vue 接入示例。先执行 `pnpm build`，再使用 `npx serve examples` 启动静态服务。React、Vue 示例使用浏览器 import map 从 CDN 加载依赖。

## AI 编程助手 Skill

仓库提供 [cyexcel Skill](skills/cyexcel/SKILL.md)，供 Claude Code 和 Codex 在接入、配置和扩展组件时使用，包含 API 示例、生命周期说明及当前能力限制。

在需要使用组件的项目目录中安装：

```bash
npx skills add tcy-xsy/cyexcel --skill cyexcel --agent claude-code codex
```

按安装提示完成配置；添加 `--global` 可安装到用户级目录。安装后重新启动对应助手的会话：

- Claude Code：输入 `/cyexcel`，并描述需要实现的功能。
- Codex：输入 `$cyexcel`，并描述需要实现的功能。

例如：“使用 cyexcel 创建一个带数量验证、金额公式和 XLSX 导出的采购工作簿。”

Skill 仅提供助手指引，不安装组件运行时，不提供远程操作表格的服务。当前仓库包仍为私有开发版本；未发布 npm 版本时，可在源码仓库执行 `pnpm build` 和 `pnpm pack`，再将生成的压缩包安装到业务项目。

## 许可证

[MIT](LICENSE)
