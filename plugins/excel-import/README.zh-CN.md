# cyexcel-plugin-excel-import

cyexcel 的可选 XLSX 导入插件，作为独立工作区包构建。主包不依赖 ExcelJS，ESM 和 UMD 产物均不包含 Excel 解析器；插件在读取文件时才加载 ExcelJS。

## 安装

需要 cyexcel 0.1.1 及以上、1.0 以下版本，适用于使用 Vite、Webpack 等构建工具的 ESM 应用。

```sh
npm install cyexcel@^0.1.1 cyexcel-plugin-excel-import@^0.1.0
```

## 本地开发

在仓库根目录执行：

```sh
pnpm install
pnpm build
pnpm build:plugins
pnpm test:plugins
```

构建产物位于 `plugins/excel-import/dist/`。

## 接入

```ts
import { createWorkbook } from 'cyexcel';
import { createExcelImportPlugin } from 'cyexcel-plugin-excel-import';

const excelImport = createExcelImportPlugin({
  onWarning: message => console.warn(message)
});
const workbook = createWorkbook(container, {
  toolbar: { items: excelImport.toolbarItems }
});
excelImport.attach(workbook);

// 也可以直接导入，跳过工具栏对话框。
const result = await excelImport.importExcel(file);
console.log(result.warnings);

// 页面卸载时：
excelImport.dispose();
workbook.dispose();
```

工具栏要求确认替换后才执行导入，替换全部工作表并清空撤销历史。错误显示在当前对话框内；关闭对话框会取消应用尚未完成的数据。解析失败时保留原有工作簿。同一插件实例同时只能执行一次导入。

仅解析文件而不替换当前工作簿：

```ts
import { readExcelWorkbook } from 'cyexcel-plugin-excel-import';

const { sheets, warnings } = await readExcelWorkbook(file);
const workbook = createWorkbook(container, { sheets });
```

每项包含独立的单工作表快照。输入支持 `Blob`（包括 `File`）、`ArrayBuffer`、`Uint8Array`。可配置 `maxFileSize`（默认 25 MiB）、`maxCells`（默认 20 万个有值或有样式的单元格）、`signal` 和 `onWarning`。

文件在本地解析，不上传服务器。ExcelJS 先在内存中解压和解析文件，再检查单元格数量，并非流式处理。取消操作可阻止应用数据，但不会中断解析器的同步运算。

## 支持范围

- 工作表顺序和名称；隐藏工作表以可见标签页导入。
- 文本、数字、布尔值、Excel 日期序列号、公式文本和缓存结果，支持共享公式展开。
- 常用字体、填充、对齐、数字格式和边框。虚线等边框近似为实线，暂不导入索引颜色。
- 合并区域、冻结行列、行高、列宽及隐藏行列。
- 换行设置按整列应用，cyexcel 当前没有单元格级换行配置。

仅支持未加密的 `.xlsx`。暂不导入旧版 `.xls`、宏、图表、图片、数据验证、条件格式、命名区域和打印设置。超链接保留显示文字，富文本转为普通文本。重算范围遵循现有公式能力；跨表引用保留文本，但暂不能计算。兼容性提示通过返回结果的 `warnings` 和 `onWarning` 提供。

## 许可证

MIT。ExcelJS 作为独立依赖分发，遵循其 MIT 许可证。
