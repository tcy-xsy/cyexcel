import type { DocElement } from "../../components/content";
import { pageHeader, paragraph, section, apiTable, callout } from "../../components/doc";

export default function coreApiPage(root: DocElement): void {
  root.append(pageHeader("API", "CyExcelCore", "表格核心类，提供渲染、交互与数据操作 API，通过 Command、Mutation 和 History 管理变更。"));

  const data = section("数据");
  root.append(data.root);
  root.append(
    apiTable([
      { name: "getCell(row, column)", type: "unknown", description: "读取一个单元格的底层值。" },
      { name: "getUsedRowCount()", type: "number", description: "实际有数据的行数，不含尾部空白行。" },
      { name: "getRowCount()", type: "number", description: "包含尾部空白行的坐标空间行数。" },
      { name: "getColumnCount()", type: "number", description: "包含尾部空白列的坐标空间列数。" },
      { name: "setCell(row, column, value)", type: "void", description: "写入一个可撤销的单元格值；null / undefined 表示清空。" },
      { name: "undo()", type: "void", description: "撤销最近一次可撤销变更。" },
      { name: "redo()", type: "void", description: "重做最近一次被撤销的变更。" }
    ])
  );

  const structure = section("结构与几何");
  root.append(structure.root);
  root.append(
    apiTable([
      { name: "insertRows(start, count)", type: "void", description: "插入空行。" },
      { name: "deleteRows(start, count)", type: "void", description: "删除连续行。" },
      { name: "moveRows(start, count, target)", type: "void", description: "插入式移动行。" },
      { name: "insertColumns(start, columns)", type: "void", description: "插入列定义。" },
      { name: "deleteColumns(start, count)", type: "void", description: "删除连续列。" },
      { name: "moveColumns(start, count, target)", type: "void", description: "移动列。" },
      { name: "setRowHeight(row, height)", type: "void", description: "设置精确行高。" },
      { name: "setColumnWidth(column, width)", type: "void", description: "设置精确列宽。" }
    ])
  );

  const styles = section("样式");
  root.append(styles.root);
  root.append(
    apiTable([
      { name: "setCellStyle(row, column, patch)", type: "void", description: "给单格应用样式 patch。" },
      { name: "setRangeStyle(range, patch)", type: "void", description: "给区域应用样式 patch，patch 语义为逐键覆盖。" },
      { name: "getCellStyle(row, column)", type: "CellStyle | undefined", description: "读取列默认样式和单元格样式合并后的结果。" },
      { name: "clearRangeStyle(range)", type: "void", description: "清除区域内已设置的样式键。" }
    ])
  );

  const interaction = section("查找、排序与剪贴板");
  root.append(interaction.root);
  root.append(
    apiTable([
      { name: "find(query, options)", type: "Promise<SearchMatch[] | null>", description: "扫描显示文本，返回匹配集合。" },
      { name: "findNext() / findPrevious()", type: "SearchMatch | null", description: "在匹配集合中前后跳转并选中单元格。" },
      { name: "replace(match, value, query?, options?)", type: "void", description: "替换命中文本片段；未传 query 时保持整格写入的兼容行为。" },
      { name: "replaceAll(query, value, options?)", type: "Promise<number | null>", description: "替换全部命中，返回替换处数。" },
      { name: "sortByColumn(column, direction)", type: "boolean", description: "按列排序，排序与行几何在同一事务中提交。" },
      { name: "filterByColumn(column, predicate)", type: "void", description: "按谓词过滤行；被过滤的行高度收缩为 0。" },
      { name: "copy(options?) / paste(options?)", type: "Promise", description: "程序化复制和粘贴。" }
    ])
  );

  const validation = section("数据验证");
  validation.body.append(
    paragraph(
      "验证规则按区域存储在快照中，支持撤销，并随行列插入、删除和移动更新覆盖范围。用法见「示例 → 数据验证」。"
    )
  );
  root.append(validation.root);
  root.append(
    apiTable([
      { name: "setDataValidation(range | ranges, rule)", type: "string", description: "给一个或多个区域设置规则，返回规则 id。每个单元格至多属于一条规则，新规则覆盖的区域会从旧规则中移除。不检查已有数据是否合规。" },
      { name: "clearDataValidation(range | ranges)", type: "void", description: "清除这些区域上的规则；覆盖区域为空的规则会被删除。可撤销。" },
      { name: "getDataValidations()", type: "readonly DataValidationRule[]", description: "当前全部规则。" },
      { name: "getDataValidationAt(row, column)", type: "DataValidationRule | undefined", description: "返回指定单元格所属的规则。" },
      { name: "validateCell(row, column)", type: "DataValidationResult", description: "检查指定单元格的值是否符合规则；未设置规则时返回有效。" },
      { name: "findInvalidCells(range?)", type: "{ row, column }[]", description: "返回未通过验证的单元格坐标。未指定范围时检查整个坐标空间。" },
      { name: "registerDataValidator(name, fn)", type: "Disposer", description: "注册自定义校验函数：同步、纯函数，返回 true / false / 错误文案。仅作用于当前实例，不写入快照。" }
    ])
  );
  root.append(
    callout(
      "只有编辑器提交、且规则是停止级时才会拒绝输入；粘贴、填充、setCell() 和警告、信息级只标记不拦截。" +
        "构造选项 dataValidation.showInvalidMarks: false 可关闭红色小三角。"
    )
  );

  root.append(callout("结构操作使用下标。插入或删除列后，原有列下标会按新坐标空间解释。"));
}
