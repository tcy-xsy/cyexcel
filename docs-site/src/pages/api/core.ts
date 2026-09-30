import { pageHeader, paragraph, section, apiTable, callout } from "../../components/doc";

export default function coreApiPage(root: HTMLElement): void {
  root.append(pageHeader("API", "CyExcelCore", "无 UI 的核心类。所有写入都经过 Command、Mutation 和 History。"));

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
      { name: "getCellStyle(row, column)", type: "CellStyle | undefined", description: "读取列默认样式和格样式合并后的结果。" },
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
      { name: "sortByColumn(column, direction)", type: "boolean", description: "按列排序，排序与行几何在同一事务里提交。" },
      { name: "filterByColumn(column, predicate)", type: "void", description: "按谓词过滤行；被过滤的行高度收缩为 0。" },
      { name: "copy(options?) / paste(options?)", type: "Promise", description: "程序化复制和粘贴。" }
    ])
  );

  const validation = section("数据验证");
  validation.body.append(
    paragraph(
      "规则是区域数据，随快照走，可撤销，也随行列的插入、删除、移动一起搬。用法与示例见「示例 → 数据验证」。"
    )
  );
  root.append(validation.root);
  root.append(
    apiTable([
      { name: "setDataValidation(range | ranges, rule)", type: "string", description: "给一个或多个区域设置规则，返回规则 id。一格至多属于一条规则，重叠部分从旧规则里挖掉。不检查已有数据是否合规。" },
      { name: "clearDataValidation(range | ranges)", type: "void", description: "清除这些区域上的规则；被挖空的规则整条消失。可撤销。" },
      { name: "getDataValidations()", type: "readonly DataValidationRule[]", description: "当前全部规则。" },
      { name: "getDataValidationAt(row, column)", type: "DataValidationRule | undefined", description: "这一格所在的规则。" },
      { name: "validateCell(row, column)", type: "DataValidationResult", description: "这一格当前的值是否符合规则；没有规则恒为有效。" },
      { name: "findInvalidCells(range?)", type: "{ row, column }[]", description: "列出违规的格子（Excel 的“圈释无效数据”），缺省扫整张坐标空间。" },
      { name: "registerDataValidator(name, fn)", type: "Disposer", description: "注册自定义校验函数：同步、纯函数，返回 true / false / 错误文案。实例级，不进快照。" }
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
