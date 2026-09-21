import type { CyExcelCore } from "../CyExcelCore";
import { parseA1, rangeToA1 } from "./a1Notation";
import * as styles from "./styles";

export interface FormulaBarHandle {
  refresh(): void;
  dispose(): void;
}

/**
 * 编辑栏：左边名称框（当前选区的 A1 地址，可输入跳转），右边内容框（当前格的值，
 * 可编辑提交）。
 *
 * 与工具栏同属 `cyexcel/toolbar` 这个 opt-in 模块——它们都是"围着表格的宿主 UI"，
 * 不 import 就都不进产物。
 *
 * # 显示的是原始值，不是格式化后的文本
 *
 * 与 Excel 一致：单元格里显示 `1,234.00`，编辑栏里是 `1234`。Step 24 的
 * `formatCellText` 只作用于绘制路径，这里直接读 `getCell()`。将来有了公式，
 * 这里显示的会是公式串本身，同一条规则。
 */
/** 作为 `CyExcelCoreOptions.topWidgets` 的一项使用。 */
export function formulaBarWidget() {
  return (sheet: CyExcelCore, host: HTMLElement) => createFormulaBar(sheet, host);
}

export function createFormulaBar(sheet: CyExcelCore, container: HTMLElement): FormulaBarHandle {
  const root = document.createElement("div");
  root.className = "cyx-formula-bar";
  styles.apply(root, styles.formulaBar);

  const nameBox = document.createElement("input");
  nameBox.className = "cyx-formula-bar-name";
  nameBox.type = "text";
  nameBox.setAttribute("aria-label", "单元格地址");
  styles.apply(nameBox, styles.formulaName);

  const label = document.createElement("span");
  label.className = "cyx-formula-bar-fx";
  label.textContent = "fx";
  styles.apply(label, styles.formulaFx);

  const input = document.createElement("input");
  input.className = "cyx-formula-bar-input";
  input.type = "text";
  input.setAttribute("aria-label", "单元格内容");
  styles.apply(input, styles.formulaInput);

  root.append(nameBox, label, input);
  container.append(root);

  /** 当前编辑的那一格。提交时用它，而不是提交那一刻的选区——见 commit()。 */
  let editing: { row: number; column: number } | null = null;

  function refresh(): void {
    const { range } = sheet.selection;
    nameBox.value = rangeToA1(range);
    const value = sheet.getCell(range.startRow, range.startColumn);
    input.value = value === null || value === undefined ? "" : String(value);
    editing = { row: range.startRow, column: range.startColumn };
  }

  function commit(): void {
    if (editing === null) {
      return;
    }
    // 用进入编辑时记下的坐标，不是此刻的选区：提交发生在 blur 或回车时，而 blur
    // 可能正是因为用户点了另一格——那时选区已经走了，写过去就是写错格子。
    sheet.setCell(editing.row, editing.column, input.value);
  }

  const onInputKeyDown = (event: KeyboardEvent): void => {
    if (event.key === "Enter") {
      event.preventDefault();
      commit();
      input.blur();
      return;
    }
    if (event.key === "Escape") {
      event.preventDefault();
      refresh(); // 丢弃改动，回到这一格现在的值
      input.blur();
    }
  };

  const onNameKeyDown = (event: KeyboardEvent): void => {
    if (event.key !== "Enter") {
      return;
    }
    event.preventDefault();
    const target = parseA1(nameBox.value);
    if (target === null) {
      refresh(); // 认不出来就把框里的文字恢复成当前地址，不报错
      return;
    }
    sheet.selectCell(target.row, target.column);
    nameBox.blur();
  };

  input.addEventListener("keydown", onInputKeyDown);
  input.addEventListener("blur", commit);
  nameBox.addEventListener("keydown", onNameKeyDown);
  nameBox.addEventListener("blur", refresh);

  const disposers = [
    sheet.events.on("selection-change", refresh),
    sheet.events.on("change", refresh)
  ];

  refresh();

  return {
    refresh,
    dispose: () => {
      input.removeEventListener("keydown", onInputKeyDown);
      input.removeEventListener("blur", commit);
      nameBox.removeEventListener("keydown", onNameKeyDown);
      nameBox.removeEventListener("blur", refresh);
      for (const dispose of disposers) {
        dispose();
      }
      root.remove();
    }
  };
}
