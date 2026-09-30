import { el } from "./dom";
import * as styles from "./styles";
import type { ToolbarContext } from "./types";
import type { SearchMatch } from "../search";
import {
  ALLOW_OPTIONS,
  ERROR_STYLE_OPTIONS,
  OPERATOR_OPTIONS,
  fieldsFor,
  formFromRule,
  ruleInputFromForm,
  type AllowKind,
  type DataValidationFormState
} from "./dataValidationForm";
import type { ValidationErrorStyle, ValidationOperator } from "../sheet/validations";

/**
 * 对话框里的内容。契约与 `panels.ts` 一样——把内容塞进调用方给的容器，返回 disposer
 * ——只是容器换成了对话框的 body，而且多拿一个 `close`。
 *
 * 这一层在 node 里测不了（没有 DOM，项目不引 jsdom——见 step25-toolbar-plan.md §5），
 * 由 playground 手验。
 */

function field(labelText: string, control: HTMLElement): HTMLElement {
  const row = el("div", "cyx-toolbar-dialog-field");
  styles.apply(row, styles.dialogField);
  const label = el("label", "cyx-toolbar-dialog-label", labelText);
  styles.apply(label, styles.dialogLabel);
  row.append(label, control);
  return row;
}

function textInput(placeholder: string, ariaLabel: string): HTMLInputElement {
  const input = el("input", "cyx-toolbar-dialog-input");
  styles.apply(input, styles.dialogInput);
  input.type = "text";
  input.placeholder = placeholder;
  input.setAttribute("aria-label", ariaLabel);
  return input;
}

function checkbox(labelText: string): { readonly wrap: HTMLElement; readonly input: HTMLInputElement } {
  const wrap = el("label", "cyx-toolbar-dialog-checkbox");
  styles.apply(wrap, styles.dialogCheckbox);
  const input = el("input");
  input.type = "checkbox";
  wrap.append(input, document.createTextNode(labelText));
  return { wrap, input };
}

function button(text: string): HTMLButtonElement {
  const node = el("button", "cyx-toolbar-dialog-button", text);
  node.type = "button";
  styles.apply(node, styles.dialogButton);
  return node;
}

/**
 * 「查找和替换」。
 *
 * 与原先那个挤在小面板里的版本相比，多了区分大小写/正则两个选项、以及「替换」单个
 * 这个动作——它们在 160px 宽的面板里排不下，这正是这件事要做成对话框的理由。
 *
 * **替换之后会重新扫一遍。** `core` 的 `replace()` 只写一格，写完之后先前 `find()`
 * 缓存的那批匹配就过时了（位置还在，但那一格的内容已经变了，后续的"下一个"会走到
 * 错的地方）——`CyExcelCore.replaceAll` 的注释里写的就是这件事。所以这里替换完不
 * 复用旧匹配集，而是拿当前条件重扫，再跳到下一个。
 */
export function findReplaceDialog(
  body: HTMLElement,
  close: () => void,
  ctx: ToolbarContext
): () => void {
  const query = textInput("查找内容…", "查找内容");
  const replacement = textInput("替换为…", "替换为");

  const matchCase = checkbox("区分大小写");
  const useRegex = checkbox("正则表达式");
  const optionRow = el("div", "cyx-toolbar-dialog-options");
  styles.apply(optionRow, styles.dialogOptions);
  optionRow.append(matchCase.wrap, useRegex.wrap);

  const status = el("div", "cyx-toolbar-dialog-status", "");
  styles.apply(status, styles.dialogStatus);

  const findAll = button("查找全部");
  const previous = button("上一个");
  const next = button("下一个");
  const replaceOne = button("替换");
  const replaceAll = button("全部替换");
  const done = button("关闭");

  const buttonRow = el("div", "cyx-toolbar-dialog-buttons");
  styles.apply(buttonRow, styles.dialogButtons);
  buttonRow.append(findAll, previous, next, replaceOne, replaceAll, done);

  /** 上一次 `findNext`/`findPrevious` 停在哪——「替换」写的就是这一格。 */
  let currentMatch: SearchMatch | null = null;
  /** 防止两次扫描叠在一起（大表上 find 是异步分批的）。 */
  let scanning = false;

  const searchOptions = (): { matchCase: boolean; useRegex: boolean } => ({
    matchCase: matchCase.input.checked,
    useRegex: useRegex.input.checked
  });

  const describe = (match: SearchMatch | null): string =>
    match === null ? "没有更多匹配" : `第 ${match.row + 1} 行，第 ${match.column + 1} 列`;

  const scan = async (): Promise<number | null> => {
    if (query.value === "") {
      status.textContent = "请先输入查找内容";
      return null;
    }
    scanning = true;
    try {
      const matches = await ctx.sheet.find(query.value, searchOptions());
      currentMatch = null;
      return matches === null ? null : matches.length;
    } finally {
      scanning = false;
    }
  };

  findAll.addEventListener("click", () => {
    if (scanning) {
      return;
    }
    void scan().then((count) => {
      if (count === null) {
        return;
      }
      status.textContent = count === 0 ? "没有匹配" : `${count} 个匹配`;
    });
  });

  next.addEventListener("click", () => {
    currentMatch = ctx.sheet.findNext();
    status.textContent = describe(currentMatch);
  });

  previous.addEventListener("click", () => {
    currentMatch = ctx.sheet.findPrevious();
    status.textContent = describe(currentMatch);
  });

  replaceOne.addEventListener("click", () => {
    if (scanning) {
      return;
    }
    const target = currentMatch;
    if (target === null) {
      status.textContent = "先用「下一个」定位到一处，再替换";
      return;
    }
    ctx.sheet.replace(target, replacement.value, query.value, searchOptions());
    // 重扫而不是复用旧匹配集——见函数头的说明。
    void scan().then((count) => {
      if (count === null || count === 0) {
        status.textContent = "已替换，没有更多匹配";
        return;
      }
      currentMatch = ctx.sheet.findNext();
      status.textContent = `已替换，${describe(currentMatch)}`;
    });
  });

  replaceAll.addEventListener("click", () => {
    if (scanning || query.value === "") {
      if (query.value === "") {
        status.textContent = "请先输入查找内容";
      }
      return;
    }
    scanning = true;
    void ctx.sheet
      .replaceAll(query.value, replacement.value, searchOptions())
      .then((count) => {
        currentMatch = null;
        status.textContent = count === null ? "已取消" : `替换了 ${count} 处`;
      })
      .finally(() => {
        scanning = false;
      });
  });

  done.addEventListener("click", close);

  // 回车 = 查找下一个，与 Excel 一致。
  const onEnter = (event: KeyboardEvent): void => {
    if (event.key === "Enter") {
      event.preventDefault();
      next.click();
    }
  };
  query.addEventListener("keydown", onEnter);
  replacement.addEventListener("keydown", onEnter);

  // 改了查找条件，旧的定位就不作数了——否则「替换」会写到一个按旧条件找到的格子。
  const invalidate = (): void => {
    currentMatch = null;
  };
  query.addEventListener("input", invalidate);
  matchCase.input.addEventListener("change", invalidate);
  useRegex.input.addEventListener("change", invalidate);

  body.append(field("查找", query), field("替换为", replacement), optionRow, buttonRow, status);

  return () => {
    // 关掉对话框就撤掉高亮：那片黄底留在表格上而没有任何界面能再操作它，会很费解。
    ctx.sheet.clearSearch();
    body.replaceChildren();
  };
}

function selectInput<T extends string>(
  options: readonly { readonly value: T; readonly label: string }[],
  ariaLabel: string
): HTMLSelectElement {
  const select = el("select", "cyx-toolbar-dialog-input");
  styles.apply(select, styles.dialogInput);
  select.setAttribute("aria-label", ariaLabel);
  for (const option of options) {
    const node = el("option", undefined, option.label);
    node.value = option.value;
    select.append(node);
  }
  return select;
}

/**
 * 「数据验证」。对选区（活动区域）设置或清除规则；打开时按选区左上角那一格的规则回填。
 *
 * 规则怎么从表单来、怎么回填表单都在 `dataValidationForm.ts`（纯函数，有单测）；这里
 * 只是把它们接到 DOM 上，node 里测不了，由 playground 手验。
 *
 * 与 Excel 一致：「任何值」+ 确定 = 清除这片区域的规则；`allow` 是「任何值」而区域上原本
 * 是自定义规则时，先提示一句"保存会替换它"，因为表单没法表达自定义规则。
 */
export function dataValidationDialog(
  body: HTMLElement,
  close: () => void,
  ctx: ToolbarContext
): () => void {
  const range = ctx.range;
  const existing = ctx.sheet.getDataValidationAt(range.startRow, range.startColumn);
  const state: DataValidationFormState = formFromRule(existing);

  const allow = selectInput<AllowKind>(ALLOW_OPTIONS, "允许");
  const operator = selectInput<ValidationOperator>(OPERATOR_OPTIONS, "数据");
  const value1 = textInput("", "值");
  const value2 = textInput("", "最大值");
  const allowBlank = checkbox("忽略空值");
  const showDropDown = checkbox("提供下拉箭头");
  const promptTitle = textInput("可选", "输入信息标题");
  const prompt = textInput("选中单元格时显示", "输入信息");
  const errorStyle = selectInput<ValidationErrorStyle>(ERROR_STYLE_OPTIONS, "错误样式");
  const errorTitle = textInput("可选", "出错警告标题");
  const errorMessage = textInput("留空使用默认文案", "出错警告信息");

  const operatorRow = field("数据", operator);
  const value1Label = el("label", "cyx-toolbar-dialog-label", "值");
  styles.apply(value1Label, styles.dialogLabel);
  const value1Row = el("div", "cyx-toolbar-dialog-field");
  styles.apply(value1Row, styles.dialogField);
  value1Row.append(value1Label, value1);
  const value2Row = field("最大值", value2);
  const optionRow = el("div", "cyx-toolbar-dialog-options");
  styles.apply(optionRow, styles.dialogOptions);
  optionRow.append(allowBlank.wrap, showDropDown.wrap);

  const status = el("div", "cyx-toolbar-dialog-status", "");
  styles.apply(status, styles.dialogStatus);

  const ok = button("确定");
  const cancel = button("取消");
  const clearAll = button("全部清除");
  const buttonRow = el("div", "cyx-toolbar-dialog-buttons");
  styles.apply(buttonRow, styles.dialogButtons);
  buttonRow.append(ok, cancel, clearAll);

  const read = (): DataValidationFormState => ({
    allow: allow.value as AllowKind,
    operator: operator.value as ValidationOperator,
    value1: value1.value,
    value2: value2.value,
    allowBlank: allowBlank.input.checked,
    showDropDown: showDropDown.input.checked,
    errorStyle: errorStyle.value as ValidationErrorStyle,
    errorTitle: errorTitle.value,
    errorMessage: errorMessage.value,
    promptTitle: promptTitle.value,
    prompt: prompt.value
  });

  /** 按当前的"允许"和"数据"显示或隐藏对应的输入框。 */
  const refresh = (): void => {
    const shown = fieldsFor(allow.value as AllowKind, operator.value as ValidationOperator);
    operatorRow.style.display = shown.operator ? "" : "none";
    value1Row.style.display = shown.value1 ? "" : "none";
    value2Row.style.display = shown.value2 ? "" : "none";
    showDropDown.wrap.style.display = shown.dropDown ? "" : "none";
    value1Label.textContent = allow.value === "list" ? "来源" : shown.value2 ? "最小值" : "值";
    value1.placeholder = allow.value === "list" ? "用逗号分隔，如 红,绿,蓝" : allow.value === "date" ? "2026-01-15" : allow.value === "time" ? "09:30" : "";
    optionRow.style.display = allow.value === "any" ? "none" : "";
  };

  allow.value = state.allow;
  operator.value = state.operator;
  value1.value = state.value1;
  value2.value = state.value2;
  allowBlank.input.checked = state.allowBlank;
  showDropDown.input.checked = state.showDropDown;
  errorStyle.value = state.errorStyle;
  errorTitle.value = state.errorTitle;
  errorMessage.value = state.errorMessage;
  promptTitle.value = state.promptTitle;
  prompt.value = state.prompt;
  refresh();
  allow.addEventListener("change", refresh);
  operator.addEventListener("change", refresh);

  if (existing?.criteria.type === "custom") {
    status.textContent = "这片区域用的是自定义规则，表单无法编辑；保存会替换它。";
  }

  ok.addEventListener("click", () => {
    const outcome = ruleInputFromForm(read());
    if (outcome.kind === "error") {
      status.textContent = outcome.message;
      return;
    }
    try {
      if (outcome.kind === "clear") {
        ctx.sheet.clearDataValidation(range);
      } else {
        ctx.sheet.setDataValidation(range, outcome.input);
      }
      close();
    } catch (error) {
      status.textContent = error instanceof Error ? error.message : String(error);
    }
  });
  cancel.addEventListener("click", close);
  clearAll.addEventListener("click", () => {
    ctx.sheet.clearDataValidation(range);
    close();
  });

  body.append(
    field("允许", allow),
    operatorRow,
    value1Row,
    value2Row,
    optionRow,
    field("输入标题", promptTitle),
    field("输入信息", prompt),
    field("错误样式", errorStyle),
    field("错误标题", errorTitle),
    field("错误信息", errorMessage),
    buttonRow,
    status
  );

  return () => {
    body.replaceChildren();
  };
}
