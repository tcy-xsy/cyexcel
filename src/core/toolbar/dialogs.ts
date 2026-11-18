import { el } from "./dom.js";
import * as styles from "./styles.js";
import type { ToolbarContext } from "./types.js";
import type { SearchMatch } from "../search/index.js";
import {
  createValidationFormOptions,
  fieldsFor,
  formFromRule,
  ruleInputFromForm,
  type AllowKind,
  type DataValidationFormState
} from "./dataValidationForm.js";
import type { DataValidationRenderMode, ValidationErrorStyle, ValidationOperator } from "../sheet/validations.js";

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
  const t = ctx.sheet.t;
  const query = textInput(t("Find..."), t("Find what"));
  const replacement = textInput(t("Replace with..."), t("Replace with"));

  const matchCase = checkbox(t("Match case"));
  const useRegex = checkbox(t("Regular expression"));
  const optionRow = el("div", "cyx-toolbar-dialog-options");
  styles.apply(optionRow, styles.dialogOptions);
  optionRow.append(matchCase.wrap, useRegex.wrap);

  const status = el("div", "cyx-toolbar-dialog-status", "");
  styles.apply(status, styles.dialogStatus);

  const findAll = button(t("Find all"));
  const previous = button(t("Previous"));
  const next = button(t("Next"));
  const replaceOne = button(t("Replace"));
  const replaceAll = button(t("Replace all"));
  const done = button(t("Close"));

  const buttonRow = el("div", "cyx-toolbar-dialog-buttons");
  styles.apply(buttonRow, styles.searchDialogButtons);
  buttonRow.append(findAll, previous, next, replaceOne, replaceAll, done);
  for (const control of [findAll, previous, next, replaceOne, replaceAll, done]) {
    styles.apply(control, { minWidth: "0", minHeight: "28px", height: "auto", padding: "4px 6px", whiteSpace: "normal" });
  }

  /** 上一次 `findNext`/`findPrevious` 停在哪——「替换」写的就是这一格。 */
  let currentMatch: SearchMatch | null = null;
  /** 防止两次扫描叠在一起（大表上 find 是异步分批的）。 */
  let scanning = false;

  const searchOptions = (): { matchCase: boolean; useRegex: boolean } => ({
    matchCase: matchCase.input.checked,
    useRegex: useRegex.input.checked
  });

  const describe = (match: SearchMatch | null): string =>
    match === null ? t("No more matches") : t("Row {0}, column {1}", match.row + 1, match.column + 1);

  const scan = async (): Promise<number | null> => {
    if (query.value === "") {
      status.textContent = t("Enter a search term first.");
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
      status.textContent = count === 0 ? t("No matches") : t("{0} matches", count);
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
      status.textContent = t("Use Next to locate a match before replacing it.");
      return;
    }
    ctx.sheet.replace(target, replacement.value, query.value, searchOptions());
    // 重扫而不是复用旧匹配集——见函数头的说明。
    void scan().then((count) => {
      if (count === null || count === 0) {
        status.textContent = t("Replaced. No more matches.");
        return;
      }
      currentMatch = ctx.sheet.findNext();
      status.textContent = t("Replaced. {0}", describe(currentMatch));
    });
  });

  replaceAll.addEventListener("click", () => {
    if (scanning || query.value === "") {
      if (query.value === "") {
        status.textContent = t("Enter a search term first.");
      }
      return;
    }
    scanning = true;
    void ctx.sheet
      .replaceAll(query.value, replacement.value, searchOptions())
      .then((count) => {
        currentMatch = null;
        status.textContent = count === null ? t("Cancelled") : t("Replaced {0} occurrences", count);
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

  body.append(field(t("Find"), query), field(t("Replace with"), replacement), optionRow, buttonRow, status);

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
  const t = ctx.sheet.t;
  const range = ctx.range;
  const { ALLOW_OPTIONS, ERROR_STYLE_OPTIONS, RENDER_MODE_OPTIONS, OPERATOR_OPTIONS } = createValidationFormOptions(t);
  const existing = ctx.sheet.getDataValidationAt(range.startRow, range.startColumn);
  const state: DataValidationFormState = formFromRule(existing);

  const allow = selectInput<AllowKind>(ALLOW_OPTIONS, t("Allow"));
  const operator = selectInput<ValidationOperator>(OPERATOR_OPTIONS, t("Data"));
  const value1 = textInput("", t("Value"));
  const value2 = textInput("", t("Maximum"));
  const allowBlank = checkbox(t("Ignore blanks"));
  const showDropDown = checkbox(t("Show dropdown arrow"));
  const renderMode = selectInput<DataValidationRenderMode>(RENDER_MODE_OPTIONS, t("Display mode"));
  const errorStyle = selectInput<ValidationErrorStyle>(ERROR_STYLE_OPTIONS, t("Error style"));
  const errorTitle = textInput(t("Optional"), t("Error alert title"));
  const errorMessage = textInput(t("Leave blank to use the default message"), t("Error alert message"));

  const operatorRow = field(t("Data"), operator);
  const value1Label = el("label", "cyx-toolbar-dialog-label", t("Value"));
  styles.apply(value1Label, styles.dialogLabel);
  const value1Row = el("div", "cyx-toolbar-dialog-field");
  styles.apply(value1Row, styles.dialogField);
  value1Row.append(value1Label, value1);
  const value2Row = field(t("Maximum"), value2);
  const sourceHelp = el("div", "cyx-toolbar-dialog-help", t("Comma-separated, with optional colors: Red=#d93025,Green,Blue"));
  styles.apply(sourceHelp, styles.dialogHelp);
  const renderModeRow = field(t("Display mode"), renderMode);
  const optionRow = el("div", "cyx-toolbar-dialog-options");
  styles.apply(optionRow, styles.dialogOptions);
  optionRow.append(allowBlank.wrap, showDropDown.wrap);

  const status = el("div", "cyx-toolbar-dialog-status", "");
  styles.apply(status, styles.dialogStatus);

  const ok = button(t("OK"));
  const cancel = button(t("Cancel"));
  const clearAll = button(t("Clear all"));
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
    renderMode: renderMode.value as DataValidationRenderMode,
    errorStyle: errorStyle.value as ValidationErrorStyle,
    errorTitle: errorTitle.value,
    errorMessage: errorMessage.value
  });

  /** 按当前的"允许"和"数据"显示或隐藏对应的输入框。 */
  const refresh = (): void => {
    const shown = fieldsFor(allow.value as AllowKind, operator.value as ValidationOperator);
    styles.setHidden(operatorRow, !shown.operator);
    styles.setHidden(value1Row, !shown.value1);
    styles.setHidden(value2Row, !shown.value2);
    styles.setHidden(showDropDown.wrap, !shown.dropDown);
    styles.setHidden(renderModeRow, !shown.renderMode);
    styles.setHidden(sourceHelp, allow.value !== "list");
    value1Label.textContent = allow.value === "list" ? t("Source") : shown.value2 ? t("Minimum") : t("Value");
    value1.placeholder = allow.value === "date" ? "2026-01-15" : allow.value === "time" ? "09:30" : "";
    styles.setHidden(optionRow, allow.value === "any");
  };

  allow.value = state.allow;
  operator.value = state.operator;
  value1.value = state.value1;
  value2.value = state.value2;
  allowBlank.input.checked = state.allowBlank;
  showDropDown.input.checked = state.showDropDown;
  renderMode.value = state.renderMode;
  errorStyle.value = state.errorStyle;
  errorTitle.value = state.errorTitle;
  errorMessage.value = state.errorMessage;
  refresh();
  allow.addEventListener("change", refresh);
  operator.addEventListener("change", refresh);

  if (existing?.criteria.type === "custom") {
    status.textContent = t("This range uses a custom rule that cannot be edited here. Saving will replace it.");
  }

  ok.addEventListener("click", () => {
    const outcome = ruleInputFromForm(read(), t);
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
    field(t("Allow"), allow),
    operatorRow,
    value1Row,
    sourceHelp,
    value2Row,
    optionRow,
    renderModeRow,
    field(t("Error style"), errorStyle),
    field(t("Error title"), errorTitle),
    field(t("Error message"), errorMessage),
    buttonRow,
    status
  );

  return () => {
    body.replaceChildren();
  };
}
