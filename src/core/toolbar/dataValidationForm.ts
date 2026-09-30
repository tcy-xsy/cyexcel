import { serialFromISODate, serialToISODate } from "../style/dateSerial";
import type {
  DataValidationRule,
  DataValidationRuleInput,
  ValidationErrorStyle,
  ValidationOperator
} from "../sheet/validations";

/**
 * 「数据验证」对话框的表单状态与它同规则之间的互转。纯函数，不碰 DOM——对话框本身
 * 在 node 里测不了（没有 DOM），把"表单怎么变成规则、规则怎么回填表单"抽到这里，
 * 才有地方放单测。
 */

export type AllowKind = "any" | "whole" | "decimal" | "textLength" | "date" | "time" | "list";

export interface DataValidationFormState {
  allow: AllowKind;
  operator: ValidationOperator;
  /** 数值/日期/时间类的"值1"；`list` 时是逗号分隔的序列。 */
  value1: string;
  value2: string;
  allowBlank: boolean;
  showDropDown: boolean;
  errorStyle: ValidationErrorStyle;
  errorTitle: string;
  errorMessage: string;
  promptTitle: string;
  prompt: string;
}

export const ALLOW_OPTIONS: readonly { readonly value: AllowKind; readonly label: string }[] = [
  { value: "any", label: "任何值" },
  { value: "whole", label: "整数" },
  { value: "decimal", label: "小数" },
  { value: "textLength", label: "文本长度" },
  { value: "date", label: "日期" },
  { value: "time", label: "时间" },
  { value: "list", label: "序列" }
];

export const OPERATOR_OPTIONS: readonly { readonly value: ValidationOperator; readonly label: string }[] = [
  { value: "between", label: "介于" },
  { value: "notBetween", label: "未介于" },
  { value: "equal", label: "等于" },
  { value: "notEqual", label: "不等于" },
  { value: "greaterThan", label: "大于" },
  { value: "greaterThanOrEqual", label: "大于或等于" },
  { value: "lessThan", label: "小于" },
  { value: "lessThanOrEqual", label: "小于或等于" }
];

export const ERROR_STYLE_OPTIONS: readonly { readonly value: ValidationErrorStyle; readonly label: string }[] = [
  { value: "stop", label: "停止" },
  { value: "warning", label: "警告" },
  { value: "information", label: "信息" }
];

export const DEFAULT_FORM: DataValidationFormState = Object.freeze({
  allow: "any",
  operator: "between",
  value1: "",
  value2: "",
  allowBlank: true,
  showDropDown: true,
  errorStyle: "stop",
  errorTitle: "",
  errorMessage: "",
  promptTitle: "",
  prompt: ""
});

/** 这个 `allow` 需要哪几个输入框。 */
export function fieldsFor(allow: AllowKind, operator: ValidationOperator): {
  readonly operator: boolean;
  readonly value1: boolean;
  readonly value2: boolean;
  readonly dropDown: boolean;
} {
  if (allow === "any") {
    return { operator: false, value1: false, value2: false, dropDown: false };
  }
  if (allow === "list") {
    return { operator: false, value1: true, value2: false, dropDown: true };
  }
  return {
    operator: true,
    value1: true,
    value2: operator === "between" || operator === "notBetween",
    dropDown: false
  };
}

export const LIST_SEPARATOR = /[,，]/;

export function parseListItems(text: string): string[] {
  const items = text
    .split(LIST_SEPARATOR)
    .map((item) => item.trim())
    .filter((item) => item !== "");
  return [...new Set(items)];
}

/** `"09:30"` / `"9:30:15"` → 一天中的小数部分；不是时间返回 `undefined`。 */
export function parseClock(text: string): number | undefined {
  const match = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(text.trim());
  if (match === null) {
    return undefined;
  }
  const [hours, minutes, seconds] = [Number(match[1]), Number(match[2]), Number(match[3] ?? 0)];
  if (hours > 23 || minutes > 59 || seconds > 59) {
    return undefined;
  }
  return (hours * 3600 + minutes * 60 + seconds) / 86400;
}

export function formatClock(fraction: number): string {
  const total = Math.round(fraction * 86400);
  const hours = Math.floor(total / 3600) % 24;
  const minutes = Math.floor((total % 3600) / 60);
  const pad = (n: number): string => String(n).padStart(2, "0");
  return `${pad(hours)}:${pad(minutes)}`;
}

export type FormOutcome =
  | { readonly kind: "clear" }
  | { readonly kind: "rule"; readonly input: DataValidationRuleInput }
  | { readonly kind: "error"; readonly message: string };

function parseNumber(text: string, label: string): number | string {
  const trimmed = text.trim();
  if (trimmed === "") {
    return `请填写${label}`;
  }
  const value = Number(trimmed);
  return Number.isFinite(value) ? value : `${label}需要是数字`;
}

/**
 * 表单 → 规则入参。`allow: "any"` 表示"清除这个区域的规则"；其余先逐字段校验，第一处
 * 问题以 `error` 返回，文案直接给用户看。日期/时间在这里就解析成序列号：解析失败在用户
 * 还能改的地方提示，不等到 `setDataValidation` 抛错。
 */
export function ruleInputFromForm(form: DataValidationFormState): FormOutcome {
  if (form.allow === "any") {
    return { kind: "clear" };
  }

  const shared = {
    allowBlank: form.allowBlank,
    errorStyle: form.errorStyle,
    ...(form.errorTitle !== "" ? { errorTitle: form.errorTitle } : {}),
    ...(form.errorMessage !== "" ? { errorMessage: form.errorMessage } : {}),
    ...(form.promptTitle !== "" ? { promptTitle: form.promptTitle } : {}),
    ...(form.prompt !== "" ? { prompt: form.prompt } : {})
  };

  if (form.allow === "list") {
    const items = parseListItems(form.value1);
    if (items.length === 0) {
      return { kind: "error", message: "请填写序列内容，用逗号分隔" };
    }
    return {
      kind: "rule",
      input: { ...shared, showDropDown: form.showDropDown, criteria: { type: "list", items } }
    };
  }

  const needsSecond = form.operator === "between" || form.operator === "notBetween";
  const read = (text: string, label: string): number | string => {
    if (form.allow === "date") {
      const serial = text.trim() === "" ? undefined : serialFromISODate(text.trim());
      return serial === undefined || serial === null ? `${label}需要是日期，例如 2026-01-15` : serial;
    }
    if (form.allow === "time") {
      const fraction = parseClock(text);
      return fraction === undefined ? `${label}需要是时间，例如 09:30` : fraction;
    }
    const value = parseNumber(text, label);
    if (typeof value === "number" && (form.allow === "whole" || form.allow === "textLength") && !Number.isInteger(value)) {
      return `${label}需要是整数`;
    }
    if (typeof value === "number" && form.allow === "textLength" && value < 0) {
      return `${label}不能是负数`;
    }
    return value;
  };

  const value1 = read(form.value1, needsSecond ? "最小值" : "值");
  if (typeof value1 === "string") {
    return { kind: "error", message: value1 };
  }
  let value2: number | undefined;
  if (needsSecond) {
    const parsed = read(form.value2, "最大值");
    if (typeof parsed === "string") {
      return { kind: "error", message: parsed };
    }
    value2 = parsed;
  }

  return {
    kind: "rule",
    input: {
      ...shared,
      criteria: {
        type: form.allow,
        operator: form.operator,
        value1,
        ...(value2 !== undefined ? { value2 } : {})
      }
    }
  };
}

/**
 * 规则 → 表单，用来打开对话框时回填。`custom` 规则没法在表单里编辑：回填成"任何值"，
 * 由调用方提示"保存会替换它"。
 */
export function formFromRule(rule: DataValidationRule | undefined): DataValidationFormState {
  if (rule === undefined) {
    return { ...DEFAULT_FORM };
  }
  const shared = {
    allowBlank: rule.allowBlank ?? true,
    showDropDown: rule.showDropDown ?? true,
    errorStyle: rule.errorStyle ?? "stop",
    errorTitle: rule.errorTitle ?? "",
    errorMessage: rule.errorMessage ?? "",
    promptTitle: rule.promptTitle ?? "",
    prompt: rule.prompt ?? ""
  } as const;
  const criteria = rule.criteria;
  switch (criteria.type) {
    case "list":
      return { ...DEFAULT_FORM, ...shared, allow: "list", value1: criteria.items.join(",") };
    case "custom":
      return { ...DEFAULT_FORM, ...shared, allow: "any" };
    case "date": {
      const show = (serial: number): string => serialToISODate(serial) ?? String(serial);
      return {
        ...DEFAULT_FORM,
        ...shared,
        allow: "date",
        operator: criteria.operator,
        value1: show(criteria.value1),
        value2: criteria.value2 === undefined ? "" : show(criteria.value2)
      };
    }
    case "time":
      return {
        ...DEFAULT_FORM,
        ...shared,
        allow: "time",
        operator: criteria.operator,
        value1: formatClock(criteria.value1),
        value2: criteria.value2 === undefined ? "" : formatClock(criteria.value2)
      };
    default:
      return {
        ...DEFAULT_FORM,
        ...shared,
        allow: criteria.type,
        operator: criteria.operator,
        value1: String(criteria.value1),
        value2: criteria.value2 === undefined ? "" : String(criteria.value2)
      };
  }
}
