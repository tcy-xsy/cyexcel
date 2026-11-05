import { CyExcelError, CyExcelErrorCode } from "../errors";
import { remapIndex } from "../layout/structuralChange";
import { serialFromISODate } from "../style/dateSerial";
import { canIndexValidationRules, findIndexedValidationRule } from "./ValidationRuleIndex";
import type { CellRange } from "../types";
import {
  axisInterval,
  rangesIntersect,
  relocateIntervalForDelete,
  relocateIntervalForInsert,
  withAxisInterval,
  type AxisInterval
} from "./merges";
import type { CellValue } from "./types";

/**
 * 数据验证规则：数据形状、区域几何、判定逻辑（设计稿 `doc/data-validation-design.md`）。
 *
 * 这个文件跟 `merges.ts` 同一个定位——只做几何/数组运算与纯判定，不碰 `SheetModel`、
 * `TableStore` 或任何 Command 层类型。`readonly DataValidationRule[]` 进，
 * `readonly DataValidationRule[]` 出。规则是**区域规则**（ADR 0004 §3.5 / §9 不变量 3），
 * 不是按 `(row, column)` 索引的侧表，所以它不进 `IndexedTableRegistry`，行列结构操作
 * 的搬移靠这里的 `relocateValidationsFor*`。
 */

export type ValidationOperator =
  | "between"
  | "notBetween"
  | "equal"
  | "notEqual"
  | "greaterThan"
  | "greaterThanOrEqual"
  | "lessThan"
  | "lessThanOrEqual";

export type ValidationErrorStyle = "stop" | "warning" | "information";

/**
 * `list` 规则在单元格里怎么显示（同 Univer 的 `DataValidationRenderMode`）：
 * - `arrow`（默认）：普通文字加右侧常驻的下拉箭头；
 * - `text`：只是文字，没有箭头；
 * - `custom`：带底色的圆角胶囊标签，箭头在胶囊里。
 */
export type DataValidationRenderMode = "text" | "arrow" | "custom";

const RENDER_MODES: ReadonlySet<string> = new Set<DataValidationRenderMode>(["text", "arrow", "custom"]);

/** 选项颜色：`#rgb` / `#rrggbb` / `#rrggbbaa`。只收十六进制，不放任意 CSS 字符串进 canvas。 */
export const COLOR_PATTERN = /^#(?:[0-9a-fA-F]{3}|[0-9a-fA-F]{6}|[0-9a-fA-F]{8})$/;

/**
 * 判据。内置类型 + 一个 `custom`：`custom` 只存**名字**和可 JSON 序列化的 `params`，
 * 判定函数由调用方按名字注册（`CyExcelCore.registerDataValidator`），不进快照。
 * 这样规则能序列化、能撤销、能随行列移动，而逻辑仍然是代码。
 */
export type DataValidationCriteria =
  | {
      readonly type: "whole" | "decimal" | "textLength" | "date" | "time";
      readonly operator: ValidationOperator;
      readonly value1: number;
      readonly value2?: number;
    }
  | {
      readonly type: "list";
      readonly items: readonly string[];
      /** 选项 → 颜色。键必须是 `items` 里的项。`custom` 模式下是胶囊底色，其余模式下是单元格底色。 */
      readonly colors?: Readonly<Record<string, string>>;
    }
  | { readonly type: "custom"; readonly name: string; readonly params?: unknown };

export interface DataValidationRule {
  /** 稳定标识：撤销、UI、xlsx 往返都要用。 */
  readonly id: string;
  /** sqref：一个或多个矩形，规则内、规则之间都不重叠。 */
  readonly ranges: readonly CellRange[];
  readonly criteria: DataValidationCriteria;
  /** 默认 `true`。 */
  readonly allowBlank?: boolean;
  /** 默认 `"stop"`。 */
  readonly errorStyle?: ValidationErrorStyle;
  readonly errorTitle?: string;
  readonly errorMessage?: string;
  /** `list` 用，默认 `true`。 */
  readonly showDropDown?: boolean;
  /** `list` 用，默认 `"arrow"`。 */
  readonly renderMode?: DataValidationRenderMode;
}

/**
 * 调用方递给 `setDataValidation` 的规则：不带 `id`（缺省时生成）和 `ranges`（单独传），
 * 日期/时间的 `value1/2` 还可以写成 `"2026-01-15"` 这样的字符串，创建时规范化成序列号。
 */
export type DataValidationCriteriaInput =
  | {
      readonly type: "whole" | "decimal" | "textLength";
      readonly operator: ValidationOperator;
      readonly value1: number;
      readonly value2?: number;
    }
  | Extract<DataValidationCriteria, { readonly type: "list" | "custom" }>
  | {
      readonly type: "date" | "time";
      readonly operator: ValidationOperator;
      readonly value1: number | string;
      readonly value2?: number | string;
    };

export interface DataValidationRuleInput extends Omit<DataValidationRule, "id" | "ranges" | "criteria"> {
  readonly id?: string;
  readonly criteria: DataValidationCriteriaInput;
}

export interface DataValidationResult {
  readonly valid: boolean;
  readonly ruleId?: string;
  readonly errorStyle?: ValidationErrorStyle;
  readonly title?: string;
  readonly message?: string;
}

/** 自定义校验函数：同步、纯函数。返回 `true` 有效；`false` 无效（用默认文案）；字符串无效并作为文案。 */
export type CustomValidatorFn = (
  value: CellValue,
  params: unknown,
  context: CustomValidatorContext
) => boolean | string;

export interface CustomValidatorContext {
  readonly row: number;
  readonly column: number;
  /** 只读，读的是当前已提交的状态。 */
  getCell(row: number, column: number): CellValue | undefined;
}

export interface ValidateOptions {
  /** 按名字查自定义校验函数；查不到按"有效"处理。 */
  readonly getCustomValidator?: (name: string) => CustomValidatorFn | undefined;
  readonly row?: number;
  readonly column?: number;
  readonly getCell?: (row: number, column: number) => CellValue | undefined;
}

// ---------------------------------------------------------------------------
// 形状校验
// ---------------------------------------------------------------------------

function invalid(message: string): never {
  throw new CyExcelError(CyExcelErrorCode.InvalidDataValidation, message);
}

const OPERATORS: ReadonlySet<string> = new Set<ValidationOperator>([
  "between",
  "notBetween",
  "equal",
  "notEqual",
  "greaterThan",
  "greaterThanOrEqual",
  "lessThan",
  "lessThanOrEqual"
]);

const TWO_VALUE_OPERATORS: ReadonlySet<string> = new Set<ValidationOperator>(["between", "notBetween"]);

const ERROR_STYLES: ReadonlySet<string> = new Set<ValidationErrorStyle>(["stop", "warning", "information"]);

function isNonNegativeInteger(value: unknown): value is number {
  return typeof value === "number" && Number.isSafeInteger(value) && value >= 0;
}

function assertValidRuleRange(range: CellRange): void {
  if (
    !isNonNegativeInteger(range?.startRow) ||
    !isNonNegativeInteger(range.startColumn) ||
    !isNonNegativeInteger(range.endRow) ||
    !isNonNegativeInteger(range.endColumn) ||
    range.startRow > range.endRow ||
    range.startColumn > range.endColumn
  ) {
    invalid(`invalid range in data validation rule: ${JSON.stringify(range)}`);
  }
}

/**
 * 校验一条规则的形状与不变量（设计稿 §3）。**不查规则之间的重叠**——那是
 * `applyRule` 的职责（后设置的挖掉先设置的）。这里查的是单条规则自己：
 * 区域合法且互不重叠、判据字段齐全、数值类型合适。
 */
export function assertValidRule(rule: DataValidationRule): void {
  if (typeof rule !== "object" || rule === null) {
    invalid("data validation rule must be an object");
  }
  if (typeof rule.id !== "string" || rule.id === "") {
    invalid("data validation rule id must be a non-empty string");
  }
  if (!Array.isArray(rule.ranges) || rule.ranges.length === 0) {
    invalid(`data validation rule "${rule.id}" must have at least one range`);
  }
  rule.ranges.forEach(assertValidRuleRange);
  for (let i = 0; i < rule.ranges.length; i += 1) {
    for (let j = i + 1; j < rule.ranges.length; j += 1) {
      if (rangesIntersect(rule.ranges[i]!, rule.ranges[j]!)) {
        invalid(`data validation rule "${rule.id}" has overlapping ranges`);
      }
    }
  }
  if (rule.errorStyle !== undefined && !ERROR_STYLES.has(rule.errorStyle)) {
    invalid(`unknown errorStyle "${String(rule.errorStyle)}"`);
  }
  if (rule.renderMode !== undefined && !RENDER_MODES.has(rule.renderMode)) {
    invalid(`unknown renderMode "${String(rule.renderMode)}"`);
  }

  const criteria = rule.criteria as DataValidationCriteria | undefined;
  if (typeof criteria !== "object" || criteria === null) {
    invalid(`data validation rule "${rule.id}" is missing criteria`);
  }
  switch (criteria.type) {
    case "whole":
    case "decimal":
    case "textLength":
    case "date":
    case "time": {
      if (!OPERATORS.has(criteria.operator)) {
        invalid(`unknown operator "${String(criteria.operator)}"`);
      }
      if (typeof criteria.value1 !== "number" || !Number.isFinite(criteria.value1)) {
        invalid(`${criteria.type} criteria needs a finite value1`);
      }
      if (TWO_VALUE_OPERATORS.has(criteria.operator)) {
        if (typeof criteria.value2 !== "number" || !Number.isFinite(criteria.value2)) {
          invalid(`operator "${criteria.operator}" needs a finite value2`);
        }
      }
      if (criteria.type === "whole") {
        const integers = [criteria.value1, criteria.value2].filter((v): v is number => v !== undefined);
        if (integers.some((v) => !Number.isInteger(v))) {
          invalid("whole criteria values must be integers");
        }
      }
      if (criteria.type === "textLength") {
        const lengths = [criteria.value1, criteria.value2].filter((v): v is number => v !== undefined);
        if (lengths.some((v) => !Number.isInteger(v) || v < 0)) {
          invalid("textLength criteria values must be non-negative integers");
        }
      }
      return;
    }
    case "list":
      if (
        !Array.isArray(criteria.items) ||
        criteria.items.length === 0 ||
        criteria.items.some((item) => typeof item !== "string")
      ) {
        invalid("list criteria needs a non-empty string array `items`");
      }
      if (criteria.colors !== undefined) {
        if (typeof criteria.colors !== "object" || criteria.colors === null || Array.isArray(criteria.colors)) {
          invalid("list criteria `colors` must be an object");
        }
        for (const [item, color] of Object.entries(criteria.colors)) {
          if (!criteria.items.includes(item)) {
            invalid(`list colors names "${item}", which is not one of the items`);
          }
          if (typeof color !== "string" || !COLOR_PATTERN.test(color)) {
            invalid(`list color for "${item}" must be a hex colour like #d93025`);
          }
        }
      }
      return;
    case "custom":
      if (typeof criteria.name !== "string" || criteria.name === "") {
        invalid("custom criteria needs a non-empty `name`");
      }
      return;
    default:
      invalid(`unknown criteria type "${String((criteria as { type: unknown }).type)}"`);
  }
}

type TemporalInput = Extract<DataValidationCriteriaInput, { readonly type: "date" | "time" }>;

function isTemporalInput(criteria: DataValidationCriteriaInput): criteria is TemporalInput {
  return criteria.type === "date" || criteria.type === "time";
}

function normalizeTemporal(value: number | string | undefined, label: string): number | undefined {
  if (value === undefined || typeof value === "number") {
    return value;
  }
  const serial = serialFromISODate(value);
  if (serial === null) {
    invalid(`${label} "${value}" is not a recognizable date/time`);
  }
  return serial;
}

/**
 * 把调用方的入参组装成一条规则：规范化日期字符串、把区域整理成互不相交的一组、
 * 补 `id`，并做形状校验。所有随机内容（`id`）都在这里定下，命令一旦构造就是纯数据。
 */
export function buildRule(
  input: DataValidationRuleInput,
  ranges: readonly CellRange[],
  generateId: () => string
): DataValidationRule {
  const source = input.criteria;
  const criteria: DataValidationCriteria = isTemporalInput(source)
    ? {
        type: source.type,
        operator: source.operator,
        value1: normalizeTemporal(source.value1, "value1") as number,
        ...(source.value2 !== undefined ? { value2: normalizeTemporal(source.value2, "value2") as number } : {})
      }
    : source;
  const { id, criteria: _ignored, ...rest } = input;
  void _ignored;
  const rule: DataValidationRule = {
    ...rest,
    id: id ?? generateId(),
    ranges: toDisjointRanges(ranges),
    criteria
  };
  assertValidRule(rule);
  return rule;
}

/** 校验整份规则列表：每条合法，且任意两条规则的区域不相交、id 不重复。 */
export function assertValidRules(rules: readonly DataValidationRule[]): void {
  const ids = new Set<string>();
  for (const rule of rules) {
    assertValidRule(rule);
    if (ids.has(rule.id)) {
      invalid(`duplicate data validation rule id "${rule.id}"`);
    }
    ids.add(rule.id);
  }
  for (let i = 0; i < rules.length; i += 1) {
    for (let j = i + 1; j < rules.length; j += 1) {
      const a = rules[i]!;
      const b = rules[j]!;
      if (a.ranges.some((ra) => b.ranges.some((rb) => rangesIntersect(ra, rb)))) {
        invalid(`data validation rules "${a.id}" and "${b.id}" overlap`);
      }
    }
  }
}

// ---------------------------------------------------------------------------
// 区域几何
// ---------------------------------------------------------------------------

/** 从 `range` 里挖掉 `cut`，返回剩下的矩形（0 到 4 个，互不相交）。 */
export function subtractRange(range: CellRange, cut: CellRange): readonly CellRange[] {
  if (!rangesIntersect(range, cut)) {
    return [range];
  }
  const pieces: CellRange[] = [];
  if (cut.startRow > range.startRow) {
    pieces.push({ ...range, endRow: cut.startRow - 1 });
  }
  if (cut.endRow < range.endRow) {
    pieces.push({ ...range, startRow: cut.endRow + 1 });
  }
  const midStart = Math.max(range.startRow, cut.startRow);
  const midEnd = Math.min(range.endRow, cut.endRow);
  if (cut.startColumn > range.startColumn) {
    pieces.push({ startRow: midStart, endRow: midEnd, startColumn: range.startColumn, endColumn: cut.startColumn - 1 });
  }
  if (cut.endColumn < range.endColumn) {
    pieces.push({ startRow: midStart, endRow: midEnd, startColumn: cut.endColumn + 1, endColumn: range.endColumn });
  }
  return pieces;
}

function subtractMany(range: CellRange, cuts: readonly CellRange[]): readonly CellRange[] {
  let pieces: readonly CellRange[] = [range];
  for (const cut of cuts) {
    pieces = pieces.flatMap((piece) => subtractRange(piece, cut));
  }
  return pieces;
}

/** 把一组可能互相重叠的矩形整理成互不相交的一组（后面的被前面的挖掉重叠部分）。 */
export function toDisjointRanges(ranges: readonly CellRange[]): readonly CellRange[] {
  const result: CellRange[] = [];
  for (const range of ranges) {
    result.push(...subtractMany(range, result));
  }
  return result;
}

function containsCell(range: CellRange, row: number, column: number): boolean {
  return row >= range.startRow && row <= range.endRow && column >= range.startColumn && column <= range.endColumn;
}

export function findRuleAt(
  rules: readonly DataValidationRule[],
  row: number,
  column: number
): DataValidationRule | undefined {
  if (rules.length > 8 && canIndexValidationRules(rules)) {
    return findIndexedValidationRule(rules, row, column);
  }
  return rules.find((rule) => rule.ranges.some((range) => containsCell(range, row, column)));
}

/** 与 `range` 相交的所有规则。绘制时每帧问一次，比逐格 `findRuleAt` 便宜。 */
export function rulesIntersecting(
  rules: readonly DataValidationRule[],
  range: CellRange
): readonly DataValidationRule[] {
  return rules.filter((rule) => rule.ranges.some((r) => rangesIntersect(r, range)));
}

/**
 * 设置一条规则：先从所有已有规则里挖掉 `rule.ranges` 的重叠部分（区域被挖空的规则
 * 整条消失），同 id 的旧规则被替换，最后把 `rule` 追加到末尾。"一格至多属于一条
 * 规则"由这里保证。
 */
export function applyRule(
  rules: readonly DataValidationRule[],
  rule: DataValidationRule
): readonly DataValidationRule[] {
  const cuts = rule.ranges;
  const next: DataValidationRule[] = [];
  for (const existing of rules) {
    if (existing.id === rule.id) {
      continue;
    }
    const ranges = existing.ranges.flatMap((range) => subtractMany(range, cuts));
    if (ranges.length > 0) {
      next.push(ranges.length === existing.ranges.length && ranges.every((r, i) => r === existing.ranges[i])
        ? existing
        : { ...existing, ranges });
    }
  }
  next.push(rule);
  return next;
}

/** 清掉 `ranges` 覆盖到的部分：只挖不追加。区域被挖空的规则整条消失。 */
export function clearRanges(
  rules: readonly DataValidationRule[],
  ranges: readonly CellRange[]
): readonly DataValidationRule[] {
  const next: DataValidationRule[] = [];
  for (const existing of rules) {
    const remaining = existing.ranges.flatMap((range) => subtractMany(range, ranges));
    if (remaining.length > 0) {
      next.push({ ...existing, ranges: remaining });
    }
  }
  return next;
}

// ---------------------------------------------------------------------------
// 结构操作搬移
// ---------------------------------------------------------------------------

function mapRules(
  rules: readonly DataValidationRule[],
  mapRange: (range: CellRange) => readonly CellRange[]
): readonly DataValidationRule[] {
  const next: DataValidationRule[] = [];
  for (const rule of rules) {
    const ranges = rule.ranges.flatMap(mapRange);
    if (ranges.length > 0) {
      next.push({ ...rule, ranges });
    }
  }
  return next;
}

/** 插入点在区域起点或之前平移，严格落在区域内部则撑大（与合并区同一口径）。 */
export function relocateValidationsForInsert(
  rules: readonly DataValidationRule[],
  axis: "row" | "column",
  start: number,
  count: number
): readonly DataValidationRule[] {
  return mapRules(rules, (range) => {
    const [s, e] = axisInterval(range, axis);
    return [withAxisInterval(range, axis, relocateIntervalForInsert(s, e, start, count))];
  });
}

/** 被删区间全覆盖则该区域消失，部分覆盖则收缩；规则的区域全没了则整条消失。 */
export function relocateValidationsForDelete(
  rules: readonly DataValidationRule[],
  axis: "row" | "column",
  start: number,
  count: number
): readonly DataValidationRule[] {
  return mapRules(rules, (range) => {
    const [s, e] = axisInterval(range, axis);
    const relocated = relocateIntervalForDelete(s, e, start, count);
    return relocated === null ? [] : [withAxisInterval(range, axis, relocated)];
  });
}

/**
 * 单轴区间在 `move(start, count, target)` 之后落在哪几段。移动是一个双射，但一个连续
 * 区间可能被移动块拆成几段（比如块从区间中间被抽走），所以返回的是**一组**区间。
 *
 * 映射的分段常数：块内整体平移；块外先摘除再插回，分界点是 `insertAt` 与
 * `insertAt + count`。在这四个分界点切开原区间，每一段的偏移量恒定，用
 * `remapIndex` 求段首的落点即可——不另起第四份移动公式。
 */
function relocateIntervalForMove(
  s: number,
  e: number,
  start: number,
  count: number,
  target: number
): readonly AxisInterval[] {
  const insertAt = target > start ? target - count : target;
  const cuts = [start, start + count, insertAt, insertAt + count]
    .filter((point) => point > s && point <= e)
    .sort((a, b) => a - b);
  const heads = [s, ...cuts.filter((point, i) => i === 0 || point !== cuts[i - 1])];
  const pieces: AxisInterval[] = heads.map((head, i) => {
    const tail = i + 1 < heads.length ? heads[i + 1]! - 1 : e;
    const delta = remapIndex(head, { operation: "move", start, count, target }) - head;
    return [head + delta, tail + delta];
  });
  pieces.sort((a, b) => a[0] - b[0]);
  // 落点相邻的段并回去，避免无谓地碎片化
  const merged: AxisInterval[] = [];
  for (const piece of pieces) {
    const last = merged[merged.length - 1];
    if (last !== undefined && piece[0] === last[1] + 1) {
      merged[merged.length - 1] = [last[0], piece[1]];
    } else {
      merged.push(piece);
    }
  }
  return merged;
}

/**
 * 移动行/列：**不拒绝**。合并区不能被拆开所以跨越块边界要拒绝，校验规则可以拆开——
 * 每个单元格跟着自己的行/列走，规则覆盖的单元格集合在移动之后不变，只是区域的
 * 划分可能变。撤销（逆移动）之后集合相同，划分不一定逐字相同。
 */
export function relocateValidationsForMove(
  rules: readonly DataValidationRule[],
  axis: "row" | "column",
  start: number,
  count: number,
  target: number
): readonly DataValidationRule[] {
  return mapRules(rules, (range) => {
    const [s, e] = axisInterval(range, axis);
    return relocateIntervalForMove(s, e, start, count, target).map((interval) =>
      withAxisInterval(range, axis, interval)
    );
  });
}

// ---------------------------------------------------------------------------
// 判定
// ---------------------------------------------------------------------------

function isBlank(value: unknown): boolean {
  return value === undefined || value === null || value === "";
}

function compare(operator: ValidationOperator, n: number, value1: number, value2: number | undefined): boolean {
  switch (operator) {
    case "between":
    case "notBetween": {
      const low = Math.min(value1, value2 ?? value1);
      const high = Math.max(value1, value2 ?? value1);
      const inside = n >= low && n <= high;
      return operator === "between" ? inside : !inside;
    }
    case "equal":
      return n === value1;
    case "notEqual":
      return n !== value1;
    case "greaterThan":
      return n > value1;
    case "greaterThanOrEqual":
      return n >= value1;
    case "lessThan":
      return n < value1;
    case "lessThanOrEqual":
      return n <= value1;
  }
}

/** 日期/时间用的数值：数字直接用；字符串按 numfmt 的日期解析；其余不是日期。 */
function toSerialNumber(value: CellValue): number | undefined {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : undefined;
  }
  if (typeof value === "string") {
    return serialFromISODate(value) ?? undefined;
  }
  return undefined;
}

const OPERATOR_TEXT: Record<ValidationOperator, string> = {
  between: "介于 {1} 和 {2} 之间",
  notBetween: "不介于 {1} 和 {2} 之间",
  equal: "等于 {1}",
  notEqual: "不等于 {1}",
  greaterThan: "大于 {1}",
  greaterThanOrEqual: "大于或等于 {1}",
  lessThan: "小于 {1}",
  lessThanOrEqual: "小于或等于 {1}"
};

const TYPE_TEXT = {
  whole: "整数",
  decimal: "数值",
  textLength: "文本长度",
  date: "日期",
  time: "时间"
} as const;

/** 没有自定义 `errorMessage` 时的默认文案。 */
const listMessages = new WeakMap<readonly string[], string>();
const listValues = new WeakMap<readonly string[], ReadonlySet<string>>();

export function defaultValidationMessage(criteria: DataValidationCriteria): string {
  switch (criteria.type) {
    case "list": {
      if (!Object.isFrozen(criteria.items)) {
        return `请从列表中选择：${criteria.items.join("、")}`;
      }
      let message = listMessages.get(criteria.items);
      if (message === undefined) {
        message = `请从列表中选择：${criteria.items.join("、")}`;
        listMessages.set(criteria.items, message);
      }
      return message;
    }
    case "custom":
      return "输入的值不符合规则";
    default: {
      const condition = OPERATOR_TEXT[criteria.operator]
        .replace("{1}", String(criteria.value1))
        .replace("{2}", String(criteria.value2 ?? ""));
      return `${TYPE_TEXT[criteria.type]}需要${condition}`;
    }
  }
}

function judge(rule: DataValidationRule, value: CellValue, options: ValidateOptions): true | string {
  const criteria = rule.criteria;
  switch (criteria.type) {
    case "whole":
      return typeof value === "number" &&
        Number.isInteger(value) &&
        compare(criteria.operator, value, criteria.value1, criteria.value2)
        ? true
        : defaultValidationMessage(criteria);
    case "decimal":
      return typeof value === "number" &&
        Number.isFinite(value) &&
        compare(criteria.operator, value, criteria.value1, criteria.value2)
        ? true
        : defaultValidationMessage(criteria);
    case "textLength":
      return compare(criteria.operator, String(value).length, criteria.value1, criteria.value2)
        ? true
        : defaultValidationMessage(criteria);
    case "date": {
      const serial = toSerialNumber(value);
      return serial !== undefined && compare(criteria.operator, serial, criteria.value1, criteria.value2)
        ? true
        : defaultValidationMessage(criteria);
    }
    case "time": {
      const serial = toSerialNumber(value);
      // 时间取序列号的小数部分：只比一天之内的钟点，与是哪一天无关
      return serial !== undefined &&
        compare(criteria.operator, serial - Math.floor(serial), criteria.value1, criteria.value2)
        ? true
        : defaultValidationMessage(criteria);
    }
    case "list": {
      if (!Object.isFrozen(criteria.items)) {
        return criteria.items.includes(String(value)) ? true : defaultValidationMessage(criteria);
      }
      let values = listValues.get(criteria.items);
      if (values === undefined) {
        values = new Set(criteria.items);
        listValues.set(criteria.items, values);
      }
      return values.has(String(value)) ? true : defaultValidationMessage(criteria);
    }
    case "custom": {
      const fn = options.getCustomValidator?.(criteria.name);
      if (fn === undefined) {
        // 没注册就当有效：别的环境打开同一份快照时不该因为缺一个函数而满屏报错
        return true;
      }
      const result = fn(value, criteria.params, {
        row: options.row ?? -1,
        column: options.column ?? -1,
        getCell: options.getCell ?? (() => undefined)
      });
      if (result === true) {
        return true;
      }
      return typeof result === "string" && result !== "" ? result : defaultValidationMessage(criteria);
    }
  }
}

/**
 * 按一条规则判定一个**已解析**的值（设计稿 §4）。同步纯函数。
 * - 空值返回 `allowBlank ?? true`。
 * - 类型不符判无效，不抛错；`whole` 不做文本到数字的隐式转换。
 */
export function validateValue(
  rule: DataValidationRule,
  value: CellValue | undefined | null,
  options: ValidateOptions = {}
): DataValidationResult {
  if (isBlank(value)) {
    return (rule.allowBlank ?? true) ? { valid: true } : invalidResult(rule, defaultBlankMessage(rule));
  }
  const verdict = judge(rule, value as CellValue, options);
  return verdict === true ? { valid: true } : invalidResult(rule, verdict);
}

function defaultBlankMessage(rule: DataValidationRule): string {
  return `不能为空。${defaultValidationMessage(rule.criteria)}`;
}

function invalidResult(rule: DataValidationRule, defaultMessage: string): DataValidationResult {
  return {
    valid: false,
    ruleId: rule.id,
    errorStyle: rule.errorStyle ?? "stop",
    ...(rule.errorTitle !== undefined ? { title: rule.errorTitle } : {}),
    message: rule.errorMessage !== undefined && rule.errorMessage !== "" ? rule.errorMessage : defaultMessage
  };
}

/** 两个矩形的交集；不相交返回 `undefined`。 */
export function intersectRanges(a: CellRange, b: CellRange): CellRange | undefined {
  const startRow = Math.max(a.startRow, b.startRow);
  const endRow = Math.min(a.endRow, b.endRow);
  const startColumn = Math.max(a.startColumn, b.startColumn);
  const endColumn = Math.min(a.endColumn, b.endColumn);
  return startRow <= endRow && startColumn <= endColumn ? { startRow, startColumn, endRow, endColumn } : undefined;
}

/**
 * 列出 `scope` 里不符合规则的格子，按行优先排序（"圈释无效数据"）。`scope` 由调用方
 * 先夹到坐标空间以内——规则区域可以是整列，不夹的话这里要扫到天边。空格子只有在规则
 * `allowBlank: false` 时才算无效。
 */
export function collectInvalidCells(
  rules: readonly DataValidationRule[],
  scope: CellRange,
  getValue: (row: number, column: number) => CellValue | undefined,
  optionsFor: (row: number, column: number) => ValidateOptions
): { readonly row: number; readonly column: number }[] {
  const found: { row: number; column: number }[] = [];
  for (const rule of rulesIntersecting(rules, scope)) {
    for (const ruleRange of rule.ranges) {
      const part = intersectRanges(ruleRange, scope);
      if (part === undefined) {
        continue;
      }
      for (let row = part.startRow; row <= part.endRow; row += 1) {
        for (let column = part.startColumn; column <= part.endColumn; column += 1) {
          if (!validateValue(rule, getValue(row, column), optionsFor(row, column)).valid) {
            found.push({ row, column });
          }
        }
      }
    }
  }
  return found.sort((a, b) => a.row - b.row || a.column - b.column);
}

/** 一个 list 单元格要怎么画：显示模式，以及当前值对应的颜色（没有则不带）。 */
export interface CellDropdownInfo {
  readonly mode: DataValidationRenderMode;
  readonly color?: string;
}

/**
 * 这一格（当前值 `value`）要不要画下拉、怎么画。非 `list` 规则、`showDropDown: false`
 * 都是 `undefined`。绘制时每个可见格每帧调一次，所以是纯查表。
 */
export function describeCellDropdown(
  rule: DataValidationRule,
  value: CellValue | undefined
): CellDropdownInfo | undefined {
  const criteria = rule.criteria;
  if (criteria.type !== "list" || rule.showDropDown === false) {
    return undefined;
  }
  const color = criteria.colors?.[String(value ?? "")];
  return { mode: rule.renderMode ?? "arrow", ...(color !== undefined ? { color } : {}) };
}

/**
 * 这一格的下拉里列哪些选项：`list` 规则且 `showDropDown !== false`，否则 `undefined`。
 * 选中带下拉的格子时，画布箭头的点击据此打开列表。
 */
export function describeDropdownItems(rule: DataValidationRule): readonly string[] | undefined {
  return rule.criteria.type === "list" && rule.showDropDown !== false ? rule.criteria.items : undefined;
}
