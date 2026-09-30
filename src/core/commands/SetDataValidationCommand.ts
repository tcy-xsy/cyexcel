import type { Command } from "../types";
import type { DataValidationRule } from "../sheet/validations";

export interface SetDataValidationCommandParams {
  readonly rule: DataValidationRule;
}

/**
 * 设置一条数据验证规则。规则（含 `id` 与规整后的互不相交区域）在**构造时**就定好，
 * plan 期不生成任何随机内容——重放同一条命令必须得到同一份规则。
 */
export class SetDataValidationCommand implements Command<"set-data-validation", SetDataValidationCommandParams> {
  readonly type = "set-data-validation";
  readonly params: SetDataValidationCommandParams;

  constructor(rule: DataValidationRule) {
    this.params = Object.freeze({
      rule: Object.freeze({ ...rule, ranges: Object.freeze(rule.ranges.map((range) => Object.freeze({ ...range }))) })
    });
    Object.freeze(this);
  }
}
