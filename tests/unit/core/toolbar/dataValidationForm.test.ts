import { describe, expect, it } from "vitest";

import {
  DEFAULT_FORM,
  fieldsFor,
  formFromRule,
  formatClock,
  parseClock,
  parseListItems,
  ruleInputFromForm,
  type DataValidationFormState
} from "../../../../src/core/toolbar/dataValidationForm";
import { buildRule, type DataValidationRule } from "../../../../src/core/sheet/validations";

function form(overrides: Partial<DataValidationFormState>): DataValidationFormState {
  return { ...DEFAULT_FORM, ...overrides };
}

const cell = [{ startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 }];

describe("fieldsFor", () => {
  it("shows nothing for 'any', only the source and arrow for a list", () => {
    expect(fieldsFor("any", "between")).toEqual({ operator: false, value1: false, value2: false, dropDown: false });
    expect(fieldsFor("list", "between")).toEqual({ operator: false, value1: true, value2: false, dropDown: true });
  });

  it("shows the second value only for between / notBetween", () => {
    expect(fieldsFor("whole", "between").value2).toBe(true);
    expect(fieldsFor("whole", "notBetween").value2).toBe(true);
    expect(fieldsFor("whole", "equal").value2).toBe(false);
    expect(fieldsFor("date", "greaterThan")).toEqual({ operator: true, value1: true, value2: false, dropDown: false });
  });
});

describe("parseListItems", () => {
  it("splits on ASCII and full-width commas, trims, drops blanks and duplicates", () => {
    expect(parseListItems(" 红 ，绿,, 蓝 ,红 ")).toEqual(["红", "绿", "蓝"]);
    expect(parseListItems("")).toEqual([]);
    expect(parseListItems(" , ,")).toEqual([]);
  });
});

describe("parseClock / formatClock", () => {
  it("parses hh:mm and hh:mm:ss into a fraction of a day", () => {
    expect(parseClock("12:00")).toBe(0.5);
    expect(parseClock("06:00")).toBe(0.25);
    expect(parseClock("9:30")).toBeCloseTo(9.5 / 24, 10);
    expect(parseClock("00:00:30")).toBeCloseTo(30 / 86400, 10);
  });

  it("rejects things that are not a time of day", () => {
    for (const bad of ["", "24:00", "12:60", "12", "ab:cd", "12:5"]) {
      expect(parseClock(bad), bad).toBeUndefined();
    }
  });

  it("formats a fraction back to hh:mm", () => {
    expect(formatClock(0.5)).toBe("12:00");
    expect(formatClock(parseClock("09:30")!)).toBe("09:30");
  });
});

describe("ruleInputFromForm", () => {
  it("'any' means clear", () => {
    expect(ruleInputFromForm(form({ allow: "any" }))).toEqual({ kind: "clear" });
  });

  it("builds a whole-number between rule with the shared options", () => {
    const outcome = ruleInputFromForm(
      form({ allow: "whole", operator: "between", value1: "1", value2: "10", errorStyle: "warning", errorMessage: "太大", prompt: "填 1~10" })
    );
    expect(outcome).toEqual({
      kind: "rule",
      input: {
        allowBlank: true,
        errorStyle: "warning",
        errorMessage: "太大",
        prompt: "填 1~10",
        criteria: { type: "whole", operator: "between", value1: 1, value2: 10 }
      }
    });
  });

  it("omits value2 for single-value operators and empty optional texts", () => {
    const outcome = ruleInputFromForm(form({ allow: "decimal", operator: "greaterThan", value1: "0.5", value2: "ignored" }));
    expect(outcome.kind).toBe("rule");
    if (outcome.kind === "rule") {
      expect(outcome.input.criteria).toEqual({ type: "decimal", operator: "greaterThan", value1: 0.5 });
      expect("errorTitle" in outcome.input).toBe(false);
      expect("prompt" in outcome.input).toBe(false);
    }
  });

  it("builds a list rule with its drop-down flag", () => {
    const outcome = ruleInputFromForm(form({ allow: "list", value1: "红,绿", showDropDown: false }));
    expect(outcome).toMatchObject({
      kind: "rule",
      input: { showDropDown: false, criteria: { type: "list", items: ["红", "绿"] } }
    });
  });

  it("parses dates into serial numbers and times into fractions", () => {
    const date = ruleInputFromForm(form({ allow: "date", operator: "between", value1: "2026-01-01", value2: "2026-01-31" }));
    expect(date).toMatchObject({ kind: "rule", input: { criteria: { type: "date", value1: 46023, value2: 46053 } } });
    const time = ruleInputFromForm(form({ allow: "time", operator: "greaterThanOrEqual", value1: "12:00" }));
    expect(time).toMatchObject({ kind: "rule", input: { criteria: { type: "time", value1: 0.5 } } });
  });

  it.each([
    ["missing number", form({ allow: "whole", operator: "equal", value1: "" }), "请填写"],
    ["non-numeric", form({ allow: "decimal", operator: "equal", value1: "abc" }), "需要是数字"],
    ["fractional whole", form({ allow: "whole", operator: "equal", value1: "1.5" }), "需要是整数"],
    ["negative length", form({ allow: "textLength", operator: "equal", value1: "-1" }), "不能是负数"],
    ["missing max", form({ allow: "whole", operator: "between", value1: "1", value2: "" }), "最大值"],
    ["bad date", form({ allow: "date", operator: "equal", value1: "tomorrow" }), "需要是日期"],
    ["bad time", form({ allow: "time", operator: "equal", value1: "noon" }), "需要是时间"],
    ["empty list", form({ allow: "list", value1: " , " }), "请填写序列"]
  ])("reports %s as a readable error", (_label, state, fragment) => {
    const outcome = ruleInputFromForm(state);
    expect(outcome.kind).toBe("error");
    if (outcome.kind === "error") {
      expect(outcome.message).toContain(fragment);
    }
  });

  it("produces input that setDataValidation would accept", () => {
    const outcome = ruleInputFromForm(form({ allow: "date", operator: "greaterThan", value1: "2026-01-01" }));
    if (outcome.kind !== "rule") {
      throw new Error("expected a rule");
    }
    expect(() => buildRule(outcome.input, cell, () => "id")).not.toThrow();
  });
});

describe("formFromRule", () => {
  function roundTrip(state: DataValidationFormState): DataValidationFormState {
    const outcome = ruleInputFromForm(state);
    if (outcome.kind !== "rule") {
      throw new Error("expected a rule");
    }
    return formFromRule(buildRule(outcome.input, cell, () => "id"));
  }

  it("gives the defaults when there is no rule", () => {
    expect(formFromRule(undefined)).toEqual(DEFAULT_FORM);
  });

  it.each([
    form({ allow: "whole", operator: "between", value1: "1", value2: "10" }),
    form({ allow: "textLength", operator: "lessThanOrEqual", value1: "5", errorStyle: "information", errorTitle: "t", errorMessage: "m" }),
    form({ allow: "list", value1: "红,绿,蓝", showDropDown: false, allowBlank: false }),
    form({ allow: "date", operator: "between", value1: "2026-01-01", value2: "2026-01-31", promptTitle: "日期", prompt: "填今年的" }),
    form({ allow: "time", operator: "notBetween", value1: "09:00", value2: "18:00" })
  ])("survives form → rule → form for %j", (state) => {
    expect(roundTrip(state)).toEqual(state);
  });

  it("falls back to 'any' for a custom rule the form cannot express", () => {
    const rule: DataValidationRule = { id: "c", ranges: cell, criteria: { type: "custom", name: "even" }, errorStyle: "warning" };
    expect(formFromRule(rule)).toMatchObject({ allow: "any", errorStyle: "warning" });
  });
});
