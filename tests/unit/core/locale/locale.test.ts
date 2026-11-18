import { describe, expect, it } from "vitest";
import { createTranslator, enUS, zhCN, type Locale } from "../../../../src/core/locale/index.js";
import { defaultValidationMessage, validateValue, type DataValidationRule } from "../../../../src/core/sheet/validations.js";
import { createDefaultToolbarSchema, defaultToolbarSchema } from "../../../../src/core/toolbar/schema.js";
import { DEFAULT_FORM, ruleInputFromForm } from "../../../../src/core/toolbar/dataValidationForm.js";

const en = createTranslator("en-US");
const zh = createTranslator();

describe("instance locale dictionaries", () => {
  it("ships matching dictionaries with matching interpolation parameters", () => {
    expect(Object.keys(zhCN.messages).sort()).toEqual(Object.keys(enUS.messages).sort());
    for (const key of Object.keys(enUS.messages) as (keyof typeof enUS.messages)[]) {
      const placeholders = (value: string) => [...value.matchAll(/\{\d+\}/g)].map(match => match[0]).sort();
      expect(placeholders(zhCN.messages[key]), key).toEqual(placeholders(enUS.messages[key]));
    }
  });

  it("keeps the Chinese default and isolates English instances", () => {
    expect(zh("Undo")).toBe("撤销");
    expect(en("Undo")).toBe("Undo");
    expect(zh("Undo")).toBe("撤销");
  });

  it("supports partial custom dictionaries, snapshots them, and falls back to English", () => {
    const messages = { Undo: "Annuler" };
    const fr = createTranslator({ name: "fr-FR", messages });
    messages.Undo = "Changed";
    expect(fr("Undo")).toBe("Annuler");
    expect(fr("Redo")).toBe("Redo");
    expect(fr("Delete sheet {0}", '<img src=x> $& {1}')).toBe('Delete sheet <img src=x> $& {1}');
  });

  it("interpolates reordered placeholders without recursively replacing user values", () => {
    const t = createTranslator({ name: "custom", messages: { "Row {0}, column {1}": "Column {1}; row {0}" } });
    expect(t("Row {0}, column {1}", 2, 5)).toBe("Column 5; row 2");
    expect(en("Delete sheet {0}")).toBe("Delete sheet {0}");
  });

  it("rejects malformed locale definitions from JavaScript callers", () => {
    for (const locale of ["unknown", null, {}, { name: "custom", messages: { Undo: 1 } }]) {
      expect(() => createTranslator(locale as Locale)).toThrow(TypeError);
    }
  });

  it("creates localized toolbar schemas while preserving the legacy Chinese export", () => {
    const english = createDefaultToolbarSchema(en);
    expect(english.find(item => item.id === "undo")?.label).toBe("Undo");
    expect(defaultToolbarSchema.find(item => item.id === "undo")?.label).toBe("撤销");
  });

  it("localizes validation criteria and keeps frozen list caches isolated", () => {
    const criteria = { type: "list", items: Object.freeze(["One", "Two"]) } as const;
    expect(defaultValidationMessage(criteria, en)).toBe("Choose from the list: One, Two");
    expect(defaultValidationMessage(criteria, zh)).toBe("请从列表中选择：One、Two");
    expect(defaultValidationMessage(criteria, en)).toBe("Choose from the list: One, Two");
    expect(defaultValidationMessage({ type: "whole", operator: "between", value1: 1, value2: 10 }, en))
      .toBe("Whole number must be between 1 and 10");
  });

  it("uses instance translations for blank values but preserves caller-supplied messages", () => {
    const rule: DataValidationRule = {
      id: "rule", ranges: [{ startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }],
      allowBlank: false, criteria: { type: "whole", operator: "greaterThan", value1: 0 }
    };
    expect(validateValue(rule, undefined, { t: en }).message).toBe("A value is required. Whole number must be greater than 0");
    expect(validateValue({ ...rule, errorMessage: "Custom message", errorTitle: "Custom title" }, -1, { t: en }))
      .toMatchObject({ message: "Custom message", title: "Custom title" });
  });

  it("localizes numeric and list form errors", () => {
    expect(ruleInputFromForm({ ...DEFAULT_FORM, allow: "whole", value1: "abc", operator: "equal" }, en))
      .toEqual({ kind: "error", message: "Value must be a number." });
    expect(ruleInputFromForm({ ...DEFAULT_FORM, allow: "list", value1: "Name=bad" }, en))
      .toEqual({ kind: "error", message: 'The color for "Name" must use #RRGGBB, for example #d93025.' });
  });
});
