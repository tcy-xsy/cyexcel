import { afterEach, describe, expect, it, vi } from "vitest";
import { currentDocRoute, docHref, parseDocRoute, setLocale, translate } from "../../../docs-site/src/i18n";

afterEach(() => {
  setLocale("zh");
  vi.unstubAllGlobals();
});

describe("documentation locale routes", () => {
  it.each([
    ["", "zh", "/"],
    ["#/", "zh", "/"],
    ["#/examples/styles", "zh", "/examples/styles"],
    ["#/en", "en", "/"],
    ["#/en/", "en", "/"],
    ["#/en/examples/styles/", "en", "/examples/styles"],
    ["#/extend/plugins", "zh", "/extend/plugins"]
  ] as const)("parses %s", (hash, locale, path) => {
    expect(parseDocRoute(hash)).toEqual({ locale, path });
  });

  it("preserves the page when switching language", () => {
    vi.stubGlobal("__DOCS_BASE_PATH__", "/cyexcel/");
    for (const path of ["/", "/quick-start", "/examples/validation", "/extend/plugins"]) {
      for (const locale of ["zh", "en"] as const) {
        vi.stubGlobal("window", { location: { pathname: docHref(locale, path), hash: "" } });
        expect(currentDocRoute()).toEqual({ locale, path });
      }
    }
  });
});

describe("documentation translation", () => {
  it("preserves original content in Chinese", () => {
    expect(translate("已撤销，当前值：42")).toBe("已撤销，当前值：42");
  });

  it("translates dynamic fragments without changing data or API names", () => {
    setLocale("en");
    expect(translate("已撤销，当前值：42")).toBe("Undone. Current value: 42");
    expect(translate("找到 3 处，已选中第一处")).toBe("Found 3 matches; selected the first match");
    expect(translate("sheet.setCell(0, 0, 42)")).toBe("sheet.setCell(0, 0, 42)");
  });

  it("uses complete phrases before shorter terms", () => {
    setLocale("en");
    expect(translate("数据验证规则")).toBe("Validation rules");
    expect(translate("stop（默认）")).toBe("stop (default)");
  });

  it("retains unknown text including object prototype names", () => {
    setLocale("en");
    for (const text of ["constructor", "__proto__", "Unknown API"]) {
      expect(translate(text)).toBe(text);
    }
  });
});
