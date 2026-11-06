import { afterEach, describe, expect, it, vi } from "vitest";
import { currentDocRoute, docHref } from "../../docs-site/src/i18n";

afterEach(() => vi.unstubAllGlobals());

describe("documentation URLs", () => {
  it.each([
    ["zh", "/", "/cyexcel/"],
    ["en", "/", "/cyexcel/en/"],
    ["zh", "/api/core", "/cyexcel/api/core/"],
    ["en", "/api/core", "/cyexcel/en/api/core/"]
  ] as const)("links %s %s to a static directory", (locale, path, expected) => {
    vi.stubGlobal("__DOCS_BASE_PATH__", "/cyexcel/");
    expect(docHref(locale, path)).toBe(expected);
  });

  it.each([
    ["/cyexcel/en/api/core/", "", "en", "/api/core"],
    ["/cyexcel/examples/basic/", "", "zh", "/examples/basic"],
    ["/cyexcel/", "#/en/api/core", "en", "/api/core"],
    ["/cyexcel/en/", "#/quick-start", "zh", "/quick-start"],
    ["/cyexcel/en/api/core/", "#methods", "en", "/api/core"]
  ])("resolves %s%s", (pathname, hash, locale, path) => {
    vi.stubGlobal("__DOCS_BASE_PATH__", "/cyexcel/");
    vi.stubGlobal("window", { location: { pathname, hash } });
    expect(currentDocRoute()).toEqual({ locale, path });
  });

  it("supports root deployment and custom subdirectories", () => {
    for (const base of ["/", "/docs/spreadsheet/"]) {
      vi.stubGlobal("__DOCS_BASE_PATH__", base);
      vi.stubGlobal("window", { location: { pathname: `${base}en/quick-start/`, hash: "" } });
      expect(docHref("en", "/quick-start")).toBe(`${base}en/quick-start/`);
      expect(currentDocRoute()).toEqual({ locale: "en", path: "/quick-start" });
    }
  });
});
