import { afterEach, describe, expect, it, vi } from "vitest";
import { DocElement } from "../../docs-site/src/components/content";
import { apiTable, codeBlock, demoBlock, element, pageHeader } from "../../docs-site/src/components/doc";
import { setLocale } from "../../docs-site/src/i18n";
import { createSeo, renderSeoHead } from "../../docs-site/src/seo";

afterEach(() => {
  vi.unstubAllGlobals();
  setLocale("zh");
});

describe("static documentation content", () => {
  it("renders documentation without DOM globals or running demo callbacks", () => {
    expect(typeof document).toBe("undefined");
    expect(typeof window).toBe("undefined");
    const page = new DocElement("article");
    const demo = demoBlock({ title: "Example" });
    const initialize = vi.fn(() => { throw new Error("Browser required"); });
    demo.mount.onMount = initialize;
    page.append(
      pageHeader("API", "createSheet()", "Create an editable spreadsheet."),
      apiTable([{ name: "setCell", type: "function", description: "Write a value." }]),
      codeBlock('sheet.setCell(0, 0, "<value>");'),
      demo.root
    );
    const html = page.toHtml();
    expect(html).toContain('<h1 class="page-title">createSheet()</h1>');
    expect(html).toContain("Write a value.");
    expect(html).toContain("&lt;value&gt;");
    expect(html).toContain('<div class="demo-mount"></div>');
    expect(initialize).not.toHaveBeenCalled();
  });

  it("escapes documentation and attribute values while preserving code whitespace", () => {
    const link = element("a", undefined, '<script>alert("x")</script> & text');
    link.attributes.href = '/?q=" onclick="alert(1)';
    expect(link.toHtml()).toBe('<a href="/?q=&quot; onclick=&quot;alert(1)">&lt;script&gt;alert(&quot;x&quot;)&lt;/script&gt; &amp; text</a>');
    expect(codeBlock("const a = 1;\n\n  a < 2;").toHtml()).toContain("const a = 1;\n\n  a &lt; 2;");
  });

  it("uses the same translated content for static HTML", () => {
    setLocale("en");
    const header = pageHeader("EXAMPLES", "基础表格", "用 createSheet() 创建一个带工具栏的可编辑表格。");
    expect(header.toHtml()).not.toContain("基础表格");
    expect(header.findByClass("page-description")?.textContent).toContain("createSheet()");
  });
});

describe("static SEO metadata", () => {
  it("keeps language and canonical URLs under a custom deployment path", () => {
    vi.stubGlobal("__DOCS_SITE_URL__", "https://example.com/docs/sheets/");
    vi.stubGlobal("__DOCS_BASE_PATH__", "/docs/sheets/");
    const seo = createSeo("en", "/api/core", "Core API", "Core documentation");
    expect(seo.url).toBe("https://example.com/docs/sheets/en/api/core/");
    expect(seo.alternates).toEqual([
      { language: "zh-CN", href: "https://example.com/docs/sheets/api/core/" },
      { language: "en", href: seo.url },
      { language: "x-default", href: "https://example.com/docs/sheets/api/core/" }
    ]);
    expect(JSON.parse(seo.schema).inLanguage).toBe("en");
  });

  it("escapes HTML and prevents descriptions from closing the structured data script", () => {
    vi.stubGlobal("__DOCS_SITE_URL__", "https://example.com/");
    vi.stubGlobal("__DOCS_BASE_PATH__", "/");
    const description = '</script><script>alert("x")</script>';
    const seo = createSeo("zh", "/api/core", "<Core>", description);
    const html = renderSeoHead(seo);
    expect(html).toContain("<title>&lt;Core&gt;</title>");
    expect(html).not.toContain(description);
    expect(html.match(/<script/g)).toHaveLength(1);
    expect(JSON.parse(seo.schema).description).toBe(description);
  });
});
