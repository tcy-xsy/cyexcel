import { describe, expect, it } from "vitest";
import {
  buildCopyHtml,
  CopyStyleCache,
  extractCopyId,
  styleAt,
  type CopiedStyles
} from "../../../../src/core/clipboard/styleClipboard";

function copied(ids: (string | undefined)[][], styles: Record<string, unknown> = {}): CopiedStyles {
  return {
    rows: ids.length,
    columns: ids[0]?.length ?? 0,
    styles: styles as CopiedStyles["styles"],
    ids
  };
}

describe("extractCopyId", () => {
  it("从我们自己生成的 HTML 里抠出 id", () => {
    const html = buildCopyHtml([["a"]], copied([[undefined]]), "cy-abc");
    expect(extractCopyId(html)).toBe("cy-abc");
  });

  it("外来 HTML 没有这个属性，认不出来", () => {
    expect(extractCopyId("<table><tr><td>from Excel</td></tr></table>")).toBeNull();
    expect(extractCopyId("")).toBeNull();
    expect(extractCopyId(undefined)).toBeNull();
  });
});

describe("buildCopyHtml", () => {
  it("样式翻成内联 CSS——跨应用互通就靠这个", () => {
    const html = buildCopyHtml(
      [["粗体"]],
      copied([["k0"]], { k0: { bold: true, color: "#f00" } }),
      "cy-1"
    );
    expect(html).toContain("font-weight:600");
    expect(html).toContain("color:#f00");
  });

  it("没有样式的格子不写 style 属性", () => {
    const html = buildCopyHtml([["纯文本"]], copied([[undefined]]), "cy-1");
    expect(html).toContain("<td>纯文本</td>");
  });

  it("转义 HTML 特殊字符——单元格里的 <script> 不能变成标签", () => {
    const html = buildCopyHtml([["<script>x</script>"]], copied([[undefined]]), "cy-1");
    expect(html).not.toContain("<script>");
    expect(html).toContain("&lt;script&gt;");
  });
});

describe("CopyStyleCache", () => {
  it("存进去再按 HTML 认领回来", () => {
    const cache = new CopyStyleCache();
    const payload = copied([["k0"]], { k0: { bold: true } });
    const id = cache.set(payload);

    const html = buildCopyHtml([["x"]], payload, id);
    expect(cache.claim(html)).toBe(payload);
  });

  it("外来 HTML 认领不到，调用方据此退回纯文本粘贴", () => {
    const cache = new CopyStyleCache();
    cache.set(copied([["k0"]], { k0: { bold: true } }));
    expect(cache.claim("<table><tr><td>from Excel</td></tr></table>")).toBeUndefined();
    expect(cache.claim(undefined)).toBeUndefined();
  });

  it("超过 10 次复制后最早那次被挤掉——刻意的内存上限，不是 bug", () => {
    const cache = new CopyStyleCache();
    const first = copied([["k0"]], { k0: { bold: true } });
    const firstId = cache.set(first);

    for (let i = 0; i < CopyStyleCache.LIMIT; i += 1) {
      cache.set(copied([[undefined]]));
    }

    expect(cache.size).toBe(CopyStyleCache.LIMIT);
    expect(cache.claim(buildCopyHtml([["x"]], first, firstId))).toBeUndefined();
  });

  it("第 10 次以内的复制仍然认领得到", () => {
    const cache = new CopyStyleCache();
    const payload = copied([["k0"]], { k0: { bold: true } });
    const id = cache.set(payload);
    for (let i = 0; i < CopyStyleCache.LIMIT - 1; i += 1) {
      cache.set(copied([[undefined]]));
    }
    expect(cache.claim(buildCopyHtml([["x"]], payload, id))).toBe(payload);
  });
});

describe("styleAt", () => {
  it("按相对坐标取样式", () => {
    const payload = copied([[undefined, "k0"]], { k0: { italic: true } });
    expect(styleAt(payload, 0, 1)).toEqual({ italic: true });
    expect(styleAt(payload, 0, 0)).toBeUndefined();
  });

  it("越界返回 undefined，不抛", () => {
    expect(styleAt(copied([[undefined]]), 5, 5)).toBeUndefined();
  });
});
