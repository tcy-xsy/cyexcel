import type { CellStyle } from "../style/types";

/**
 * 一次复制里随值一起带走的样式。
 *
 * 存的是**样式对象**而不是 id：样式 id 是随机串（见 `StyleRegistry`），同一个
 * "加粗12px"在两个文档里是两个不同的 id，带 id 过去没有意义。粘贴时在目标表的
 * registry 重新 intern。
 *
 * 表示成"子表 + 引用矩阵"而不是每格一个对象：复制十万格时前者是十万个字符串引用
 * 加一张小表，后者是十万个对象。
 */
export interface CopiedStyles {
  readonly rows: number;
  readonly columns: number;
  /** id → 样式对象，只含这次复制真正用到的那些 */
  readonly styles: Readonly<Record<string, CellStyle>>;
  /** `[r][c]` 是相对坐标上那一格的样式 id，没有样式则 `undefined` */
  readonly ids: readonly (readonly (string | undefined)[])[];
}

/** 取相对坐标 `(r, c)` 的样式对象。越界或无样式都返回 `undefined`。 */
export function styleAt(copied: CopiedStyles, r: number, c: number): CellStyle | undefined {
  const id = copied.ids[r]?.[c];
  return id === undefined ? undefined : copied.styles[id];
}

const COPY_ID_PATTERN = /data-copy-id="([^"\s]+)"/;

/**
 * 从粘贴进来的 HTML 里抠出 copyId。
 *
 * 这是"这份剪贴板内容是不是我刚才拷的"唯一可靠的判据——系统剪贴板是唯一能把一个
 * 标识带过上下文边界（另一个标签页、另一次刷新）的东西。Univer 用的是同一招，
 * 正则也几乎一样。
 */
export function extractCopyId(html: string | undefined): string | null {
  if (html === undefined || html === "") {
    return null;
  }
  return COPY_ID_PATTERN.exec(html)?.[1] ?? null;
}

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/** 样式对象翻成内联 CSS——跨应用互通就靠它（我们复制 → Excel/Sheets 粘贴有样式）。 */
function inlineCss(style: CellStyle | undefined): string {
  if (style === undefined) {
    return "";
  }
  const parts: string[] = [];
  if (style.bold === true) parts.push("font-weight:600");
  if (style.italic === true) parts.push("font-style:italic");
  if (style.underline === true || style.strike === true) {
    const decorations = [
      style.underline === true ? "underline" : "",
      style.strike === true ? "line-through" : ""
    ].filter(Boolean);
    parts.push(`text-decoration:${decorations.join(" ")}`);
  }
  if (style.color !== undefined) parts.push(`color:${style.color}`);
  if (style.background !== undefined) parts.push(`background-color:${style.background}`);
  if (style.fontSize !== undefined) parts.push(`font-size:${style.fontSize}px`);
  if (style.fontFamily !== undefined) parts.push(`font-family:${style.fontFamily}`);
  if (style.align !== undefined) parts.push(`text-align:${style.align}`);
  if (style.verticalAlign !== undefined) parts.push(`vertical-align:${style.verticalAlign}`);
  return parts.join(";");
}

/**
 * 生成随 `text/plain` 一起写进剪贴板的那份 HTML。
 *
 * 它有两个作用，**主次分明**：
 * 1. 载着 `data-copy-id`，让同一个应用内的粘贴能认领到内存里那份完整数据；
 * 2. 内联样式让 Excel / Google Sheets 粘过去也带格式。
 *
 * 反过来**不解析外来 HTML**：从 Excel 拷来的内容我们只取纯文本。那一半
 * （Univer 的 `html-to-usm` + 按来源特化的三个插件 + 消毒 + 体积防护）明确不做。
 */
export function buildCopyHtml(
  grid: readonly (readonly string[])[],
  copied: CopiedStyles,
  copyId: string
): string {
  const rows = grid
    .map((row, r) => {
      const cells = row
        .map((text, c) => {
          const css = inlineCss(styleAt(copied, r, c));
          return `<td${css === "" ? "" : ` style="${css}"`}>${escapeHtml(text)}</td>`;
        })
        .join("");
      return `<tr>${cells}</tr>`;
    })
    .join("");
  return `<table data-copy-id="${copyId}"><tbody>${rows}</tbody></table>`;
}

/**
 * 最近几次复制的样式，按 copyId 索引。
 *
 * **上限 10 条**（与 Univer 的 `COPY_CONTENT_CACHE_LIMIT` 同）：复制十次以上，
 * 最早那次的保真能力过期，粘贴退回纯文本。这是刻意的内存上限，不是 bug——
 * 一次大范围复制的引用矩阵不小，无上限地留着才是问题。
 */
export class CopyStyleCache {
  static readonly LIMIT = 10;
  readonly #entries = new Map<string, CopiedStyles>();
  #nextId = 0;

  /** 记下一次复制，返回它的 copyId（写进 HTML 的那个）。 */
  set(copied: CopiedStyles): string {
    const id = `cy-${Date.now().toString(36)}-${(this.#nextId += 1).toString(36)}`;
    this.#entries.set(id, copied);
    // Map 的迭代顺序就是插入顺序，所以第一个 key 一定是最老的那次复制。
    while (this.#entries.size > CopyStyleCache.LIMIT) {
      const oldest = this.#entries.keys().next().value as string;
      this.#entries.delete(oldest);
    }
    return id;
  }

  /**
   * 认领：给定粘贴时拿到的 HTML，看它是不是我们自己拷出去的、而且那次的数据还在。
   * 认不出来就返回 `undefined`，调用方退回纯文本粘贴。
   */
  claim(html: string | undefined): CopiedStyles | undefined {
    const id = extractCopyId(html);
    return id === null ? undefined : this.#entries.get(id);
  }

  get size(): number {
    return this.#entries.size;
  }
}
