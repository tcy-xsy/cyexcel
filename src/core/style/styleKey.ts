import type { CellStyle, CellStyleKey } from "./types";
import { STYLE_KEYS } from "./types";

/**
 * 规范化 intern key：按 `STYLE_KEYS` 的固定顺序序列化，跳过 `undefined`。
 *
 * **键顺序必须由这里决定，不能用 `JSON.stringify(style)`。** Univer 的 `Styles`
 * 就是直接 stringify 的，于是 `{bold,color}` 与 `{color,bold}` 得到两个不同的 key，
 * 同一个样式换个写法就必然 cache miss——它靠"遍历整张 styles 表做深比较"兜底
 * （`_getExistingStyleId`，O(n) 每次）。规范化之后同一个样式只有一种 key 表示，
 * O(1) 命中，那条兜底路径根本不需要存在。
 *
 * `border` 是唯一的嵌套字段，四条边同样按固定顺序展开，理由相同。
 */
export function styleKey(style: CellStyle): string {
  const parts: string[] = [];
  for (const key of STYLE_KEYS) {
    const value = style[key];
    if (value === undefined) {
      continue;
    }
    parts.push(key === "border" ? `border:${borderKey(style.border)}` : `${key}:${JSON.stringify(value)}`);
  }
  return parts.join("|");
}

function borderKey(border: CellStyle["border"]): string {
  if (!border) {
    return JSON.stringify(border);
  }
  // 固定 top/right/bottom/left 顺序——与 CSS 的简写顺序一致，省得读的人再记一套。
  const sides = (["top", "right", "bottom", "left"] as const).map((side) => {
    const value = border[side];
    return value === undefined ? "" : `${value.color},${value.width ?? 1},${value.style ?? "solid"}`;
  });
  return sides.join(";");
}

/**
 * 合并语义：`patch` 里的 `null` 清除该键，`undefined`（键不存在）保留原值。
 *
 * 这是 Step 25 工具栏 toggle 按钮唯一能用的语义——"点一下加粗"只该改 `bold`
 * 一个键，不该把这一格的颜色、边框、数字格式一起抹掉。
 *
 * 返回 `undefined` 表示合并结果是空样式，调用方应当把 `CellRecord.s` 整个删掉
 * 而不是指向一条空记录（与"未占用的格子不存空记录"是同一条不变量）。
 */
export function mergeStyle(
  base: CellStyle | undefined,
  patch: CellStylePatch
): CellStyle | undefined {
  const merged: Record<string, unknown> = { ...base };
  for (const key of STYLE_KEYS) {
    if (!(key in patch)) {
      continue;
    }
    const value = patch[key];
    if (value === null) {
      delete merged[key];
    } else {
      merged[key] = value;
    }
  }
  return Object.keys(merged).length === 0 ? undefined : (merged as CellStyle);
}

/** `setCellStyle` 的入参：省略的键保留，`null` 的键清除。 */
export type CellStylePatch = {
  readonly [K in CellStyleKey]?: CellStyle[K] | null;
};
