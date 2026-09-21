import { describe, expect, it } from "vitest";
import { StyleRegistry } from "../../../../src/core/style/StyleRegistry";
import { mergeStyle, styleKey } from "../../../../src/core/style/styleKey";
import type { CellStyle } from "../../../../src/core/style/types";

/** 确定性 id 生成器——生产用随机串，测试要能断言具体 id。 */
function sequentialIds(prefix = "s"): () => string {
  let next = 1;
  return () => `${prefix}${next++}`;
}

describe("styleKey", () => {
  it("与键的书写顺序无关", () => {
    const a: CellStyle = { bold: true, color: "#f00" };
    const b: CellStyle = { color: "#f00", bold: true };
    expect(styleKey(a)).toBe(styleKey(b));
  });

  it("区分不同的值", () => {
    expect(styleKey({ bold: true })).not.toBe(styleKey({ bold: false }));
    expect(styleKey({ fontSize: 12 })).not.toBe(styleKey({ fontSize: 13 }));
  });

  it("跳过 undefined 的键——显式写 undefined 和根本不写是同一个样式", () => {
    // `exactOptionalPropertyTypes` 下 TS 不允许显式 undefined，但 JS 调用方照样
    // 传得进来（比如把一个字段全是可选的表单对象直接铺开），所以运行时要兜住。
    const withUndefined = { bold: true, color: undefined } as unknown as CellStyle;
    expect(styleKey(withUndefined)).toBe(styleKey({ bold: true }));
  });

  it("边框按固定的 top/right/bottom/left 顺序展开", () => {
    const a: CellStyle = { border: { left: { color: "#000" }, top: { color: "#fff" } } };
    const b: CellStyle = { border: { top: { color: "#fff" }, left: { color: "#000" } } };
    expect(styleKey(a)).toBe(styleKey(b));
  });

  it("边框的默认 width/style 与显式写出来是同一个 key", () => {
    const implicit: CellStyle = { border: { top: { color: "#000" } } };
    const explicit: CellStyle = { border: { top: { color: "#000", width: 1, style: "solid" } } };
    expect(styleKey(implicit)).toBe(styleKey(explicit));
  });

  it("空样式的 key 是空串", () => {
    expect(styleKey({})).toBe("");
  });
});

describe("mergeStyle", () => {
  it("只改传进来的那个键，其余保留", () => {
    const base: CellStyle = { bold: true, color: "#f00", numFmt: "0.00" };
    expect(mergeStyle(base, { italic: true })).toEqual({
      bold: true,
      color: "#f00",
      numFmt: "0.00",
      italic: true
    });
  });

  it("null 清除单个键", () => {
    const base: CellStyle = { bold: true, color: "#f00" };
    expect(mergeStyle(base, { bold: null })).toEqual({ color: "#f00" });
  });

  it("清掉最后一个键时返回 undefined，而不是空对象", () => {
    expect(mergeStyle({ bold: true }, { bold: null })).toBeUndefined();
  });

  it("base 为 undefined 时就是 patch 本身", () => {
    expect(mergeStyle(undefined, { bold: true })).toEqual({ bold: true });
  });

  it("清一个本来就没有的键是无操作，不会凭空造出键", () => {
    expect(mergeStyle({ bold: true }, { italic: null })).toEqual({ bold: true });
  });

  it("不认识的键不会被塞进结果里", () => {
    const patch = { bold: true, nonsense: 1 } as unknown as Parameters<typeof mergeStyle>[1];
    expect(mergeStyle(undefined, patch)).toEqual({ bold: true });
  });
});

describe("StyleRegistry", () => {
  it("同一个样式 intern 两次得到同一个 id，表里只有一条", () => {
    const registry = new StyleRegistry({}, sequentialIds());
    const first = registry.intern({ bold: true });
    const second = registry.intern({ bold: true });
    expect(second).toBe(first);
    expect(registry.size).toBe(1);
  });

  it("键序不同的同一个样式也复用同一条", () => {
    const registry = new StyleRegistry({}, sequentialIds());
    const first = registry.intern({ bold: true, color: "#f00" });
    const second = registry.intern({ color: "#f00", bold: true });
    expect(second).toBe(first);
    expect(registry.size).toBe(1);
  });

  it("不同样式得到不同 id", () => {
    const registry = new StyleRegistry({}, sequentialIds());
    expect(registry.intern({ bold: true })).not.toBe(registry.intern({ italic: true }));
    expect(registry.size).toBe(2);
  });

  it("get 取回原样式，未知 id 返回 undefined", () => {
    const registry = new StyleRegistry({}, sequentialIds());
    const id = registry.intern({ fontSize: 16 });
    expect(registry.get(id)).toEqual({ fontSize: 16 });
    expect(registry.get("nope")).toBeUndefined();
  });

  it("拒绝 intern 空样式——没有样式的格子根本不该有 s 字段", () => {
    const registry = new StyleRegistry({}, sequentialIds());
    expect(() => registry.intern({})).toThrow(/empty style/);
  });

  it("id 撞上时重新生成，不覆盖已有条目", () => {
    // 前两次都吐 "dup"，第三次才给新的——模拟随机串撞车。
    const ids = ["dup", "dup", "fresh"];
    let index = 0;
    const registry = new StyleRegistry({}, () => ids[index++] ?? "exhausted");

    const first = registry.intern({ bold: true });
    const second = registry.intern({ italic: true });

    expect(first).toBe("dup");
    expect(second).toBe("fresh");
    expect(registry.get("dup")).toEqual({ bold: true });
    expect(registry.get("fresh")).toEqual({ italic: true });
  });

  it("从快照加载已有样式，并能继续 intern 命中它们", () => {
    const registry = new StyleRegistry({ existing: { bold: true } }, sequentialIds());
    expect(registry.get("existing")).toEqual({ bold: true });
    expect(registry.intern({ bold: true })).toBe("existing");
    expect(registry.size).toBe(1);
  });

  it("快照里同一样式存了两个 id 时，两个都还能取，但 intern 一律收敛到第一个", () => {
    // 随机 id 不幂等，这种快照是可能存在的。
    const registry = new StyleRegistry(
      { first: { bold: true }, second: { bold: true } },
      sequentialIds()
    );
    expect(registry.get("first")).toEqual({ bold: true });
    expect(registry.get("second")).toEqual({ bold: true });
    expect(registry.intern({ bold: true })).toBe("first");
  });

  describe("toJSON", () => {
    it("不传 usedIds 时整表导出", () => {
      const registry = new StyleRegistry({}, sequentialIds());
      registry.intern({ bold: true });
      registry.intern({ italic: true });
      expect(Object.keys(registry.toJSON())).toHaveLength(2);
    });

    it("传 usedIds 时剔除无人引用的条目——这是唯一的回收时机", () => {
      const registry = new StyleRegistry({}, sequentialIds());
      const used = registry.intern({ bold: true });
      registry.intern({ italic: true }); // 没人引用

      expect(registry.toJSON([used])).toEqual({ s1: { bold: true } });
    });

    it("usedIds 里的未知 id 被忽略，不会写出 undefined 条目", () => {
      const registry = new StyleRegistry({}, sequentialIds());
      const used = registry.intern({ bold: true });
      expect(registry.toJSON([used, "ghost"])).toEqual({ s1: { bold: true } });
    });

    it("剔除只影响导出，不改内存里的表——同一次会话里 undo 回去还能用", () => {
      const registry = new StyleRegistry({}, sequentialIds());
      const kept = registry.intern({ bold: true });
      const dropped = registry.intern({ italic: true });

      registry.toJSON([kept]);

      expect(registry.get(dropped)).toEqual({ italic: true });
      expect(registry.size).toBe(2);
    });
  });
});
