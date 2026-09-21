import { describe, expect, it } from "vitest";

import { setHidden } from "../../../../src/core/toolbar/styles";

/**
 * `setHidden` 只碰三样东西：`hidden` 属性、内联 `display`、以及存档用的 `dataset`。
 * 这个环境没有 DOM（vitest `environment: "node"`），但这三样用一个最小替身就能完整
 * 表达。
 *
 * **为什么值得单独钉一条**：把 `setHidden` "简化"回一句 `element.hidden = true` 在
 * node 里看不出任何区别，在浏览器里却会静默失效——UA 样式表的 `[hidden]{display:none}`
 * 优先级低于内联样式，而工具栏每个控件都被 `apply()` 写了内联 `display`。那次失效
 * 的表现是工具栏被裁掉一截、"更多"按钮却不出现（收起三十几项宽度纹丝不动）。真实的
 * 优先级行为只有浏览器能验，但"有没有去改 display"这件事在这里就能拦住。
 */
function fakeElement(display = ""): HTMLElement {
  return { hidden: false, style: { display }, dataset: {} } as unknown as HTMLElement;
}

describe("setHidden", () => {
  it("隐藏时不只设 hidden 属性，还把内联 display 改成 none", () => {
    const element = fakeElement("inline-flex");

    setHidden(element, true);

    expect(element.hidden).toBe(true);
    // 这一条就是整个 helper 存在的理由——只有上面那句的话，浏览器里元素照常占宽度。
    expect(element.style.display).toBe("none");
  });

  it("恢复时把 display 还原成原值，而不是清成空串", () => {
    const element = fakeElement("inline-flex");

    setHidden(element, true);
    setHidden(element, false);

    expect(element.hidden).toBe(false);
    // 清成 "" 会连 apply() 设下的 inline-flex 一起抹掉，把 flex 项退回成 block。
    expect(element.style.display).toBe("inline-flex");
  });

  it("重复隐藏不会把存档覆盖成 none", () => {
    // applyOverflow 会在一次收敛里反复调用，第二次进来时 display 已经是 "none"。
    // 少了那道 `=== undefined` 守卫，存档就被覆盖，之后永远恢复不回来。
    const element = fakeElement("inline-block");

    setHidden(element, true);
    setHidden(element, true);
    setHidden(element, false);

    expect(element.style.display).toBe("inline-block");
  });

  it("没被隐藏过的元素，显示一次不会动它的 display", () => {
    // applyOverflow 每轮开头会把所有项都设成可见，其中大多数本来就没被收起过。
    const element = fakeElement("inline-flex");

    setHidden(element, false);

    expect(element.hidden).toBe(false);
    expect(element.style.display).toBe("inline-flex");
  });

  it("原本就没有内联 display 的元素，还原后仍然没有", () => {
    const element = fakeElement();

    setHidden(element, true);
    expect(element.style.display).toBe("none");

    setHidden(element, false);
    expect(element.style.display).toBe("");
  });
});
