/**
 * 工具栏的默认外观。
 *
 * 用**内联样式**而不是注入一张 stylesheet——与 `DefaultMenuRenderer` 一致（它也是
 * 逐个 `element.style.xxx`）。理由：一个 opt-in 模块往文档里塞全局 CSS 是副作用，
 * 而内联样式只影响它自己创建的节点。元素上仍然带 `cyx-toolbar-*` 类名，宿主想整体
 * 换皮肤就用自己的 CSS 覆盖（CSS 的优先级规则下需要 `!important`，这是内联样式的
 * 代价，写在这里免得以后有人困惑）。
 *
 * 配色取自 `render/theme.ts`，保持和画布同一套灰阶与那个选中蓝。
 */

export const COLORS = {
  text: "#111827",
  border: "#d1d5db",
  hover: "#f3f4f6",
  activeBackground: "#e0edff",
  activeText: "#1d4ed8",
  panelShadow: "0 2px 8px rgba(0, 0, 0, 0.15)",
  surface: "#ffffff"
} as const;

type Style = Partial<CSSStyleDeclaration>;

export function apply(element: HTMLElement, style: Style): void {
  Object.assign(element.style, style);
}

export const toolbarRoot: Style = {
  position: "relative",
  display: "flex",
  alignItems: "center",
  flexWrap: "nowrap",
  gap: "2px",
  padding: "4px 6px",
  background: COLORS.surface,
  borderBottom: `1px solid ${COLORS.border}`,
  fontSize: "13px",
  color: COLORS.text,
  overflow: "hidden"
};

export const slot: Style = { display: "inline-flex", flex: "0 0 auto" };

export const button: Style = {
  display: "inline-flex",
  alignItems: "center",
  justifyContent: "center",
  minWidth: "28px",
  height: "28px",
  padding: "0 6px",
  border: "1px solid transparent",
  borderRadius: "4px",
  background: "transparent",
  color: COLORS.text,
  cursor: "pointer",
  lineHeight: "1"
};

/** 分段按钮：图标区与箭头区挨在一起，视觉上是一颗按钮的两半。 */
export const split: Style = {
  display: "inline-flex",
  alignItems: "center",
  borderRadius: "4px",
  overflow: "hidden"
};

export const splitMain: Style = {
  position: "relative",
  minWidth: "28px",
  paddingRight: "2px",
  borderRadius: "4px 0 0 4px",
  flexDirection: "column",
  gap: "1px"
};

/** 图标下方那条色带——"当前会套用哪个颜色"要一眼看见，Excel 的 A 和油漆桶都有。 */
export const splitSwatch: Style = {
  display: "block",
  width: "14px",
  height: "3px",
  borderRadius: "1px"
};

export const splitCaret: Style = {
  minWidth: "14px",
  padding: "0",
  borderRadius: "0 4px 4px 0"
};

export const select: Style = {
  height: "28px",
  padding: "0 4px",
  border: `1px solid ${COLORS.border}`,
  borderRadius: "4px",
  background: COLORS.surface,
  color: COLORS.text,
  fontSize: "13px",
  cursor: "pointer"
};

export const separator: Style = {
  display: "inline-block",
  flex: "0 0 auto",
  width: "1px",
  height: "18px",
  margin: "0 4px",
  background: COLORS.border
};

export const panel: Style = {
  position: "absolute",
  zIndex: "20",
  minWidth: "160px",
  padding: "6px",
  background: COLORS.surface,
  border: `1px solid ${COLORS.border}`,
  borderRadius: "4px",
  boxShadow: COLORS.panelShadow,
  display: "flex",
  flexDirection: "column",
  gap: "4px"
};

export const panelAction: Style = {
  display: "block",
  width: "100%",
  padding: "4px 8px",
  border: "none",
  borderRadius: "3px",
  background: "transparent",
  color: COLORS.text,
  fontSize: "13px",
  textAlign: "left",
  cursor: "pointer"
};

export const panelInput: Style = {
  width: "100%",
  boxSizing: "border-box",
  height: "26px",
  padding: "0 6px",
  border: `1px solid ${COLORS.border}`,
  borderRadius: "3px",
  fontSize: "13px"
};

/** 编辑栏：名称框 + fx + 内容框，一条横带，贴在工具栏下面。 */
export const formulaBar: Style = {
  display: "flex",
  alignItems: "center",
  gap: "0",
  height: "28px",
  background: COLORS.surface,
  borderBottom: `1px solid ${COLORS.border}`,
  fontSize: "13px",
  color: COLORS.text
};

export const formulaName: Style = {
  width: "96px",
  height: "100%",
  padding: "0 8px",
  border: "none",
  borderRight: `1px solid ${COLORS.border}`,
  background: "transparent",
  color: COLORS.text,
  fontSize: "13px",
  textAlign: "center",
  outline: "none"
};

export const formulaFx: Style = {
  padding: "0 10px",
  borderRight: `1px solid ${COLORS.border}`,
  color: "#6b7280",
  fontStyle: "italic",
  lineHeight: "28px"
};

export const formulaInput: Style = {
  flex: "1 1 auto",
  height: "100%",
  padding: "0 8px",
  border: "none",
  background: "transparent",
  color: COLORS.text,
  fontSize: "13px",
  outline: "none"
};

/** 溢出菜单里"标签 + 下拉"的一行。 */
export const panelRow: Style = {
  display: "flex",
  alignItems: "center",
  justifyContent: "space-between",
  gap: "8px",
  padding: "2px 8px",
  fontSize: "13px"
};

export const panelStatus: Style = { padding: "2px 8px", fontSize: "12px", color: "#6b7280" };

export const swatchGrid: Style = {
  display: "grid",
  gridTemplateColumns: "repeat(5, 20px)",
  gap: "4px",
  padding: "2px"
};

export const swatch: Style = {
  width: "20px",
  height: "20px",
  border: `1px solid ${COLORS.border}`,
  borderRadius: "3px",
  cursor: "pointer",
  padding: "0"
};

/**
 * 自绘 tooltip。
 *
 * 不用浏览器原生 `title`：它要悬停一两秒才出、样式不可控、而且在触摸设备上根本不
 * 出现。Excel 与 Univer 的工具栏都是自绘的小气泡，因为工具栏几乎全是图标按钮——
 * "这个图标是干什么的"是这里最高频的问题，答案慢一秒都嫌久。
 *
 * 气泡挂在 `document.body` 上而不是按钮里：工具栏本身是 `overflow: hidden`
 * （溢出计算要靠它），挂在里面会被裁掉。
 */
export function attachTooltip(element: HTMLElement, text: string): () => void {
  let bubble: HTMLElement | null = null;
  let timer: ReturnType<typeof setTimeout> | null = null;

  const hide = (): void => {
    if (timer !== null) {
      clearTimeout(timer);
      timer = null;
    }
    bubble?.remove();
    bubble = null;
  };

  const show = (): void => {
    hide();
    // 300ms 而不是立刻：鼠标划过一排按钮时不该沿途炸出一串气泡。
    timer = setTimeout(() => {
      timer = null;
      const rect = element.getBoundingClientRect();
      bubble = document.createElement("div");
      bubble.className = "cyx-toolbar-tooltip";
      bubble.textContent = text;
      Object.assign(bubble.style, {
        position: "fixed",
        zIndex: "9999",
        padding: "4px 8px",
        borderRadius: "4px",
        background: "#111827",
        color: "#ffffff",
        fontSize: "12px",
        lineHeight: "1.4",
        whiteSpace: "nowrap",
        pointerEvents: "none",
        transform: "translateX(-50%)",
        left: `${rect.left + rect.width / 2}px`,
        top: `${rect.bottom + 6}px`
      });
      document.body.append(bubble);
      // 贴着视口右边的按钮（溢出菜单那颗）气泡会被切掉，量一次再往回收。
      const bubbleRect = bubble.getBoundingClientRect();
      if (bubbleRect.right > window.innerWidth - 4) {
        bubble.style.left = `${window.innerWidth - 4 - bubbleRect.width / 2}px`;
      }
      if (bubbleRect.left < 4) {
        bubble.style.left = `${4 + bubbleRect.width / 2}px`;
      }
    }, 300);
  };

  element.addEventListener("mouseenter", show);
  element.addEventListener("mouseleave", hide);
  // 点下去就收起来：点击之后气泡还赖着，会挡住刚刚发生变化的那片单元格。
  element.addEventListener("mousedown", hide);
  element.addEventListener("blur", hide);
  return () => {
    hide();
    element.removeEventListener("mouseenter", show);
    element.removeEventListener("mouseleave", hide);
    element.removeEventListener("mousedown", hide);
    element.removeEventListener("blur", hide);
  };
}

/**
 * 压掉焦点边框。
 *
 * 项目方明确不要工具栏上的选中边框。但**只是"不画"是不够的**：控件会落进宿主页面
 * 的 `:focus-visible` 规则（playground 那条是 `2px solid #b93620` 砖红），所以必须
 * 显式把 outline 设成 none 才盖得住——内联样式的优先级高于宿主的普通 CSS 规则。
 *
 * 键盘焦点改用一层浅背景表示：不是边框，但键盘用户仍然看得见自己停在哪颗按钮上。
 * 完全无提示的话，Tab 过一遍工具栏就等于在黑暗里摸索。
 */
export function attachFocusStyle(element: HTMLElement): () => void {
  element.style.outline = "none";
  const onFocus = (): void => {
    let visible: boolean;
    try {
      visible = element.matches(":focus-visible");
    } catch {
      visible = true;
    }
    if (visible) {
      element.style.background = COLORS.hover;
    }
  };
  const onBlur = (): void => {
    element.style.background = "";
  };
  element.addEventListener("focus", onFocus);
  element.addEventListener("blur", onBlur);
  return () => {
    element.removeEventListener("focus", onFocus);
    element.removeEventListener("blur", onBlur);
  };
}

/** 悬停与激活态：没有 stylesheet 就没有 `:hover`，只能自己挂事件。 */
export function attachButtonStates(button: HTMLButtonElement, isActivated: () => boolean): () => void {
  const paint = (hovering: boolean): void => {
    const activated = isActivated();
    button.style.background = activated
      ? COLORS.activeBackground
      : hovering && !button.disabled
        ? COLORS.hover
        : "transparent";
    button.style.color = activated ? COLORS.activeText : COLORS.text;
    button.style.opacity = button.disabled ? "0.4" : "1";
    button.style.cursor = button.disabled ? "default" : "pointer";
  };
  const onEnter = (): void => paint(true);
  const onLeave = (): void => paint(false);
  button.addEventListener("mouseenter", onEnter);
  button.addEventListener("mouseleave", onLeave);
  paint(false);
  return () => {
    button.removeEventListener("mouseenter", onEnter);
    button.removeEventListener("mouseleave", onLeave);
  };
}
