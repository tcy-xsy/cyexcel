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
