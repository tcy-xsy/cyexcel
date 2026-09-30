import type { CellRect } from "../layout";
import type { ValidationPresenter, ValidationPresentRequest } from "./editValidation";

const COLORS = {
  reject: { border: "#d93025", accent: "#d93025", label: "输入无效" },
  confirm: { border: "#f29900", accent: "#b06000", label: "请确认" },
  notify: { border: "#1a73e8", accent: "#1a73e8", label: "提示" }
} as const;

/** 信息级提示自动消失的时间；停止和警告不自动消失，用户必须处理。 */
const NOTIFY_MS = 4000;

/**
 * 挂在编辑层里的 DOM 气泡，贴在被编辑单元格正下方。
 *
 * 挂在编辑层（`EditorManager` 的 `container`）而不是 `document.body`：编辑器的
 * `focusout` 判断用的是 `container.contains(relatedTarget)`，点气泡上的按钮时焦点落在
 * 容器内部，不会被当成"点到别处去了"而触发提交。
 */
export function createDomValidationPresenter(container: HTMLElement): ValidationPresenter {
  let root: HTMLDivElement | undefined;
  let timer: ReturnType<typeof setTimeout> | undefined;

  const place = (rect: CellRect): void => {
    if (root === undefined) {
      return;
    }
    root.style.left = `${rect.x}px`;
    root.style.top = `${rect.y + rect.height + 4}px`;
  };

  const hide = (): void => {
    if (timer !== undefined) {
      clearTimeout(timer);
      timer = undefined;
    }
    root?.remove();
    root = undefined;
  };

  const button = (label: string, primary: boolean, onClick: () => void): HTMLButtonElement => {
    const element = document.createElement("button");
    element.type = "button";
    element.textContent = label;
    element.style.cssText = [
      "font:inherit",
      "padding:2px 10px",
      "border-radius:4px",
      "cursor:pointer",
      `border:1px solid ${primary ? "#1a73e8" : "#c4c7cc"}`,
      `background:${primary ? "#1a73e8" : "#fff"}`,
      `color:${primary ? "#fff" : "#202124"}`
    ].join(";");
    element.addEventListener("click", onClick);
    return element;
  };

  return {
    show(request: ValidationPresentRequest): void {
      hide();
      const palette = COLORS[request.kind];
      const bubble = document.createElement("div");
      bubble.setAttribute("role", request.kind === "reject" ? "alert" : "status");
      bubble.dataset.cyValidation = request.kind;
      bubble.style.cssText = [
        "position:absolute",
        "z-index:20",
        "box-sizing:border-box",
        "max-width:280px",
        "padding:8px 10px",
        "background:#fff",
        `border:1px solid ${palette.border}`,
        "border-radius:6px",
        "box-shadow:0 2px 8px rgba(0,0,0,.18)",
        "font:12px/1.5 system-ui,sans-serif",
        "color:#202124",
        // 编辑层本身不接指针事件；气泡自己要能点按钮
        "pointer-events:auto"
      ].join(";");
      // 点气泡不该把焦点从编辑器上抢走（按钮除外，它们在容器内，不触发提交）
      bubble.addEventListener("mousedown", (event) => {
        if (!(event.target instanceof HTMLButtonElement)) {
          event.preventDefault();
        }
      });

      const heading = document.createElement("div");
      heading.style.cssText = `font-weight:600;color:${palette.accent}`;
      heading.textContent = request.title ?? palette.label;
      const body = document.createElement("div");
      body.textContent = request.message;
      bubble.append(heading, body);

      if (request.kind === "confirm") {
        const actions = document.createElement("div");
        actions.style.cssText = "display:flex;gap:6px;justify-content:flex-end;margin-top:8px";
        actions.append(
          button("取消", false, () => request.onCancel?.()),
          button("继续", true, () => request.onContinue?.())
        );
        bubble.append(actions);
      }

      container.append(bubble);
      root = bubble;
      place(request.rect);
      if (request.kind === "notify") {
        timer = setTimeout(hide, NOTIFY_MS);
      }
    },
    reposition: place,
    hide,
    dispose: hide
  };
}
