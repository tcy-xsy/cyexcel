import type { ValidationPresenter, ValidationPresentRequest } from "./editValidation";

const DEFAULT_TITLE = "提示";
const OK_LABEL = "确定";

/**
 * 拒绝输入时的模态对话框，形式同 Univer 的 `reject-input-dialog`：宽 368、标题、一段
 * 错误信息、底部一个主按钮「确定」，带遮罩。关掉（点确定或按 Esc）后由调用方决定
 * 回到哪里——这里只负责说清楚为什么被拒绝。
 *
 * 挂在 `document.body` 上而不是编辑层：模态框要盖住整个页面（含工具栏），而且不该被
 * 表格的滚动、裁剪影响。代价是编辑器会在对话框拿到焦点时失焦，所以 `EditorManager` 在
 * 对话框开着的整个期间挡掉提交，关闭时放弃无效修改并退出编辑——见它的注释。
 *
 * 这一层在 node 里测不了（没有 DOM），由 playground 手验；"什么时候弹"是
 * `EditorManager` 的事，那部分用假呈现器测。
 */
export function createDomValidationPresenter(): ValidationPresenter {
  let backdrop: HTMLDivElement | undefined;
  let releaseFocus: (() => void) | undefined;

  const hide = (): void => {
    releaseFocus?.();
    releaseFocus = undefined;
    backdrop?.remove();
    backdrop = undefined;
  };

  return {
    show(request: ValidationPresentRequest): void {
      hide();
      const overlay = document.createElement("div");
      overlay.dataset.cyValidation = "reject-dialog";
      overlay.style.cssText = [
        "position:fixed",
        "inset:0",
        "z-index:2000",
        "display:flex",
        "align-items:center",
        "justify-content:center",
        "background:rgba(32,33,36,.32)"
      ].join(";");

      const dialog = document.createElement("div");
      dialog.setAttribute("role", "alertdialog");
      dialog.setAttribute("aria-modal", "true");
      dialog.style.cssText = [
        "box-sizing:border-box",
        "width:368px",
        "max-width:calc(100vw - 32px)",
        "padding:16px 20px",
        "background:#fff",
        "border-radius:8px",
        "box-shadow:0 8px 24px rgba(0,0,0,.24)",
        "font:14px/1.6 system-ui,sans-serif",
        "color:#202124"
      ].join(";");

      const heading = document.createElement("div");
      heading.id = "cy-validation-dialog-title";
      heading.style.cssText = "font-size:16px;font-weight:600;margin-bottom:8px";
      heading.textContent = request.title ?? DEFAULT_TITLE;
      dialog.setAttribute("aria-labelledby", heading.id);

      const body = document.createElement("div");
      body.textContent = request.message;
      const hint = document.createElement("div");
      hint.textContent = "本次修改未保存，关闭后保留原值。";
      hint.style.cssText = "margin-top:8px;color:#5f6368;font-size:12px";

      const footer = document.createElement("div");
      footer.style.cssText = "display:flex;justify-content:flex-end;margin-top:16px";
      const ok = document.createElement("button");
      ok.type = "button";
      ok.textContent = OK_LABEL;
      ok.style.cssText = [
        "font:inherit",
        "padding:4px 16px",
        "border:1px solid #1a73e8",
        "border-radius:4px",
        "background:#1a73e8",
        "color:#fff",
        "cursor:pointer"
      ].join(";");

      let closed = false;
      const close = (): void => {
        if (closed) {
          return;
        }
        closed = true;
        hide();
        request.onClose();
      };
      ok.addEventListener("click", close);
      const onKeyDown = (event: KeyboardEvent): void => {
        if (event.key === "Escape") {
          event.preventDefault();
          event.stopPropagation();
          close();
        } else if (event.key === "Tab") {
          // 对话框里只有一个可聚焦的东西：Tab 也别让焦点漏到后面的表格里
          event.preventDefault();
        }
      };
      const onFocusIn = (event: FocusEvent): void => {
        // A click that triggered focusout can still focus the canvas after
        // show() returns. Keep keyboard actions inside the modal as well.
        if (event.target instanceof Node && !dialog.contains(event.target)) {
          ok.focus();
        }
      };
      document.addEventListener("keydown", onKeyDown, true);
      document.addEventListener("focusin", onFocusIn);
      releaseFocus = () => {
        document.removeEventListener("keydown", onKeyDown, true);
        document.removeEventListener("focusin", onFocusIn);
      };

      footer.append(ok);
      dialog.append(heading, body, hint, footer);
      overlay.append(dialog);
      document.body.append(overlay);
      backdrop = overlay;
      ok.focus();
    },
    hide,
    dispose: hide
  };
}
