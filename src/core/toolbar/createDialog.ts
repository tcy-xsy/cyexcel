import { el } from "./dom";
import * as styles from "./styles";

/**
 * 非模态、可拖动的对话框——Excel 的「查找和替换」那一类。
 *
 * **为什么不是面板。** 工具栏的 `panel` 是个贴着按钮的小浮层，点外面就关。查找替换
 * 塞不进去：它要两个输入框、几个选项、四个按钮，而且**开着的时候必须还能操作表格**
 * ——"查找下一个"的意义就在于看见它跳到哪一格，点一下单元格就把框关掉的话这个功能
 * 没法用。所以这里既不做遮罩也不做焦点陷阱，更不监听外部点击。
 *
 * 代价是关闭只能靠标题栏的 ×、内容自己调 `close()`、或者焦点在框内时按 Esc。Esc 特意
 * 只在框内生效：非模态框不该把整个文档的 Esc 抢走——那是编辑单元格时取消编辑用的键。
 *
 * 同 `panel`：`fixed` + 挂 `document.body`。工具栏 root 是 `overflow: hidden`，挂进去
 * 会被整个裁掉（见 styles.ts 的 `panel`）。
 *
 * 这一层在 node 里测不了（没有 DOM，项目不引 jsdom——见 step25-toolbar-plan.md §5），
 * 拖动与收边由 playground 手验。
 */
export interface DialogOptions {
  readonly title: string;
  /**
   * 内容自绘，返回 disposer。`close` 交给内容用——比如「关闭」按钮，或者替换完成后
   * 自行收起。
   */
  readonly fill: (body: HTMLElement, close: () => void) => () => void;
  /** 关掉之后回调一次，用于让调用方把它从"开着的对话框"里摘掉。 */
  readonly onClosed?: () => void;
}

export interface DialogHandle {
  readonly root: HTMLElement;
  close(): void;
  /** 把焦点送进框里——打开时聚到第一个输入框，比让用户再点一下强。 */
  focus(): void;
}

/** 拖动时至少留这么多像素在视口内，免得标题栏被拖出去就再也抓不回来。 */
const KEPT_VISIBLE = 48;
const EDGE = 4;

function clamp(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max);
}

export function createDialog(options: DialogOptions): DialogHandle {
  const root = el("div", "cyx-toolbar-dialog");
  styles.apply(root, styles.dialog);
  root.setAttribute("role", "dialog");
  // aria-modal 明确是 false：它真的不是模态的，读屏不该把外面的内容藏起来。
  root.setAttribute("aria-modal", "false");
  root.setAttribute("aria-label", options.title);

  const header = el("div", "cyx-toolbar-dialog-header");
  styles.apply(header, styles.dialogHeader);
  const titleNode = el("span", "cyx-toolbar-dialog-title", options.title);

  const closeButton = el("button", "cyx-toolbar-dialog-close", "✕");
  closeButton.type = "button";
  closeButton.setAttribute("aria-label", "关闭");
  styles.apply(closeButton, styles.dialogClose);
  const disposeCloseStyle = styles.attachFocusStyle(closeButton);

  header.append(titleNode, closeButton);

  const body = el("div", "cyx-toolbar-dialog-body");
  styles.apply(body, styles.dialogBody);

  root.append(header, body);

  let closed = false;
  const close = (): void => {
    if (closed) {
      return;
    }
    closed = true;
    disposeContent();
    disposeCloseStyle();
    root.remove();
    options.onClosed?.();
  };

  closeButton.addEventListener("click", close);

  // Esc 只在焦点落在框内时生效——见文件头的说明。
  root.addEventListener("keydown", (event: KeyboardEvent) => {
    if (event.key === "Escape") {
      event.stopPropagation();
      close();
    }
  });

  // ---- 拖动 ----
  // 用 pointer capture 而不是往 document 上挂 mousemove：指针移出窗口、或者滑过
  // canvas 时事件照样回到标题栏，不会拖到一半"脱手"。
  let dragPointer: number | null = null;
  let grabX = 0;
  let grabY = 0;

  header.addEventListener("pointerdown", (event: PointerEvent) => {
    // 关闭按钮不起拖，否则按下 × 会先变成一次拖动。
    if (closeButton.contains(event.target as Node)) {
      return;
    }
    const rect = root.getBoundingClientRect();
    grabX = event.clientX - rect.left;
    grabY = event.clientY - rect.top;
    dragPointer = event.pointerId;
    header.setPointerCapture(event.pointerId);
    event.preventDefault();
  });

  header.addEventListener("pointermove", (event: PointerEvent) => {
    if (dragPointer !== event.pointerId) {
      return;
    }
    const rect = root.getBoundingClientRect();
    // 收边按"还剩多少可见"算，不是按整个框——框比视口宽时也得能拖。
    const left = clamp(
      event.clientX - grabX,
      KEPT_VISIBLE - rect.width,
      window.innerWidth - KEPT_VISIBLE
    );
    const top = clamp(event.clientY - grabY, 0, window.innerHeight - KEPT_VISIBLE);
    root.style.left = `${left}px`;
    root.style.top = `${top}px`;
  });

  const endDrag = (event: PointerEvent): void => {
    if (dragPointer !== event.pointerId) {
      return;
    }
    header.releasePointerCapture(event.pointerId);
    dragPointer = null;
  };
  header.addEventListener("pointerup", endDrag);
  header.addEventListener("pointercancel", endDrag);

  const disposeContent = options.fill(body, close);
  document.body.append(root);

  // 先落位再收边：要量到真实尺寸才知道会不会顶出视口。同 openLayer 的做法。
  // 居中而不是锚定在触发按钮下方——按钮可能在工具栏最左或最右，锚定过去经常
  // 大半个框飘在视口外；居中不需要知道按钮在哪，也不会有这问题。
  const rect = root.getBoundingClientRect();
  root.style.left = `${clamp((window.innerWidth - rect.width) / 2, EDGE, Math.max(EDGE, window.innerWidth - EDGE - rect.width))}px`;
  root.style.top = `${clamp((window.innerHeight - rect.height) / 2, EDGE, Math.max(EDGE, window.innerHeight - EDGE - rect.height))}px`;

  return {
    root,
    close,
    focus: () => {
      const first = root.querySelector("input, select, button");
      if (first instanceof HTMLElement) {
        first.focus();
      }
    }
  };
}
