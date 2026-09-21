/**
 * 工具栏这一层共用的 DOM 小工具。
 *
 * 只有 `el()` 一个函数，但它原先在 `createToolbar.ts` 与 `panels.ts` 各有一份拷贝
 * （后者多一个 `text` 参数），`createDialog.ts` 又要用——与其加第三份，不如收在这里。
 */
export function el<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  className?: string,
  text?: string
): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  if (className !== undefined) {
    node.className = className;
  }
  if (text !== undefined) {
    node.textContent = text;
  }
  return node;
}
